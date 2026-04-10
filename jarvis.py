"""
╔══════════════════════════════════════════╗
║   JARVIS - LINKS Mark II                 ║
║   Lightweight Background Voice Assistant ║
╚══════════════════════════════════════════╝
Requires Python 3.10  (py -3.10 jarvis.py)
Install (run with py -3.10 -m pip install ...):
         speechrecognition pyttsx3 requests pyaudio pystray pillow
         edge-tts pygame psutil
         pyaudiowpatch            ← system audio loopback for music recognition
         demucs torch torchaudio  ← anime OST isolation (vocal removal)
         "numpy<2.0"              ← keep numpy 1.x; torch 2.2.2 crashes on numpy 2.x
Requires Ollama running locally: https://ollama.com  (ollama pull llama3)
"""

import os
import sys
import re
import time
import queue
import socket
import hashlib
import asyncio
import tempfile
import datetime
import subprocess
import threading
import webbrowser
import glob
import ctypes
import logging
import random

# ─────────────────────────────────────────
#  LOGGING — capture all output to file
# ─────────────────────────────────────────
_LOG_FILE  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "jarvis.log")
_NO_WINDOW = getattr(subprocess, 'CREATE_NO_WINDOW', 0x08000000)

class HourlyClearHandler(logging.FileHandler):
    """File handler that wipes the log once per hour to keep it small."""
    def __init__(self, filename, mode='a', encoding=None, delay=False):
        super().__init__(filename, mode, encoding, delay)
        self.next_clear = time.time() + 3600   # actually hourly

    def emit(self, record):
        if time.time() >= self.next_clear:
            self.close()
            with open(self.baseFilename, 'w', encoding=self.encoding): pass
            self.stream = self._open()
            self.next_clear = time.time() + 3600
        super().emit(record)

_log_handlers = [HourlyClearHandler(_LOG_FILE, encoding="utf-8")]
try:
    # Only attach a console handler when a real console is available.
    # sys.stdout.fileno() raises io.UnsupportedOperation when running
    # as a background/GUI process (no attached terminal), which would
    # silently break the entire logging config including the file handler.
    _console_stream = open(
        sys.stdout.fileno(), mode='w', encoding='utf-8', errors='replace', closefd=False
    )
    _log_handlers.append(logging.StreamHandler(_console_stream))
except Exception:
    pass  # No console — file-only logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=_log_handlers,
)
log = logging.getLogger("JARVIS")

# Redirect unhandled exceptions to the log file
def _exc_hook(exc_type, exc_value, exc_tb):
    log.critical("Unhandled exception", exc_info=(exc_type, exc_value, exc_tb))
sys.excepthook = _exc_hook

try:
    import requests
    import pyttsx3
    import speech_recognition as sr
    import pystray
    from PIL import Image, ImageDraw
except ImportError as e:
    log.critical(f"Missing dependency: {e}. Run: pip install speechrecognition pyttsx3 requests pyaudio pystray pillow")
    sys.exit(1)

# ─────────────────────────────────────────
#  CONFIG  — edit these freely
# ─────────────────────────────────────────
OLLAMA_URL    = "http://localhost:11434/api/generate"
OLLAMA_MODEL  = "llama3"
WAKE_WORD     = "jarvis"
TTS_RATE      = 175
TTS_VOLUME    = 1.0
TTS_VOICE     = "en-GB-RyanNeural"           # Microsoft neural voice — British male, very Jarvis-like
TTS_EDGE_RATE = "+30%"                       # edge-tts speed tweak — faster delivery, more fluid

# ─────────────────────────────────────────
#  QUICK-ACCESS FOLDERS
# ─────────────────────────────────────────
FOLDER_MAP = {
    "downloads":  os.path.expanduser("~/Downloads"),
    "download":   os.path.expanduser("~/Downloads"),
    "desktop":    os.path.expanduser("~/Desktop"),
    "documents":  os.path.expanduser("~/Documents"),
    "pictures":   os.path.expanduser("~/Pictures"),
    "music":      os.path.expanduser("~/Music"),
    "videos":     os.path.expanduser("~/Videos"),
    "appdata":    os.path.expandvars("%APPDATA%"),
    "jarvis":     os.path.dirname(os.path.abspath(__file__)),
}

NOTES_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "jarvis_notes.txt")

# Stopwatch state
_stopwatch_start = None   # float timestamp when running, None otherwise
_last_spoken     = None   # last text spoken — for "say that again"

# ─────────────────────────────────────────
#  UNIT CONVERSION TABLE  (base SI unit per category)
# ─────────────────────────────────────────
_UNIT_GROUPS = {
    "length": {
        "m": 1, "meter": 1, "meters": 1,
        "km": 1000, "kilometer": 1000, "kilometers": 1000,
        "mile": 1609.344, "miles": 1609.344,
        "foot": 0.3048, "feet": 0.3048, "ft": 0.3048,
        "inch": 0.0254, "inches": 0.0254,
        "cm": 0.01, "centimeter": 0.01, "centimeters": 0.01,
        "yard": 0.9144, "yards": 0.9144,
    },
    "mass": {
        "kg": 1, "kilogram": 1, "kilograms": 1,
        "g": 0.001, "gram": 0.001, "grams": 0.001,
        "lb": 0.453592, "pound": 0.453592, "pounds": 0.453592, "lbs": 0.453592,
        "oz": 0.0283495, "ounce": 0.0283495, "ounces": 0.0283495,
    },
    "volume": {
        "l": 1, "liter": 1, "liters": 1,
        "ml": 0.001, "milliliter": 0.001, "milliliters": 0.001,
        "gallon": 3.78541, "gallons": 3.78541,
        "pint": 0.473176, "pints": 0.473176,
        "cup": 0.236588, "cups": 0.236588,
    },
    "speed": {
        "mps": 1,
        "mph": 0.44704,
        "kph": 0.277778, "kmh": 0.277778,
        "knot": 0.514444, "knots": 0.514444,
    },
}

# ─────────────────────────────────────────
#  CITY → TIMEZONE MAP
# ─────────────────────────────────────────
_CITY_TIMEZONES = {
    "london": "Europe/London", "uk": "Europe/London", "england": "Europe/London",
    "paris": "Europe/Paris", "france": "Europe/Paris",
    "berlin": "Europe/Berlin", "germany": "Europe/Berlin",
    "amsterdam": "Europe/Amsterdam", "netherlands": "Europe/Amsterdam",
    "rome": "Europe/Rome", "italy": "Europe/Rome",
    "madrid": "Europe/Madrid", "spain": "Europe/Madrid",
    "moscow": "Europe/Moscow", "russia": "Europe/Moscow",
    "istanbul": "Europe/Istanbul", "turkey": "Europe/Istanbul",
    "dubai": "Asia/Dubai", "uae": "Asia/Dubai",
    "riyadh": "Asia/Riyadh", "saudi arabia": "Asia/Riyadh",
    "mumbai": "Asia/Kolkata", "delhi": "Asia/Kolkata", "india": "Asia/Kolkata",
    "karachi": "Asia/Karachi", "pakistan": "Asia/Karachi",
    "bangkok": "Asia/Bangkok", "thailand": "Asia/Bangkok",
    "singapore": "Asia/Singapore",
    "jakarta": "Asia/Jakarta", "indonesia": "Asia/Jakarta",
    "hong kong": "Asia/Hong_Kong",
    "beijing": "Asia/Shanghai", "shanghai": "Asia/Shanghai", "china": "Asia/Shanghai",
    "seoul": "Asia/Seoul", "korea": "Asia/Seoul",
    "tokyo": "Asia/Tokyo", "japan": "Asia/Tokyo",
    "sydney": "Australia/Sydney", "australia": "Australia/Sydney",
    "new york": "America/New_York", "nyc": "America/New_York",
    "chicago": "America/Chicago",
    "los angeles": "America/Los_Angeles", "la": "America/Los_Angeles",
    "toronto": "America/Toronto", "ottawa": "America/Toronto",
    "vancouver": "America/Vancouver",
    "mexico city": "America/Mexico_City", "mexico": "America/Mexico_City",
    "sao paulo": "America/Sao_Paulo", "brazil": "America/Sao_Paulo",
    "buenos aires": "America/Argentina/Buenos_Aires", "argentina": "America/Argentina/Buenos_Aires",
    "cairo": "Africa/Cairo", "egypt": "Africa/Cairo",
    "nairobi": "Africa/Nairobi", "kenya": "Africa/Nairobi",
    "johannesburg": "Africa/Johannesburg", "south africa": "Africa/Johannesburg",
}

SITE_MAP = {
    "youtube":               "https://www.youtube.com",
    "gmail":                 "https://mail.google.com",
    "google":                "https://www.google.com",
    "github":                "https://www.github.com",
    "reddit":                "https://www.reddit.com",
    "twitter":               "https://www.twitter.com",
    "netflix":               "https://www.netflix.com",
    "twitch":                "https://www.twitch.tv",
    "spotify web":           "https://open.spotify.com",
    "chatgpt":               "https://chat.openai.com",
    "claude":                "https://claude.ai",
    "amazon":                "https://www.amazon.com",
    "wikipedia":             "https://www.wikipedia.org",
    "brightspace carleton":  "https://brightspace.carleton.ca/d2l/home",
    "brightspace algonquin": "https://brightspace.algonquincollege.com/d2l/home",
    "nokia":                 "https://nokia.fileopen.com/?id=Nokia%20Multiprotocol%20Label%20Switching%20Student%20Guide%20v3.2.3.pdf",
    "nokia services":        "https://nokia.fileopen.com/?id=Nokia%20Services%20Architecture%20Student%20Guide%20v4.1.3.pdf",
    "anime":                 "https://aniwatchtv.to/home",
    "Micheal":               "https://web.ncf.ca/andersonm/",
    
}

# ─────────────────────────────────────────
#  APP DETECTION
# ─────────────────────────────────────────
APP_NAMES = [
    "spotify", "brave", "firefox", "edge",
    "notepad", "calculator", "explorer", "task manager", "paint",
    "word", "excel", "powerpoint",
    "vs code", "discord", "steam", "vlc", "obs", "league of legends",
    "antigravity", "soundswitch", "osu",
]

APP_EXTRA_CANDIDATES = {
    "spotify":    [r"%APPDATA%\Spotify\Spotify.exe",
                   r"%LOCALAPPDATA%\Microsoft\WindowsApps\Spotify.exe"],
    "brave":      [r"%LOCALAPPDATA%\BraveSoftware\Brave-Browser\Application\brave.exe"],
    "discord":    [r"%LOCALAPPDATA%\Discord\app-*\Discord.exe",
                   r"%APPDATA%\discord\Discord.exe"],
    "vs code":    [r"%LOCALAPPDATA%\Programs\Microsoft VS Code\Code.exe"],
    "steam":      [r"C:\Program Files (x86)\Steam\Steam.exe"],
    "vlc":        [r"C:\Program Files\VideoLAN\VLC\vlc.exe",
                   r"C:\Program Files (x86)\VideoLAN\VLC\vlc.exe"],
    "league of legends":        [r"%USERPROFILE%\Desktop\League of Legends.lnk",
                                 r"%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\Riot Games\League of Legends.lnk"],
    "word":       [r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
                   r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"],
    "excel":      [r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
                   r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"],
    "powerpoint": [r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"],
    "antigravity":  [r"%APPDATA%\Microsoft\Windows\Start Menu\Programs\Antigravity\Antigravity.lnk"],
    "osu":          [r"%LOCALAPPDATA%\osulazer\current\osu!.exe",
                     r"%USERPROFILE%\Desktop\osu!(lazer).lnk"],
    "soundswitch":  [r"%LOCALAPPDATA%\Programs\SoundSwitch\SoundSwitch.exe",
                     r"%LOCALAPPDATA%\SoundSwitch\SoundSwitch.exe",
                     r"C:\Program Files\SoundSwitch\SoundSwitch.exe",
                     r"C:\Program Files (x86)\SoundSwitch\SoundSwitch.exe"],
}

APP_SHELL_ALIAS = {
    "notepad":      "notepad.exe",
    "calculator":   "calc.exe",
    "explorer":     "explorer.exe",
    "task manager": "taskmgr.exe",
    "paint":        "mspaint.exe",
}

SEQUENCES = {
    "work mode":       [APP_EXTRA_CANDIDATES["discord"], APP_EXTRA_CANDIDATES["brave"], APP_EXTRA_CANDIDATES["spotify"]],
    "gaming mode":     [APP_EXTRA_CANDIDATES["discord"], APP_EXTRA_CANDIDATES["spotify"], APP_EXTRA_CANDIDATES["steam"], APP_EXTRA_CANDIDATES["league of legends"]],
    "research mode":   ["https://www.wikipedia.org","https://scholar.google.com", "https://www.google.com"],
    "exam mode":       ["https://www.youtube.com",  "https://brightspace.carleton.ca/d2l/home",
                        "https://brightspace.algonquincollege.com/d2l/home", "https://nokia.fileopen.com/?id=Nokia%20Multiprotocol%20Label%20Switching%20Student%20Guide%20v3.2.3.pdf",
                        "https://nokia.fileopen.com/?id=Nokia%20Services%20Architecture%20Student%20Guide%20v4.1.3.pdf",
                        "https://web.ncf.ca/andersonm/"],
    "movie mode":      ["https://xprime.stream/", "https://aether.mom/"],
    "anime mode":      ["https://aniwatchtv.to/home", APP_EXTRA_CANDIDATES["discord"]],
}
# Always lowercase so voice (which Google returns lowercase) matches
SEQUENCES = {k.lower(): v for k, v in SEQUENCES.items()}

def _registry_lookup(app_keyword):
    try:
        import winreg
        roots = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
            (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        ]
        kw = app_keyword.lower().replace(" ", "")
        for hive, path in roots:
            try:
                reg = winreg.OpenKey(hive, path)
            except OSError:
                continue
            for i in range(winreg.QueryInfoKey(reg)[0]):
                try:
                    sub_name = winreg.EnumKey(reg, i)
                    sub = winreg.OpenKey(reg, sub_name)
                    try:
                        disp, _ = winreg.QueryValueEx(sub, "DisplayName")
                        if kw in disp.lower().replace(" ", ""):
                            try:
                                loc, _ = winreg.QueryValueEx(sub, "InstallLocation")
                                if loc:
                                    for exe in [f"{disp}.exe", f"{sub_name}.exe", "brave.exe",
                                                "firefox.exe", "Code.exe", "Discord.exe", "Spotify.exe"]:
                                        candidate = os.path.join(loc, exe)
                                        if os.path.exists(candidate):
                                            return candidate
                            except OSError:
                                pass
                    except OSError:
                        pass
                except OSError:
                    continue
    except ImportError:
        pass
    return None

def resolve_app_path(name):
    if name in APP_SHELL_ALIAS:
        return APP_SHELL_ALIAS[name]
    for raw in APP_EXTRA_CANDIDATES.get(name, []):
        expanded = os.path.expandvars(raw)
        if '*' in expanded:
            matches = glob.glob(expanded)
            if matches:
                return sorted(matches)[-1]
        elif os.path.exists(expanded):
            return expanded
    reg = _registry_lookup(name)
    if reg:
        return reg
    try:
        result = subprocess.run(["where", name.replace(" ", "") + ".exe"],
                                capture_output=True, text=True)
        if result.returncode == 0:
            found = result.stdout.strip().splitlines()[0]
            if found and os.path.exists(found):
                return found
    except Exception:
        pass
    return None

log.info("Scanning for apps...")
APP_MAP = {}
for _n in APP_NAMES:
    _p = resolve_app_path(_n)
    if _p:
        APP_MAP[_n] = _p
        log.info(f"  ✓ {_n}")
    else:
        log.info(f"  ✗ {_n}")
log.info(f"  📋 Sequences: {list(SEQUENCES.keys())}")

# ─────────────────────────────────────────
#  OLLAMA STARTUP CHECK
# ─────────────────────────────────────────
def _check_ollama_startup():
    try:
        if _ollama_is_running():
            log.info("  ✓ Ollama is running.")
            if _ollama_model_available(OLLAMA_MODEL):
                log.info(f"  ✓ Model '{OLLAMA_MODEL}' is ready.")
            else:
                log.warning(f"  ✗ Model '{OLLAMA_MODEL}' not pulled yet — will auto-pull on first AI query.")
        else:
            log.warning("  ✗ Ollama is not running — will attempt auto-start on first AI query.")
    except Exception:
        pass

# ─────────────────────────────────────────
#  LOWER PROCESS PRIORITY
# ─────────────────────────────────────────
try:
    handle = ctypes.windll.kernel32.GetCurrentProcess()
    ctypes.windll.kernel32.SetPriorityClass(handle, 0x4000)  # BELOW_NORMAL
    log.info("  ✓ Priority: below-normal")
except Exception as e:
    log.warning(f"Could not set priority: {e}")

# ─────────────────────────────────────────
#  TTS  — Microsoft Edge neural voices via edge-tts
# ─────────────────────────────────────────
try:
    import edge_tts
    import pygame
    pygame.mixer.pre_init(frequency=24000, size=-16, channels=2, buffer=512)
    pygame.mixer.init()
    _USE_EDGE_TTS = True
    log.info(f"  ✓ Edge TTS ready — voice: {TTS_VOICE}")
except ImportError as _tts_err:
    _USE_EDGE_TTS = False
    log.warning(f"  ✗ Edge TTS unavailable ({_tts_err}) — falling back to pyttsx3 (robotic)")

_tts_queue = queue.Queue()

# ── TTS audio cache — persist to disk so common phrases are instant on repeat ──
_TTS_CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".tts_cache")
os.makedirs(_TTS_CACHE_DIR, exist_ok=True)
_audio_cache: dict = {}  # text → cached file path

def _tts_cache_key(text):
    return hashlib.md5(f"{TTS_VOICE}|{TTS_EDGE_RATE}|{text}".encode()).hexdigest()

def _get_cached_tts(text):
    key = _tts_cache_key(text)
    if key in _audio_cache:
        p = _audio_cache[key]
        if os.path.exists(p) and os.path.getsize(p) > 0:
            return p
        del _audio_cache[key]
    path = os.path.join(_TTS_CACHE_DIR, key + ".mp3")
    if os.path.exists(path) and os.path.getsize(path) > 0:
        _audio_cache[key] = path
        return path
    return None

async def _cache_tts_async(text):
    key = _tts_cache_key(text)
    path = os.path.join(_TTS_CACHE_DIR, key + ".mp3")
    if os.path.exists(path) and os.path.getsize(path) > 0:
        _audio_cache[key] = path
        return path
    tmp = path + ".tmp"
    try:
        await edge_tts.Communicate(text, voice=TTS_VOICE, rate=TTS_EDGE_RATE).save(tmp)
        os.replace(tmp, path)  # atomic: only visible once fully written
    except Exception:
        try: os.unlink(tmp)
        except OSError: pass
        raise
    _audio_cache[key] = path
    return path

def _tts_worker():
    while True:
        text = _tts_queue.get()
        if text is None:
            break
        try:
            if _USE_EDGE_TTS:
                cached = _get_cached_tts(text)
                if cached:
                    audio_path, is_temp = cached, False
                else:
                    with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
                        audio_path = f.name
                    asyncio.run(edge_tts.Communicate(text, voice=TTS_VOICE, rate=TTS_EDGE_RATE).save(audio_path))
                    is_temp = True
                pygame.mixer.music.load(audio_path)
                pygame.mixer.music.play()
                while pygame.mixer.music.get_busy():
                    time.sleep(0.05)
                pygame.mixer.music.unload()
                if is_temp:
                    try:
                        os.unlink(audio_path)
                    except Exception:
                        pass
            else:
                # pyttsx3 fallback if edge-tts not installed
                code = f"""
import pyttsx3
try:
    engine = pyttsx3.init()
    engine.setProperty('rate', {TTS_RATE})
    engine.setProperty('volume', {TTS_VOLUME})
    voices = engine.getProperty('voices')
    if voices: engine.setProperty('voice', voices[0].id)
    engine.say({repr(text)})
    engine.runAndWait()
except Exception: pass
"""
                flags = _NO_WINDOW
                subprocess.run([sys.executable, "-c", code], shell=False, creationflags=flags)
        except Exception as ex:
            log.error(f"TTS error: {ex}")
        _tts_queue.task_done()

threading.Thread(target=_tts_worker, daemon=True).start()

def speak(text):
    global _last_spoken
    log.info(f"🔊 {text}")
    _last_spoken = text
    _tts_queue.put(text)

# ─────────────────────────────────────────
#  BRAVE CONTROLLER
# ─────────────────────────────────────────
def open_in_brave(url, new_tab=True):
    brave = APP_MAP.get("brave")
    try:
        if brave and os.path.exists(brave):
            args = [brave]
            if new_tab:
                args.append("--new-tab")
            args.append(url)
            subprocess.Popen(args)
        else:
            webbrowser.open(url)
    except Exception as e:
        log.error(f"Brave error: {e}")
        webbrowser.open(url)

def __launch_sequence_item(item, is_new_tab):
    if isinstance(item, list):
        app_name = next((k for k, v in APP_EXTRA_CANDIDATES.items() if v == item), None)
        if app_name and app_name in APP_MAP:
            try: os.startfile(APP_MAP[app_name])
            except Exception as e: log.error(f"Failed to launch {app_name}: {e}")
        else:
            log.warning(f"Could not resolve app for sequence item: {item}")
    elif isinstance(item, str):
        if item.startswith("http") or "www." in item:
            open_in_brave(item, new_tab=is_new_tab)
        elif item in APP_MAP:
            try: os.startfile(APP_MAP[item])
            except Exception as e: log.error(f"Failed to launch app {item}: {e}")
        else:
            open_in_brave(item, new_tab=is_new_tab)

def run_sequence(name):
    items = SEQUENCES.get(name, [])
    for i, item in enumerate(items):
        threading.Timer(i * 0.8, __launch_sequence_item, args=[item, i > 0]).start()

# ─────────────────────────────────────────
#  TIMER STORE
# ─────────────────────────────────────────
_active_timers = {}   # label -> threading.Timer

def _parse_duration(text):
    """Parse a natural-language duration into seconds. Returns None if unparseable."""
    total = 0
    h = re.search(r'(\d+)\s*hour', text)
    m = re.search(r'(\d+)\s*min',  text)
    s = re.search(r'(\d+)\s*sec',  text)
    if h: total += int(h.group(1)) * 3600
    if m: total += int(m.group(1)) * 60
    if s: total += int(s.group(1))
    return total if total > 0 else None

# ─────────────────────────────────────────
#  INTENT PARSER
# ─────────────────────────────────────────
def parse_intent(text):
    t = text.lower().strip()

    # Match sequences case-insensitively, and allow synonyms for "mode"
    # e.g. "exam time", "gaming session", "work setup" all resolve correctly
    _MODE_SYNONYMS = ["mode", "time", "session", "setup", "settings", "screen"]
    def _seq_matches(seq, text):
        if seq in text:
            return True
        # Strip the trailing word from the seq key and check if the core name
        # appears in text alongside any synonym
        parts = seq.rsplit(" ", 1)
        if len(parts) == 2:
            core, _ = parts  # e.g. "exam", "gaming"
            if core in text and any(syn in text for syn in _MODE_SYNONYMS):
                return True
        return False

    matched_sequences = [seq for seq in SEQUENCES if _seq_matches(seq, t)]
    if len(matched_sequences) > 1:
        return ("multi_sequence", matched_sequences)
    elif len(matched_sequences) == 1:
        return ("sequence", matched_sequences[0])

    for app_name in APP_NAMES:
        if app_name in t and any(w in t for w in ["open", "launch", "start", "run"]):
            return ("open_app", app_name, APP_MAP.get(app_name))

    for site_name, url in SITE_MAP.items():
        if site_name.lower() in t and any(w in t for w in ["open","go to","launch","show","load"]):
            return ("open_site", site_name, url)

    for w in t.split():
        if "." in w and len(w) > 4 and any(w.endswith(tld) for tld in
                [".com",".co.uk",".org",".net",".io",".tv",".gg",".ai",".dev",".uk",".ca"]):
            url = w if w.startswith("http") else "https://" + w
            return ("open_site", w, url)

    if any(w in t for w in ["search for", "search", "look up", "google"]):
        query = t
        for w in ["search for", "look up", "google", "search"]:
            query = query.replace(w, "").strip()
        return ("web_search", query)

    if any(w in t for w in ["play music","pause music","play pause","pause","play"]):
        return ("media", "playpause")
    if any(w in t for w in ["next song","next track","skip"]):
        return ("media", "next")
    if any(w in t for w in ["previous song","previous track","go back"]):
        return ("media", "prev")
    if any(w in t for w in ["stop music","stop playing"]):
        return ("media", "stop")

    if any(w in t for w in ["what time","what's the time"]):
        return ("time", None)
    if any(w in t for w in ["what day","what's the date","today's date"]):
        return ("date", None)
    if any(w in t for w in ["shutdown","shut down"]):
        return ("shutdown", None)
    if any(w in t for w in ["restart jarvis","reboot jarvis","reload jarvis","reset jarvis","refresh jarvis"]):
        return ("restart_jarvis", None)
    if any(w in t for w in ["restart","reboot"]):
        return ("restart", None)
    if any(w in t for w in ["sleep","hibernate"]):
        return ("sleep", None)
    if any(w in t for w in ["lock screen","lock the screen","lock computer","lock the computer"]):
        return ("lock", None)

    # Volume: "set volume to 50" / "volume at 70 percent"
    _vol_match = re.search(r'(?:set\s+)?volume\s+(?:to|at)\s+(\d{1,3})\s*(?:percent|%)?', t)
    if _vol_match:
        return ("volume", "set", int(_vol_match.group(1)))
    if "volume up"   in t: return ("volume", "up",   None)
    if "volume down" in t: return ("volume", "down", None)
    if "mute"        in t: return ("volume", "mute", None)

    # Screenshot — optional name and/or folder
    # e.g. "screenshot named report in downloads"
    #       "take a screenshot called diagram and save it in a new folder called diagrams"
    #       "screenshot and put it in documents"
    if any(w in t for w in ["screenshot", "take a screenshot", "capture screen", "screen capture"]):
        _name_match       = re.search(r'(?:named?|called?)\s+([\w\-]+)', t)
        _custom_name      = _name_match.group(1) if _name_match else None
        _new_folder_match = re.search(r'new folder(?:\s+called?|\s+named?)?\s+([\w\-]+)', t)
        _in_folder_match  = re.search(r'(?:in|into|to|save (?:it )?(?:in|to))\s+(?:my\s+)?([\w\-]+)(?:\s+folder)?', t)
        if _new_folder_match:
            _folder_arg = ("new", _new_folder_match.group(1))
        elif _in_folder_match:
            _folder_arg = ("known", _in_folder_match.group(1))
        else:
            _folder_arg = None
        return ("screenshot", _custom_name, _folder_arg)

    # Weather
    if any(w in t for w in ["weather","forecast","temperature outside"]):
        return ("weather", None)

    # Close app: "close spotify" / "kill discord"
    for app_name in APP_NAMES:
        if app_name in t and any(w in t for w in ["close","kill","quit","exit","stop"]):
            return ("close_app", app_name)

    # Open folder: "open downloads" / "show my desktop"
    for folder_name in FOLDER_MAP:
        if folder_name in t and any(w in t for w in ["open","show","go to","navigate to"]):
            return ("open_folder", folder_name)

    # Quick note: "note that the meeting is at 3pm"
    if t.startswith("note ") or t.startswith("take a note") or t.startswith("remember that") or "make a note" in t:
        note_text = t
        for prefix in ["note that", "take a note that", "take a note", "remember that", "make a note that", "make a note", "note"]:
            if note_text.startswith(prefix):
                note_text = note_text[len(prefix):].strip()
                break
        return ("note", note_text)

    # Read notes back
    if any(w in t for w in ["read my notes","what are my notes","show my notes","read notes"]):
        return ("read_notes", None)

    # System status
    if any(w in t for w in ["system status","how's the system","system info","cpu","memory usage","ram usage"]):
        return ("system_status", None)

    # Calculator / quick math
    _math = re.search(r'(?:what(?:\'s|\s+is)\s+)?(\d[\d\s\+\-\*\/\.\(\)]+)(?:\s*=\s*\?)?$', t)
    if _math and any(op in _math.group(1) for op in ['+','-','*','/','^']):
        return ("calculate", _math.group(1).strip())

    # Clipboard operations
    if any(w in t for w in ["what's in my clipboard","read clipboard","paste that","clipboard content"]):
        return ("clipboard_read", None)

    # Clear notes
    if any(w in t for w in ["clear my notes","delete my notes","wipe my notes","clear notes"]):
        return ("clear_notes", None)

    # Timer — "set a timer for 5 minutes" / "set a 30 second timer"
    if any(w in t for w in ["set a timer","set timer","timer for","start a timer","start timer"]):
        return ("timer", t)
    # Cancel timer
    if any(w in t for w in ["cancel timer","stop timer","cancel the timer","kill the timer"]):
        return ("cancel_timer", None)

    # Reminder — "remind me in 10 minutes to check the oven"
    _remind = re.search(r'remind me\s+in\s+(.+?)\s+to\s+(.+)', t)
    if _remind:
        return ("reminder", _remind.group(1), _remind.group(2))

    # Window controls
    if any(w in t for w in ["minimize window","minimise window","minimize this","minimise this"]):
        return ("window", "minimize")
    if any(w in t for w in ["maximize window","maximise window","maximize this","maximise this","fullscreen"]):
        return ("window", "maximize")
    if any(w in t for w in ["close window","close this window","close this tab"]):
        return ("window", "close")

    # Empty recycle bin
    if any(w in t for w in ["empty recycle bin","empty the recycle bin","clear recycle bin"]):
        return ("recycle_bin", None)

    # IP address
    if any(w in t for w in ["what's my public ip","public ip","external ip","what's my ip address"]):
        return ("ip_address", "public")
    if any(w in t for w in ["what's my ip","my ip","local ip","ip address"]):
        return ("ip_address", "local")

    # Internet / network check
    if any(w in t for w in ["check my internet","am i connected","internet connection","check connection","check network"]):
        return ("internet_check", None)

    # Battery
    if any(w in t for w in ["battery","battery level","battery status","how much battery"]):
        return ("battery", None)

    # Brightness
    if "brightness up"   in t or "increase brightness" in t: return ("brightness", "up")
    if "brightness down" in t or "decrease brightness" in t or "lower brightness" in t: return ("brightness", "down")

    # Dark / light mode toggle
    if any(w in t for w in ["dark mode","enable dark mode","turn on dark mode"]):
        return ("theme", "dark")
    if any(w in t for w in ["light mode","enable light mode","turn on light mode"]):
        return ("theme", "light")
    if any(w in t for w in ["toggle dark mode","toggle theme","switch theme","toggle mode"]):
        return ("theme", "toggle")

    # Windows Settings shortcuts
    _settings_map = {
        "sound settings":      "ms-settings:sound",
        "display settings":    "ms-settings:display",
        "bluetooth settings":  "ms-settings:bluetooth",
        "wifi settings":       "ms-settings:network-wifi",
        "network settings":    "ms-settings:network",
        "update settings":     "ms-settings:windowsupdate",
        "privacy settings":    "ms-settings:privacy",
        "power settings":      "ms-settings:powersleep",
        "storage settings":    "ms-settings:storagesense",
        "app settings":        "ms-settings:appsfeatures",
    }
    for label, uri in _settings_map.items():
        if label in t:
            return ("open_settings", label, uri)
    if any(w in t for w in ["open settings","windows settings","system settings"]):
        return ("open_settings", "settings", "ms-settings:")

    # Coin flip / dice / random number
    if any(w in t for w in ["flip a coin","toss a coin","heads or tails"]):
        return ("random", "coin")
    if any(w in t for w in ["roll a die","roll a dice","roll the dice","roll dice"]):
        return ("random", "dice")
    _rand_match = re.search(r'random number(?:\s+between\s+(\d+)\s+and\s+(\d+))?', t)
    if _rand_match:
        lo = int(_rand_match.group(1)) if _rand_match.group(1) else 1
        hi = int(_rand_match.group(2)) if _rand_match.group(2) else 100
        return ("random", "number", lo, hi)

    # F1 news briefing
    if any(w in t for w in ["f1 news", "formula 1 news", "formula one news",
                             "f1 update", "formula 1 update", "f1 headlines",
                             "what's happening in f1", "what's going on in f1",
                             "formula 1", "latest f1", "f1 latest"]):
        return ("f1_news", None)

    # World news briefing
    if any(w in t for w in ["world news", "current events", "latest news", "news update",
                             "what's going on in the world", "what is going on in the world",
                             "what's happening in the world", "catch me up on the news",
                             "what's in the news"]):
        return ("world_news", None)

    # Audio device switching via SoundSwitch / nircmd
    # "switch to headphones", "switch to speakers", "switch to [device name]"
    _sw_specific = re.search(r'switch(?:ing)?\s+(?:audio\s+)?(?:to|output\s+to)\s+(?:my\s+)?([\w\s]+)', t)
    if _sw_specific:
        device = _sw_specific.group(1).strip()
        return ("audio_switch", device)
    # "switch audio", "switch output", "switch sound", "cycle audio"
    if any(w in t for w in ["switch audio", "switch output", "switch sound", "cycle audio",
                             "switch my audio", "change audio output", "change sound output"]):
        return ("audio_switch", None)

    # Now playing (Spotify + browser)
    if any(w in t for w in ["what's playing", "what is playing", "now playing",
                             "what song is playing", "current song",
                             "what's the song", "what's on"]):
        return ("now_playing", None)

    # Shazam-style audio recognition from mic
    if any(w in t for w in ["what song is this", "what's this song", "identify this song",
                             "identify the song", "shazam this", "shazam",
                             "what music is this", "what's this music",
                             "recognize this song", "song recognition",
                             "what is this song", "find this song"]):
        return ("identify_music", None)

    # Anime OST recognition — vocal removal + fingerprint
    if any(w in t for w in ["what anime song is this", "find the anime music",
                             "what's the ost", "what is the ost", "identify the ost",
                             "anime ost", "anime music lookup", "what's the background music",
                             "find the background song", "what's playing in the background",
                             "isolate the background", "remove the voices",
                             "what's the soundtrack", "identify the soundtrack"]):
        return ("anime_ost", None)

    # List active timers
    if any(w in t for w in ["list timers", "active timers", "what timers", "how many timers",
                             "check timers", "show timers", "my timers"]):
        return ("list_timers", None)

    # Stopwatch
    if any(w in t for w in ["start stopwatch", "start the stopwatch", "begin stopwatch", "stopwatch start"]):
        return ("stopwatch", "start")
    if any(w in t for w in ["stop stopwatch", "pause stopwatch", "end stopwatch", "stopwatch stop"]):
        return ("stopwatch", "stop")
    if any(w in t for w in ["reset stopwatch", "clear stopwatch", "stopwatch reset"]):
        return ("stopwatch", "reset")
    if any(w in t for w in ["elapsed time", "stopwatch time", "check stopwatch", "how long has it been",
                             "how long have i been"]):
        return ("stopwatch", "check")

    # Tell a joke
    if any(w in t for w in ["tell me a joke", "tell a joke", "make me laugh", "say something funny",
                             "give me a joke"]):
        return ("joke", None)

    # Repeat last
    if any(w in t for w in ["say that again", "repeat that", "repeat last", "what did you say",
                             "can you repeat that"]):
        return ("repeat_last", None)

    # Time in another city
    _time_in_m = re.search(r"what(?:'s|\s+is)\s+(?:the\s+)?time\s+in\s+(.+)", t)
    if _time_in_m:
        return ("time_in", _time_in_m.group(1).strip().rstrip("?"))

    # Unit / temperature conversion
    _conv_m = re.search(r"convert\s+([\d\.]+)\s+([\w]+)\s+(?:to|into|in)\s+([\w]+)", t)
    if _conv_m:
        return ("unit_convert", float(_conv_m.group(1)), _conv_m.group(2).lower(), _conv_m.group(3).lower())

    # Word definition
    _def_word = None
    if t.startswith("define "):
        _def_word = t[7:].split()[0].rstrip("?")
    else:
        _dm = re.search(r"what does (\w+) mean", t)
        if _dm:
            _def_word = _dm.group(1)
        else:
            _dm2 = re.search(r"(?:meaning|definition) of (\w+)", t)
            if _dm2:
                _def_word = _dm2.group(1)
    if _def_word:
        return ("define", _def_word)

    # Wikipedia quick lookup — "who is X" / "who was X" / "tell me about X"
    _wiki_m = re.search(r"^(?:who\s+(?:is|was|are|were)|tell me about)\s+(.+)", t)
    if _wiki_m:
        return ("wikipedia", _wiki_m.group(1).strip().rstrip("?"))

    return ("ai", t)

# ─────────────────────────────────────────
#  PERSONALITY RESPONSES
# ─────────────────────────────────────────
_QUIPS_LAUNCH = [
    "Right away, sir.",
    "Of course, sir.",
    "As you wish, sir.",
    "On it, sir.",
    "Certainly, sir.",
    "Consider it done, sir.",
    "Initiating now, sir.",
]
_QUIPS_DONE = [
    "Done, sir.",
    "All set, sir.",
    "Completed, sir.",
    "There you go, sir.",
    "Finished, sir.",
]
_QUIPS_SEARCH = [
    "Searching now, sir.",
    "Pulling that up for you, sir.",
    "On it, sir.",
    "Querying the web, sir.",
]
_QUIPS_MEDIA = {
    "playpause": ["Toggling playback, sir.", "Play-pause, sir.", "Done, sir."],
    "next":      ["Next track, sir.", "Skipping ahead, sir.", "Moving on, sir."],
    "prev":      ["Going back, sir.", "Previous track, sir.", "Rewinding, sir."],
    "stop":      ["Stopping playback, sir.", "Music off, sir.", "Silencing, sir."],
}
_QUIPS_VOLUME = {
    "up":   ["Volume up, sir.", "Turning it up, sir.", "Louder, sir."],
    "down": ["Volume down, sir.", "Turning it down, sir.", "Quieter, sir."],
    "mute": ["Muted, sir.", "Silencing audio, sir.", "Going quiet, sir."],
}

def _q(pool): return random.choice(pool)

def _precache_quips():
    """Pre-generate TTS audio for all short personality quips in a background thread."""
    if not _USE_EDGE_TTS:
        return
    phrases = list(dict.fromkeys(
        _QUIPS_LAUNCH + _QUIPS_DONE + _QUIPS_SEARCH
        + [q for pool in _QUIPS_MEDIA.values() for q in pool]
        + [q for pool in _QUIPS_VOLUME.values() for q in pool]
        + ["JARVIS online. Good to be of service, sir.",
           "Yes, sir?", "At your service, sir.", "How can I help, sir?", "Standing by, sir.",
           "My apologies, sir.", "No worries, sir.",
           "On it, sir.", "Let me think on that, sir.", "One moment, sir.", "Calculating, sir.",
           "Scanning the feeds, sir.", "Pulling the latest headlines, sir.", "Checking global activity, sir.",
           "Stopwatch started, sir.", "Stopwatch stopped.", "Stopwatch reset, sir.",
           "The stopwatch isn't running, sir.", "You have no active timers, sir.",
           "Let me find one for you, sir.", "Looking that up, sir.",
           "Nothing appears to be playing on Spotify at the moment, sir.",
           "Checking the paddock feeds, sir.", "Pulling the latest from the pit lane, sir.", "Scanning the grid, sir."]
    ))

    async def _warm():
        for phrase in phrases:
            try:
                await _cache_tts_async(phrase)
            except Exception as e:
                log.debug(f"Pre-cache skipped '{phrase[:30]}': {e}")
        log.info(f"  ✓ TTS cache warm — {len(_audio_cache)} phrases ready.")

    threading.Thread(target=lambda: asyncio.run(_warm()), daemon=True, name="tts-precache").start()

_precache_quips()

def _find_folder(name):
    """Search common user directories for an existing folder named *name* (case-insensitive).
    Returns the full path if found, otherwise None."""
    search_roots = [
        os.path.expanduser("~/Desktop"),
        os.path.expanduser("~/Documents"),
        os.path.expanduser("~/Downloads"),
        os.path.expanduser("~/Pictures"),
        os.path.expanduser("~/Videos"),
        os.path.expanduser("~/Music"),
        os.path.expanduser("~"),
        os.path.expandvars("%USERPROFILE%/OneDrive"),
    ]
    name_lower = name.lower()
    for root in search_roots:
        if not os.path.isdir(root):
            continue
        try:
            for entry in os.scandir(root):
                if entry.is_dir() and entry.name.lower() == name_lower:
                    return entry.path
        except PermissionError:
            pass
    return None

# ─────────────────────────────────────────
#  NEWS FETCHER
# ─────────────────────────────────────────
def _fetch_f1_headlines(max_per_feed=3):
    """Fetch top F1 headlines from RSS feeds. Returns list of (title, url) tuples."""
    import xml.etree.ElementTree as ET
    feeds = [
        "https://www.autosport.com/rss/f1/news/",
        "https://www.motorsport.com/rss/f1/news/",
        "https://feeds.feedburner.com/SkySportsF1",
        "https://www.racefans.net/feed/",
        "https://www.bbc.co.uk/sport/formula1/rss.xml",
    ]
    headlines = []
    for feed_url in feeds:
        try:
            r = requests.get(feed_url, timeout=5, headers={"User-Agent": "Mozilla/5.0"})
            root = ET.fromstring(r.content)
            for item in root.findall(".//item")[:max_per_feed]:
                title = item.findtext("title", "").strip()
                link  = item.findtext("link",  "").strip()
                if title and link:
                    headlines.append((title, link))
            if headlines:
                break  # one working feed is enough for a sharp briefing
        except Exception as e:
            log.warning(f"F1 RSS fetch failed ({feed_url}): {e}")
    return headlines

def _fetch_headlines(max_per_feed=2):
    """Fetch top headlines from RSS feeds. Returns list of (title, url) tuples."""
    import xml.etree.ElementTree as ET
    feeds = [
        "http://feeds.bbci.co.uk/news/rss.xml",
        "https://feeds.reuters.com/reuters/topNews",
        "https://www.aljazeera.com/xml/rss/all.xml",
        "https://www.theguardian.com/world/rss",
        "https://feeds.npr.org/1001/rss.xml",
    ]
    headlines = []
    for feed_url in feeds:
        try:
            r = requests.get(feed_url, timeout=5, headers={"User-Agent": "Mozilla/5.0"})
            root = ET.fromstring(r.content)
            for item in root.findall(".//item")[:max_per_feed]:
                title = item.findtext("title", "").strip()
                link  = item.findtext("link",  "").strip()
                if title and link:
                    headlines.append((title, link))
        except Exception as e:
            log.warning(f"RSS fetch failed ({feed_url}): {e}")
    return headlines

def _get_spotify_now_playing():
    """Return 'Track - Artist' from Spotify's window title, or None if nothing is playing."""
    user32 = ctypes.windll.user32
    try:
        output = subprocess.check_output(
            ['tasklist', '/fi', 'imagename eq Spotify.exe', '/fo', 'csv', '/nh'],
            text=True, creationflags=_NO_WINDOW
        )
        pids = {int(line.split(',')[1].strip('"')) for line in output.strip().split('\n')
                if line and 'Spotify.exe' in line}
    except Exception:
        return None
    if not pids:
        return None
    result = []
    def _enum_proc(hwnd, lParam):
        pid = ctypes.c_ulong()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value in pids:
            length = user32.GetWindowTextLengthW(hwnd)
            if length > 4:
                buf = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buf, length + 1)
                title = buf.value.strip()
                if " - " in title and title not in ("Spotify", "Spotify Free", "Spotify Premium"):
                    result.append(title)
        return True
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    user32.EnumWindows(WNDENUMPROC(_enum_proc), 0)
    return result[0] if result else None

def _get_music_from_browser():
    """Scan visible browser window titles for currently playing media.
    Returns (source, display_string) or None."""
    user32 = ctypes.windll.user32
    candidates = []

    def _enum(hwnd, lParam):
        if not user32.IsWindowVisible(hwnd):
            return True
        length = user32.GetWindowTextLengthW(hwnd)
        if length < 3:
            return True
        buf = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buf, length + 1)
        title = buf.value.strip()
        if title:
            candidates.append(title)
        return True

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    user32.EnumWindows(WNDENUMPROC(_enum), 0)

    for title in candidates:
        # YouTube: "Song Name - Artist Name - YouTube"  or  "Video Title - YouTube"
        if title.endswith("- YouTube") or "- YouTube" in title:
            name = re.sub(r'\s*[-–]\s*YouTube\s*$', '', title).strip()
            if name and name.lower() not in ("youtube", "youtube music", "home"):
                return ("YouTube", name)
        # YouTube Music: "Song - Artist - YouTube Music"
        if title.endswith("- YouTube Music") or "- YouTube Music" in title:
            name = re.sub(r'\s*[-–]\s*YouTube Music\s*$', '', title).strip()
            if name and name.lower() not in ("youtube music", "home", "library"):
                return ("YouTube Music", name)
        # Spotify Web Player: "Song - Artist | Spotify"
        if "Spotify" in title and (" - " in title or " | " in title):
            name = re.sub(r'\s*[|]\s*Spotify.*$', '', title)
            name = re.sub(r'\s*[-–]\s*Spotify.*$', '', name).strip()
            if name and name.lower() not in ("spotify", "home", "search"):
                return ("Spotify Web", name)
        # Netflix: "Show Name | Netflix" — less useful but readable
        if title.endswith("| Netflix") or "| Netflix" in title:
            name = re.sub(r'\s*[|]\s*Netflix\s*$', '', title).strip()
            if name and name.lower() != "netflix":
                return ("Netflix", name)
        # Twitch: "StreamerName - Twitch"
        if title.endswith("- Twitch") or "- Twitch" in title:
            name = re.sub(r'\s*[-–]\s*Twitch\s*$', '', title).strip()
            if name and name.lower() not in ("twitch", "home", "browse"):
                return ("Twitch", name)

    return None

def _record_system_audio(duration=6):
    """Capture system audio via WASAPI loopback (what's playing through speakers).
    Returns (frames, channels, rate, sample_width) or None."""
    try:
        import pyaudiowpatch as pyaudio
    except ImportError:
        log.warning("pyaudiowpatch not installed — run: pip install pyaudiowpatch")
        return None

    p = pyaudio.PyAudio()
    try:
        try:
            wasapi_info = p.get_host_api_info_by_type(pyaudio.paWASAPI)
        except OSError:
            log.warning("WASAPI not available on this system.")
            return None

        default_out = p.get_device_info_by_index(wasapi_info["defaultOutputDevice"])
        loopback = None

        # Prefer the generator helper (pyaudiowpatch >= 0.2.12)
        try:
            for dev in p.get_loopback_device_info_generator():
                if default_out["name"] in dev["name"]:
                    loopback = dev
                    break
            if not loopback:
                for dev in p.get_loopback_device_info_generator():
                    loopback = dev   # any loopback
                    break
        except AttributeError:
            # Older pyaudiowpatch — scan manually
            for i in range(p.get_device_count()):
                dev = p.get_device_info_by_index(i)
                if dev.get("isLoopbackDevice"):
                    if not loopback or default_out["name"] in dev["name"]:
                        loopback = dev

        if not loopback:
            log.warning("No WASAPI loopback device found.")
            return None

        # WASAPI loopback devices are output devices captured as input.
        # maxInputChannels may be 0; use maxOutputChannels as fallback.
        ch_in  = int(loopback.get("maxInputChannels")  or 0)
        ch_out = int(loopback.get("maxOutputChannels") or 2)
        channels = max(1, min(ch_in if ch_in > 0 else ch_out, 2))
        rate  = int(loopback["defaultSampleRate"])
        chunk = 1024
        log.info(f"Loopback: '{loopback['name']}' | {channels}ch | {rate}Hz")

        stream = p.open(
            format=pyaudio.paInt16,
            channels=channels,
            rate=rate,
            input=True,
            input_device_index=int(loopback["index"]),
            frames_per_buffer=chunk,
        )
        frames = [stream.read(chunk, exception_on_overflow=False)
                  for _ in range(int(rate / chunk * duration))]
        stream.stop_stream()
        stream.close()
        sample_width = p.get_sample_size(pyaudio.paInt16)
        return frames, channels, rate, sample_width

    except Exception as e:
        log.error(f"System audio capture failed: {e}")
        return None
    finally:
        p.terminate()

def _recognize_audio(duration=6):
    """Capture system audio (loopback) and identify via AudD.
    Falls back to microphone if loopback is unavailable.
    Returns dict with title/artist/album/spotify or None."""
    import wave

    frames = channels = rate = sample_width = None
    source = None

    # ── System audio (loopback) ──────────────────────────────────────
    result = _record_system_audio(duration)
    if result:
        frames, channels, rate, sample_width = result
        source = "system audio"

    # ── Microphone fallback ──────────────────────────────────────────
    if frames is None:
        log.info("Falling back to microphone for audio capture.")
        try:
            import pyaudio
            RATE, CHUNK, CHANNELS = 44100, 1024, 1
            p = pyaudio.PyAudio()
            stream = p.open(format=pyaudio.paInt16, channels=CHANNELS, rate=RATE,
                            input=True, frames_per_buffer=CHUNK)
            frames = [stream.read(CHUNK, exception_on_overflow=False)
                      for _ in range(int(RATE / CHUNK * duration))]
            stream.stop_stream()
            stream.close()
            sample_width = p.get_sample_size(pyaudio.paInt16)
            p.terminate()
            channels, rate = CHANNELS, RATE
            source = "microphone"
        except ImportError:
            log.warning("Neither pyaudiowpatch nor pyaudio available.")
            return None
        except Exception as e:
            log.error(f"Mic capture failed: {e}")
            return None

    # ── Write temp WAV ────────────────────────────────────────────────
    with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as f:
        tmp_path = f.name
    try:
        with wave.open(tmp_path, 'wb') as wf:
            wf.setnchannels(channels)
            wf.setsampwidth(sample_width)
            wf.setframerate(rate)
            wf.writeframes(b''.join(frames))
    except Exception as e:
        log.error(f"WAV write error: {e}")
        try: os.unlink(tmp_path)
        except OSError: pass
        return None

    log.info(f"Sending {source} audio to AudD ({duration}s, {rate}Hz, {channels}ch)...")

    # ── AudD recognition ─────────────────────────────────────────────
    try:
        with open(tmp_path, 'rb') as audio_file:
            r = requests.post(
                "https://api.audd.io/",
                data={"return": "spotify,apple_music"},
                files={"file": ("audio.wav", audio_file, "audio/wav")},
                timeout=20,
            )
        try: os.unlink(tmp_path)
        except OSError: pass

        if r.status_code == 200:
            data = r.json()
            if data.get("status") == "success" and data.get("result"):
                res = data["result"]
                spotify_url = (res.get("spotify") or {}).get("external_urls", {}).get("spotify", "")
                return {
                    "title":   res.get("title", ""),
                    "artist":  res.get("artist", ""),
                    "album":   res.get("album", ""),
                    "spotify": spotify_url,
                }
            log.info(f"AudD: no match — {data.get('status')} {data.get('error', {})}")
        else:
            log.warning(f"AudD HTTP {r.status_code}: {r.text[:200]}")
        return None
    except Exception as e:
        log.error(f"AudD request error: {e}")
        try: os.unlink(tmp_path)
        except OSError: pass
        return None

def _isolate_background_music(wav_path):
    """Run Demucs CLI (--two-stems vocals) to strip dialogue/vocals.
    Uses subprocess so every line of Demucs output — including download
    progress and errors — is routed into jarvis.log.
    Returns path to the no_vocals WAV or None on failure."""
    import shutil

    out_dir = tempfile.mkdtemp(prefix="jarvis_demucs_")
    try:
        # --two-stems vocals: only separates into vocals vs no_vocals
        # (much faster than 4-stem htdemucs and exactly what we need)
        cmd = [
            sys.executable, "-m", "demucs",
            "--two-stems", "vocals",
            "-n", "htdemucs",
            "--out", out_dir,
            wav_path,
        ]
        log.info(f"Demucs cmd: {' '.join(cmd)}")

        proc = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,   # merge stderr → stdout so we capture everything
            text=True,
            timeout=300,                # 5-minute hard cap
            creationflags=_NO_WINDOW,
        )

        # Pipe every line of Demucs output into our log
        for line in (proc.stdout or "").splitlines():
            if line.strip():
                log.info(f"[demucs] {line}")

        if proc.returncode != 0:
            log.error(f"Demucs exited with code {proc.returncode}")
            return None

        # Expected output: {out_dir}/htdemucs/{trackname}/no_vocals.wav
        track_name   = os.path.splitext(os.path.basename(wav_path))[0]
        no_vocals_path = os.path.join(out_dir, "htdemucs", track_name, "no_vocals.wav")

        if not os.path.exists(no_vocals_path):
            # Log every file Demucs actually created so we can debug
            for root, _, files in os.walk(out_dir):
                for f in files:
                    log.info(f"[demucs] created: {os.path.join(root, f)}")
            log.error(f"no_vocals.wav not found at expected path: {no_vocals_path}")
            return None

        # Copy out before we delete the demucs temp dir
        with tempfile.NamedTemporaryFile(suffix="_no_vocals.wav", delete=False) as f:
            bg_path = f.name
        shutil.copy2(no_vocals_path, bg_path)
        log.info(f"Background isolated → {bg_path}")
        return bg_path

    except subprocess.TimeoutExpired:
        log.error("Demucs timed out (>5 min). Try a shorter capture duration.")
        return None
    except Exception as e:
        log.error(f"Demucs error: {e}", exc_info=True)
        return None
    finally:
        shutil.rmtree(out_dir, ignore_errors=True)

def _identify_anime_ost(duration=10):
    """Full pipeline: capture system audio → strip vocals → fingerprint via AudD.
    Returns result dict or None."""
    import wave

    # ── 1. Capture system audio (loopback preferred) ──────────────────
    captured = _record_system_audio(duration)
    if not captured:
        log.warning("System audio capture unavailable for anime OST recognition.")
        return None
    frames, channels, rate, sample_width = captured

    with tempfile.NamedTemporaryFile(suffix="_raw.wav", delete=False) as f:
        raw_path = f.name
    with wave.open(raw_path, "wb") as wf:
        wf.setnchannels(channels)
        wf.setsampwidth(sample_width)
        wf.setframerate(rate)
        wf.writeframes(b"".join(frames))
    log.info(f"System audio saved: {raw_path} ({duration}s, {rate}Hz, {channels}ch)")

    # ── 2. Isolate background music with Demucs ────────────────────────
    bg_path = _isolate_background_music(raw_path)
    try: os.unlink(raw_path)
    except OSError: pass

    if not bg_path:
        return None

    # ── 3. Send isolated background to AudD ───────────────────────────
    try:
        with open(bg_path, "rb") as audio_file:
            r = requests.post(
                "https://api.audd.io/",
                data={"return": "spotify,apple_music"},
                files={"file": ("background.wav", audio_file, "audio/wav")},
                timeout=20,
            )
        try: os.unlink(bg_path)
        except OSError: pass

        if r.status_code == 200:
            data = r.json()
            if data.get("status") == "success" and data.get("result"):
                res = data["result"]
                spotify_url = (res.get("spotify") or {}).get("external_urls", {}).get("spotify", "")
                return {
                    "title":   res.get("title", ""),
                    "artist":  res.get("artist", ""),
                    "album":   res.get("album", ""),
                    "spotify": spotify_url,
                }
            log.info(f"AudD no match — {data.get('status')} {data.get('error', '')}")
        else:
            log.warning(f"AudD HTTP {r.status_code}: {r.text[:200]}")
        return None
    except Exception as e:
        log.error(f"AudD request error: {e}")
        try: os.unlink(bg_path)
        except OSError: pass
        return None

def _fetch_joke():
    """Fetch a random joke. Returns (setup, punchline) or None."""
    try:
        r = requests.get("https://official-joke-api.appspot.com/random_joke", timeout=5)
        if r.status_code == 200:
            data = r.json()
            return data.get("setup", ""), data.get("punchline", "")
    except Exception as e:
        log.warning(f"Joke API failed: {e}")
    return None

def _convert_units(value, from_unit, to_unit):
    """Convert between units. Returns (result, category) or (None, None) if incompatible."""
    temp_map = {"f": "fahrenheit", "c": "celsius", "k": "kelvin",
                "fahrenheit": "fahrenheit", "celsius": "celsius", "kelvin": "kelvin"}
    if from_unit in temp_map and to_unit in temp_map:
        f, t = temp_map[from_unit], temp_map[to_unit]
        if f == t:                                   return value, "temperature"
        if f == "celsius"    and t == "fahrenheit":  return value * 9/5 + 32, "temperature"
        if f == "fahrenheit" and t == "celsius":     return (value - 32) * 5/9, "temperature"
        if f == "celsius"    and t == "kelvin":      return value + 273.15, "temperature"
        if f == "kelvin"     and t == "celsius":     return value - 273.15, "temperature"
        if f == "fahrenheit" and t == "kelvin":      return (value - 32) * 5/9 + 273.15, "temperature"
        if f == "kelvin"     and t == "fahrenheit":  return (value - 273.15) * 9/5 + 32, "temperature"
    for category, table in _UNIT_GROUPS.items():
        if from_unit in table and to_unit in table:
            return value * table[from_unit] / table[to_unit], category
    return None, None

def _get_time_in_city(city):
    """Return current time string for the given city, or None if unknown."""
    tz_name = _CITY_TIMEZONES.get(city.lower().strip())
    if not tz_name:
        return None
    try:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo(tz_name)
        now = datetime.datetime.now(tz)
        hour = now.strftime("%I").lstrip("0") or "12"
        return f"{hour}:{now.strftime('%M')} {now.strftime('%p')}"
    except Exception as e:
        log.warning(f"Timezone lookup error: {e}")
        return None

# ─────────────────────────────────────────
#  COMMAND EXECUTOR
# ─────────────────────────────────────────
def __control_spotify(action):
    try:
        output = subprocess.check_output(
            ['tasklist', '/fi', 'imagename eq Spotify.exe', '/fo', 'csv', '/nh'],
            text=True, creationflags=_NO_WINDOW
        )
        pids = [int(line.split(',')[1].strip('"')) for line in output.strip().split('\n') if line and 'Spotify.exe' in line]
    except Exception: return False
    if not pids: return False

    user32 = ctypes.windll.user32
    spotify_hwnd = None
    def enum_windows_proc(hwnd, lParam):
        nonlocal spotify_hwnd
        pid = ctypes.c_ulong()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value in pids and user32.GetWindowTextLengthW(hwnd) > 0:
            spotify_hwnd = hwnd
            return False
        return True
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    user32.EnumWindows(WNDENUMPROC(enum_windows_proc), 0)
    
    if not spotify_hwnd: return False
    cmd = {"playpause": 14, "next": 11, "prev": 12, "stop": 13}.get(action, 14)
    # SendMessageW (synchronous) is more reliable than PostMessageW for Spotify
    user32.SendMessageW(spotify_hwnd, 0x0319, 0, cmd << 16)
    return True

def process_command(text):
    log.info(f"📥 {text}")
    intent, *args = parse_intent(text)
    log.info(f"🎯 {intent} | {args}")

    if intent == "open_app":
        name, path = args
        if not path:
            speak(f"I'm afraid I couldn't locate {name} on this system, sir.")
            return
        speak(f"Opening {name}. {_q(_QUIPS_LAUNCH)}")
        try:
            os.startfile(path)
            log.info(f"Opened {name}.")
        except Exception:
            speak(f"I encountered an error launching {name}, sir.")

    elif intent == "sequence":
        name = args[0]
        urls = SEQUENCES[name]
        speak(f"Activating {name}, sir. Spinning up {len(urls)} systems.")
        log.info(f"{name}. {len(urls)} items.")
        run_sequence(name)

    elif intent == "multi_sequence":
        seqs = args[0]
        label = " and ".join(seqs)
        speak(f"Running {label} simultaneously, sir.")
        log.info(f"Running sequences: {', '.join(seqs)}.")
        all_items = []
        for name in seqs:
            all_items.extend(SEQUENCES.get(name, []))
        for i, item in enumerate(all_items):
            threading.Timer(i * 0.8, __launch_sequence_item, args=[item, i > 0]).start()

    elif intent == "open_site":
        name, url = args
        speak(f"Navigating to {name}, sir.")
        open_in_brave(url)
        log.info(f"Opened {name}.")

    elif intent == "web_search":
        query = args[0]
        speak(f"{_q(_QUIPS_SEARCH)}")
        open_in_brave(f"https://www.google.com/search?q={query.replace(' ', '+')}")
        log.info(f"Searching for {query}.")

    elif intent == "media":
        action = args[0]
        speak(_q(_QUIPS_MEDIA.get(action, _QUIPS_DONE)))
        try:
            sent = __control_spotify(action)
            if sent:
                log.info(f"Spotify media command: {action}")
            else:
                # Spotify not running — launch it, wait, then send command
                spotify_path = APP_MAP.get("spotify")
                if spotify_path:
                    log.info("Spotify not running — launching it.")
                    os.startfile(spotify_path)
                    time.sleep(5)
                    sent = __control_spotify(action)
                    if sent:
                        log.info(f"Launched Spotify and sent command: {action}")
                    else:
                        log.warning("Launched Spotify but couldn't send media command yet.")
                else:
                    log.warning("Spotify not found on this system.")
        except Exception as e:
            log.error(f"Media command error: {e}")
            speak("I'm having trouble with that media command, sir.")

    elif intent == "volume":
        d, level = args[0], args[1]
        if d == "set" and level is not None:
            speak(f"Setting volume to {level} percent, sir.")
            nircmd_val = int(level / 100 * 65535)  # nircmd range is 0–65535
            subprocess.run(["nircmd.exe", "setsysvolume", str(nircmd_val)], creationflags=_NO_WINDOW)
            log.info(f"Volume set to {level}%.")
        else:
            speak(_q(_QUIPS_VOLUME.get(d, _QUIPS_DONE)))
            if d == "up":
                for _ in range(5): subprocess.run(["nircmd.exe", "changesysvolume", "5000"], creationflags=_NO_WINDOW)
            elif d == "down":
                for _ in range(5): subprocess.run(["nircmd.exe", "changesysvolume", "-5000"], creationflags=_NO_WINDOW)
            elif d == "mute":
                subprocess.run(["nircmd.exe", "mutesysvolume", "2"], creationflags=_NO_WINDOW)
            log.info(f"Volume {d}. Done.")

    elif intent == "time":
        speak(f"It is currently {time.strftime('%I:%M %p')}, sir.")
    elif intent == "date":
        speak(f"Today is {time.strftime('%A, %B %d, %Y')}, sir.")
    elif intent == "shutdown":
        speak("Initiating shutdown sequence, sir. You have ten seconds to abort.")
        time.sleep(10)
        subprocess.run(["shutdown", "/s", "/t", "1"], creationflags=_NO_WINDOW)
    elif intent == "restart_jarvis":
        speak("Restarting JARVIS. Back in a moment, sir.")
        log.info("Self-restart initiated.")
        flags = _NO_WINDOW
        subprocess.Popen([sys.executable, os.path.abspath(__file__)], creationflags=flags)
        time.sleep(2)
        os._exit(0)

    elif intent == "restart":
        speak("Restarting all systems, sir. Back in a moment.")
        time.sleep(10)
        subprocess.run(["shutdown", "/r", "/t", "1"], creationflags=_NO_WINDOW)
    elif intent == "sleep":
        speak("Putting the system to sleep, sir.")
        time.sleep(2)
        subprocess.run(["rundll32.exe", "powrprof.dll,SetSuspendState", "0,1,0"], creationflags=_NO_WINDOW)

    elif intent == "lock":
        speak("Locking your workstation, sir.")
        ctypes.windll.user32.LockWorkStation()
        log.info("Screen locked.")

    elif intent == "screenshot":
        custom_name = args[0] if args else None
        folder_arg  = args[1] if len(args) > 1 else None

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

        if folder_arg is None:
            save_dir = os.path.expanduser("~/Desktop")
            dir_label = "Desktop"
        elif folder_arg[0] == "new":
            # Create a brand-new folder on the Desktop
            new_name = folder_arg[1]
            save_dir = os.path.join(os.path.expanduser("~/Desktop"), new_name)
            os.makedirs(save_dir, exist_ok=True)
            dir_label = f"new folder '{new_name}' on your Desktop"
        else:
            # Known folder alias or bare name
            folder_key = folder_arg[1].lower()
            if folder_key in FOLDER_MAP:
                save_dir  = FOLDER_MAP[folder_key]
                dir_label = folder_key
            else:
                # Search common locations for an existing folder with that name
                found = _find_folder(folder_key)
                if found:
                    save_dir  = found
                    dir_label = f"folder '{folder_key}'"
                else:
                    # Not found anywhere — create it on the Desktop
                    save_dir = os.path.join(os.path.expanduser("~/Desktop"), folder_key)
                    os.makedirs(save_dir, exist_ok=True)
                    dir_label = f"new folder '{folder_key}' on your Desktop"

        file_name = (custom_name if custom_name else f"screenshot_{ts}") + ".png"
        save_path = os.path.join(save_dir, file_name)

        speak("Taking a screenshot, sir.")
        try:
            from PIL import ImageGrab
            img = ImageGrab.grab()
            img.save(save_path)
            if custom_name:
                speak(f"Screenshot saved as {file_name} in your {dir_label}, sir.")
            else:
                speak(f"Screenshot saved in your {dir_label}, sir.")
            log.info(f"Screenshot saved: {save_path}")
        except Exception as e:
            speak("I had trouble capturing the screen, sir.")
            log.error(f"Screenshot error: {e}")

    elif intent == "weather":
        speak("Pulling up the weather for you, sir.")
        open_in_brave("https://wttr.in/?lang=en")
        log.info("Opened weather.")

    elif intent == "close_app":
        app_name = args[0]
        exe_map = {
            "spotify": "Spotify.exe", "brave": "brave.exe", "discord": "Discord.exe",
            "firefox": "firefox.exe", "edge": "msedge.exe", "steam": "steam.exe",
            "vlc": "vlc.exe", "obs": "obs64.exe", "vs code": "Code.exe",
            "notepad": "notepad.exe", "calculator": "Calculator.exe",
            "league of legends": "LeagueClient.exe", "osu": "osu!.exe",
        }
        exe = exe_map.get(app_name, app_name.replace(" ", "") + ".exe")
        speak(f"Closing {app_name}, sir.")
        flags = _NO_WINDOW
        result = subprocess.run(["taskkill", "/f", "/im", exe], capture_output=True, creationflags=flags)
        if result.returncode == 0:
            log.info(f"Closed {app_name} ({exe}).")
        else:
            speak(f"I couldn't find a running instance of {app_name}, sir.")
            log.warning(f"taskkill failed for {exe}: {result.stderr.strip()}")

    elif intent == "open_folder":
        folder_name = args[0]
        path = FOLDER_MAP.get(folder_name)
        if path and os.path.exists(path):
            speak(f"Opening your {folder_name} folder, sir.")
            os.startfile(path)
            log.info(f"Opened folder: {path}")
        else:
            speak(f"I couldn't locate the {folder_name} folder, sir.")

    elif intent == "note":
        note_text = args[0]
        if not note_text:
            speak("What would you like me to note, sir?")
        else:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
            with open(NOTES_FILE, "a", encoding="utf-8") as f:
                f.write(f"[{timestamp}] {note_text}\n")
            speak(f"Noted, sir. I've saved that for you.")
            log.info(f"Note saved: {note_text}")

    elif intent == "read_notes":
        if not os.path.exists(NOTES_FILE):
            speak("You have no notes saved, sir.")
        else:
            with open(NOTES_FILE, "r", encoding="utf-8") as f:
                lines = [l.strip() for l in f.readlines() if l.strip()]
            if not lines:
                speak("Your notes are empty, sir.")
            else:
                recent = lines[-3:]  # read last 3 notes
                speak(f"Your last {len(recent)} note{'s' if len(recent) > 1 else ''}, sir.")
                for line in recent:
                    speak(line)
                log.info(f"Read {len(recent)} notes.")

    elif intent == "system_status":
        speak("Running diagnostics, sir.")
        try:
            import psutil
            cpu = psutil.cpu_percent(interval=1)
            ram_pct = psutil.virtual_memory().percent
            speak(f"CPU is at {cpu:.0f} percent. RAM usage is {ram_pct:.0f} percent, sir.")
            log.info(f"System status: CPU {cpu}%, RAM {ram_pct}%")
        except ImportError:
            speak("System diagnostics require psutil. Run pip install psutil to enable that, sir.")

    elif intent == "calculate":
        expr = args[0]
        try:
            safe_expr = re.sub(r'[^\d\s\+\-\*\/\.\(\)\^]', '', expr).replace('^', '**')
            result = eval(safe_expr, {"__builtins__": {}})
            speak(f"That comes to {result}, sir.")
            log.info(f"Calculated: {expr} = {result}")
        except Exception:
            speak("I couldn't compute that, sir.")

    elif intent == "clipboard_read":
        try:
            import tkinter as tk
            root = tk.Tk()
            root.withdraw()
            content = root.clipboard_get()
            root.destroy()
            if content:
                short = content[:100] + ("..." if len(content) > 100 else "")
                speak(f"Your clipboard contains: {short}, sir.")
            else:
                speak("Your clipboard is empty, sir.")
        except Exception:
            speak("I couldn't read the clipboard, sir.")

    elif intent == "audio_switch":
        speak(random.choice(["Switching audio output, sir.", "Cycling to the next audio device, sir.",
                             "Switching output device, sir."]))
        # Simulate Alt+F11 — SoundSwitch hotkey for cycling devices
        ALT  = 0x12
        F11  = 0x7A
        KEYDOWN, KEYUP = 0, 0x0002
        ctypes.windll.user32.keybd_event(ALT, 0, KEYDOWN, 0)
        ctypes.windll.user32.keybd_event(F11, 0, KEYDOWN, 0)
        ctypes.windll.user32.keybd_event(F11, 0, KEYUP,   0)
        ctypes.windll.user32.keybd_event(ALT, 0, KEYUP,   0)
        log.info("Audio switch triggered via Alt+F11 (SoundSwitch hotkey).")

    elif intent == "clear_notes":
        if os.path.exists(NOTES_FILE):
            open(NOTES_FILE, "w", encoding="utf-8").close()
            speak("Your notes have been cleared, sir.")
            log.info("Notes file cleared.")
        else:
            speak("There are no notes to clear, sir.")

    elif intent == "timer":
        raw = args[0]
        secs = _parse_duration(raw)
        if not secs:
            speak("I couldn't work out that duration, sir. Try something like: set a timer for 5 minutes.")
        else:
            mins, s = divmod(secs, 60)
            hrs,  m = divmod(mins, 60)
            label = f"{hrs}h {m}m {s}s".strip() if hrs else (f"{m}m {s}s".strip() if m else f"{s}s")
            speak(f"Timer set for {label}, sir.")
            def _timer_done(lbl):
                del _active_timers[lbl]
                speak(f"Sir, your {lbl} timer is up.")
            t_obj = threading.Timer(secs, _timer_done, args=[label])
            _active_timers[label] = t_obj
            t_obj.start()
            log.info(f"Timer set: {label} ({secs}s).")

    elif intent == "cancel_timer":
        if not _active_timers:
            speak("There are no active timers to cancel, sir.")
        else:
            label, t_obj = next(iter(_active_timers.items()))
            t_obj.cancel()
            del _active_timers[label]
            speak(f"Timer cancelled, sir.")
            log.info(f"Timer cancelled: {label}.")

    elif intent == "reminder":
        duration_text, message = args[0], args[1]
        secs = _parse_duration(duration_text)
        if not secs:
            speak("I couldn't parse that duration, sir.")
        else:
            mins, s = divmod(secs, 60)
            hrs,  m = divmod(mins, 60)
            label = f"{hrs}h {m}m {s}s".strip() if hrs else (f"{m}m {s}s".strip() if m else f"{s}s")
            speak(f"I'll remind you in {label}, sir.")
            def _remind_done(msg):
                speak(f"Sir, you asked me to remind you: {msg}")
            t_obj = threading.Timer(secs, _remind_done, args=[message])
            _active_timers[f"reminder:{label}"] = t_obj
            t_obj.start()
            log.info(f"Reminder set in {label}: {message}")

    elif intent == "window":
        action = args[0]
        WIN, DOWN, UP, F4, ALT = 0x5B, 0x28, 0x26, 0x73, 0x12
        KEYDOWN, KEYUP = 0, 0x0002
        if action == "minimize":
            speak("Minimizing, sir.")
            ctypes.windll.user32.keybd_event(WIN,  0, KEYDOWN, 0)
            ctypes.windll.user32.keybd_event(DOWN, 0, KEYDOWN, 0)
            ctypes.windll.user32.keybd_event(DOWN, 0, KEYUP,   0)
            ctypes.windll.user32.keybd_event(WIN,  0, KEYUP,   0)
        elif action == "maximize":
            speak("Maximizing, sir.")
            ctypes.windll.user32.keybd_event(WIN, 0, KEYDOWN, 0)
            ctypes.windll.user32.keybd_event(UP,  0, KEYDOWN, 0)
            ctypes.windll.user32.keybd_event(UP,  0, KEYUP,   0)
            ctypes.windll.user32.keybd_event(WIN, 0, KEYUP,   0)
        elif action == "close":
            speak("Closing the window, sir.")
            ctypes.windll.user32.keybd_event(ALT, 0, KEYDOWN, 0)
            ctypes.windll.user32.keybd_event(F4,  0, KEYDOWN, 0)
            ctypes.windll.user32.keybd_event(F4,  0, KEYUP,   0)
            ctypes.windll.user32.keybd_event(ALT, 0, KEYUP,   0)
        log.info(f"Window action: {action}.")

    elif intent == "recycle_bin":
        speak("Emptying the recycle bin, sir.")
        # SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND
        ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, 0x0007)
        speak("Done, sir.")
        log.info("Recycle bin emptied.")

    elif intent == "ip_address":
        mode = args[0]
        if mode == "local":
            try:
                local_ip = socket.gethostbyname(socket.gethostname())
                speak(f"Your local IP address is {local_ip}, sir.")
                log.info(f"Local IP: {local_ip}")
            except Exception as e:
                speak("I couldn't retrieve your local IP, sir.")
                log.error(f"Local IP error: {e}")
        else:
            try:
                r = requests.get("https://api.ipify.org", timeout=5)
                public_ip = r.text.strip()
                speak(f"Your public IP address is {public_ip}, sir.")
                log.info(f"Public IP: {public_ip}")
            except Exception as e:
                speak("I couldn't retrieve your public IP at the moment, sir.")
                log.error(f"Public IP error: {e}")

    elif intent == "internet_check":
        speak("Checking your connection, sir.")
        try:
            socket.create_connection(("8.8.8.8", 53), timeout=3)
            start = time.time()
            requests.get("https://www.google.com", timeout=5)
            ms = int((time.time() - start) * 1000)
            speak(f"You're online, sir. Response time is approximately {ms} milliseconds.")
            log.info(f"Internet check: online, {ms}ms.")
        except OSError:
            speak("You don't appear to be connected to the internet, sir.")
            log.warning("Internet check: offline.")

    elif intent == "battery":
        try:
            import psutil
            battery = psutil.sensors_battery()
            if battery is None:
                speak("I couldn't detect a battery on this system, sir.")
            else:
                status = "charging" if battery.power_plugged else "not charging"
                speak(f"Battery is at {battery.percent:.0f} percent and {status}, sir.")
                log.info(f"Battery: {battery.percent:.0f}%, {status}.")
        except ImportError:
            speak("Battery check requires psutil. Run pip install psutil, sir.")

    elif intent == "brightness":
        direction = args[0]
        val = "10" if direction == "up" else "-10"
        speak(f"Brightness {direction}, sir.")
        flags = _NO_WINDOW
        subprocess.run(["nircmd.exe", "changebrightness", val], creationflags=flags)
        log.info(f"Brightness {direction}.")

    elif intent == "theme":
        mode = args[0]
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                0, winreg.KEY_READ | winreg.KEY_WRITE
            )
            current, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            if mode == "dark":
                new_val = 0
            elif mode == "light":
                new_val = 1
            else:  # toggle
                new_val = 0 if current == 1 else 1
            winreg.SetValueEx(key, "AppsUseLightTheme",    0, winreg.REG_DWORD, new_val)
            winreg.SetValueEx(key, "SystemUsesLightTheme", 0, winreg.REG_DWORD, new_val)
            label = "dark" if new_val == 0 else "light"
            speak(f"Switched to {label} mode, sir.")
            log.info(f"Theme set to {label} mode.")
        except Exception as e:
            speak("I couldn't change the theme, sir.")
            log.error(f"Theme error: {e}")

    elif intent == "open_settings":
        label, uri = args[0], args[1]
        speak(f"Opening {label}, sir.")
        os.startfile(uri)
        log.info(f"Opened settings: {uri}")

    elif intent == "random":
        kind = args[0]
        if kind == "coin":
            result = random.choice(["Heads", "Tails"])
            speak(f"{result}, sir.")
            log.info(f"Coin flip: {result}.")
        elif kind == "dice":
            result = random.randint(1, 6)
            speak(f"You rolled a {result}, sir.")
            log.info(f"Dice roll: {result}.")
        elif kind == "number":
            lo, hi = args[1], args[2]
            result = random.randint(lo, hi)
            speak(f"Your random number is {result}, sir.")
            log.info(f"Random number ({lo}-{hi}): {result}.")

    elif intent == "f1_news":
        speak(random.choice(["Checking the paddock feeds, sir.", "Pulling the latest from the pit lane, sir.", "Scanning the grid, sir."]))
        headlines = _fetch_f1_headlines()
        if not headlines:
            speak("I'm unable to reach the Formula 1 feeds at the moment, sir.")
        else:
            titles = [title for title, _ in headlines]
            prompt = (
                "Here are the latest Formula 1 news headlines:\n"
                + "\n".join(f"- {t}" for t in titles)
                + "\n\nBriefly summarise what is happening in Formula 1 based on these. "
                  "Speak naturally in 3 to 4 sentences as JARVIS would. Keep it punchy and on-topic."
            )
            speak(ask_ai(prompt))
            links = [link for _, link in headlines[:2] if link.startswith("http")]
            for i, link in enumerate(links):
                open_in_brave(link, new_tab=i > 0)
            log.info(f"F1 news: {len(headlines)} headlines, {len(links)} articles opened.")

    elif intent == "world_news":
        speak(random.choice(["Scanning the feeds, sir.", "Pulling the latest headlines, sir.", "Checking global activity, sir."]))
        headlines = _fetch_headlines()
        if not headlines:
            speak("I'm unable to reach the news feeds at the moment, sir.")
        else:
            titles = [title for title, _ in headlines]
            prompt = (
                "Here are today's top news headlines:\n"
                + "\n".join(f"- {t}" for t in titles)
                + "\n\nBriefly summarise what is going on in the world based on these. "
                  "Speak naturally in 3 to 4 sentences as JARVIS would."
            )
            speak(ask_ai(prompt))
            links = [link for _, link in headlines[:2] if link.startswith("http")]
            for i, link in enumerate(links):
                open_in_brave(link, new_tab=i > 0)
            log.info(f"World news: {len(headlines)} headlines, {len(links)} articles opened.")

    elif intent == "now_playing":
        track = _get_spotify_now_playing()
        if track:
            speak(f"Currently playing on Spotify: {track}, sir.")
            log.info(f"Now playing (Spotify): {track}")
        else:
            browser = _get_music_from_browser()
            if browser:
                source, name = browser
                speak(f"Currently on {source}: {name}, sir.")
                log.info(f"Now playing ({source}): {name}")
            else:
                speak("I couldn't detect anything playing, sir.")

    elif intent == "identify_music":
        speak("Sampling your system audio, sir. One moment.")
        log.info("Music recognition: capturing system audio...")
        result = _recognize_audio(duration=6)
        if result and result.get("title"):
            title  = result["title"]
            artist = result["artist"]
            album  = result.get("album", "")
            speak(f"That's {title} by {artist}{', from the album ' + album if album else ''}, sir.")
            if result.get("spotify"):
                open_in_brave(result["spotify"])
                log.info(f"Opened Spotify: {result['spotify']}")
            else:
                query = f"{title} {artist}".replace(" ", "+")
                open_in_brave(f"https://www.youtube.com/results?search_query={query}")
            log.info(f"Identified: {title} – {artist}")
        else:
            browser = _get_music_from_browser()
            if browser:
                source, name = browser
                speak(f"Audio recognition came up empty, but {source} is showing: {name}, sir.")
            else:
                speak("I wasn't able to identify that one, sir. Make sure something is audible and try again.")

    elif intent == "anime_ost":
        import wave

        # ── 0. Dependency check — fail loudly before making promises ──
        import importlib.util
        missing = []
        if importlib.util.find_spec("pyaudiowpatch") is None:
            missing.append("pyaudiowpatch  →  pip install pyaudiowpatch")
        if importlib.util.find_spec("demucs") is None:
            missing.append("demucs  →  pip install demucs torch torchaudio")
        elif importlib.util.find_spec("torch") is None:
            missing.append("torch  →  pip install torch torchaudio")
        if missing:
            speak("I'm missing required packages, sir: " + ", and ".join(missing))
            log.warning(f"Anime OST: missing deps: {missing}")
            return

        # ── 1. Capture system audio (10 s — speech plays during recording) ──
        speak("Capturing your system audio now, sir. Recording ten seconds.")
        log.info("Anime OST: capturing system audio...")
        captured = _record_system_audio(duration=10)
        if not captured:
            speak("I couldn't capture the system audio, sir. "
                  "Check that pyaudiowpatch is installed and that audio is playing.")
            return
        frames, channels, rate, sample_width = captured

        # Quick energy check — make sure something is actually audible
        import numpy as np
        audio_np = np.frombuffer(b"".join(frames), dtype=np.int16)
        energy = float(np.abs(audio_np).mean())
        log.info(f"Audio energy: {energy:.1f}")
        if energy < 50:
            speak("The audio appears to be silent, sir. Make sure the volume is up and try again.")
            return

        # Write captured audio to a temp WAV
        with tempfile.NamedTemporaryFile(suffix="_raw.wav", delete=False) as f:
            raw_path = f.name
        with wave.open(raw_path, "wb") as wf:
            wf.setnchannels(channels)
            wf.setsampwidth(sample_width)
            wf.setframerate(rate)
            wf.writeframes(b"".join(frames))
        log.info(f"Raw audio saved: {raw_path}")

        # ── 2. Demucs separation (20-60 s on CPU — speech plays during it) ──
        speak("Removing character voices now, sir. Thirty seconds or so.")
        log.info("Anime OST: running Demucs separation...")
        bg_path = _isolate_background_music(raw_path)
        try: os.unlink(raw_path)
        except OSError: pass

        if not bg_path:
            speak("Demucs separation failed, sir. "
                  "Make sure torch and demucs are properly installed.")
            return

        # ── 3. AudD fingerprint (speech plays during the request) ──────────
        speak("Checking the database, sir.")
        log.info("Anime OST: sending isolated background to AudD...")
        try:
            with open(bg_path, "rb") as audio_file:
                r = requests.post(
                    "https://api.audd.io/",
                    data={"return": "spotify,apple_music"},
                    files={"file": ("background.wav", audio_file, "audio/wav")},
                    timeout=20,
                )
            try: os.unlink(bg_path)
            except OSError: pass

            ost = None
            if r.status_code == 200:
                data = r.json()
                log.info(f"AudD: status={data.get('status')} result={bool(data.get('result'))}")
                if data.get("status") == "success" and data.get("result"):
                    res = data["result"]
                    spotify_url = (res.get("spotify") or {}).get("external_urls", {}).get("spotify", "")
                    ost = {
                        "title":   res.get("title", ""),
                        "artist":  res.get("artist", ""),
                        "album":   res.get("album", ""),
                        "spotify": spotify_url,
                    }
            else:
                log.warning(f"AudD HTTP {r.status_code}: {r.text[:200]}")

            if ost and ost.get("title"):
                title  = ost["title"]
                artist = ost["artist"]
                album  = ost.get("album", "")
                speak(f"Found it, sir. That's {title} by {artist}"
                      f"{', from ' + album if album else ''}.")
                if ost.get("spotify"):
                    open_in_brave(ost["spotify"])
                else:
                    query = f"{title} {artist}".replace(" ", "+")
                    open_in_brave(f"https://www.youtube.com/results?search_query={query}")
                log.info(f"Anime OST identified: {title} – {artist}")
            else:
                speak("I couldn't match that one in the database, sir. "
                      "The OST may not be catalogued yet, or try during a louder section.")
                log.info(f"Anime OST: no match. Raw AudD: {r.text[:300]}")

        except Exception as e:
            speak("I had trouble reaching the AudD database, sir.")
            log.error(f"AudD request error: {e}")
            try: os.unlink(bg_path)
            except OSError: pass

    elif intent == "list_timers":
        if not _active_timers:
            speak("You have no active timers, sir.")
        else:
            labels = list(_active_timers.keys())
            if len(labels) == 1:
                speak(f"You have one active timer: {labels[0]}, sir.")
            else:
                speak(f"You have {len(labels)} active timers: {', '.join(labels)}, sir.")
        log.info(f"Active timers: {list(_active_timers.keys())}")

    elif intent == "stopwatch":
        global _stopwatch_start
        action = args[0]
        if action == "start":
            _stopwatch_start = time.time()
            speak("Stopwatch started, sir.")
            log.info("Stopwatch started.")
        elif action == "stop":
            if _stopwatch_start is None:
                speak("The stopwatch isn't running, sir.")
            else:
                elapsed = time.time() - _stopwatch_start
                _stopwatch_start = None
                m, s = divmod(int(elapsed), 60)
                h, m = divmod(m, 60)
                if h:
                    speak(f"Stopwatch stopped. Elapsed time: {h} hours, {m} minutes and {s} seconds, sir.")
                elif m:
                    speak(f"Stopwatch stopped. Elapsed time: {m} minutes and {s} seconds, sir.")
                else:
                    speak(f"Stopwatch stopped. Elapsed time: {s} seconds, sir.")
                log.info(f"Stopwatch stopped: {elapsed:.1f}s")
        elif action == "check":
            if _stopwatch_start is None:
                speak("The stopwatch isn't running, sir.")
            else:
                elapsed = time.time() - _stopwatch_start
                m, s = divmod(int(elapsed), 60)
                h, m = divmod(m, 60)
                if h:
                    speak(f"Elapsed time: {h} hours, {m} minutes and {s} seconds, sir.")
                elif m:
                    speak(f"Elapsed time: {m} minutes and {s} seconds, sir.")
                else:
                    speak(f"Elapsed time: {s} seconds, sir.")
        elif action == "reset":
            _stopwatch_start = None
            speak("Stopwatch reset, sir.")
            log.info("Stopwatch reset.")

    elif intent == "joke":
        speak("Let me find one for you, sir.")
        joke = _fetch_joke()
        if joke:
            setup, punchline = joke
            speak(setup)
            time.sleep(1.5)
            speak(punchline)
        else:
            speak(ask_ai("Tell me a short clever joke in character as JARVIS. One setup line, one punchline. No markdown."))
        log.info("Joke delivered.")

    elif intent == "repeat_last":
        if _last_spoken:
            speak(_last_spoken)
        else:
            speak("I don't have anything to repeat, sir.")

    elif intent == "time_in":
        city = args[0]
        result = _get_time_in_city(city)
        if result:
            speak(f"It is {result} in {city.title()}, sir.")
        else:
            speak(f"I'm afraid I don't have timezone data for {city}, sir.")
        log.info(f"Time in {city}: {result}")

    elif intent == "unit_convert":
        value, from_unit, to_unit = args[0], args[1], args[2]
        result, category = _convert_units(value, from_unit, to_unit)
        if result is None:
            speak(f"I can't convert {from_unit} to {to_unit}, sir. Those units may be incompatible.")
        else:
            formatted = f"{result:.6f}".rstrip("0").rstrip(".")
            speak(f"{value} {from_unit} is {formatted} {to_unit}, sir.")
        log.info(f"Unit convert: {value} {from_unit} → {to_unit} = {result}")

    elif intent == "define":
        word = args[0]
        speak(f"Looking up {word}, sir.")
        try:
            r = requests.get(f"https://api.dictionaryapi.dev/api/v2/entries/en/{word}", timeout=5)
            if r.status_code == 200:
                data = r.json()
                meanings = data[0].get("meanings", [])
                if meanings:
                    part = meanings[0].get("partOfSpeech", "")
                    defn = meanings[0].get("definitions", [{}])[0].get("definition", "")
                    if defn:
                        speak(f"{word.capitalize()}, {part}: {defn}")
                    else:
                        speak(f"I found {word} but couldn't extract a clean definition, sir.")
                else:
                    speak(f"No definition found for {word}, sir.")
            else:
                speak(f"I couldn't find a definition for {word} in my dictionary, sir.")
        except Exception as e:
            speak("I had trouble reaching the dictionary, sir.")
            log.error(f"Dictionary lookup error: {e}")
        log.info(f"Define: {word}")

    elif intent == "wikipedia":
        query = args[0]
        speak("Looking that up, sir.")
        try:
            search_term = query.replace(" ", "_")
            r = requests.get(
                f"https://en.wikipedia.org/api/rest_v1/page/summary/{search_term}",
                timeout=5, headers={"User-Agent": "JARVIS/2.0"}
            )
            if r.status_code == 200:
                data = r.json()
                extract = data.get("extract", "")
                if extract:
                    sentences = re.split(r'(?<=[.!?])\s+', extract)
                    summary = " ".join(sentences[:2])
                    speak(summary)
                    page_url = data.get("content_urls", {}).get("desktop", {}).get("page", "")
                    if page_url:
                        open_in_brave(page_url)
                else:
                    speak(f"I found a page for {query} but couldn't extract a summary, sir.")
            else:
                speak(f"I couldn't find anything on Wikipedia for {query}, sir.")
        except Exception as e:
            speak("I had trouble reaching Wikipedia, sir.")
            log.error(f"Wikipedia error: {e}")
        log.info(f"Wikipedia lookup: {query}")

    elif intent == "ai":
        speak(random.choice(["On it, sir.", "Let me think on that, sir.", "One moment, sir.", "Calculating, sir."]))
        reply = ask_ai(args[0])
        speak(reply)

# ─────────────────────────────────────────
#  AI FALLBACK — Ollama first, Claude if Ollama is unavailable
# ─────────────────────────────────────────
_SYSTEM_PROMPT = (
    "You are JARVIS — Just A Rather Very Intelligent System — the AI assistant "
    "built by Tony Stark. You are sophisticated, loyal, and have a dry British wit. "
    "Always address the user as 'sir'. Reply in 1-3 short spoken sentences. "
    "No markdown, no bullet points, no lists. This is a voice interface — be concise and natural."
)

def _ollama_is_running():
    """Return True if the Ollama server is up."""
    try:
        r = requests.get("http://localhost:11434/", timeout=3)
        return r.status_code == 200
    except Exception:
        return False

def _ollama_model_available(model):
    """Return True if the model is already pulled locally."""
    try:
        r = requests.get("http://localhost:11434/api/tags", timeout=5)
        if r.status_code == 200:
            names = [m.get("name", "").split(":")[0] for m in r.json().get("models", [])]
            return model in names or model.split(":")[0] in names
    except Exception:
        pass
    return False

def _ensure_ollama():
    """Start Ollama in the background if it isn't already running."""
    if _ollama_is_running():
        return True
    log.warning("Ollama not running — attempting to start it...")
    try:
        flags = _NO_WINDOW
        subprocess.Popen(["ollama", "serve"], creationflags=flags,
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        # Wait up to 10 seconds for it to come up
        for _ in range(10):
            time.sleep(1)
            if _ollama_is_running():
                log.info("  ✓ Ollama started successfully.")
                return True
        log.error("  ✗ Ollama did not start in time.")
    except FileNotFoundError:
        log.error("  ✗ Ollama not found — is it installed? https://ollama.com")
    except Exception as e:
        log.error(f"  ✗ Failed to start Ollama: {e}")
    return False

def _ensure_ollama_model():
    """Pull the model if it isn't downloaded yet."""
    if _ollama_model_available(OLLAMA_MODEL):
        return True
    log.warning(f"Model '{OLLAMA_MODEL}' not found — pulling now (this may take a while)...")
    speak(f"Downloading the {OLLAMA_MODEL} model for the first time, sir. This may take a few minutes.")
    try:
        flags = _NO_WINDOW
        result = subprocess.run(["ollama", "pull", OLLAMA_MODEL],
                                creationflags=flags, timeout=600)
        if result.returncode == 0:
            log.info(f"  ✓ Model '{OLLAMA_MODEL}' pulled successfully.")
            speak("Model downloaded. AI core is ready, sir.")
            return True
        else:
            log.error(f"  ✗ Failed to pull model '{OLLAMA_MODEL}'.")
    except subprocess.TimeoutExpired:
        log.error("  ✗ Model pull timed out.")
    except FileNotFoundError:
        log.error("  ✗ Ollama CLI not found.")
    except Exception as e:
        log.error(f"  ✗ Model pull error: {e}")
    return False

def _ask_ollama(prompt):
    """Returns a response string, or None if Ollama is unreachable/broken."""
    # Make sure Ollama is running before we try
    if not _ensure_ollama():
        return None
    _ensure_ollama_model()
    try:
        data = {"model": OLLAMA_MODEL, "prompt": prompt, "stream": False,
                "system": _SYSTEM_PROMPT}
        r = requests.post(OLLAMA_URL, json=data, timeout=30)
        if r.status_code == 200:
            reply = r.json().get("response", "").strip()
            if reply:
                return reply
            log.warning("Ollama returned an empty response.")
        else:
            log.error(f"Ollama HTTP {r.status_code}: {r.text[:200]}")
    except requests.exceptions.ConnectionError:
        log.error("Ollama connection refused — server may have crashed.")
    except requests.exceptions.Timeout:
        log.error("Ollama request timed out.")
    except Exception as e:
        log.error(f"Ollama error: {e}")
    return None

def ask_ai(prompt):
    """Send prompt to Ollama."""
    reply = _ask_ollama(prompt)
    if reply:
        log.info("AI answered via Ollama.")
        return reply
    log.info("Ollama unavailable.")
    return "I'm afraid my AI core is offline at the moment, sir. Please ensure Ollama is running."

# ─────────────────────────────────────────
#  SYSTEM TRAY ICON
# ─────────────────────────────────────────
def make_icon():
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d   = ImageDraw.Draw(img)
    d.ellipse([0, 0, 63, 63], fill=(0, 20, 40))
    d.text((20, 14), "J", fill=(0, 212, 255))
    return img

def build_tray(on_ready=None):
    """Build the system tray icon.  *on_ready* is called once the tray
    is fully visible — use it to start the voice listener thread."""
    def on_quit(icon, _):
        log.info("👋 Shutting down.")
        icon.stop()
        os._exit(0)

    def on_restart_jarvis(icon, _):
        def _do():
            speak("Restarting JARVIS. Back in a moment, sir.")
            log.info("Self-restart initiated from tray.")
            icon.stop()
            flags = _NO_WINDOW
            subprocess.Popen([sys.executable, os.path.abspath(__file__)], creationflags=flags)
            time.sleep(2)
            os._exit(0)
        threading.Thread(target=_do, daemon=True).start()

    def on_status(icon, _):
        threading.Thread(target=speak, args=("All systems are fully operational, sir.",), daemon=True).start()

    def on_open_log(icon, _):
        try:
            os.startfile(_LOG_FILE)
        except Exception as e:
            log.error(f"Could not open log: {e}")

    def on_open_notes(icon, _):
        if not os.path.exists(NOTES_FILE):
            open(NOTES_FILE, "w", encoding="utf-8").close()
        try:
            os.startfile(NOTES_FILE)
        except Exception as e:
            log.error(f"Could not open notes: {e}")

    def on_open_downloads(icon, _):
        path = FOLDER_MAP.get("downloads")
        if path:
            os.startfile(path)

    def on_open_desktop(icon, _):
        path = FOLDER_MAP.get("desktop")
        if path:
            os.startfile(path)

    def on_screenshot(icon, _):
        threading.Thread(target=process_command, args=("take a screenshot",), daemon=True).start()

    def on_system_status(icon, _):
        threading.Thread(target=process_command, args=("system status",), daemon=True).start()

    def on_lock(icon, _):
        threading.Thread(target=process_command, args=("lock screen",), daemon=True).start()

    def on_restart_ollama(icon, _):
        def _do():
            speak("Restarting Ollama, sir. One moment.")
            try:
                flags = _NO_WINDOW
                subprocess.run(["taskkill", "/f", "/im", "ollama.exe"],
                               capture_output=True, creationflags=flags)
                time.sleep(1)
            except Exception:
                pass
            if _ensure_ollama():
                speak("Ollama is back online, sir.")
            else:
                speak("I was unable to restart Ollama, sir. Please check your installation.")
        threading.Thread(target=_do, daemon=True).start()

    def on_help(icon, _):
        help_text = (
            "JARVIS Voice Commands:\n\n"
            "APPS: 'open [app]', 'close [app]'\n"
            "SITES: 'open youtube', 'go to github'\n"
            "SEARCH: 'search for [query]'\n"
            "MEDIA: 'play', 'pause', 'next song', 'previous song'\n"
            "         'what's playing', 'now playing'\n"
            "         'what song is this', 'shazam this' (system audio)\n"
            "         'what's the OST', 'anime OST' (vocal removal + fingerprint)\n"
            "VOLUME: 'volume up/down', 'mute', 'set volume to 50'\n"
            "SYSTEM: 'system status', 'lock screen', 'screenshot'\n"
            "         'shutdown', 'restart', 'sleep'\n"
            "FOLDERS: 'open downloads/desktop/documents/pictures'\n"
            "NOTES: 'note that [text]', 'read my notes'\n"
            "WEATHER: 'what's the weather'\n"
            "MATH: 'what is 5 + 3 * 2'\n"
            "CONVERT: 'convert 5 miles to kilometers'\n"
            "         'convert 100 fahrenheit to celsius'\n"
            "DEFINE: 'define [word]', 'what does [word] mean'\n"
            "LOOKUP: 'who is [person]', 'tell me about [topic]'\n"
            "TIMERS: 'set a timer for 5 minutes', 'cancel timer'\n"
            "         'list timers', 'my timers'\n"
            "STOPWATCH: 'start stopwatch', 'stop stopwatch'\n"
            "            'check stopwatch', 'reset stopwatch'\n"
            "TIME: 'what time is it in Tokyo'\n"
            "CLIPBOARD: 'what's in my clipboard'\n"
            "MODES: 'work mode', 'gaming mode', 'anime mode', etc.\n"
            "NEWS: 'world news', 'f1 news', 'formula 1 update'\n"
            "FUN: 'tell me a joke', 'flip a coin', 'roll a dice'\n"
            "      'say that again', 'repeat that'\n"
            "AI: anything else → Ollama (llama3)"
        )
        try:
            ctypes.windll.user32.MessageBoxW(0, help_text, "JARVIS — Voice Commands", 0x40)
        except Exception:
            log.info(help_text)

    def _setup(icon):
        icon.visible = True
        log.info("System tray icon is now visible.")
        if on_ready:
            on_ready()

    menu = pystray.Menu(
        pystray.MenuItem("JARVIS — LINKS Mark II", None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Status",           on_status),
        pystray.MenuItem("Screenshot",       on_screenshot),
        pystray.MenuItem("System Info",      on_system_status),
        pystray.MenuItem("Lock Screen",      on_lock),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Open Downloads",   on_open_downloads),
        pystray.MenuItem("Open Desktop",     on_open_desktop),
        pystray.MenuItem("Open Notes",       on_open_notes),
        pystray.MenuItem("View Log",         on_open_log),
        pystray.MenuItem("Restart Ollama",   on_restart_ollama),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Voice Commands (Help)", on_help),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Restart JARVIS",   on_restart_jarvis),
        pystray.MenuItem("Quit JARVIS",      on_quit),
    )
    tray = pystray.Icon("JARVIS", make_icon(), "JARVIS — Listening", menu)
    tray._jarvis_setup = _setup
    return tray

# ─────────────────────────────────────────
#  VOICE LISTENER
# ─────────────────────────────────────────
def listen_loop():
    log.info("Voice listener thread started.")
    recognizer = sr.Recognizer()
    recognizer.dynamic_energy_threshold = True
    recognizer.pause_threshold          = 0.8
    recognizer.energy_threshold         = 300

    # Log available mics and which one we're using
    mic_list = sr.Microphone.list_microphone_names()
    for i, name in enumerate(mic_list):
        log.info(f"  🎤 [{i}] {name}")

    try:
        mic = sr.Microphone()
        default_idx = mic.device_index if mic.device_index is not None else sr.Microphone.list_microphone_names().index(mic_list[0]) if mic_list else 0
        log.info(f"  ➜ Using: [{default_idx}] {mic_list[default_idx] if default_idx < len(mic_list) else 'unknown'}")
    except Exception as e:
        log.critical(f"No microphone available: {e}")
        return

    with mic as source:
        log.info("🎙  Calibrating mic...")
        recognizer.adjust_for_ambient_noise(source, duration=2)
        log.info(f"✅  Listening for '{WAKE_WORD.upper()}'")
        speak("JARVIS online. Good to be of service sir.")

        while True:
            try:
                audio = recognizer.listen(source, timeout=None, phrase_time_limit=10)
                text  = recognizer.recognize_google(audio).lower()
                log.info(f"🎤 {text}")

                if WAKE_WORD in text:
                    command = text.replace(WAKE_WORD, "", 1).strip(" ,.")
                    if not command:
                        speak(random.choice(["Yes, sir?", "Standing by, sir."]))
                        try:
                            audio2  = recognizer.listen(source, timeout=8, phrase_time_limit=12)
                            command = recognizer.recognize_google(audio2).lower()
                        except (sr.WaitTimeoutError, sr.UnknownValueError):
                            speak(random.choice(["My apologies, sir.", "No worries, sir."]))
                            continue
                    log.info(f"⚡ Command: {command}")
                    threading.Thread(target=process_command, args=(command,), daemon=True).start()

            except sr.WaitTimeoutError:
                pass
            except sr.UnknownValueError:
                pass
            except sr.RequestError as e:
                log.error(f"Speech API error: {e}")
                time.sleep(2)
            except KeyboardInterrupt:
                sys.exit(0)
            except Exception as e:
                log.error(f"Listen loop error: {e}", exc_info=True)
                time.sleep(1)

# ─────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────
if __name__ == "__main__":
    log.info("═" * 50)
    log.info("JARVIS — LINKS Mark II starting up")
    log.info("═" * 50)
    _check_ollama_startup()

    def _start_listener():
        threading.Thread(target=listen_loop, daemon=True).start()

    try:
        tray = build_tray(on_ready=_start_listener)
        log.info("Starting system tray (main thread)...")
        tray.run(setup=tray._jarvis_setup)   # blocks — tray must run on main thread
    except Exception as e:
        log.critical(f"Fatal error: {e}", exc_info=True)
        sys.exit(1)
