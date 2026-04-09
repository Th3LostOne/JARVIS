"""
JARVIS Setup — LINKS Mark II
Standalone installer: detects Python, installs all packages,
optionally installs Ollama and nircmd.
"""
import os
import re
import sys
import shutil
import subprocess
import threading
import tempfile
import zipfile
import urllib.request
import tkinter as tk
from tkinter import ttk

# ── Resolve paths whether running as .py or bundled .exe ─────────────────────
if getattr(sys, "frozen", False):
    _INSTALLER_DIR = os.path.dirname(sys.executable)
    _ASSETS_DIR    = os.path.join(sys._MEIPASS, "assets")
else:
    _INSTALLER_DIR = os.path.dirname(os.path.abspath(__file__))
    _ASSETS_DIR    = os.path.join(_INSTALLER_DIR, "assets")

# ── Theme ─────────────────────────────────────────────────────────────────────
BG       = "#0d1b2a"
BG_PANEL = "#1a2b3c"
CYAN     = "#00d4ff"
WHITE    = "#e8f0fe"
GREY     = "#5a6a7a"
GREEN    = "#00c853"
RED      = "#ff4d4d"
YELLOW   = "#ffd600"

# ── Required pip packages ─────────────────────────────────────────────────────
PACKAGES = [
    ("speechrecognition", "SpeechRecognition  —  voice input"),
    ("pyaudio",           "PyAudio            —  microphone"),
    ("pyttsx3",           "pyttsx3            —  TTS fallback"),
    ("requests",          "Requests           —  HTTP"),
    ("pystray",           "pystray            —  system tray"),
    ("pillow",            "Pillow             —  images"),
    ("edge-tts",          "edge-tts           —  neural voice"),
    ("pygame",            "pygame             —  audio playback"),
    ("psutil",            "psutil             —  system stats"),
]


# ─────────────────────────────────────────────────────────────────────────────
#  Installer GUI
# ─────────────────────────────────────────────────────────────────────────────
class InstallerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("JARVIS Setup — LINKS Mark II")
        self.root.geometry("660x560")
        self.root.configure(bg=BG)
        self.root.resizable(False, False)
        self._set_icon()
        self._build_ui()

    def _set_icon(self):
        try:
            ico = os.path.join(_ASSETS_DIR, "jarvis.ico")
            if os.path.exists(ico):
                self.root.iconbitmap(ico)
        except Exception:
            pass

    # ── UI construction ───────────────────────────────────────────────────────
    def _build_ui(self):
        # Header
        hdr = tk.Frame(self.root, bg=BG, pady=18)
        hdr.pack(fill="x")
        tk.Label(hdr, text="J.A.R.V.I.S", font=("Courier New", 30, "bold"),
                 fg=CYAN, bg=BG).pack()
        tk.Label(hdr, text="LINKS Mark II  —  Setup", font=("Courier New", 11),
                 fg=GREY, bg=BG).pack()

        tk.Frame(self.root, bg=CYAN, height=1).pack(fill="x", padx=24)

        # Options
        opts = tk.Frame(self.root, bg=BG_PANEL, padx=20, pady=14)
        opts.pack(fill="x", padx=24, pady=(14, 0))
        tk.Label(opts, text="COMPONENTS TO INSTALL", font=("Courier New", 8, "bold"),
                 fg=GREY, bg=BG_PANEL).pack(anchor="w", pady=(0, 6))

        # Python packages — always on, greyed out
        pkg_var = tk.BooleanVar(value=True)
        cb_pkg = tk.Checkbutton(
            opts, text="Python packages  (required)",
            variable=pkg_var, state="disabled",
            fg=WHITE, bg=BG_PANEL, selectcolor=BG_PANEL,
            disabledforeground=WHITE, activebackground=BG_PANEL,
            font=("Courier New", 10),
        )
        cb_pkg.pack(anchor="w")

        self.var_ollama = tk.BooleanVar(value=True)
        tk.Checkbutton(
            opts, text="Ollama  (local AI — llama3 auto-pulled on first use)",
            variable=self.var_ollama,
            fg=WHITE, bg=BG_PANEL, selectcolor=BG,
            activebackground=BG_PANEL, activeforeground=WHITE,
            font=("Courier New", 10),
        ).pack(anchor="w")

        self.var_nircmd = tk.BooleanVar(value=True)
        tk.Checkbutton(
            opts, text="nircmd  (volume & brightness control — freeware)",
            variable=self.var_nircmd,
            fg=WHITE, bg=BG_PANEL, selectcolor=BG,
            activebackground=BG_PANEL, activeforeground=WHITE,
            font=("Courier New", 10),
        ).pack(anchor="w")

        # Log output
        log_frame = tk.Frame(self.root, bg=BG, padx=24, pady=10)
        log_frame.pack(fill="both", expand=True)

        self.log = tk.Text(
            log_frame, bg="#060d14", fg=WHITE,
            font=("Courier New", 9), state="disabled",
            relief="flat", padx=10, pady=8, height=13,
        )
        self.log.pack(fill="both", expand=True)
        self.log.tag_config("ok",   foreground=GREEN)
        self.log.tag_config("err",  foreground=RED)
        self.log.tag_config("info", foreground=CYAN)
        self.log.tag_config("warn", foreground=YELLOW)
        self.log.tag_config("dim",  foreground=GREY)

        # Progress bar
        style = ttk.Style()
        style.theme_use("default")
        style.configure("J.Horizontal.TProgressbar",
                        troughcolor=BG_PANEL, background=CYAN,
                        thickness=6, borderwidth=0)
        self.progress = ttk.Progressbar(
            self.root, style="J.Horizontal.TProgressbar",
            length=612, mode="determinate",
        )
        self.progress.pack(padx=24, pady=(0, 4))

        # Buttons
        btn_frame = tk.Frame(self.root, bg=BG, pady=12)
        btn_frame.pack(fill="x", padx=24)

        self.btn_close = tk.Button(
            btn_frame, text="CLOSE", command=self.root.destroy,
            bg=BG_PANEL, fg=GREY, font=("Courier New", 11),
            relief="flat", padx=20, pady=6, cursor="hand2",
            activebackground=BG_PANEL, activeforeground=WHITE,
        )
        self.btn_close.pack(side="right", padx=(8, 0))

        self.btn_install = tk.Button(
            btn_frame, text="INSTALL", command=self._start,
            bg=CYAN, fg=BG, font=("Courier New", 11, "bold"),
            relief="flat", padx=24, pady=6, cursor="hand2",
            activebackground="#00aad4", activeforeground=BG,
        )
        self.btn_install.pack(side="right")

    # ── Logging helpers ───────────────────────────────────────────────────────
    def _log(self, msg, tag=""):
        def _write():
            self.log.config(state="normal")
            self.log.insert("end", msg + "\n", tag)
            self.log.see("end")
            self.log.config(state="disabled")
        self.root.after(0, _write)

    def _set_progress(self, pct):
        self.root.after(0, lambda: self.progress.config(value=pct))

    # ── Entry point ───────────────────────────────────────────────────────────
    def _start(self):
        self.btn_install.config(state="disabled")
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        install_ollama = self.var_ollama.get()
        install_nircmd = self.var_nircmd.get()

        total = 2 + len(PACKAGES) \
              + (3 if install_ollama else 0) \
              + (2 if install_nircmd else 0)
        done = [0]

        def tick(n=1):
            done[0] += n
            self._set_progress(min(int(done[0] / total * 100), 99))

        # 1 ── Find Python ─────────────────────────────────────────────────────
        self._log("Locating Python 3.10+...", "info")
        python = _find_python()
        if not python:
            self._log("✗  Python 3.10+ not found.", "err")
            self._log("   Download from https://www.python.org", "warn")
            self._log("   Tick 'Add Python to PATH', then re-run Setup.", "warn")
            self.root.after(0, lambda: self.btn_install.config(state="normal"))
            return
        self._log(f"✓  {python}", "ok")
        tick()

        # 2 ── Upgrade pip ─────────────────────────────────────────────────────
        self._log("Upgrading pip...", "info")
        r = subprocess.run(
            [python, "-m", "pip", "install", "--upgrade", "pip", "-q"],
            capture_output=True, text=True,
        )
        self._log("✓  pip ready" if r.returncode == 0
                  else f"⚠  pip upgrade skipped ({r.stderr.strip()[:60]})", "ok" if r.returncode == 0 else "warn")
        tick()

        # 3 ── pip packages ────────────────────────────────────────────────────
        for pkg, label in PACKAGES:
            self._log(f"Installing {label}...", "dim")
            r = subprocess.run(
                [python, "-m", "pip", "install", pkg, "-q", "--prefer-binary"],
                capture_output=True, text=True,
            )
            if r.returncode == 0:
                self._log(f"✓  {label}", "ok")
            else:
                last = (r.stderr.strip().splitlines() or ["unknown"])[-1]
                self._log(f"✗  {label}  →  {last[:70]}", "err")
            tick()

        # 4 ── Ollama ──────────────────────────────────────────────────────────
        if install_ollama:
            _install_ollama(self._log, tick)

        # 5 ── nircmd ─────────────────────────────────────────────────────────
        if install_nircmd:
            _install_nircmd(self._log, tick, _INSTALLER_DIR)

        # Done ─────────────────────────────────────────────────────────────────
        self._set_progress(100)
        self._log("", "")
        self._log("══════════════════════════════════════════", "info")
        self._log("  All done.  Launch START_JARVIS.bat", "ok")
        self._log("══════════════════════════════════════════", "info")
        self.root.after(0, lambda: self.btn_close.config(fg=WHITE))

    def run(self):
        self.root.mainloop()


# ─────────────────────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────────────────────
def _find_python():
    """Return path to Python 3.10+ or None."""
    for cmd in ["python", "python3", "py"]:
        path = shutil.which(cmd)
        if not path:
            continue
        try:
            r = subprocess.run([path, "--version"], capture_output=True, text=True)
            ver = (r.stdout + r.stderr).strip()
            m = re.search(r"(\d+)\.(\d+)", ver)
            if m and (int(m.group(1)), int(m.group(2))) >= (3, 10):
                return path
        except Exception:
            pass
    return None


def _install_ollama(log, tick):
    log("Checking Ollama...", "info")
    if shutil.which("ollama"):
        log("✓  Ollama already installed", "ok")
        tick(3)
        return
    log("Downloading Ollama installer...", "dim")
    try:
        url = "https://ollama.com/download/OllamaSetup.exe"
        tmp = os.path.join(tempfile.gettempdir(), "OllamaSetup.exe")
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=60) as resp, open(tmp, "wb") as f:
            f.write(resp.read())
        log("✓  OllamaSetup.exe downloaded", "ok")
        tick()
        log("Running Ollama installer...", "dim")
        subprocess.run([tmp], check=False)
        log("✓  Ollama installed — llama3 will auto-pull on first JARVIS AI query", "ok")
        tick(2)
    except Exception as e:
        log(f"✗  Ollama download failed: {e}", "err")
        log("   Download manually: https://ollama.com/download", "warn")
        tick(3)


def _install_nircmd(log, tick, jarvis_dir):
    """
    nircmd is freeware by NirSoft, freely redistributable at no charge.
    License: https://www.nirsoft.net/utils/nircmd.html
    Downloaded directly from the official source.
    """
    dest = os.path.join(jarvis_dir, "nircmd.exe")
    if os.path.exists(dest):
        log("✓  nircmd already present", "ok")
        tick(2)
        return
    log("Downloading nircmd (NirSoft freeware)...", "dim")
    try:
        url = "https://www.nirsoft.net/utils/nircmd-x64.zip"
        tmp_zip = os.path.join(tempfile.gettempdir(), "nircmd.zip")
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as resp, open(tmp_zip, "wb") as f:
            f.write(resp.read())
        tick()
        with zipfile.ZipFile(tmp_zip, "r") as z:
            names = [n for n in z.namelist() if os.path.basename(n).lower() == "nircmd.exe"]
            if names:
                with z.open(names[0]) as src, open(dest, "wb") as dst:
                    dst.write(src.read())
                log(f"✓  nircmd installed → {dest}", "ok")
            else:
                log("✗  nircmd.exe not found in zip", "err")
        os.unlink(tmp_zip)
        tick()
    except Exception as e:
        log(f"✗  nircmd download failed: {e}", "err")
        log("   Download manually: https://www.nirsoft.net/utils/nircmd.html", "warn")
        tick(2)


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    InstallerApp().run()
