# J.A.R.V.I.S — LINKS Mark II
I AM VIBE CODING ALL OF THIS SO I AM TAKING NO CREDIT
A lightweight Windows voice assistant that runs silently in the system tray. Wake it with **"Jarvis"**, give a command, and it handles the rest — opening apps, controlling media, taking screenshots, running system commands, and falling back to a local AI (Ollama) for anything else.

---

## Setup

**1. Install Python dependencies**

Run `setup.bat` or manually:
```
pip install speechrecognition pyttsx3 requests pyaudio pystray pillow edge-tts pygame psutil
```

> If PyAudio fails: `pip install pipwin` then `pipwin install pyaudio`

**2. Install Ollama** (for AI fallback)

Download from [ollama.com](https://ollama.com), then:
```
ollama pull llama3
```
JARVIS will auto-start Ollama and auto-pull the model on first use if needed.

**3. Start JARVIS**

Double-click `START_JARVIS.bat`. A **J** icon appears in the system tray — JARVIS is now listening.

---

## Usage

Say **"Jarvis"** followed by a command. Saying "Jarvis" alone prompts for a follow-up within 8 seconds.

---

## Commands

### Apps
```
open spotify / brave / discord / vs code / steam / vlc / obs
open notepad / calculator / explorer / task manager / paint
open word / excel / powerpoint / league of legends / osu / antigravity

close spotify / brave / discord  (etc.)
```

### Websites
All open in Brave Browser.
```
open youtube / gmail / google / github / reddit / twitter
open netflix / twitch / spotify web / chatgpt / claude / amazon / wikipedia
open brightspace carleton / brightspace algonquin
open anime / nokia / nokia services
```
Or say any domain directly: *"go to twitch.tv"*, *"open github.com"*

### Modes — launch multiple apps/sites at once
```
work mode       → Discord + Brave + Spotify
gaming mode     → Discord + Spotify + Steam + League
research mode   → Wikipedia + Google Scholar + Google
exam mode       → YouTube + Brightspace (both) + Nokia guides
movie mode      → xprime.stream + aether.mom
anime mode      → AniWatch + Discord
```
> "mode" is flexible — *"gaming session"*, *"exam time"*, *"work setup"* all work.

### Web Search
```
search for [query]
look up [query]
google [query]
```

### Media — always targets Spotify (launches it if not running)
```
play / pause
next song / skip
previous song
stop music
```

### Volume
```
volume up / volume down
mute
set volume to 50
set volume to 75 percent
```

### Audio Output
Requires [SoundSwitch](https://soundswitch.aaflalo.me/) running in the tray.
```
switch audio / switch output / cycle audio
switch to headphones / switch to speakers
```

### Timers & Reminders
```
set a timer for 5 minutes
set a timer for 1 hour 30 minutes
set a 30 second timer
cancel timer

remind me in 10 minutes to check the oven
remind me in 2 hours to submit the assignment
```

### Screenshots
```
take a screenshot                         → Desktop
screenshot named report                   → named file
screenshot in downloads                   → specific folder
screenshot named budget in documents      → named + folder
screenshot in a new folder called charts  → creates folder on Desktop
```
Folders searched: Downloads, Desktop, Documents, Pictures, Videos, Music, home, OneDrive.

### Folders
```
open downloads / desktop / documents / pictures / music / videos / appdata / jarvis
```

### Notes
```
note that [text]
remember that [text]
make a note [text]
read my notes          → speaks your last 3 notes
clear my notes
```
Saved to `jarvis_notes.txt` in the JARVIS folder.

### System
```
what time is it
what day is it
system status          → CPU % and RAM %
battery status
check my internet      → connection test with latency
what's my ip           → local IP
what's my public ip    → external IP
what's the weather     → opens wttr.in
lock screen
sleep / hibernate
shutdown               → 10 sec delay
restart / reboot       → 10 sec delay
```

### Display & Theme
```
dark mode / light mode / toggle dark mode
brightness up / brightness down
```

### Windows Settings
```
open settings / sound settings / display settings / bluetooth settings
open wifi settings / network settings / power settings / update settings
open privacy settings / storage settings / app settings
```

### Window Management
```
minimize window
maximize window
close window
```

### Quick Math
```
what is 5 + 3
what is 100 / 4 * 2
```

### Fun
```
flip a coin
roll a die
random number
random number between 1 and 50
```

### Clipboard
```
what's in my clipboard
read clipboard
```

### Misc
```
empty recycle bin
restart jarvis
```

### World News Briefing
```
what's going on in the world
what's happening in the world
world news / latest news / current events
news update / catch me up on the news
```
Fetches live headlines from BBC, Reuters, Al Jazeera, The Guardian, and NPR. Ollama narrates a spoken summary, then the top 2 articles open in Brave.

### AI Fallback
Anything not matched above is sent to **Ollama (llama3)** running locally. JARVIS responds in character as the AI assistant from Iron Man.

---

## System Tray

Right-click the **J** icon in the taskbar:

| Item | Action |
|---|---|
| Status | Speaks "all systems operational" |
| Screenshot | Takes a screenshot to Desktop |
| System Info | Speaks CPU % and RAM % |
| Lock Screen | Locks the workstation |
| Open Downloads / Desktop | Opens folder |
| Open Notes | Opens `jarvis_notes.txt` |
| View Log | Opens `jarvis.log` |
| Restart Ollama | Kills and restarts Ollama |
| Voice Commands (Help) | Shows command list popup |
| Restart JARVIS | Reloads the assistant |
| Quit JARVIS | Exits |

---

## Requirements

- Windows 10/11
- Python 3.10+
- Microphone
- [Ollama](https://ollama.com) — for AI fallback
- [SoundSwitch](https://soundswitch.aaflalo.me/) — optional, for audio output switching
- [nircmd](https://www.nirsoft.net/utils/nircmd.html) — optional, for volume/brightness control
