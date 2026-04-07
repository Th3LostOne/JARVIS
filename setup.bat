@echo off
title JARVIS Setup — LINKS Mark II

echo.
echo  ============================================================
echo   JARVIS -- LINKS MARK II  //  SETUP
echo  ============================================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Python not found. Download from https://www.python.org
    pause & exit /b 1
)
echo [+] Python found.

echo [+] Installing packages...
pip install speechrecognition pyttsx3 requests pyaudio pystray pillow edge-tts pygame --quiet

if %errorlevel% neq 0 (
    echo.
    echo [!] PyAudio failed? Try:
    echo     pip install pipwin
    echo     pipwin install pyaudio
    echo.
)

echo.
echo [+] Done.
echo.
echo  ============================================================
echo   NEXT STEPS:
echo  ============================================================
echo.
echo  1. Install Ollama:   https://ollama.com
echo  2. Pull a model:     ollama pull llama3
echo  3. Start Ollama:     ollama serve
echo  4. Run JARVIS:       double-click START_JARVIS.bat
echo.
echo  JARVIS will run silently in the background.
echo  Look for the J icon in your system tray (bottom-right).
echo  Right-click the tray icon to check status or quit.
echo.
pause
