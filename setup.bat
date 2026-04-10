@echo off
title JARVIS Setup — LINKS Mark II

echo.
echo  ============================================================
echo   JARVIS -- LINKS MARK II  //  SETUP
echo  ============================================================
echo.

set PYTHON=C:\Users\Morpheus\AppData\Local\Programs\Python\Python310\python.exe
set PIP=%PYTHON% -m pip

"%PYTHON%" --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Python 3.10 not found at expected path.
    echo     Expected: %PYTHON%
    echo     Download from https://www.python.org
    pause & exit /b 1
)
echo [+] Python 3.10 found.

echo [+] Installing core packages...
"%PIP%" install speechrecognition pyttsx3 requests pystray pillow --quiet

echo [+] Installing audio packages...
"%PIP%" install pyaudiowpatch pygame edge-tts --quiet

echo [+] Installing AI / media packages...
"%PIP%" install wikipedia-api pytz wolframalpha --quiet

echo [+] Installing Demucs + torch (may take a few minutes)...
"%PIP%" install torch torchaudio --index-url https://download.pytorch.org/whl/cpu --quiet
"%PIP%" install demucs soundfile --quiet
"%PIP%" install "numpy<2.0" --quiet

if %errorlevel% neq 0 (
    echo.
    echo [!] Some packages may have failed. Check output above.
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
