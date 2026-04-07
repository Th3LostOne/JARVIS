@echo off
title JARVIS — LINKS Mark II
cd /d "%~dp0"

echo  [JARVIS] Starting...
echo  [JARVIS] Watch for the J icon in your system tray.
echo  [JARVIS] Logs written to jarvis.log
echo.


:: Run with python (not pythonw) but hide the console window via VBScript
:: This keeps pystray working while hiding the window completely

:: Create a temporary VBScript to launch hidden (explicitly passes env vars)
(
    echo Set sh = CreateObject^("WScript.Shell"^)
    echo sh.Run "python jarvis.py", 0, False
) > "%~dp0_launch.vbs"
cscript //nologo "%~dp0_launch.vbs"
del "%~dp0_launch.vbs"

timeout /t 2 /nobreak >nul
exit
