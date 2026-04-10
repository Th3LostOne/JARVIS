@echo off
cd /d "%~dp0"

:: Launch JARVIS with Python 3.10 (windowless, no console)
powershell -WindowStyle Hidden -Command "Start-Process 'C:\Users\Morpheus\AppData\Local\Programs\Python\Python310\pythonw.exe' -ArgumentList '%~dp0jarvis.py' -WorkingDirectory '%~dp0'"

timeout /t 2 /nobreak >nul
exit
