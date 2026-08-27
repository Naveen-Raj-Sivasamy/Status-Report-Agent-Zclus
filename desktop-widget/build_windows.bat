@echo off
REM Run this ON a Windows machine (PyInstaller doesn't cross-compile).
REM Produces dist\StatusUpdate.exe — ship that plus a config.json next to it.

pip install -r requirements.txt pyinstaller
pyinstaller --onefile --noconsole --name StatusUpdate status_widget.py

echo.
echo Build done. Give teammates dist\StatusUpdate.exe + a config.json
echo (copied from config.example.json with your WEBHOOK_URL and TOKEN filled in)
echo placed in the SAME folder as the exe.
