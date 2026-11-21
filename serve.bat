@echo off
rem Starts the local server, then opens the app in Edge. Needs Python 3 (python.org) - or use serve.ps1 if you do not have it.
cd /d "%~dp0"
where python >nul 2>nul && (start "" msedge http://localhost:8765 & python serve.py & goto :eof)
where py >nul 2>nul && (start "" msedge http://localhost:8765 & py serve.py & goto :eof)
echo Python was not found. Right-click serve.ps1 and choose "Run with PowerShell" instead.
pause
