@echo off
cd /d "%~dp0"
echo Iniciando Consulta Petlove em http://127.0.0.1:5000
start "" http://127.0.0.1:5000
"%~dp0.python\python.exe" app.py
pause
