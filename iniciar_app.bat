@echo off
REM Doble clic para abrir la aplicación (requiere Python instalado y "pip install -r requirements.txt")
cd /d "%~dp0"
pythonw app.py || python app.py
