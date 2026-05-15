@echo off
chcp 65001 >nul
cd /d "%~dp0"
call myenv\Scripts\activate.bat
python app.py
call myenv\Scripts\deactivate.bat
pause
