@echo off
chcp 65001 >nul
cd /d "%~dp0"

call myenv\Scripts\activate.bat
pip install -r requirements.txt
python -m PyInstaller --noconfirm --clean --windowed --name "ExcelCompareHighlighter" app.py
call myenv\Scripts\deactivate.bat

echo.
echo 打包完成：dist\ExcelCompareHighlighter\ExcelCompareHighlighter.exe
pause
