@echo off
chcp 65001 >nul
pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --onefile --windowed --name "TeacherFileManager" --add-data "template.docx;." main.py
echo.
echo الملف جاهز في: dist\TeacherFileManager.exe
pause
