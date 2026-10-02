@echo off
REM ============================================================
REM  بناء نسخة EXE 32-bit — يجب تشغيل هذا الملف على Windows
REM  المتطلبات: Python 3.x بنسخة 32-bit (من python.org) + pip
REM ============================================================

REM 1) تحقق أن بايثون المثبت هو 32-bit فعليًا
python -c "import struct,sys; print('bits:', struct.calcsize('P')*8); sys.exit(0 if struct.calcsize('P')==4 else 1)" || (echo خطأ: شغّل هذا الملف بـ Python 32-bit & exit /b 1)

REM 2) تثبيت الاعتماديات (نسخ 32-bit تلقائيًا عبر pip الحالي)
python -m pip install --upgrade pip
python -m pip install openpyxl xlrd pillow pyinstaller

REM 3) بناء ملف تنفيذي واحد 32-bit بدون نافذة كونسول
pyinstaller --onefile --noconsole --name ExamMonitoring2026 ^
  --add-data "data;data" --add-data "assets;assets" ^
  main.py

echo.
echo تم البناء: dist\ExamMonitoring2026.exe  (نسخة 32-bit حسب بايثون المستخدم)
pause
