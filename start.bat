@echo off
echo ========================================
echo   نظام التنبؤ المبكر بالتعثر الأكاديمي
echo ========================================
echo.

cd backend

echo [1/2] تثبيت المتطلبات...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
echo [2/2] تشغيل الخادم...
echo.
echo افتح المتصفح على: http://localhost:5000
echo.
python app.py

pause
