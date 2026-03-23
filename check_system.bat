@echo off
echo ===============================================
echo    Cdimex Mail Merge - Kiểm tra hệ thống
echo ===============================================
echo.

echo [INFO] Kiểm tra Python...
python --version
if %errorlevel% neq 0 (
    echo [ERROR] Python chưa được cài đặt!
    echo Vui lòng cài đặt Python từ: https://python.org
    pause
    exit /b 1
)
echo [OK] Python đã sẵn sàng
echo.

echo [INFO] Kiểm tra virtual environment...
if exist ".venv" (
    echo [OK] Virtual environment đã tồn tại
    call .venv\Scripts\activate.bat
    echo [INFO] Kiểm tra các thư viện...
    python -c "import pandas, openpyxl, streamlit, requests; print('[OK] Tất cả thư viện đã sẵn sàng')"
    if %errorlevel% neq 0 (
        echo [WARNING] Một số thư viện chưa được cài đặt
        echo Chạy install.bat để cài đặt lại
    )
) else (
    echo [WARNING] Virtual environment chưa được tạo
    echo Chạy install.bat để cài đặt
)
echo.

echo [INFO] Kiểm tra file cần thiết...
if exist "streamlit_app.py" (
    echo [OK] streamlit_app.py
) else (
    echo [ERROR] Không tìm thấy streamlit_app.py
)

if exist "recipients.xlsx" (
    echo [OK] recipients.xlsx
) else (
    echo [WARNING] Không tìm thấy recipients.xlsx (có thể tạo mới)
)

if exist "template.html" (
    echo [OK] template.html
) else (
    echo [WARNING] Không tìm thấy template.html (có thể tạo mới)
)

echo.
echo ===============================================
echo    Kiểm tra hoàn tất
echo ===============================================
echo.
pause


