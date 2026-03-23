@echo off
echo ===============================================
echo    Cdimex Mail Merge - Cài đặt tự động
echo ===============================================
echo.

REM Kiểm tra Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python chưa được cài đặt!
    echo Vui lòng tải và cài đặt Python từ: https://python.org
    echo Sau đó chạy lại file này.
    pause
    exit /b 1
)

echo [OK] Python đã được cài đặt
echo.

REM Tạo virtual environment
echo [INFO] Tạo môi trường ảo...
python -m venv .venv
if %errorlevel% neq 0 (
    echo [ERROR] Không thể tạo môi trường ảo
    pause
    exit /b 1
)

REM Kích hoạt virtual environment và cài đặt packages
echo [INFO] Cài đặt các thư viện cần thiết...
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install pandas openpyxl streamlit requests streamlit-quill

if %errorlevel% neq 0 (
    echo [ERROR] Không thể cài đặt thư viện
    pause
    exit /b 1
)

echo.
echo ===============================================
echo    Cài đặt hoàn tất!
echo ===============================================
echo.
echo Để chạy ứng dụng:
echo 1. Double-click file "run.bat"
echo 2. Hoặc chạy lệnh: run.bat
echo.
pause


