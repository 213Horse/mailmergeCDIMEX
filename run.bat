@echo off
echo ===============================================
echo    Cdimex Mail Merge - Khởi động ứng dụng
echo ===============================================
echo.

REM Kiểm tra virtual environment
if not exist ".venv" (
    echo [ERROR] Môi trường ảo chưa được tạo!
    echo Vui lòng chạy file "install.bat" trước.
    pause
    exit /b 1
)

REM Kích hoạt virtual environment và chạy ứng dụng
echo [INFO] Khởi động ứng dụng...
call .venv\Scripts\activate.bat
streamlit run streamlit_app.py

echo.
echo [INFO] Ứng dụng đã dừng.
pause


