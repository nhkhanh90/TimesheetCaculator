@echo off
echo ==================================================
echo    HE THONG TINH GIO CONG - BUILD SCRIPT
echo ==================================================
echo.

echo [1/3] Cai dat cac thu vien can thiet...
pip install -r requirements_desktop.txt
if %errorlevel% neq 0 (
    echo FAILED: Khong the cai dat thu vien
    pause
    exit /b 1
)
echo     --> Thanh cong!
echo.

echo [2/3] Tao file .exe bang PyInstaller...
pyinstaller --onefile --windowed --name="TimesheetCalculator" --icon=icon.ico desktop_app.py
if %errorlevel% neq 0 (
    echo FAILED: Khong the tao file .exe
    pause
    exit /b 1
)
echo     --> Thanh cong!
echo.

echo [3/3] Sao chep file .exe den thu muc hien tai...
if exist "dist\TimesheetCalculator.exe" (
    copy "dist\TimesheetCalculator.exe" "TimesheetCalculator.exe"
    echo     --> File .exe da duoc tao: TimesheetCalculator.exe
) else (
    echo FAILED: Khong tim thay file .exe
)
echo.

echo ==================================================
echo    BUILD HOAN TAT!
echo ==================================================
echo File executable: TimesheetCalculator.exe
echo Ban co the chia se file nay cho nguoi dung khac.
echo.
pause
