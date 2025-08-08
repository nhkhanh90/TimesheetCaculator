@echo off
echo ==================================================
echo    CAI DAT PYINSTALLER VA BUILD FILE .EXE
echo ==================================================
echo.

echo [1/2] Cai dat PyInstaller...
pip install pyinstaller
if %errorlevel% neq 0 (
    echo FAILED: Khong the cai dat PyInstaller
    pause
    exit /b 1
)
echo     --> Da cai dat PyInstaller thanh cong!
echo.

echo [2/2] Tao file .exe...
pyinstaller --onefile --windowed --name="TimesheetCalculator" desktop_app.py
if %errorlevel% neq 0 (
    echo FAILED: Khong the tao file .exe
    pause
    exit /b 1
)
echo     --> Da tao file .exe thanh cong!
echo.

echo Sao chep file .exe...
if exist "dist\TimesheetCalculator.exe" (
    copy "dist\TimesheetCalculator.exe" "TimesheetCalculator.exe"
    echo     --> File .exe: TimesheetCalculator.exe
) else (
    echo FAILED: Khong tim thay file .exe
)
echo.

echo ==================================================
echo    BUILD HOAN TAT!
echo ==================================================
echo.
echo File executable da duoc tao: TimesheetCalculator.exe
echo Ban co the chia se file nay cho nguoi khac su dung.
echo.
echo Nhan Enter de dong cua so...
pause > nul
