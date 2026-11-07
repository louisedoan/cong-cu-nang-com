@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem Tim Python: lenh "python", neu khong co thi thu "py"
set "PY=python"
%PY% --version >nul 2>&1 || set "PY=py"
%PY% --version >nul 2>&1 || goto :nopython
%PY% "%~dp0phieu_xuat_kho.py"
pause
exit /b

:nopython
echo.
echo  ============================================================
echo   MAY NAY CHUA CAI PYTHON - can cai truoc khi lap phieu
echo  ============================================================
echo   1. Tai Python tai: https://www.python.org/downloads/
echo   2. Khi cai, NHO TICK o "Add python.exe to PATH"
echo   3. Cai xong, chay lai Tao-phieu-xuat-kho.bat
echo.
pause
