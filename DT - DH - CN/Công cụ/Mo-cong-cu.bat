@echo off
chcp 65001 >nul
title Cong cu Doanh thu - Don hang - Cong no (DUNG DONG cua so nay khi dang dung)
rem Tim Python: lenh "python", neu khong co thi thu "py"
set "PY=python"
%PY% --version >nul 2>&1 || set "PY=py"
%PY% --version >nul 2>&1 || goto :nopython
%PY% "%~dp0app.py"
pause
exit /b

:nopython
echo.
echo  MAY NAY CHUA CAI PYTHON
echo   1. Tai Python tai: https://www.python.org/downloads/
echo   2. Khi cai, NHO TICK o "Add python.exe to PATH"
echo   3. Cai xong, chay lai Mo-cong-cu.bat
echo.
pause
