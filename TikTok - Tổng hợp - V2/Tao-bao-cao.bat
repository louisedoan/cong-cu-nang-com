@echo off
chcp 65001 >nul
set "PY=python"
%PY% --version >nul 2>&1 || set "PY=py"
%PY% --version >nul 2>&1 || goto :nopython
if "%~1"=="" (
    %PY% "%~dp0bao_cao_so_luong.py"
) else (
    %PY% "%~dp0bao_cao_so_luong.py" "%~1"
)
pause
exit /b

:nopython
echo Can cai Python va them vao PATH truoc khi chay cong cu.
pause
