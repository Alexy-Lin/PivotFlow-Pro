@echo off
setlocal
chcp 65001 >nul

set "SCRIPT_DIR=%~dp0"
set "PYTHON_EXE=C:\anaconda3\python.exe"

if not exist "%PYTHON_EXE%" (
    echo 未找到 %PYTHON_EXE%
    echo 将尝试使用系统 PATH 中的 python。
    set "PYTHON_EXE=python"
)

if "%~1"=="" (
    "%PYTHON_EXE%" "%SCRIPT_DIR%create_pivot_gui.py"
) else (
    "%PYTHON_EXE%" "%SCRIPT_DIR%create_pivot_gui.py" "%~1"
)

if errorlevel 1 (
    echo.
    echo 程序执行失败，请查看上方错误信息。
    pause
)

endlocal
