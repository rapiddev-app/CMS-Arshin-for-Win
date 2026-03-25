@echo off
chcp 65001 >nul

cd /d "%~dp0"

REM Используем встроенный Python если есть
if exist "python_embedded\python.exe" (
    python_embedded\python.exe installer_simple.py
) else (
    python installer_simple.py
)

pause
