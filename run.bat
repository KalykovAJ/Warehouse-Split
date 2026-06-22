@echo off
chcp 1251 >nul

if "%1" neq "maximized" (
    start "" /max "%~f0" maximized
    exit /b
)

color 0B
title Разделитель черновиков по складам

cd /d "%~dp0"

.venv\Scripts\python.exe split.py

echo.
echo Нажмите любую клавишу для выхода...
pause >nul
exit 0