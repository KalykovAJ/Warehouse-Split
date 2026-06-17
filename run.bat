@echo off
:: Включаем поддержку кириллицы (UTF-8) в консоли
chcp 65001 >nul

if "%1" neq "maximized" (
    start "" /max "%~f0" maximized
    exit /b
)

color 0B
title Разделитель черновиков по складам

:: Переходим в папку со скриптом
cd /d "C:\Users\Пользователь\Desktop\Python Scripts\Warehouse Split"

:: Запускаем скрипт
.venv\Scripts\python.exe split.py

:: Оставляем окно открытым
echo.
echo Нажмите любую клавишу для выхода...
pause >nul