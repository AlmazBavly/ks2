@echo off
REM Скрипт для запуска веб-приложения проверки КС-2 (Windows)
REM Запускается двойным кликом

cd /d "%~dp0"

REM Проверяем Python
python --version >nul 2>&1
if errorlevel 1 (
    msg * "Ошибка: Python не найден!`n`nУстановите Python 3.7 или выше с сайта python.org"
    exit /b 1
)

REM Проверяем наличие зависимостей
python -c "import flask; import pandas; import openpyxl" 2>nul
if errorlevel 1 (
    echo Установка зависимостей...
    pip install -r requirements.txt
    if errorlevel 1 (
        msg * "Ошибка при установке зависимостей!"
        pause
        exit /b 1
    )
)

REM Запускаем приложение через универсальный скрипт
python ЗАПУСТИТЬ.py

REM Если приложение закрылось, показываем сообщение
if errorlevel 1 (
    echo.
    echo Приложение завершилось с ошибкой.
    pause
)

