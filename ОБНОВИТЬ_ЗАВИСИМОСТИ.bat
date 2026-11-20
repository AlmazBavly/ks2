@echo off
REM Скрипт для создания/обновления виртуального окружения и зависимостей (Windows)
REM Файл: ОБНОВИТЬ_ЗАВИСИМОСТИ.bat
REM Запуск: ОБНОВИТЬ_ЗАВИСИМОСТИ.bat

echo ========================================
echo   Обновление зависимостей проекта
echo   Создание изолированного окружения
echo ========================================
echo.

REM Проверяем Python
echo Проверка Python 3...
python --version >nul 2>&1
if errorlevel 1 (
    echo Ошибка: Python 3 не найден!
    echo Установите Python 3 с https://www.python.org/downloads/
    pause
    exit /b 1
)

for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo Python найден: %PYTHON_VERSION%
echo.

REM Создаем виртуальное окружение если его нет
set VENV_DIR=venv
if not exist "%VENV_DIR%" (
    echo Создание виртуального окружения...
    python -m venv %VENV_DIR%
    if errorlevel 1 (
        echo Ошибка при создании виртуального окружения
        pause
        exit /b 1
    )
    echo Виртуальное окружение создано в папке: %VENV_DIR%
) else (
    echo Виртуальное окружение уже существует: %VENV_DIR%
)
echo.

REM Активируем виртуальное окружение
echo Активация виртуального окружения...
call %VENV_DIR%\Scripts\activate.bat
if errorlevel 1 (
    echo Ошибка при активации виртуального окружения
    pause
    exit /b 1
)
echo Виртуальное окружение активировано
echo.

REM Обновляем pip
echo Обновление pip...
python -m pip install --upgrade pip --quiet
echo pip обновлен
echo.

REM Устанавливаем/обновляем зависимости
echo Установка/обновление зависимостей из requirements.txt...
if exist "requirements.txt" (
    pip install --upgrade -r requirements.txt
    if errorlevel 1 (
        echo Ошибка при установке зависимостей
        call %VENV_DIR%\Scripts\deactivate.bat
        pause
        exit /b 1
    )
    echo Все зависимости установлены/обновлены
) else (
    echo Ошибка: файл requirements.txt не найден!
    call %VENV_DIR%\Scripts\deactivate.bat
    pause
    exit /b 1
)
echo.

REM Показываем установленные пакеты
echo Установленные пакеты:
pip list | findstr /i "pandas openpyxl flask werkzeug pyinstaller pytest"
echo.

REM Создаем файл с зафиксированными версиями
echo Создание requirements-lock.txt с точными версиями...
pip freeze > requirements-lock.txt 2>nul
echo Версии зависимостей сохранены в requirements-lock.txt
echo.

REM Деактивируем виртуальное окружение
call %VENV_DIR%\Scripts\deactivate.bat

echo ========================================
echo   Готово!
echo ========================================
echo.
echo Виртуальное окружение находится в папке: %VENV_DIR%
echo Все зависимости изолированы и не конфликтуют с системными
echo.
echo Для использования виртуального окружения:
echo   %VENV_DIR%\Scripts\activate.bat
echo.
pause

