@echo off
REM Скрипт для сборки приложения в .exe (Windows)
REM Файл: СБОРКА_WINDOWS.bat
REM Просто дважды кликните на этот файл для запуска сборки!

echo ========================================
echo Сборка приложения KS2Validator
echo Платформа: WINDOWS
echo Создание одного исполняемого .exe файла
echo ========================================
echo.

REM Проверяем Python
python --version >nul 2>&1
if errorlevel 1 (
    echo Ошибка: Python не найден!
    pause
    exit /b 1
)

REM Проверяем виртуальное окружение
set VENV_DIR=venv
if not exist "%VENV_DIR%" (
    echo ⚠️  Виртуальное окружение не найдено!
    echo Создание виртуального окружения...
    python -m venv %VENV_DIR%
    if errorlevel 1 (
        echo ❌ Ошибка при создании виртуального окружения
        echo.
        echo Попробуйте запустить сначала:
        echo   ОБНОВИТЬ_ЗАВИСИМОСТИ.bat
        pause
        exit /b 1
    )
    echo ✅ Виртуальное окружение создано
    echo.
    echo Установка зависимостей...
    call %VENV_DIR%\Scripts\activate.bat
    python -m pip install --upgrade pip --quiet
    pip install -r requirements.txt
    call %VENV_DIR%\Scripts\deactivate.bat
    echo ✅ Зависимости установлены
) else (
    echo ✅ Виртуальное окружение найдено: %VENV_DIR%
)
echo.

REM Активируем виртуальное окружение
echo Активация виртуального окружения...
call %VENV_DIR%\Scripts\activate.bat
if errorlevel 1 (
    echo ❌ Ошибка при активации виртуального окружения
    pause
    exit /b 1
)
echo ✅ Виртуальное окружение активировано
echo.

REM Проверяем PyInstaller
echo Проверка PyInstaller...
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo Установка PyInstaller...
    pip install pyinstaller
)

REM Запускаем сборку
echo.
echo Начало сборки...
echo.
python build_exe.py

echo.
echo ========================================
echo ✅ Сборка завершена!
echo ========================================
echo.
echo Готовый файл: dist\KS2Validator.exe
echo.
echo Этот файл можно:
echo   - Скопировать на USB-флешку
echo   - Запустить на любом Windows ПК
echo   - Использовать без установки Python
echo.
echo Размер файла: ~50-100 МБ (это нормально!)
echo.

REM Деактивируем виртуальное окружение
call %VENV_DIR%\Scripts\deactivate.bat

pause

