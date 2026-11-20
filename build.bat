@echo off
REM Скрипт для сборки приложения в .exe (Windows)
REM Команда: build.bat

echo ========================================
echo Сборка приложения KS2Validator
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

REM Устанавливаем PyInstaller если нужно
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
pause

