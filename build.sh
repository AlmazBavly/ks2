#!/bin/bash
# Скрипт для сборки приложения в исполняемый файл (macOS/Linux)
# Команда: ./build.sh

echo "========================================"
echo "Сборка приложения KS2Validator"
echo "Создание одного исполняемого файла"
echo "========================================"
echo ""

# Проверяем Python
if ! command -v python3 &> /dev/null; then
    echo "Ошибка: Python 3 не найден!"
    exit 1
fi

# Устанавливаем PyInstaller если нужно
echo "Проверка PyInstaller..."
python3 -c "import PyInstaller" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Установка PyInstaller..."
    pip3 install pyinstaller
fi

# Запускаем сборку
echo ""
echo "Начало сборки..."
echo ""
python3 build_exe.py

echo ""
echo "========================================"
echo "Сборка завершена!"
echo "Исполняемый файл: dist/KS2Validator"
echo "========================================"

