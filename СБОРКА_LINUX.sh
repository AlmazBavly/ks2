#!/bin/bash
# Скрипт для сборки приложения в исполняемый файл (Linux)
# Файл: СБОРКА_LINUX.sh
# Запуск: ./СБОРКА_LINUX.sh

# Переходим в директорию скрипта
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR" || exit 1

echo "========================================"
echo "Сборка приложения KS2Validator"
echo "Платформа: Linux"
echo "Создание одного исполняемого файла"
echo "========================================"
echo ""

# Проверяем Python
if ! command -v python3 &> /dev/null; then
    echo "Ошибка: Python 3 не найден!"
    exit 1
fi

# Проверяем виртуальное окружение
VENV_DIR="venv"
if [ ! -d "$VENV_DIR" ]; then
    echo "⚠️  Виртуальное окружение не найдено!"
    echo "Создание виртуального окружения..."
    if ! python3 -m venv "$VENV_DIR"; then
        echo "❌ Ошибка при создании виртуального окружения"
        echo ""
        echo "Попробуйте запустить сначала:"
        echo "  ./ОБНОВИТЬ_ЗАВИСИМОСТИ.sh"
        exit 1
    fi
    echo "✅ Виртуальное окружение создано"
    echo ""
    echo "Установка зависимостей..."
    source "$VENV_DIR/bin/activate"
    pip install --upgrade pip --quiet
    pip install -r requirements.txt
    deactivate
    echo "✅ Зависимости установлены"
else
    echo "✅ Виртуальное окружение найдено: $VENV_DIR"
fi
echo ""

# Активируем виртуальное окружение
echo "Активация виртуального окружения..."
source "$VENV_DIR/bin/activate" || exit 1
echo "✅ Виртуальное окружение активировано"
echo ""

# Проверяем PyInstaller
echo "Проверка PyInstaller..."
if ! python -c "import PyInstaller" 2>/dev/null; then
    echo "Установка PyInstaller..."
    pip install pyinstaller
fi
echo ""

# Запускаем сборку
echo ""
echo "Начало сборки..."
echo ""
python build_exe.py

echo ""
echo "========================================"
echo "✅ Сборка завершена!"
echo "========================================"
echo ""
echo "Готовый файл: dist/KS2Validator"
echo ""
echo "Этот файл можно:"
echo "  - Скопировать на USB-флешку"
echo "  - Запустить на любом Linux компьютере"
echo "  - Использовать без установки Python"
echo ""
echo "Размер файла: ~50-100 МБ (это нормально!)"
echo ""

# Деактивируем виртуальное окружение
deactivate

