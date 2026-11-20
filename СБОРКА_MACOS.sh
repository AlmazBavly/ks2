#!/bin/bash
# Скрипт для сборки приложения в исполняемый файл (macOS)
# Файл: СБОРКА_MACOS.sh
# Запуск во встроенном терминале: ./СБОРКА_MACOS.sh

# Переходим в директорию скрипта (для работы из любой директории)
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR" || exit 1

# Цвета для вывода (для лучшей читаемости в терминале)
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

clear
echo "========================================"
echo "  Сборка приложения KS2Validator"
echo "  Платформа: macOS"
echo "  Создание одного исполняемого файла"
echo "========================================"
echo ""

# Проверяем Python
echo -e "${BLUE}Проверка Python 3...${NC}"
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}❌ Ошибка: Python 3 не найден!${NC}"
    echo "Установите Python 3 с https://www.python.org/downloads/"
    echo "Или через Homebrew: brew install python3"
    exit 1
fi

PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
echo -e "${GREEN}✅ Python найден: $PYTHON_VERSION${NC}"
echo ""

# Проверяем виртуальное окружение
VENV_DIR="venv"
if [ ! -d "$VENV_DIR" ]; then
    echo -e "${YELLOW}⚠️  Виртуальное окружение не найдено!${NC}"
    echo "Создание виртуального окружения..."
    if ! python3 -m venv "$VENV_DIR"; then
        echo -e "${RED}❌ Ошибка при создании виртуального окружения${NC}"
        echo ""
        echo "Попробуйте запустить сначала:"
        echo "  ./ОБНОВИТЬ_ЗАВИСИМОСТИ.sh"
        exit 1
    fi
    echo -e "${GREEN}✅ Виртуальное окружение создано${NC}"
    echo ""
    echo -e "${BLUE}Установка зависимостей...${NC}"
    source "$VENV_DIR/bin/activate"
    pip install --upgrade pip --quiet
    pip install -r requirements.txt
    deactivate
    echo -e "${GREEN}✅ Зависимости установлены${NC}"
else
    echo -e "${GREEN}✅ Виртуальное окружение найдено: $VENV_DIR${NC}"
fi
echo ""

# Активируем виртуальное окружение
echo -e "${BLUE}Активация виртуального окружения...${NC}"
source "$VENV_DIR/bin/activate" || exit 1
echo -e "${GREEN}✅ Виртуальное окружение активировано${NC}"
echo ""

# Проверяем зависимости в venv
echo -e "${BLUE}Проверка зависимостей в виртуальном окружении...${NC}"
MISSING_DEPS=()
for dep in pandas openpyxl flask werkzeug; do
    if ! python -c "import $dep" 2>/dev/null; then
        MISSING_DEPS+=("$dep")
    fi
done

if [ ${#MISSING_DEPS[@]} -gt 0 ]; then
    echo -e "${YELLOW}⚠️  Обнаружены отсутствующие зависимости: ${MISSING_DEPS[*]}${NC}"
    echo "Установка зависимостей..."
    pip install --upgrade pip --quiet
    if ! pip install -r requirements.txt; then
        echo -e "${RED}❌ Ошибка при установке зависимостей${NC}"
        deactivate
        exit 1
    fi
    echo -e "${GREEN}✅ Зависимости установлены${NC}"
else
    echo -e "${GREEN}✅ Все зависимости установлены${NC}"
fi
echo ""

# Проверяем PyInstaller
echo -e "${BLUE}Проверка PyInstaller...${NC}"
if ! python -c "import PyInstaller" 2>/dev/null; then
    echo -e "${YELLOW}Установка PyInstaller...${NC}"
    if ! pip install pyinstaller; then
        echo -e "${RED}❌ Ошибка при установке PyInstaller${NC}"
        deactivate
        exit 1
    fi
    echo -e "${GREEN}✅ PyInstaller установлен${NC}"
else
    echo -e "${GREEN}✅ PyInstaller установлен${NC}"
fi
echo ""

# Запускаем сборку
echo "========================================"
echo -e "${BLUE}Начало сборки...${NC}"
echo "========================================"
echo ""
echo "Это может занять 2-5 минут..."
echo ""

if python build_exe.py; then
    echo ""
    echo "========================================"
    echo -e "${GREEN}✅ Сборка завершена успешно!${NC}"
    echo "========================================"
    echo ""
    
    # Проверяем созданные файлы
    if [ -d "dist" ]; then
        echo -e "${GREEN}Готовые файлы в папке dist/:${NC}"
        ls -lh dist/ | grep -E "KS2Validator" || echo "Файлы не найдены"
        echo ""
        
        if [ -f "dist/KS2Validator" ]; then
            FILE_SIZE=$(du -h "dist/KS2Validator" | cut -f1)
            echo -e "${BLUE}📦 Исполняемый файл: dist/KS2Validator${NC}"
            echo -e "${BLUE}📊 Размер: $FILE_SIZE${NC}"
        fi
        
        if [ -d "dist/KS2Validator.app" ]; then
            APP_SIZE=$(du -sh "dist/KS2Validator.app" | cut -f1)
            echo -e "${BLUE}📦 Приложение: dist/KS2Validator.app${NC}"
            echo -e "${BLUE}📊 Размер: $APP_SIZE${NC}"
        fi
        
        echo ""
        echo -e "${GREEN}Эти файлы можно:${NC}"
        echo "  ✅ Скопировать на USB-флешку"
        echo "  ✅ Запустить на любом macOS компьютере"
        echo "  ✅ Использовать без установки Python"
        echo ""
    else
        echo -e "${RED}⚠️  Папка dist/ не найдена${NC}"
    fi
else
    echo ""
    echo "========================================"
    echo -e "${RED}❌ Ошибка при сборке!${NC}"
    echo "========================================"
    echo ""
    echo "Проверьте сообщения об ошибках выше"
    deactivate
    exit 1
fi

# Деактивируем виртуальное окружение
deactivate

echo ""
echo "Нажмите Enter для выхода..."
read -r

