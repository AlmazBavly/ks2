#!/bin/bash
# Скрипт для создания/обновления виртуального окружения и зависимостей (macOS/Linux)
# Файл: ОБНОВИТЬ_ЗАВИСИМОСТИ.sh
# Запуск: ./ОБНОВИТЬ_ЗАВИСИМОСТИ.sh

# Переходим в директорию скрипта
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR" || exit 1

# Цвета для вывода
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

clear
echo "========================================"
echo "  Обновление зависимостей проекта"
echo "  Создание изолированного окружения"
echo "========================================"
echo ""

# Проверяем Python
echo -e "${BLUE}Проверка Python 3...${NC}"
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}❌ Ошибка: Python 3 не найден!${NC}"
    echo "Установите Python 3 с https://www.python.org/downloads/"
    exit 1
fi

PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
echo -e "${GREEN}✅ Python найден: $PYTHON_VERSION${NC}"
echo ""

# Создаем виртуальное окружение если его нет
VENV_DIR="venv"
if [ ! -d "$VENV_DIR" ]; then
    echo -e "${BLUE}Создание виртуального окружения...${NC}"
    if ! python3 -m venv "$VENV_DIR"; then
        echo -e "${RED}❌ Ошибка при создании виртуального окружения${NC}"
        exit 1
    fi
    echo -e "${GREEN}✅ Виртуальное окружение создано в папке: $VENV_DIR${NC}"
else
    echo -e "${GREEN}✅ Виртуальное окружение уже существует: $VENV_DIR${NC}"
fi
echo ""

# Активируем виртуальное окружение
echo -e "${BLUE}Активация виртуального окружения...${NC}"
source "$VENV_DIR/bin/activate" || exit 1
echo -e "${GREEN}✅ Виртуальное окружение активировано${NC}"
echo ""

# Обновляем pip
echo -e "${BLUE}Обновление pip...${NC}"
python -m pip install --upgrade pip --quiet
echo -e "${GREEN}✅ pip обновлен${NC}"
echo ""

# Устанавливаем/обновляем зависимости
echo -e "${BLUE}Установка/обновление зависимостей из requirements.txt...${NC}"
if [ -f "requirements.txt" ]; then
    if pip install --upgrade -r requirements.txt; then
        echo -e "${GREEN}✅ Все зависимости установлены/обновлены${NC}"
    else
        echo -e "${RED}❌ Ошибка при установке зависимостей${NC}"
        deactivate
        exit 1
    fi
else
    echo -e "${RED}❌ Файл requirements.txt не найден!${NC}"
    deactivate
    exit 1
fi
echo ""

# Показываем установленные пакеты
echo -e "${BLUE}Установленные пакеты:${NC}"
pip list | grep -E "pandas|openpyxl|flask|werkzeug|pyinstaller|pytest" || echo "Пакеты не найдены"
echo ""

# Создаем файл с зафиксированными версиями
echo -e "${BLUE}Создание requirements-lock.txt с точными версиями...${NC}"
pip freeze > requirements-lock.txt 2>/dev/null
echo -e "${GREEN}✅ Версии зависимостей сохранены в requirements-lock.txt${NC}"
echo ""

# Деактивируем виртуальное окружение
deactivate

echo "========================================"
echo -e "${GREEN}✅ Готово!${NC}"
echo "========================================"
echo ""
echo "Виртуальное окружение находится в папке: $VENV_DIR"
echo "Все зависимости изолированы и не конфликтуют с системными"
echo ""
echo "Для использования виртуального окружения:"
echo "  source $VENV_DIR/bin/activate"
echo ""
echo "Для деактивации:"
echo "  deactivate"
echo ""

