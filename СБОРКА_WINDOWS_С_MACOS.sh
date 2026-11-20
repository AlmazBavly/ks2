#!/bin/bash
# Скрипт для сборки Windows .exe на macOS
# Файл: СБОРКА_WINDOWS_С_MACOS.sh
# Запуск: ./СБОРКА_WINDOWS_С_MACOS.sh

# Переходим в директорию скрипта
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR" || exit 1

# Цвета для вывода
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

clear
echo "========================================"
echo "  Сборка Windows .exe на macOS"
echo "  KS2Validator для Windows"
echo "========================================"
echo ""

# ⚠️ Важное предупреждение
echo -e "${YELLOW}⚠️  ВАЖНО: PyInstaller не поддерживает кросс-компиляцию!${NC}"
echo ""
echo "PyInstaller создает исполняемые файлы только для той платформы,"
echo "на которой он запущен. Для сборки Windows .exe нужна Windows."
echo ""
echo -e "${CYAN}Доступные варианты:${NC}"
echo ""
echo "1. ${GREEN}GitHub Actions (РЕКОМЕНДУЕТСЯ)${NC} - автоматическая сборка на Windows"
echo "2. ${GREEN}Виртуальная машина Windows${NC} - используйте Parallels/VMware"
echo "3. ${YELLOW}Wine (экспериментально)${NC} - может не работать со всеми библиотеками"
echo ""
echo "========================================"
echo ""

# Проверяем наличие Wine
WINE_AVAILABLE=false
if command -v wine &> /dev/null; then
    WINE_AVAILABLE=true
    echo -e "${GREEN}✅ Wine найден${NC}"
else
    echo -e "${YELLOW}⚠️  Wine не найден${NC}"
fi

echo ""
echo -e "${CYAN}Выберите вариант:${NC}"
echo ""
echo "1) Создать GitHub Actions workflow для автоматической сборки"
echo "2) Показать инструкции по использованию виртуальной машины"
if [ "$WINE_AVAILABLE" = true ]; then
    echo "3) Попробовать собрать через Wine (экспериментально)"
fi
echo "4) Выход"
echo ""
read -p "Ваш выбор (1-4): " choice

case $choice in
    1)
        echo ""
        echo -e "${BLUE}Создание GitHub Actions workflow...${NC}"
        echo ""
        
        # Создаем директорию .github/workflows если её нет
        mkdir -p .github/workflows
        
        # Создаем workflow файл
        cat > .github/workflows/build-windows.yml << 'EOF'
name: Build Windows Executable

on:
  workflow_dispatch:  # Ручной запуск
  push:
    branches: [ main, master ]
    paths:
      - 'ks2_validator_web.py'
      - 'build_exe.py'
      - 'requirements.txt'
      - '.github/workflows/build-windows.yml'

jobs:
  build:
    runs-on: windows-latest
    
    steps:
    - uses: actions/checkout@v3
    
    - name: Set up Python
      uses: actions/setup-python@v4
      with:
        python-version: '3.9'
    
    - name: Install dependencies
      run: |
        python -m pip install --upgrade pip
        pip install -r requirements.txt
        pip install pyinstaller
    
    - name: Build executable
      run: python build_exe.py
    
    - name: Upload artifact
      uses: actions/upload-artifact@v3
      with:
        name: KS2Validator-Windows
        path: dist/KS2Validator.exe
        retention-days: 30
EOF
        
        echo -e "${GREEN}✅ GitHub Actions workflow создан!${NC}"
        echo ""
        echo "Файл: .github/workflows/build-windows.yml"
        echo ""
        echo -e "${CYAN}Как использовать:${NC}"
        echo ""
        echo "1. Загрузите проект на GitHub:"
        echo "   git add .github/workflows/build-windows.yml"
        echo "   git commit -m 'Add Windows build workflow'"
        echo "   git push"
        echo ""
        echo "2. Перейдите на GitHub → Actions → Build Windows Executable"
        echo ""
        echo "3. Нажмите 'Run workflow' → выберите ветку → Run"
        echo ""
        echo "4. После завершения скачайте .exe из Artifacts"
        echo ""
        echo -e "${GREEN}✅ Готово! Теперь вы можете собирать .exe автоматически на GitHub${NC}"
        ;;
        
    2)
        echo ""
        echo -e "${CYAN}Инструкции по использованию виртуальной машины:${NC}"
        echo ""
        echo "1. Установите виртуальную машину:"
        echo "   - Parallels Desktop (платно, но удобно)"
        echo "   - VMware Fusion (платно)"
        echo "   - VirtualBox (бесплатно)"
        echo ""
        echo "2. Установите Windows на виртуальную машину"
        echo ""
        echo "3. Скопируйте проект на виртуальную машину:"
        echo "   - Через общую папку"
        echo "   - Или через git clone"
        echo ""
        echo "4. В Windows выполните:"
        echo "   СБОРКА_WINDOWS.bat"
        echo ""
        echo "5. Скопируйте готовый .exe обратно на macOS"
        echo ""
        echo -e "${GREEN}✅ Это самый надежный способ!${NC}"
        ;;
        
    3)
        if [ "$WINE_AVAILABLE" = false ]; then
            echo -e "${RED}❌ Wine не установлен!${NC}"
            echo ""
            echo "Установите Wine:"
            echo "  brew install --cask wine-stable"
            echo ""
            echo "Или используйте другой вариант сборки."
            exit 1
        fi
        
        echo ""
        echo -e "${YELLOW}⚠️  ВНИМАНИЕ: Сборка через Wine экспериментальна!${NC}"
        echo ""
        echo "Wine может не работать со всеми библиотеками Python."
        echo "Особенно проблематичны: pandas, numpy, openpyxl"
        echo ""
        read -p "Продолжить? (y/n): " confirm
        if [ "$confirm" != "y" ] && [ "$confirm" != "Y" ]; then
            exit 0
        fi
        
        echo ""
        echo -e "${BLUE}Проверка Wine...${NC}"
        wine --version
        echo ""
        
        # Проверяем наличие Python для Windows в Wine
        echo -e "${BLUE}Проверка Python для Windows...${NC}"
        if ! wine python --version 2>/dev/null; then
            echo -e "${YELLOW}⚠️  Python для Windows не найден в Wine${NC}"
            echo ""
            echo "Вам нужно:"
            echo "1. Скачать Python для Windows: https://www.python.org/downloads/"
            echo "2. Установить его через Wine:"
            echo "   wine python-3.x.x.exe"
            echo ""
            echo "Или используйте другой вариант сборки."
            exit 1
        fi
        
        echo ""
        echo -e "${BLUE}Установка зависимостей в Wine...${NC}"
        wine python -m pip install --upgrade pip
        wine python -m pip install -r requirements.txt
        wine python -m pip install pyinstaller
        
        echo ""
        echo -e "${BLUE}Запуск сборки через Wine...${NC}"
        echo "Это может занять много времени и может не сработать..."
        echo ""
        
        # Модифицируем build_exe.py для работы через Wine
        # Но лучше просто запустить PyInstaller напрямую
        wine python -m PyInstaller \
            --onefile \
            --name=KS2Validator \
            --hidden-import=pandas \
            --hidden-import=openpyxl \
            --hidden-import=flask \
            --hidden-import=werkzeug \
            --hidden-import=webbrowser \
            --collect-all=pandas \
            --collect-all=openpyxl \
            --collect-all=flask \
            --add-data="templates;templates" \
            --clean \
            ks2_validator_web.py
        
        if [ $? -eq 0 ]; then
            echo ""
            echo -e "${GREEN}✅ Сборка завершена!${NC}"
            echo ""
            echo "Файл должен быть в: dist/KS2Validator.exe"
            echo ""
            echo -e "${YELLOW}⚠️  ВАЖНО: Проверьте файл на Windows!${NC}"
            echo "Wine может создать нерабочий .exe файл."
        else
            echo ""
            echo -e "${RED}❌ Сборка не удалась${NC}"
            echo ""
            echo "Wine часто не может правильно собрать Python приложения."
            echo "Рекомендуется использовать GitHub Actions или виртуальную машину."
        fi
        ;;
        
    4)
        echo "Выход..."
        exit 0
        ;;
        
    *)
        echo -e "${RED}Неверный выбор${NC}"
        exit 1
        ;;
esac

echo ""
echo "Нажмите Enter для выхода..."
read -r

