#!/bin/bash
# Скрипт для запуска веб-приложения проверки КС-2

echo "Запуск веб-приложения проверки КС-2..."
echo ""

# Проверяем наличие Python
if ! command -v python3 &> /dev/null; then
    echo "Ошибка: Python 3 не найден!"
    echo "Установите Python 3.7 или выше"
    exit 1
fi

# Проверяем наличие зависимостей
echo "Проверка зависимостей..."
python3 -c "import flask; import pandas; import openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Установка зависимостей..."
    pip3 install -r requirements.txt
fi

# Запускаем приложение
echo "Запуск веб-сервера..."
echo "Браузер откроется автоматически..."
echo ""
python3 ЗАПУСТИТЬ.py

