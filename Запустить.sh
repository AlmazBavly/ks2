#!/bin/bash
# Файл для запуска приложения двойным кликом на Linux

cd "$(dirname "$0")"

# Проверяем наличие Python
if ! command -v python3 &> /dev/null; then
    zenity --error --text="Ошибка: Python 3 не найден!\n\nУстановите Python 3.7 или выше" 2>/dev/null || \
    kdialog --error "Ошибка: Python 3 не найден!\n\nУстановите Python 3.7 или выше" 2>/dev/null || \
    xmessage "Ошибка: Python 3 не найден! Установите Python 3.7 или выше"
    exit 1
fi

# Проверяем зависимости
python3 -c "import flask; import pandas; import openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    zenity --info --text="Установка зависимостей..." 2>/dev/null || \
    kdialog --msgbox "Установка зависимостей..." 2>/dev/null
    pip3 install -r requirements.txt
fi

# Запускаем приложение через универсальный скрипт
python3 ЗАПУСТИТЬ.py

