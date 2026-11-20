#!/bin/bash
# Файл для запуска приложения двойным кликом на macOS

# Переходим в директорию скрипта
cd "$(dirname "$0")"

# Открываем окно терминала и запускаем приложение через универсальный скрипт
osascript <<EOF
tell application "Terminal"
    activate
    do script "cd '$(pwd)' && python3 ЗАПУСТИТЬ.py"
end tell
EOF

