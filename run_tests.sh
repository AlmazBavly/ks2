#!/bin/bash
# Скрипт для запуска тестов

echo "Запуск юнит-тестов..."
echo ""

python3 -m unittest discover tests -v

