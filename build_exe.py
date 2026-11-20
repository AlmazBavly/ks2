#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для сборки приложения в один исполняемый файл
"""
import subprocess
import sys
import os

def build_exe():
    """Собирает приложение в .exe файл"""
    print("Сборка исполняемого файла...")
    print("=" * 60)
    
    # Проверяем наличие PyInstaller
    try:
        import PyInstaller
    except ImportError:
        print("Установка PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # Команда для сборки
    cmd = [
        "pyinstaller",
        "--onefile",  # Один файл
        "--name=KS2Validator",  # Имя исполняемого файла
        "--icon=NONE",  # Без иконки
        "--hidden-import=pandas",
        "--hidden-import=openpyxl",
        "--hidden-import=flask",
        "--hidden-import=werkzeug",
        "--hidden-import=webbrowser",
        "--collect-all=pandas",
        "--collect-all=openpyxl",
        "--collect-all=flask",
        "--add-data=templates;templates",  # Включаем шаблоны
        "--clean",  # Очистить кэш перед сборкой
        "ks2_validator_web.py"
    ]
    
    # Для Windows используем --windowed
    # Для macOS НЕ используем --noconsole, чтобы видеть вывод и ошибки
    # Для Linux можно использовать --noconsole
    if sys.platform == "win32":
        cmd.insert(2, "--windowed")
    elif sys.platform.startswith("linux"):
        cmd.insert(2, "--noconsole")
    # Для macOS оставляем консоль, чтобы видеть вывод
    
    # Исправляем разделитель для --add-data в зависимости от платформы
    if sys.platform != "win32":
        cmd = [c.replace(";", ":") for c in cmd]
    
    print(f"Выполняется команда: {' '.join(cmd)}")
    print()
    
    try:
        subprocess.check_call(cmd)
        print()
        print("=" * 60)
        print("✅ Сборка завершена успешно!")
        print("=" * 60)
        print()
        
        if sys.platform == "win32":
            exe_path = os.path.join("dist", "KS2Validator.exe")
            print(f"📦 Готовый файл: {exe_path}")
            print()
            print("Этот файл можно:")
            print("  ✅ Скопировать на USB-флешку")
            print("  ✅ Запустить на любом Windows ПК")
            print("  ✅ Использовать без установки Python")
            print()
            if os.path.exists(exe_path):
                file_size = os.path.getsize(exe_path) / (1024 * 1024)  # Размер в МБ
                print(f"📊 Размер файла: {file_size:.1f} МБ (это нормально!)")
        else:
            exe_path = os.path.join("dist", "KS2Validator")
            print(f"📦 Готовый файл: {exe_path}")
            print()
            print("Этот файл можно:")
            print("  ✅ Скопировать на USB-флешку")
            print("  ✅ Запустить на любом компьютере с той же ОС")
            print("  ✅ Использовать без установки Python")
            print()
            if os.path.exists(exe_path):
                file_size = os.path.getsize(exe_path) / (1024 * 1024)  # Размер в МБ
                print(f"📊 Размер файла: {file_size:.1f} МБ (это нормально!)")
        
        print()
        print("=" * 60)
    except subprocess.CalledProcessError as e:
        print(f"❌ Ошибка при сборке: {e}")
        sys.exit(1)

if __name__ == "__main__":
    build_exe()

