#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Универсальный скрипт запуска приложения для всех операционных систем
Работает на Windows, macOS и Linux
"""
import sys
import os
import subprocess
import platform

def check_python():
    """Проверяет версию Python"""
    if sys.version_info < (3, 7):
        print("❌ Ошибка: Требуется Python 3.7 или выше")
        print(f"   Текущая версия: {sys.version}")
        return False
    return True

def check_dependencies():
    """Проверяет и устанавливает зависимости"""
    required_modules = ['pandas', 'openpyxl', 'flask', 'werkzeug']
    missing = []
    
    for module in required_modules:
        try:
            __import__(module)
        except ImportError:
            missing.append(module)
    
    if missing:
        print("⚠️  Обнаружены отсутствующие зависимости:")
        for module in missing:
            print(f"   - {module}")
        print()
        print("Установка зависимостей...")
        try:
            subprocess.check_call([
                sys.executable, "-m", "pip", "install", "-r", "requirements.txt"
            ])
            print("✅ Зависимости установлены!")
            return True
        except subprocess.CalledProcessError:
            print("❌ Ошибка при установке зависимостей")
            print("   Попробуйте установить вручную:")
            print(f"   {sys.executable} -m pip install -r requirements.txt")
            return False
    return True

def main():
    """Главная функция запуска"""
    print("=" * 60)
    print("  Сверка КС-2 со сметой")
    print("=" * 60)
    print()
    print(f"ОС: {platform.system()} {platform.release()}")
    print(f"Python: {sys.version.split()[0]}")
    print()
    
    # Проверяем Python
    if not check_python():
        input("Нажмите Enter для выхода...")
        sys.exit(1)
    
    # Проверяем зависимости
    if not check_dependencies():
        input("Нажмите Enter для выхода...")
        sys.exit(1)
    
    print()
    print("Запуск приложения...")
    print()
    
    # Запускаем веб-приложение через subprocess
    try:
        script_path = os.path.join(os.path.dirname(__file__), 'ks2_validator_web.py')
        if not os.path.exists(script_path):
            script_path = 'ks2_validator_web.py'
        
        # Запускаем через subprocess, чтобы правильно обработать __main__
        subprocess.run([sys.executable, script_path])
    
    except FileNotFoundError:
        print("❌ Ошибка: файл ks2_validator_web.py не найден")
        print("   Убедитесь, что вы запускаете скрипт из папки проекта")
        if platform.system() == 'Windows':
            input("Нажмите Enter для выхода...")
        sys.exit(1)
    except ImportError as e:
        print(f"❌ Ошибка импорта: {e}")
        print("   Убедитесь, что все зависимости установлены:")
        print(f"   {sys.executable} -m pip install -r requirements.txt")
        if platform.system() == 'Windows':
            input("Нажмите Enter для выхода...")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Ошибка при запуске: {e}")
        import traceback
        traceback.print_exc()
        if platform.system() == 'Windows':
            input("Нажмите Enter для выхода...")
        sys.exit(1)

if __name__ == '__main__':
    main()

