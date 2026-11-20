#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Веб-приложение для автоматической проверки КС-2
"""
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from werkzeug.serving import make_server
import os
import sys
import tempfile
import webbrowser
import threading
import time
import urllib.request
import urllib.error
import socket
import subprocess
import signal
import uuid
import logging
from logging.handlers import RotatingFileHandler
from ks2_validator import SmetaParser, KS2Parser, KS2Validator, normalize_smeta_table, auto_normalize_smeta
import pandas as pd
import openpyxl
from datetime import datetime

# Определяем путь к шаблонам для PyInstaller
if getattr(sys, 'frozen', False):
    # Если запущено из .exe
    template_dir = os.path.join(sys._MEIPASS, 'templates')
    app = Flask(__name__, template_folder=template_dir)
else:
    # Если запущено из исходников
    app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

# Для портативности: создаем папку для временных файлов рядом с исполняемым файлом
if getattr(sys, 'frozen', False):
    # Если запущено из .exe - используем папку рядом с исполняемым файлом
    base_dir = os.path.dirname(sys.executable)
    temp_dir = os.path.join(base_dir, 'temp')
    os.makedirs(temp_dir, exist_ok=True)
    app.config['UPLOAD_FOLDER'] = temp_dir
else:
    # Если запущено из исходников - используем системную temp папку
    app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

app.config['SECRET_KEY'] = 'ks2-validator-secret-key'

# Разрешенные расширения
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# Глобальная переменная для хранения сервера (для возможности остановки)
server = None
shutdown_event = threading.Event()
shutdown_called = False
shutdown_lock = threading.Lock()

# Настройка логирования
def setup_logging():
    """Настраивает систему логирования"""
    # Определяем путь к файлу логов
    if getattr(sys, 'frozen', False):
        # Если запущено из .exe, логи в папку рядом с исполняемым файлом
        log_dir = os.path.dirname(sys.executable)
    else:
        # Если запущено из исходников, логи в папку проекта
        log_dir = os.path.dirname(os.path.abspath(__file__))
    
    log_file = os.path.join(log_dir, 'ks2_validator_web.log')
    
    # Настройка формата логирования
    log_format = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Создаем корневой логгер
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    
    # Очищаем существующие обработчики
    logger.handlers = []
    
    # Файловый обработчик с ротацией (макс. 10MB, 5 файлов резерва)
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=10*1024*1024,  # 10MB
        backupCount=5,
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(log_format)
    logger.addHandler(file_handler)
    
    # Консольный обработчик (только INFO и выше)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_format = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%H:%M:%S')
    console_handler.setFormatter(console_format)
    logger.addHandler(console_handler)
    
    # Логируем запуск приложения
    logger.info(f"="*60)
    logger.info(f"Запуск приложения проверки КС-2")
    logger.info(f"Лог-файл: {log_file}")
    logger.info(f"Версия Python: {sys.version}")
    logger.info(f"Платформа: {sys.platform}")
    logger.info(f"="*60)
    
    return logger

# Инициализируем логирование
logger = setup_logging()

# Регистрируем обработчики ошибок Flask
@app.errorhandler(404)
def not_found(error):
    logger.warning(f"404 ошибка: {request.url}")
    return jsonify({'error': 'Страница не найдена'}), 404

@app.errorhandler(500)
def internal_error(error):
    logger.error(f"500 ошибка: {str(error)}", exc_info=True)
    return jsonify({'error': 'Внутренняя ошибка сервера'}), 500

@app.errorhandler(Exception)
def handle_exception(e):
    logger.error(f"Необработанное исключение: {str(e)}", exc_info=True)
    return jsonify({'error': f'Произошла ошибка: {str(e)}'}), 500

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def kill_process_on_port(port):
    """Останавливает процесс, использующий указанный порт"""
    try:
        if sys.platform == 'darwin' or sys.platform.startswith('linux'):
            # Для macOS и Linux используем lsof
            result = subprocess.run(
                ['lsof', '-ti', f':{port}'],
                capture_output=True,
                text=True
            )
            if result.returncode == 0 and result.stdout.strip():
                pids = result.stdout.strip().split('\n')
                for pid in pids:
                    try:
                        pid_int = int(pid)
                        logger.info(f"Найден процесс с PID {pid_int} на порту {port}")
                        print(f"  Найден процесс с PID {pid_int} на порту {port}")
                        os.kill(pid_int, signal.SIGTERM)
                        time.sleep(0.5)  # Даем процессу время на завершение
                        # Проверяем, завершился ли процесс
                        try:
                            os.kill(pid_int, 0)  # Проверка существования процесса
                            # Если процесс еще жив, отправляем SIGKILL
                            os.kill(pid_int, signal.SIGKILL)
                            logger.info(f"Процесс {pid_int} принудительно завершен")
                            print(f"  Процесс {pid_int} принудительно завершен")
                        except ProcessLookupError:
                            logger.info(f"Процесс {pid_int} успешно остановлен")
                            print(f"  Процесс {pid_int} успешно остановлен")
                    except (ValueError, ProcessLookupError, PermissionError) as e:
                        logger.warning(f"Не удалось остановить процесс {pid}: {e}")
                        print(f"  Не удалось остановить процесс {pid}: {e}")
                return True
        elif sys.platform == 'win32':
            # Для Windows используем netstat и taskkill
            result = subprocess.run(
                ['netstat', '-ano'],
                capture_output=True,
                text=True
            )
            if result.returncode == 0:
                lines = result.stdout.split('\n')
                for line in lines:
                    if f':{port}' in line and 'LISTENING' in line:
                        parts = line.split()
                        if len(parts) > 0:
                            try:
                                pid = int(parts[-1])
                                subprocess.run(['taskkill', '/F', '/PID', str(pid)], 
                                             capture_output=True)
                                logger.info(f"Процесс {pid} на порту {port} остановлен (Windows)")
                                print(f"  Процесс {pid} на порту {port} остановлен")
                                return True
                            except (ValueError, Exception) as e:
                                logger.warning(f"Ошибка при остановке процесса на Windows: {e}")
                                pass
    except Exception as e:
        logger.warning(f"Не удалось проверить процессы на порту {port}: {e}")
        print(f"  Предупреждение: не удалось проверить процессы на порту {port}: {e}")
    return False

def find_free_port(start_port=5000, max_attempts=10, kill_existing=False):
    """Находит свободный порт, начиная с start_port
    
    Args:
        start_port: Порт, с которого начинать поиск
        max_attempts: Максимальное количество попыток
        kill_existing: Если True, пытается остановить процесс на занятом порту
    """
    for port in range(start_port, start_port + max_attempts):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(('127.0.0.1', port))
                return port
        except OSError:
            # Порт занят
            if kill_existing and port == start_port:
                # Пытаемся остановить процесс на первом порту
                logger.info(f"Порт {port} занят, пытаемся остановить предыдущий процесс...")
                print(f"Порт {port} занят, пытаемся остановить предыдущий процесс...")
                if kill_process_on_port(port):
                    time.sleep(1)  # Даем время на освобождение порта
                    # Пробуем снова
                    try:
                        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                            s.bind(('127.0.0.1', port))
                            logger.info(f"Порт {port} успешно освобожден")
                            print(f"✓ Порт {port} освобожден")
                            return port
                    except OSError as e:
                        logger.warning(f"Порт {port} все еще занят после попытки освобождения: {e}")
                        print(f"  Порт {port} все еще занят, пробуем следующий...")
            continue
    # Если не нашли свободный порт в диапазоне, пробуем случайный
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(('127.0.0.1', 0))
            return s.getsockname()[1]
    except OSError:
        return 5000  # Fallback

@app.route('/')
def index():
    """Главная страница"""
    logger.debug(f"Запрос главной страницы от {request.remote_addr}")
    return render_template('index.html')

@app.route('/api/validate', methods=['POST'])
def validate():
    """API для проверки файлов"""
    try:
        logger.info("Получен запрос на валидацию файлов")
        
        if 'smeta_file' not in request.files or 'ks2_file' not in request.files:
            logger.warning("Запрос на валидацию: отсутствуют один или оба файла")
            return jsonify({'error': 'Необходимо загрузить оба файла'}), 400
        
        smeta_file = request.files['smeta_file']
        ks2_file = request.files['ks2_file']
        
        if smeta_file.filename == '' or ks2_file.filename == '':
            logger.warning("Запрос на валидацию: файлы не выбраны")
            return jsonify({'error': 'Файлы не выбраны'}), 400
        
        if not (allowed_file(smeta_file.filename) and allowed_file(ks2_file.filename)):
            logger.warning(f"Запрос на валидацию: неподдерживаемый формат файлов - {smeta_file.filename}, {ks2_file.filename}")
            return jsonify({'error': 'Неподдерживаемый формат файла. Используйте .xlsx или .xls'}), 400
        
        logger.info(f"Начало валидации: смета={smeta_file.filename}, КС-2={ks2_file.filename}")
        
        # Сохраняем файлы во временную директорию
        smeta_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(smeta_file.filename))
        ks2_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(ks2_file.filename))
        
        smeta_file.save(smeta_path)
        ks2_file.save(ks2_path)
        
        # Парсим файлы
        logger.debug(f"Парсинг файла сметы: {smeta_path}")
        smeta_parser = SmetaParser(smeta_path)
        smeta_positions = smeta_parser.parse()
        logger.debug(f"Распарсено позиций сметы: {len(smeta_positions)}")
        
        logger.debug(f"Парсинг файла КС-2: {ks2_path}")
        ks2_parser = KS2Parser(ks2_path)
        ks2_positions = ks2_parser.parse()
        ks2_header = ks2_parser.parse_header()
        logger.debug(f"Распарсено позиций КС-2: {len(ks2_positions)}")
        
        # Валидируем
        logger.debug("Запуск валидации")
        validator = KS2Validator(ks2_positions, smeta_positions, ks2_header)
        results = validator.validate()
        summary = validator.get_summary()
        logger.info(f"Валидация завершена: ошибок={summary.get('errors', 0)}, предупреждений={summary.get('warnings', 0)}, успешных={summary.get('ok', 0)}")
        
        # Сортируем результаты по номеру позиции КС-2 (первая графа)
        # Разделяем результаты позиций и ошибки шапки
        position_results = [r for r in results if r.ks2_pos is not None]
        header_results = [r for r in results if r.ks2_pos is None]
        
        # Сортируем позиции по номеру КС-2
        position_results.sort(key=lambda x: int(x.ks2_pos.number) if x.ks2_pos.number.isdigit() else 9999)
        
        # Объединяем: сначала позиции (по порядку), потом ошибки шапки
        sorted_results = position_results + header_results
        
        # Форматируем результаты для отправки
        formatted_results = []
        for idx, result in enumerate(sorted_results):
            # Отладка для первых 3 результатов
            if idx < 3:
                print(f"[ОТЛАДКА] Результат {idx}:")
                print(f"  ks2_pos: {result.ks2_pos}")
                print(f"  smeta_pos: {result.smeta_pos}")
                if result.smeta_pos:
                    print(f"  smeta_pos.quantity: {result.smeta_pos.quantity}")
                    print(f"  smeta_pos.price: {result.smeta_pos.price}")
                    print(f"  smeta_pos.cost: {result.smeta_pos.cost}")
            ks2_data = None
            smeta_data = None
            
            if result.ks2_pos:
                ks2_data = {
                    'number': result.ks2_pos.number,
                    'smeta_number': result.ks2_pos.smeta_number,
                    'name': result.ks2_pos.name,
                    'justification': result.ks2_pos.justification,
                    'unit': result.ks2_pos.unit,
                    'quantity': float(result.ks2_pos.quantity) if result.ks2_pos.quantity else 0.0,
                    'price': float(result.ks2_pos.price) if result.ks2_pos.price else 0.0,
                    'cost': float(result.ks2_pos.cost) if result.ks2_pos.cost else 0.0,
                }
            
            # КРИТИЧНО: Всегда создаем smeta_data из result.smeta_pos
            # Если result.smeta_pos равен None, создаем пустой объект
            if result.smeta_pos:
                try:
                    smeta_data = {
                        'number': str(result.smeta_pos.number) if result.smeta_pos.number is not None else None,
                        'name': str(result.smeta_pos.name) if result.smeta_pos.name is not None else None,
                        'justification': str(result.smeta_pos.justification) if result.smeta_pos.justification is not None else None,
                        'unit': str(result.smeta_pos.unit) if result.smeta_pos.unit is not None else None,
                        'quantity': float(result.smeta_pos.quantity) if result.smeta_pos.quantity is not None and result.smeta_pos.quantity != '' else None,
                        'price': float(result.smeta_pos.price) if result.smeta_pos.price is not None and result.smeta_pos.price != '' else None,
                        'cost': float(result.smeta_pos.cost) if result.smeta_pos.cost is not None and result.smeta_pos.cost != '' else None,
                    }
                    # Отладка для первых 3 результатов
                    if idx < 3:
                        print(f"  [ВЕБ] smeta_data создан: quantity={smeta_data['quantity']}, price={smeta_data['price']}, cost={smeta_data['cost']}")
                except (ValueError, TypeError, AttributeError) as e:
                    # Если ошибка при преобразовании, создаем с None значениями
                    print(f"  [ВЕБ] Ошибка при создании smeta_data: {e}")
                    smeta_data = {
                        'number': None,
                        'name': None,
                        'justification': None,
                        'unit': None,
                        'quantity': None,
                        'price': None,
                        'cost': None,
                    }
            else:
                # Если smeta_pos равен None, создаем пустой объект для отображения
                smeta_data = {
                    'number': None,
                    'name': None,
                    'justification': None,
                    'unit': None,
                    'quantity': None,
                    'price': None,
                    'cost': None,
                }
                if idx < 3:
                    print(f"  [ВЕБ] smeta_pos равен None, создан пустой smeta_data")
            
            # Формируем статус для отображения
            status_text = 'OK' if result.issue_type.name == 'OK' else result.message
            
            formatted_results.append({
                'number': result.ks2_pos.number if result.ks2_pos else '-',
                'smeta_number': result.ks2_pos.smeta_number if result.ks2_pos else '-',
                'issue_type': result.issue_type.value,
                'issue_type_name': result.issue_type.name,
                'message': result.message,
                'status': status_text,  # Добавляем поле status для совместимости с новым интерфейсом
                'quantity_diff': result.quantity_diff,
                'cost_diff': result.cost_diff,
                'price_diff': getattr(result, 'price_diff', 0.0),
                'severity': getattr(result, 'severity', 'error'),
                'confidence': getattr(result, 'confidence', 1.0),
                'suggested_fix': getattr(result, 'suggested_fix', ''),
                'details': getattr(result, 'details', {}),
                'warnings': getattr(result, 'warnings', []),
                'ks2_pos': ks2_data,  # Полные данные позиции КС-2
                'smeta_pos': smeta_data,  # Полные данные позиции сметы
            })
        
        # Удаляем временные файлы
        try:
            os.remove(smeta_path)
            os.remove(ks2_path)
            logger.debug("Временные файлы удалены после валидации")
        except Exception as e:
            logger.warning(f"Не удалось удалить временные файлы после валидации: {e}")
        
        # Формируем рекомендации на основе результатов
        recommendations = []
        error_count = summary.get('errors', 0)
        warning_count = summary.get('warnings', 0)
        
        if error_count > 0:
            recommendations.append({
                'title': 'Критические ошибки обнаружены',
                'severity': 'error',
                'items': [
                    f'Найдено {error_count} ошибок, требующих исправления',
                    'Проверьте превышения количества, стоимости или цены',
                    'Убедитесь, что все позиции КС-2 присутствуют в смете'
                ]
            })
        
        if warning_count > 0:
            recommendations.append({
                'title': 'Предупреждения',
                'severity': 'warning',
                'items': [
                    f'Обнаружено {warning_count} предупреждений',
                    'Проверьте экономию по позициям - возможно, требуется корректировка',
                    'Убедитесь в корректности данных'
                ]
            })
        
        if error_count == 0 and warning_count == 0:
            recommendations.append({
                'title': 'Проверка пройдена успешно',
                'severity': 'success',
                'items': [
                    'Все позиции соответствуют смете',
                    'Превышений не обнаружено',
                    'Документы готовы к использованию'
                ]
            })
        
        logger.info("Валидация успешно завершена, отправка результатов")
        return jsonify({
            'success': True,
            'summary': summary,
            'results': formatted_results,
            'recommendations': recommendations
        })
        
    except Exception as e:
        logger.error(f"Ошибка при валидации файлов: {str(e)}", exc_info=True)
        return jsonify({'error': f'Ошибка при проверке: {str(e)}'}), 500

@app.route('/api/test-tables', methods=['POST'])
def test_tables():
    """API для получения трех тестовых таблиц (нормализованные КС-2, Смета, объединенная)"""
    try:
        logger.info("Получен запрос на получение тестовых таблиц")
        
        if 'smeta_file' not in request.files or 'ks2_file' not in request.files:
            logger.warning("Запрос на тестовые таблицы: отсутствуют один или оба файла")
            return jsonify({'error': 'Необходимо загрузить оба файла'}), 400
        
        smeta_file = request.files['smeta_file']
        ks2_file = request.files['ks2_file']
        
        if smeta_file.filename == '' or ks2_file.filename == '':
            return jsonify({'error': 'Файлы не выбраны'}), 400
        
        if not (allowed_file(smeta_file.filename) and allowed_file(ks2_file.filename)):
            return jsonify({'error': 'Неподдерживаемый формат файла. Используйте .xlsx или .xls'}), 400
        
        # Сохраняем файлы во временную директорию
        smeta_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(smeta_file.filename))
        ks2_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(ks2_file.filename))
        
        smeta_file.save(smeta_path)
        ks2_file.save(ks2_path)
        
        # Парсим файлы
        smeta_parser = SmetaParser(smeta_path)
        smeta_positions = smeta_parser.parse()
        
        ks2_parser = KS2Parser(ks2_path)
        ks2_positions = ks2_parser.parse()
        ks2_header = ks2_parser.parse_header()
        
        # Создаем валидатор и получаем тестовые таблицы
        print(f"[API] Создание валидатора. Позиций КС-2: {len(ks2_positions)}, позиций Сметы: {len(smeta_positions)}")
        validator = KS2Validator(ks2_positions, smeta_positions, ks2_header)
        test_tables = validator.get_test_tables()
        
        # Проверяем данные перед отправкой
        print(f"[API] Тестовые таблицы получены:")
        print(f"  - Нормализованная КС-2: {len(test_tables.get('normalized_ks2', []))} записей")
        print(f"  - Нормализованная Смета: {len(test_tables.get('normalized_smeta', []))} записей")
        print(f"  - Объединенная таблица: {len(test_tables.get('comparison_table', []))} записей")
        
        if len(test_tables.get('normalized_smeta', [])) == 0:
            print(f"[API ОШИБКА] Таблица нормализованной Сметы пустая!")
            print(f"  - Проверяем validator.normalized_smeta: {len(validator.normalized_smeta)} записей")
            print(f"  - Проверяем validator.smeta_positions: {len(validator.smeta_positions)} записей")
        
        # Удаляем временные файлы
        try:
            os.remove(smeta_path)
            os.remove(ks2_path)
            logger.debug("Временные файлы удалены после получения тестовых таблиц")
        except Exception as e:
            logger.warning(f"Не удалось удалить временные файлы после получения тестовых таблиц: {e}")
        
        logger.info(f"Тестовые таблицы успешно получены: КС-2={len(test_tables.get('normalized_ks2', []))}, Смета={len(test_tables.get('normalized_smeta', []))}, Объединенная={len(test_tables.get('comparison_table', []))}")
        return jsonify({
            'success': True,
            'tables': test_tables
        })
        
    except Exception as e:
        logger.error(f"Ошибка при получении тестовых таблиц: {str(e)}", exc_info=True)
        return jsonify({'error': f'Ошибка при получении тестовых таблиц: {str(e)}'}), 500

@app.route('/api/smeta-table', methods=['POST'])
def smeta_table():
    """API для получения нормализованной таблицы только по смете"""
    try:
        logger.info("Получен запрос на формирование нормализованной таблицы сметы")
        
        if 'smeta_file' not in request.files:
            logger.warning("Запрос на нормализацию сметы: файл не загружен")
            return jsonify({'error': 'Необходимо загрузить файл сметы'}), 400
        
        smeta_file = request.files['smeta_file']
        
        if smeta_file.filename == '':
            return jsonify({'error': 'Файл не выбран'}), 400
        
        if not allowed_file(smeta_file.filename):
            return jsonify({'error': 'Неподдерживаемый формат файла. Используйте .xlsx или .xls'}), 400
        
        # Сохраняем файл во временную директорию
        smeta_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(smeta_file.filename))
        smeta_file.save(smeta_path)
        
        print(f"\n{'='*80}")
        print("ФОРМИРОВАНИЕ НОРМАЛИЗОВАННОЙ ТАБЛИЦЫ СМЕТЫ")
        print(f"{'='*80}")
        print(f"Файл: {smeta_file.filename}")
        
        # Используем автоматическую нормализацию с автопоиском начала данных
        print("\nАвтоматическая нормализация сметы...")
        print("  - Поиск строки с номерами граф (1, 2, 3, 4, 5, 6, 7, 8)")
        print("  - Автоматическое определение индексов столбцов")
        print("  - Чтение данных начиная со строки после номеров граф")
        
        normalized_table = auto_normalize_smeta(smeta_path)
        logger.info(f"Сформировано записей в нормализованной таблице сметы: {len(normalized_table)}")
        print(f"✓ Сформировано записей: {len(normalized_table)}")
        
        # Вывод отладочной информации (первые 3 позиции)
        print(f"\n{'='*80}")
        print("ПЕРВЫЕ 3 ПОЗИЦИИ ИЗ НОРМАЛИЗОВАННОЙ ТАБЛИЦЫ СМЕТЫ:")
        print(f"{'='*80}")
        
        for i, row in enumerate(normalized_table[:3], 1):
            print(f"\nПозиция {i}:")
            print(f"  Номер: {row['number']}")
            print(f"  Обоснование: {row['justification'][:50] if row['justification'] else ''}...")
            print(f"  Наименование: {row['name'][:50] if row['name'] else ''}...")
            print(f"  Единица: {row['unit']}")
            print(f"  Количество: {row['quantity']}")
            print(f"  Цена: {row['price']}")
            print(f"  Стоимость: {row['cost']}")
        
        # Удаляем временный файл
        try:
            os.remove(smeta_path)
            logger.debug("Временный файл сметы удален")
            print(f"\n✓ Временный файл удален")
        except Exception as e:
            logger.warning(f"Не удалось удалить временный файл сметы: {e}")
            print(f"\n⚠ Ошибка при удалении временного файла: {e}")
        
        print(f"{'='*80}\n")
        
        # Возвращаем результат
        logger.info("Нормализованная таблица сметы успешно сформирована и отправлена")
        return jsonify({
            'success': True,
            'count': len(normalized_table),
            'table': normalized_table
        })
        
    except Exception as e:
        import traceback
        error_msg = f'Ошибка при формировании таблицы сметы: {str(e)}'
        logger.error(f"Ошибка при формировании таблицы сметы: {error_msg}", exc_info=True)
        print(f"\n{'='*80}")
        print("ОШИБКА:")
        print(f"{'='*80}")
        print(error_msg)
        print(traceback.format_exc())
        print(f"{'='*80}\n")
        
        return jsonify({'error': error_msg}), 500

@app.route('/api/export-test-tables', methods=['POST'])
def export_test_tables():
    """API для экспорта тестовых таблиц в Excel"""
    try:
        logger.info("Получен запрос на экспорт тестовых таблиц в Excel")
        
        data = request.json
        if not data or 'tables' not in data:
            logger.warning("Запрос на экспорт тестовых таблиц: данные не предоставлены")
            return jsonify({'error': 'Данные таблиц не предоставлены'}), 400
        
        tables = data['tables']
        
        # Создаем Excel файл
        wb = openpyxl.Workbook()
        
        # Удаляем дефолтный лист
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])
        
        # Лист 1: Нормализованная КС-2
        ws1 = wb.create_sheet("Нормализованная КС-2")
        ws1.append(['№ п/п', '№ по смете', 'Наименование', 'Единица', 'Количество', 'Цена', 'Стоимость', 'Источник'])
        for row in tables.get('normalized_ks2', []):
            ws1.append([
                row.get('number', ''),
                row.get('justification_or_smeta_number', ''),
                row.get('name', ''),
                row.get('unit', ''),
                row.get('quantity', 0),
                row.get('price', 0),
                row.get('cost', 0),
                row.get('source', '')
            ])
        
        # Лист 2: Нормализованная Смета
        ws2 = wb.create_sheet("Нормализованная Смета")
        ws2.append(['№ п/п', 'Обоснование', 'Наименование', 'Единица', 'Количество', 'Цена', 'Стоимость', 'Источник'])
        for row in tables.get('normalized_smeta', []):
            ws2.append([
                row.get('number', ''),
                row.get('justification_or_smeta_number', ''),
                row.get('name', ''),
                row.get('unit', ''),
                row.get('quantity', 0),
                row.get('price', 0),
                row.get('cost', 0),
                row.get('source', '')
            ])
        
        # Лист 3: Объединенная таблица
        ws3 = wb.create_sheet("Объединенная таблица")
        ws3.append([
            '№ КС-2', '№ Сметы',
            'КС-2: Наименование', 'КС-2: Обоснование', 'КС-2: Единица', 'КС-2: Количество', 'КС-2: Цена', 'КС-2: Стоимость',
            'Смета: Наименование', 'Смета: Обоснование', 'Смета: Единица', 'Смета: Количество', 'Смета: Цена', 'Смета: Стоимость',
            'Есть КС-2', 'Есть Смета'
        ])
        for row in tables.get('comparison_table', []):
            ws3.append([
                row.get('ks2_number', ''),
                row.get('smeta_number', ''),
                row.get('ks2_name', ''),
                row.get('ks2_justification', ''),
                row.get('ks2_unit', ''),
                row.get('ks2_quantity', 0),
                row.get('ks2_price', 0),
                row.get('ks2_cost', 0),
                row.get('smeta_name', ''),
                row.get('smeta_justification', ''),
                row.get('smeta_unit', ''),
                row.get('smeta_quantity', 0),
                row.get('smeta_price', 0),
                row.get('smeta_cost', 0),
                'Да' if row.get('has_ks2', False) else 'Нет',
                'Да' if row.get('has_smeta', False) else 'Нет'
            ])
        
        # Создаем уникальное имя файла
        filename = f'test_tables_{uuid.uuid4().hex[:8]}.xlsx'
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        try:
            wb.save(output_path)
            
            # Отправляем файл
            logger.info(f"Тестовые таблицы успешно экспортированы в файл: {filename}")
            return send_file(
                output_path,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='test_tables.xlsx'
            )
        except Exception as e:
            # Удаляем файл в случае ошибки
            logger.error(f"Ошибка при сохранении файла экспорта тестовых таблиц: {e}")
            try:
                if os.path.exists(output_path):
                    os.remove(output_path)
            except Exception as del_e:
                logger.warning(f"Не удалось удалить файл экспорта после ошибки: {del_e}")
            raise e
        
    except Exception as e:
        logger.error(f"Ошибка при экспорте тестовых таблиц: {str(e)}", exc_info=True)
        return jsonify({'error': f'Ошибка при экспорте: {str(e)}'}), 500

@app.route('/api/export', methods=['POST'])
def export():
    """API для экспорта результатов в Excel"""
    try:
        logger.info("Получен запрос на экспорт результатов в Excel")
        
        data = request.json
        if not data:
            logger.warning("Запрос на экспорт результатов: данные не предоставлены")
            return jsonify({'error': 'Нет данных для экспорта'}), 400
            
        results = data.get('results', [])
        logger.debug(f"Количество результатов для экспорта: {len(results)}")
        
        if not results:
            logger.warning("Запрос на экспорт результатов: результаты пусты")
            return jsonify({'error': 'Нет результатов для экспорта'}), 400
        
        # Создаем таблицу аналогичную интерфейсу
        # Структура: № КС-2, № сметы, Наименование, Обоснование, Ед., 
        # Количество (Смета, КС-2, Остаток), Цена (Смета, КС-2, Разн.), 
        # Стоимость (Смета, КС-2, Остаток), Статус
        
        # Создаем Excel файл напрямую через openpyxl для более точного контроля
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Детальные результаты проверки'
        
        # Заголовки таблицы (аналогично интерфейсу)
        # Первая строка заголовков
        headers_row1 = [
            '№ КС-2', '№ сметы', 'Наименование', 'Обоснование', 'Ед.',
            'Количество', '', '',  # Количество: Смета, КС-2, Остаток
            'Цена, руб.', '', '',  # Цена: Смета, КС-2, Разн.
            'Стоимость, руб.', '', '',  # Стоимость: Смета, КС-2, Остаток
            'Статус'
        ]
        ws.append(headers_row1)
        
        # Вторая строка заголовков (подзаголовки для группированных колонок)
        headers_row2 = [
            '', '', '', '', '',
            'Смета', 'КС-2', 'Остаток',  # Количество
            'Смета', 'КС-2', 'Разн.',  # Цена
            'Смета', 'КС-2', 'Остаток',  # Стоимость
            ''
        ]
        ws.append(headers_row2)
        
        # Объединяем ячейки для группированных заголовков
        from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
        from openpyxl.utils import get_column_letter
        
        # Объединяем ячейки для "Количество"
        ws.merge_cells(f'F1:H1')
        ws['F1'].alignment = Alignment(horizontal='center', vertical='center')
        
        # Объединяем ячейки для "Цена, руб."
        ws.merge_cells(f'I1:K1')
        ws['I1'].alignment = Alignment(horizontal='center', vertical='center')
        
        # Объединяем ячейки для "Стоимость, руб."
        ws.merge_cells(f'L1:N1')
        ws['L1'].alignment = Alignment(horizontal='center', vertical='center')
        
        # Стили для заголовков
        header_fill = PatternFill(start_color='E5E7EB', end_color='E5E7EB', fill_type='solid')
        header_font = Font(bold=True, size=10)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Применяем стили к заголовкам
        for row in [1, 2]:
            for col in range(1, len(headers_row1) + 1):
                cell = ws.cell(row=row, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.border = border
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        # Заполняем данные
        for result in results:
            # Получаем данные позиций
            ks2_pos = result.get('ks2_pos', {})
            smeta_pos = result.get('smeta_pos', {})
            
            # Вычисляем значения
            smeta_quantity = smeta_pos.get('quantity') if smeta_pos and smeta_pos.get('quantity') is not None else 0
            ks2_quantity = ks2_pos.get('quantity') if ks2_pos and ks2_pos.get('quantity') is not None else 0
            quantity_diff = smeta_quantity - ks2_quantity
            
            smeta_price = smeta_pos.get('price') if smeta_pos and smeta_pos.get('price') is not None else 0
            ks2_price = ks2_pos.get('price') if ks2_pos and ks2_pos.get('price') is not None else 0
            price_diff = smeta_price - ks2_price
            
            smeta_cost = smeta_pos.get('cost') if smeta_pos and smeta_pos.get('cost') is not None else 0
            ks2_cost = ks2_pos.get('cost') if ks2_pos and ks2_pos.get('cost') is not None else 0
            cost_diff = smeta_cost - ks2_cost
            
            # Форматируем статус
            issue_type_name = result.get('issue_type_name', '')
            message = result.get('message', '')
            if issue_type_name == 'OK':
                status = 'OK'
            else:
                status = message[:100] if len(message) > 100 else message
            
            # Форматируем значения
            def format_num(num, decimals=2):
                if num is None or (isinstance(num, (int, float)) and (num == 0 or pd.isna(num))):
                    return '-'
                try:
                    return round(float(num), decimals) if decimals == 2 else round(float(num), decimals)
                except:
                    return '-'
            
            row_data = [
                ks2_pos.get('number', '-') if ks2_pos else '-',
                ks2_pos.get('smeta_number', '-') if ks2_pos else '-',
                (ks2_pos.get('name', '-') if ks2_pos else '-')[:100],  # Ограничиваем длину
                (ks2_pos.get('justification', '-') if ks2_pos else '-')[:50],
                ks2_pos.get('unit', '-') if ks2_pos else '-',
                format_num(smeta_quantity, 4) if smeta_pos else '-',
                format_num(ks2_quantity, 4),
                format_num(quantity_diff, 4) if smeta_pos else '-',
                format_num(smeta_price, 2) if smeta_pos else '-',
                format_num(ks2_price, 2),
                format_num(price_diff, 2) if smeta_pos else '-',
                format_num(smeta_cost, 2) if smeta_pos else '-',
                format_num(ks2_cost, 2),
                format_num(cost_diff, 2) if smeta_pos else '-',
                status[:200]  # Ограничиваем длину статуса
            ]
            
            ws.append(row_data)
            
            # Применяем стили к строке данных
            row_num = ws.max_row
            for col in range(1, len(row_data) + 1):
                cell = ws.cell(row=row_num, column=col)
                cell.border = border
                
                # Выравнивание для числовых колонок
                if col in [6, 7, 8, 9, 10, 11, 12, 13, 14]:  # Числовые колонки
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    if isinstance(cell.value, (int, float)):
                        # Цвет для разниц
                        if col in [8, 11, 14]:  # Колонки с разницами
                            if isinstance(cell.value, (int, float)) and cell.value != 0:
                                if cell.value > 0.01:
                                    cell.font = Font(bold=True, color='006600')  # Зеленый для экономии
                                elif cell.value < -0.01:
                                    cell.font = Font(bold=True, color='CC0000')  # Красный для превышения
                else:
                    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            # Цветовое кодирование строки по статусу
            row_fill = None
            if issue_type_name == 'OK':
                row_fill = PatternFill(start_color='D4EDDA', end_color='D4EDDA', fill_type='solid')
            elif 'EXCEED' in issue_type_name or 'ПРЕВЫШЕНИЕ' in message.upper():
                row_fill = PatternFill(start_color='F8D7DA', end_color='F8D7DA', fill_type='solid')
            elif 'NOT_IN_SMETA' in issue_type_name:
                row_fill = PatternFill(start_color='F8D7DA', end_color='F8D7DA', fill_type='solid')
            elif 'MISMATCH' in issue_type_name:
                row_fill = PatternFill(start_color='FFF3CD', end_color='FFF3CD', fill_type='solid')
            
            if row_fill:
                for col in range(1, len(row_data) + 1):
                    ws.cell(row=row_num, column=col).fill = row_fill
        
        # Настраиваем ширину колонок
        column_widths = {
            'A': 8,   # № КС-2
            'B': 10,  # № сметы
            'C': 40,  # Наименование
            'D': 15,  # Обоснование
            'E': 6,   # Ед.
            'F': 12,  # Количество Смета
            'G': 12,  # Количество КС-2
            'H': 12,  # Количество Остаток
            'I': 12,  # Цена Смета
            'J': 12,  # Цена КС-2
            'K': 12,  # Цена Разн.
            'L': 12,  # Стоимость Смета
            'M': 12,  # Стоимость КС-2
            'N': 12,  # Стоимость Остаток
            'O': 30,  # Статус
        }
        
        for col_letter, width in column_widths.items():
            ws.column_dimensions[col_letter].width = width
        
        # Замораживаем первую строку заголовков
        ws.freeze_panes = 'A3'
        
        # Сохраняем во временный файл
        temp_file = os.path.join(app.config['UPLOAD_FOLDER'], f'ks2_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
        
        try:
            wb.save(temp_file)
            
            # Проверяем, что файл создан
            if not os.path.exists(temp_file):
                return jsonify({'error': 'Не удалось создать файл для экспорта'}), 500
            
            file_size = os.path.getsize(temp_file)
            if file_size == 0:
                logger.error("Создан пустой файл экспорта")
                return jsonify({'error': 'Создан пустой файл'}), 500
            
            logger.info(f"Результаты успешно экспортированы в файл: {temp_file} (размер: {file_size} байт)")
            return send_file(
                temp_file, 
                as_attachment=True, 
                download_name=f'ks2_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
        except Exception as e:
            # Удаляем временный файл при ошибке
            logger.error(f"Ошибка при экспорте результатов: {e}", exc_info=True)
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except Exception as del_e:
                    logger.warning(f"Не удалось удалить временный файл экспорта после ошибки: {del_e}")
            raise e
        
    except Exception as e:
        import traceback
        error_msg = f'Ошибка при экспорте: {str(e)}'
        logger.error(f"Критическая ошибка при экспорте результатов: {error_msg}", exc_info=True)
        print(f"Export error: {error_msg}")
        print(traceback.format_exc())
        return jsonify({'error': error_msg}), 500

@app.route('/api/shutdown', methods=['POST', 'GET'])
def shutdown():
    """API для остановки сервера"""
    global server
    try:
        logger.info("Получен запрос на остановку сервера")
        # Устанавливаем событие остановки
        shutdown_event.set()
        
        # Останавливаем сервер в отдельном потоке
        def stop_server():
            global server, shutdown_called
            time.sleep(0.5)  # Даем время на отправку ответа клиенту
            try:
                with shutdown_lock:
                    if not shutdown_called and server:
                        # Используем shutdown() для остановки сервера
                        # Это должно остановить accept() цикл в serve_forever()
                        server.shutdown()
                        shutdown_called = True
                        logger.info("Команда остановки отправлена серверу")
                        print("✓ Команда остановки отправлена серверу")
            except Exception as e:
                logger.error(f"Ошибка при вызове server.shutdown(): {e}", exc_info=True)
                print(f"⚠ Ошибка при вызове server.shutdown(): {e}")
                import traceback
                traceback.print_exc()
        
        shutdown_thread = threading.Thread(target=stop_server, daemon=False)
        shutdown_thread.start()
        
        return jsonify({'success': True, 'message': 'Сервер будет остановлен'})
    except Exception as e:
        import traceback
        error_msg = f'Ошибка при остановке сервера: {str(e)}'
        logger.error(f"Критическая ошибка при остановке сервера: {error_msg}", exc_info=True)
        print(f"Shutdown error: {error_msg}")
        print(traceback.format_exc())
        return jsonify({'error': error_msg}), 500

def open_browser(port):
    """Открывает браузер после запуска сервера"""
    url = f'http://127.0.0.1:{port}'
    
    # Ждем, пока сервер запустится (проверяем доступность)
    max_attempts = 20  # Проверяем до 20 раз (6 секунд)
    
    for attempt in range(max_attempts):
        try:
            time.sleep(0.3)  # Проверяем каждые 0.3 секунды
            response = urllib.request.urlopen(url, timeout=2)
            if response.getcode() == 200:
                # Сервер доступен, открываем браузер
                print("✓ Сервер доступен, открываем браузер...")
                
                # Для macOS используем 'open' команду (работает лучше в .app bundle)
                if sys.platform == 'darwin':
                    try:
                        subprocess.Popen(['open', url],
                                        stdout=subprocess.DEVNULL,
                                        stderr=subprocess.DEVNULL)
                        logger.info(f"Браузер открыт автоматически (macOS): {url}")
                        print("✓ Браузер открыт автоматически!")
                        return
                    except Exception as e:
                        logger.warning(f"Не удалось открыть браузер через 'open': {e}")
                        # Пробуем стандартный способ
                        try:
                            webbrowser.open(url)
                            logger.info(f"Браузер открыт через webbrowser: {url}")
                            print("✓ Браузер открыт автоматически!")
                            return
                        except Exception as e2:
                            logger.error(f"Не удалось открыть браузер: {e2}")
                            print(f"⚠ Ошибка открытия браузера: {e2}")
                            print(f"   Пожалуйста, откройте вручную: {url}")
                            return
                else:
                    # Для других платформ используем стандартный способ
                    try:
                        webbrowser.open(url)
                        logger.info(f"Браузер открыт автоматически: {url}")
                        print("✓ Браузер открыт автоматически!")
                        return
                    except Exception as e:
                        logger.error(f"Не удалось открыть браузер: {e}")
                        print(f"⚠ Ошибка открытия браузера: {e}")
                        print(f"   Пожалуйста, откройте вручную: {url}")
                        return
        except (urllib.error.URLError, OSError, socket.timeout, Exception) as e:
            # Сервер еще не запустился, продолжаем ждать
            if attempt == max_attempts - 1:
                # Последняя попытка - все равно пытаемся открыть браузер
                print("⚠ Не удалось проверить доступность сервера, открываем браузер...")
                if sys.platform == 'darwin':
                    try:
                        subprocess.Popen(['open', url],
                                        stdout=subprocess.DEVNULL,
                                        stderr=subprocess.DEVNULL)
                        logger.info(f"Браузер открыт (macOS, последняя попытка): {url}")
                        print("✓ Браузер открыт!")
                    except Exception as e2:
                        logger.error(f"Не удалось открыть браузер на macOS: {e2}")
                        print(f"⚠ Не удалось открыть браузер автоматически: {e2}")
                        print(f"   Пожалуйста, откройте вручную: {url}")
                else:
                    try:
                        webbrowser.open(url)
                        logger.info(f"Браузер открыт (последняя попытка): {url}")
                        print("✓ Браузер открыт!")
                    except Exception as e_open:
                        logger.error(f"Не удалось открыть браузер автоматически: {e_open}")
                        print(f"⚠ Не удалось открыть браузер автоматически: {e_open}")
                        print(f"   Пожалуйста, откройте вручную: {url}")
            continue

if __name__ == '__main__':
    # Запускаем Flask сервер
    # Примечание: server и shutdown_called уже объявлены как глобальные переменные на уровне модуля
    print("=" * 60)
    print("  Сверка КС-2 со сметой")
    print("=" * 60)
    print()
    
    # Сбрасываем флаг при новом запуске
    shutdown_called = False
    
    # Находим свободный порт (пытаемся остановить процесс на порту 5000, если он занят)
    port = find_free_port(5000, kill_existing=True)
    if port != 5000:
        logger.warning(f"Порт 5000 был занят, используем порт {port}")
        print(f"⚠ Порт 5000 был занят, используем порт {port}")
        print()
    
    logger.info(f"Запуск веб-сервера на порту {port}")
    print("Запуск веб-сервера...")
    print(f"Браузер откроется автоматически через несколько секунд...")
    print(f"Адрес: http://127.0.0.1:{port}")
    print()
    
    # Запускаем Flask сервер с возможностью остановки
    try:
        server = make_server('127.0.0.1', port, app, threaded=True)
        logger.info("Сервер создан, запускаем...")
        print("✓ Сервер создан, запускаем...")
        
        # Запускаем сервер в отдельном потоке
        def run_server():
            try:
                logger.info("Сервер запущен в отдельном потоке")
                server.serve_forever()
            except Exception as e:
                logger.error(f"Ошибка в потоке сервера: {e}", exc_info=True)
                print(f"⚠ Ошибка в потоке сервера: {e}")
        
        server_thread = threading.Thread(target=run_server, daemon=False)
        server_thread.start()
        
        # Ждем, пока сервер запустится и станет доступен
        print("Ожидание запуска сервера...")
        server_ready = False
        for attempt in range(20):  # Проверяем до 20 раз (6 секунд)
            try:
                time.sleep(0.3)
                response = urllib.request.urlopen(f'http://127.0.0.1:{port}', timeout=1)
                if response.getcode() == 200:
                    server_ready = True
                    break
            except:
                continue
        
        if server_ready:
            logger.info("Сервер запущен и доступен")
            print("✓ Сервер запущен и доступен")
        else:
            logger.warning("Сервер запущен, но проверка доступности не удалась")
            print("⚠ Сервер запущен, но проверка доступности не удалась")
        
        logger.info("Сервер готов к работе")
        print("  Для остановки используйте кнопку 'Завершить процесс' на странице")
        print("  или закройте страницу в браузере")
        print()
        
        # Открываем браузер ПОСЛЕ запуска сервера
        # Для macOS .app bundle используем subprocess напрямую
        url = f'http://127.0.0.1:{port}'
        try:
            # Для macOS всегда используем 'open' команду (работает лучше в .app bundle)
            if sys.platform == 'darwin':
                # Даем серверу немного времени на полный запуск
                time.sleep(1)
                try:
                    # Используем 'open' команду macOS - она работает надежнее в .app bundle
                    subprocess.Popen(['open', url], 
                                    stdout=subprocess.DEVNULL, 
                                    stderr=subprocess.DEVNULL)
                    logger.info(f"Браузер открыт автоматически (macOS): {url}")
                    print("✓ Браузер открыт автоматически!")
                except Exception as e:
                    logger.warning(f"Не удалось открыть браузер через 'open': {e}")
                    # Пробуем стандартный способ
                    try:
                        webbrowser.open(url)
                        logger.info(f"Браузер открыт через webbrowser: {url}")
                        print("✓ Браузер открыт автоматически!")
                    except Exception as e2:
                        logger.error(f"Не удалось открыть браузер: {e2}")
                        print(f"⚠ Не удалось открыть браузер автоматически")
                        print(f"   Пожалуйста, откройте вручную: {url}")
            else:
                # Для других платформ используем стандартный способ
                if server_ready:
                    try:
                        webbrowser.open(url)
                        logger.info(f"Браузер открыт автоматически: {url}")
                        print("✓ Браузер открыт автоматически!")
                    except Exception as e:
                        logger.error(f"Не удалось открыть браузер: {e}")
                        print(f"⚠ Не удалось открыть браузер автоматически: {e}")
                        print(f"   Пожалуйста, откройте вручную: {url}")
                else:
                    # Сервер не готов, запускаем поток для проверки и открытия
                    browser_thread = threading.Thread(target=open_browser, args=(port,), daemon=False)
                    browser_thread.start()
        except Exception as e:
            logger.error(f"Ошибка при открытии браузера: {e}", exc_info=True)
            print(f"⚠ Ошибка при открытии браузера: {e}")
            print(f"   Пожалуйста, откройте вручную: {url}")
        
        # Ждем события остановки
        try:
            shutdown_event.wait()
        except KeyboardInterrupt:
            pass
        
        # Останавливаем сервер (если еще не остановлен)
        logger.info("Остановка сервера...")
        print("\nОстановка сервера...")
        try:
            with shutdown_lock:
                if not shutdown_called and server:
                    server.shutdown()
                    shutdown_called = True
                    logger.info("Команда остановки отправлена серверу")
                    print("✓ Команда остановки отправлена серверу")
        except Exception as e:
            logger.error(f"Ошибка при остановке сервера: {e}", exc_info=True)
            print(f"⚠ Ошибка при остановке сервера: {e}")
            import traceback
            traceback.print_exc()
        
        # Ждем завершения потока сервера
        try:
            server_thread.join(timeout=3)
            if server_thread.is_alive():
                logger.warning("Поток сервера не завершился в течение 3 секунд")
                print("⚠ Поток сервера не завершился, принудительное завершение...")
            else:
                logger.info("Поток сервера успешно завершен")
        except Exception as e:
            logger.error(f"Ошибка при ожидании завершения потока: {e}", exc_info=True)
            print(f"⚠ Ошибка при ожидании завершения потока: {e}")
        
    except KeyboardInterrupt:
        logger.info("Получено прерывание от пользователя (Ctrl+C)")
        print()
        print("=" * 60)
        print("Приложение остановлено")
        print("=" * 60)
        with shutdown_lock:
            if not shutdown_called and server:
                try:
                    server.shutdown()
                    shutdown_called = True
                    logger.info("Сервер остановлен после прерывания")
                except Exception as e:
                    logger.warning(f"Ошибка при остановке сервера после прерывания: {e}")
    except Exception as e:
        logger.critical(f"Критическая ошибка при запуске сервера: {e}", exc_info=True)
        print(f"Ошибка при запуске сервера: {e}")
        print()
        import traceback
        traceback.print_exc()
        with shutdown_lock:
            if not shutdown_called and server:
                try:
                    server.shutdown()
                    shutdown_called = True
                except Exception as e:
                    logger.error(f"Ошибка при попытке остановки сервера после критической ошибки: {e}")
    finally:
        logger.info("Приложение завершает работу")
        print("Сервер остановлен")
        print("Если браузер не открылся, перейдите по адресу:")
        print(f"  http://127.0.0.1:{port}")
        logger.info("="*60)
        logger.info("Приложение завершило работу")
        logger.info("="*60)

