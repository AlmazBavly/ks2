#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Автоматическая проверка формы КС-2 и сравнение с локальной сметой

Смета является эталоном (нормативом), с ней сверяются значения КС-2.

Критерии проверки:
- Графа 3 КС-2 (наименование) должна совпадать с графой 3 сметы
- Графа 4 КС-2 (обоснование) должна совпадать с графой 2 сметы
- Графа 5 КС-2 (единица измерения) должна совпадать с графой 4 сметы
- Графа 6 КС-2 (количество) НЕ ДОЛЖНА превышать графу 5 сметы
- Графа 7 КС-2 (цена) НЕ ДОЛЖНА превышать графу 6 сметы
- Графа 8 КС-2 (стоимость) НЕ ДОЛЖНА превышать графу 7 сметы
"""
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import re
import logging
from typing import Dict, List, Tuple, Optional, Any
from dataclasses import dataclass
from enum import Enum


class IssueType(Enum):
    """Типы несоответствий"""
    EXCEED_QUANTITY = "Превышение количества"
    EXCEED_COST = "Превышение стоимости"
    EXCEED_PRICE = "Превышение цены"
    NOT_IN_SMETA = "Позиция отсутствует в смете"
    UNIT_MISMATCH = "Несоответствие единиц измерения"
    NAME_MISMATCH = "Несоответствие наименования"
    JUSTIFICATION_MISMATCH = "Несоответствие обоснования"
    QUANTITY_MISMATCH = "Несоответствие объёма работ"
    COST_MISMATCH = "Несоответствие стоимости"
    PRICE_MISMATCH = "Несоответствие расценки"
    DUPLICATE = "Дублирование работ"
    HEADER_MISMATCH = "Несоответствие реквизитов"
    WARNING = "Предупреждение"  # Для предупреждений и экономии
    OK = "Соответствует"


@dataclass
class SmetaPosition:
    """Позиция из сметы"""
    number: str  # Номер позиции
    name: str  # Наименование
    unit: str  # Единица измерения
    quantity: float  # Количество
    price: float  # Цена за единицу
    cost: float  # Стоимость
    justification: str  # Обоснование


@dataclass
class KS2Position:
    """Позиция из КС-2
    
    Структура КС-2:
    Столбец 1: № п/п
    Столбец 2: № по смете
    Столбец 3: Наименование работ
    Столбец 4: Номер единичной расценки
    Столбец 5: Единица измерения
    Столбец 6: Выполнено работ (количество)
    Столбец 7: Цена за единицу
    Столбец 8: Стоимость
    """
    number: str  # Номер п/п в КС-2 (столбец 1)
    smeta_number: str  # Номер позиции по смете (столбец 2)
    name: str  # Наименование работ (столбец 3)
    justification: str  # Номер единичной расценки (столбец 4)
    unit: str  # Единица измерения (столбец 5)
    quantity: float  # Выполнено работ - количество (столбец 6)
    price: float  # Цена за единицу (столбец 7)
    cost: float  # Стоимость (столбец 8)


@dataclass
class KS2Header:
    """Реквизиты из шапки КС-2"""
    customer: str = ""  # Заказчик
    customer_okpo: str = ""  # ОКПО заказчика
    contractor: str = ""  # Подрядчик
    contractor_okpo: str = ""  # ОКПО подрядчика
    contract_number: str = ""  # Номер договора
    contract_date: str = ""  # Дата договора
    object_name: str = ""  # Наименование объекта


@dataclass
class NormalizedPosition:
    """Нормализованная позиция для сравнения
    
    Единый формат для позиций из КС-2 и Сметы.
    Содержит только необходимые графы для сверки:
    - Графа 1: № п/п (номер позиции)
    - Графа 2: Обоснование (для сметы) / № по смете (для КС-2)
    - Графа 3: Наименование работ
    - Графа 4: Единица измерения
    - Графа 5: Количество (для сметы - "всего", для КС-2 - "выполнено")
    - Графа 6: Цена за единицу
    - Графа 7: Стоимость
    """
    number: str  # Графа 1: № п/п
    justification_or_smeta_number: str  # Графа 2: Обоснование (смета) или № по смете (КС-2)
    name: str  # Графа 3: Наименование работ
    unit: str  # Графа 4: Единица измерения
    quantity: float  # Графа 5/6: Количество
    price: float  # Графа 6/7: Цена за единицу
    cost: float  # Графа 7/8: Стоимость
    source: str  # "smeta" или "ks2" - источник данных


@dataclass
class ComparisonRow:
    """Строка объединенной таблицы для сверки
    
    Содержит данные из КС-2 и Сметы рядом друг с другом для сравнения.
    Это третья таблица, которая формируется из двух нормализованных таблиц.
    """
    # Идентификаторы
    ks2_number: str  # № п/п в КС-2
    smeta_number: str  # № п/п в смете (из графы 2 КС-2)
    
    # Данные из КС-2
    ks2_name: str
    ks2_justification: str
    ks2_unit: str
    ks2_quantity: float
    ks2_price: float
    ks2_cost: float
    
    # Данные из Сметы
    smeta_name: str
    smeta_justification: str
    smeta_unit: str
    smeta_quantity: float
    smeta_price: float
    smeta_cost: float
    
    # Флаги наличия данных
    has_ks2: bool  # Есть ли данные из КС-2
    has_smeta: bool  # Есть ли данные из сметы
    
    # Оригинальные объекты для формирования ValidationResult
    ks2_orig: Optional[KS2Position] = None
    smeta_orig: Optional[SmetaPosition] = None


@dataclass
class UnifiedNormalizedPosition:
    """Унифицированная нормализованная позиция для единого представления КС-2 и Сметы
    
    НОВЫЙ ПОДХОД: Сразу объединяет данные из КС-2 и Сметы в одну структуру
    """
    # Общие идентификаторы
    ks2_number: str              # № п/п в КС-2 (Графа 1)
    smeta_ref: str               # № по смете в КС-2 (Графа 2)
    
    # Данные КС-2
    ks2_name: str                # Наименование КС-2 (Графа 3)
    ks2_justification: str       # Обоснование КС-2 (Графа 4)
    ks2_unit: str                # Единица КС-2 (Графа 5)
    ks2_quantity: float          # Количество КС-2 (Графа 6)
    ks2_price: float             # Цена КС-2 (Графа 7)
    ks2_cost: float              # Стоимость КС-2 (Графа 8)
    
    # Данные СМЕТЫ
    smeta_number: Optional[str] = None          # № п/п в смете (Графа 1)
    smeta_name: Optional[str] = None            # Наименование сметы (Графа 3)
    smeta_justification: Optional[str] = None   # Обоснование сметы (Графа 2)
    smeta_unit: Optional[str] = None            # Единица сметы (Графа 4)
    smeta_quantity: float = 0.0                 # Количество сметы (Графа 6)
    smeta_price: float = 0.0                    # Цена сметы (Графа 7)
    smeta_cost: float = 0.0                     # Стоимость сметы (Графа 8)
    
    # ОСТАТКИ (Смета - КС-2)
    quantity_remainder: float = 0.0
    price_remainder: float = 0.0
    cost_remainder: float = 0.0
    
    # Статус и проблемы
    status: str = "OK"
    issues: List[str] = None
    
    def __post_init__(self):
        if self.issues is None:
            self.issues = []


def safe_float(val, debug: bool = False):
    """Безопасное преобразование значения в float."""
    if pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        try:
            return float(val)
        except (ValueError, TypeError):
            return None
    s = str(val)
    # Игнорируем формулы и многострочные выражения
    if any(ch in s for ch in ('*', '/', '\n')):
        return None
    normalized = (
        s.replace(',', '.')
         .replace(' ', '')
         .replace('\xa0', '')
         .strip()
    )
    if normalized == '':
        return None
    try:
        return float(normalized)
    except (ValueError, TypeError):
        return None


class Parser:
    """Базовый парсер для Excel-файлов."""

    def __init__(self, filepath: str):
        self.filepath = filepath
        self.df = None
        self.header_row: Optional[int] = None
        self.data_start: Optional[int] = None
        self.cols: Dict[str, int] = {}

    def load(self):
        """Загружает первый лист Excel-файла без заголовков."""
        self.df = pd.read_excel(self.filepath, sheet_name=0, header=None)

    def find_header_row(self, search_terms: List[str]) -> Optional[int]:
        """Находит строку с заголовками, содержащую указанные термины."""
        if self.df is None:
            return None
        for i in range(len(self.df)):
            row = self.df.iloc[i]
            row_str = ' '.join([str(v).lower() for v in row if pd.notna(v)])
            if all(term.lower() in row_str for term in search_terms):
                return i
        return None

    def skip_graph_numbers_row(self):
        """Пропускает строку с номерами граф (1, 2, 3 ...) и строки "Раздел".
        
        КРИТИЧНО: Строка с номерами граф содержит "1", "2", "3", "4", "5", "6", "7", "8"
        в первых 8 колонках. Эта строка НЕ является данными и должна быть пропущена.
        """
        if self.header_row is None or self.df is None:
            return

        # Начальная позиция после заголовков
        self.data_start = self.header_row + 2
        
        # КРИТИЧНО: Пропускаем строку с номерами граф ("1", "2", "3", "4", "5", "6", "7", "8")
        # Проверяем следующие несколько строк на наличие строки с номерами граф
        # Начинаем проверку С САМОЙ СТРОКИ ЗАГОЛОВКА (header_row), так как номера граф могут быть там
        for check_idx in range(self.header_row, min(self.header_row + 10, len(self.df))):
            row_check = self.df.iloc[check_idx]
            
            # Проверяем первые 8 ячеек - если они содержат "1", "2", "3", "4", "5", "6", "7", "8"
            # это строка с номерами граф
            first_eight = []
            for col in range(min(8, len(row_check))):
                val = row_check.iloc[col]
                if pd.notna(val):
                    val_str = str(val).strip()
                    first_eight.append(val_str)
            
            # Строгая проверка: первые 8 значений должны быть "1", "2", "3", "4", "5", "6", "7", "8"
            # Или хотя бы первые 3 должны быть "1", "2", "3" и остальные должны быть числами
            is_graph_numbers_row = False
            if len(first_eight) >= 8:
                # Проверяем, что первые 8 значений - это последовательные числа от 1 до 8
                try:
                    nums = [int(x) for x in first_eight[:8]]
                    if nums == [1, 2, 3, 4, 5, 6, 7, 8]:
                        is_graph_numbers_row = True
                except (ValueError, TypeError):
                    pass
            
            # Альтернативная проверка: первые 3 значения = "1", "2", "3" и все они - числа
            if not is_graph_numbers_row and len(first_eight) >= 3:
                try:
                    if first_eight[0] == "1" and first_eight[1] == "2" and first_eight[2] == "3":
                        # Проверяем, что остальные значения тоже числа
                        all_numbers = True
                        for val_str in first_eight[3:min(8, len(first_eight))]:
                            try:
                                int(val_str)
                            except (ValueError, TypeError):
                                all_numbers = False
                                break
                        if all_numbers:
                            is_graph_numbers_row = True
                except (ValueError, TypeError):
                    pass
            
            if is_graph_numbers_row:
                # Начинаем со следующей строки после номеров граф
                self.data_start = check_idx + 1
                # Проверяем следующую строку - если это "Раздел", пропускаем и её
                if self.data_start < len(self.df):
                    next_row = self.df.iloc[self.data_start]
                    next_row_str = ' '.join([str(x).lower() for x in next_row if pd.notna(x)])
                    if 'раздел' in next_row_str:
                        self.data_start += 1
                break
        
        # Дополнительная проверка: убеждаемся, что мы не на строке "Раздел" или другой технической строке
        if self.data_start is not None and self.data_start < len(self.df):
            row_check = self.df.iloc[self.data_start]
            row_str = ' '.join([str(x).lower() for x in row_check if pd.notna(x)])
            if 'раздел' in row_str:
                # Пропускаем строку "Раздел"
                self.data_start += 1


@dataclass
class ValidationResult:
    """Результат проверки позиции"""
    ks2_pos: Optional[KS2Position] = None
    smeta_pos: Optional[SmetaPosition] = None
    issue_type: IssueType = IssueType.OK
    message: str = ""
    quantity_diff: float = 0.0
    cost_diff: float = 0.0
    price_diff: float = 0.0
    ks2_row: int = 0  # Номер строки в КС-2 для удобства
    smeta_row: int = 0  # Номер строки в смете для удобства
    # Дополнительные поля для подробностей
    details: Dict = None  # Дополнительные детали ошибки
    severity: str = "error"  # Уровень серьезности: "error", "warning", "info"
    suggested_fix: str = ""  # Предложение по исправлению
    confidence: float = 1.0  # Уровень уверенности в ошибке (0.0-1.0)
    warnings: List[str] = None  # Список предупреждений (например, экономия)
    
    def __post_init__(self):
        """Инициализация дополнительных полей"""
        if self.details is None:
            self.details = {}
        if self.warnings is None:
            self.warnings = []


class SmetaParser(Parser):
    """Парсер для чтения сметы"""
    
    def __init__(self, filepath: str):
        super().__init__(filepath)
        self.positions: Dict[str, SmetaPosition] = {}
        # Атрибуты для работы с openpyxl напрямую (для поддержки 14 граф)
        self.wb: Optional[openpyxl.Workbook] = None
        self.ws: Optional[openpyxl.worksheet.worksheet.Worksheet] = None
        self.graph_to_column: Dict[int, int] = {}  # Маппинг: номер графы → физический индекс колонки
        self.data_start_row: Optional[int] = None  # Строка, с которой начинаются данные
        self.column_mapping: Dict[str, Optional[int]] = {}  # Маппинг полей к колонкам
    
    def parse(self) -> Dict[str, SmetaPosition]:
        """Парсит смету и возвращает словарь позиций
        
        Порядок поиска:
        1. НОВЫЙ МЕТОД: Пробует найти 14 граф с помощью openpyxl (с поддержкой merged cells)
        2. СТАРЫЙ МЕТОД: Если не получилось, ищет 8 граф через pandas
        3. Начинает парсить данные
        """
        # ПОПЫТКА 1: Используем новый метод с поддержкой 14 граф и объединенных ячеек
        try:
            print("[СМЕТА] Попытка использовать новый парсер с поддержкой 14 граф...")
            
            # Загружаем файл через openpyxl
            if self.wb is None:
                self.wb = openpyxl.load_workbook(self.filepath, data_only=True)
                self.ws = self.wb.active
            
            # Пробуем определить структуру с 14 графами
            if self.determine_smeta_structure():
                print(f"[СМЕТА] ✓ Найдена структура с {len(self.graph_to_column)} графами!")
                
                # Нормализуем таблицу
                normalized_data = self.normalize_smeta_table()
                
                # Конвертируем в формат SmetaPosition для обратной совместимости
                self.positions = {}
                for pos_data in normalized_data:
                    pos = SmetaPosition(
                        number=pos_data['number'],
                        name=pos_data['name'],
                        justification=pos_data['justification'],
                        unit=pos_data['unit'],
                        quantity=pos_data['quantity'],
                        price=pos_data['price'],
                        cost=pos_data['cost']
                    )
                    self.positions[pos.number] = pos
                
                print(f"[СМЕТА] ✓ Найдено позиций (новый парсер): {len(self.positions)}")
                if self.positions:
                    first_pos = list(self.positions.values())[0]
                    print(f"[СМЕТА] Первая позиция: №{first_pos.number}, name='{first_pos.name[:50]}', qty={first_pos.quantity}")
                
                return self.positions
            else:
                print("[СМЕТА] ⚠ Новый парсер не нашел структуру с 14 графами")
                print("[СМЕТА] → Переключаемся на старый парсер с 8 графами...")
                
        except Exception as e:
            print(f"[СМЕТА] ⚠ Ошибка в новом парсере: {e}")
            print("[СМЕТА] → Переключаемся на старый парсер с 8 графами...")
        
        # ПОПЫТКА 2: Используем старый метод с pandas (для обратной совместимости)
        self.load()
        self.positions = {}
        self.cols = {}
        
        # Ищем заголовок по ключевым словам - графа "№ п/п" обычно первая
        # Пробуем разные варианты поиска
        self.header_row = self.find_header_row(['№', 'п/п'])
        if self.header_row is None:
            self.header_row = self.find_header_row(['п/п'])
        if self.header_row is None:
            self.header_row = self.find_header_row(['номер'])
        if self.header_row is None:
            # Пробуем найти по комбинации "наименование" и "единиц"
            self.header_row = self.find_header_row(['наименование', 'единиц'])
        if self.header_row is None:
            self.header_row = 22  # Fallback
        
        if self.header_row is None:
            raise ValueError("Не найдена строка с заголовками граф в смете")
        
        self._find_columns()
        
        # КРИТИЧНО: Найти строку с номерами граф (1, 2, 3, 4, 5, 6, 7, 8) и использовать её для точного определения колонок
        graph_numbers_row = None
        cols_map = {}  # Словарь: номер графы → индекс колонки
        
        for check_idx in range(self.header_row, min(self.header_row + 10, len(self.df))):
            row_check = self.df.iloc[check_idx]
            
            # Проверяем первые 8 колонок
            first_eight = []
            for col in range(min(8, len(row_check))):
                val = row_check.iloc[col]
                if pd.notna(val):
                    val_str = str(val).strip()
                    # Проверяем, что это число от 1 до 8
                    try:
                        num = int(val_str)
                        if 1 <= num <= 8:
                            first_eight.append(num)
                    except (ValueError, TypeError):
                        pass
            
            # Если нашли последовательность 1, 2, 3, 4, 5, 6, 7, 8
            if len(first_eight) >= 8 and first_eight[:8] == [1, 2, 3, 4, 5, 6, 7, 8]:
                graph_numbers_row = check_idx
                
                # Создаем маппинг: номер графы → индекс колонки
                for col_idx in range(min(8, len(row_check))):
                    val = row_check.iloc[col_idx]
                    if pd.notna(val):
                        try:
                            graph_num = int(str(val).strip())
                            if 1 <= graph_num <= 8:
                                cols_map[graph_num] = col_idx
                        except (ValueError, TypeError):
                            pass
                
                print(f"[СМЕТА] Найдена строка с номерами граф: {check_idx}")
                print(f"[СМЕТА] Маппинг граф: {cols_map}")
                break
        
        # Если нашли строку с номерами граф, обновляем self.cols точными значениями
        if cols_map:
            if 1 in cols_map:
                self.cols['number'] = cols_map[1]
            if 2 in cols_map:
                self.cols['justification'] = cols_map[2]
            if 3 in cols_map:
                self.cols['name'] = cols_map[3]
            if 4 in cols_map:
                self.cols['unit'] = cols_map[4]
            if 6 in cols_map:
                self.cols['qty'] = cols_map[6]  # Графа 6: Количество "всего"
            if 7 in cols_map:
                self.cols['price'] = cols_map[7]  # Графа 7: Цена "на ед."
            if 8 in cols_map:
                self.cols['cost'] = cols_map[8]  # Графа 8: Стоимость "Всего"
            print(f"[СМЕТА] Обновленные колонки по номерам граф: {self.cols}")
        
        self.skip_graph_numbers_row()
        if self.data_start is None:
            self.data_start = self.header_row + 2
        
        # Отладочный вывод для диагностики
        import sys
        debug_mode = '--debug-smeta' in sys.argv  # Отладка только при флаге --debug-smeta
        if debug_mode:
            print(f"\n=== DEBUG SmetaParser ===")
            print(f"header_row={self.header_row}, data_start={self.data_start}")
            print(f"Найденные колонки: {self.cols}")
            if self.df is not None and self.data_start < len(self.df):
                print(f"\nПервые 3 строки данных:")
                for i in range(self.data_start, min(self.data_start + 3, len(self.df))):
                    print(f"\nСтрока {i}:")
                    row = self.df.iloc[i]
                    for col_idx in range(min(10, len(row))):
                        val = row.iloc[col_idx]
                        print(f"  Колонка {col_idx}: {val}")
        
        for row_idx in range(self.data_start, len(self.df)):
            pos = self._parse_row(row_idx, debug_mode=debug_mode and row_idx < self.data_start + 3)
            if pos and pos.number not in self.positions:
                self.positions[pos.number] = pos
        
        # Всегда выводим сводку (даже без debug_mode)
        print(f"\n[СМЕТА] Найдено позиций: {len(self.positions)}")
        if self.positions:
            first_pos = list(self.positions.values())[0]
            print(f"[СМЕТА] Первая позиция: №{first_pos.number}, name='{first_pos.name[:50]}', qty={first_pos.quantity}, price={first_pos.price}, cost={first_pos.cost}")
        else:
            print(f"[СМЕТА] ВНИМАНИЕ: Позиции не найдены! Проверьте:")
            print(f"[СМЕТА]   - header_row={self.header_row}")
            print(f"[СМЕТА]   - data_start={self.data_start}")
            print(f"[СМЕТА]   - cols={self.cols}")
        
        if debug_mode:
            print("=== END DEBUG SmetaParser ===\n")
        
        return self.positions
    
    def normalize(self) -> Dict[str, NormalizedPosition]:
        """Приводит позиции сметы к нормализованному виду.
        
        Возвращает словарь нормализованных позиций, содержащих только необходимые графы:
        - Графа 1: № п/п
        - Графа 2: Обоснование
        - Графа 3: Наименование работ
        - Графа 4: Единица измерения
        - Графа 6: Количество "всего"
        - Графа 7: Цена "на ед."
        - Графа 8: Стоимость "Всего"
        """
        normalized = {}
        for number, pos in self.positions.items():
            normalized[number] = NormalizedPosition(
                number=pos.number,
                justification_or_smeta_number=pos.justification,
                name=pos.name,
                unit=pos.unit,
                quantity=pos.quantity,
                price=pos.price,
                cost=pos.cost,
                source="smeta"
            )
        return normalized
    
    def _find_columns(self):
        """Определяет индексы колонок, анализируя несколько строк заголовков.
        
        Анализирует 3 строки (row-1, row, row+1) для более точного определения колонок.
        """
        if self.df is None or self.header_row is None:
            return
        
        for col_idx in range(self.df.shape[1]):
            val0 = self.df.iat[self.header_row - 1, col_idx] if (self.header_row - 1) >= 0 else None
            val1 = self.df.iat[self.header_row, col_idx]
            val2 = self.df.iat[self.header_row + 1, col_idx] if (self.header_row + 1) < len(self.df) else None
            val_combined = f"{str(val0).lower() if pd.notna(val0) else ''} {str(val1).lower() if pd.notna(val1) else ''} {str(val2).lower() if pd.notna(val2) else ''}"
            
            # Графа 1 СМЕТА: № п/п (номер позиции)
            if ('№' in val_combined or 'п/п' in val_combined or 'номер' in val_combined) and 'number' not in self.cols:
                # Проверяем, что это не "№ по смете" или другой номер
                if 'по смете' not in val_combined and 'расцен' not in val_combined:
                    self.cols['number'] = col_idx
            # Графа 2 СМЕТА: Обоснование (шифр, расценка, код)
            elif ('обосн' in val_combined or 'шифр' in val_combined or 'код' in val_combined) and 'justification' not in self.cols:
                self.cols['justification'] = col_idx
            # Графа 3 СМЕТА: Наименование работ и затрат
            elif ('наименование' in val_combined or 'наименов' in val_combined) and 'name' not in self.cols:
                if 'работ' in val_combined or 'затрат' in val_combined or 'name' not in self.cols:
                    self.cols['name'] = col_idx
            # Графа 4 СМЕТА: Единица измерения
            elif 'unit' not in self.cols:
                # Проверяем разные варианты: "Единица измерения", "Ед. изм.", просто "Единица"
                if ('единиц' in val_combined) or ('ед.' in val_combined and 'изм' in val_combined):
                    self.cols['unit'] = col_idx
            # Графа 6 СМЕТА: Количество / всего
            elif 'qty' not in self.cols:
                if ('количество' in val_combined or 'кол-во' in val_combined or 'объем' in val_combined):
                    # Предпочитаем колонку со словом "всего" если есть выбор
                    if 'всего' in val_combined or 'qty' not in self.cols:
                        self.cols['qty'] = col_idx
            # Графа 7 СМЕТА: Сметная стоимость / на ед.
            elif 'price' not in self.cols:
                if ('цена' in val_combined or 'расценка' in val_combined or 'на ед' in val_combined) and 'всего' not in val_combined:
                    self.cols['price'] = col_idx
            # Графа 8 СМЕТА: Сметная стоимость / Всего
            elif 'cost' not in self.cols:
                if ('стоимость' in val_combined and 'всего' in val_combined) or ('сумма' in val_combined):
                    self.cols['cost'] = col_idx
    
    def _parse_row(self, row_idx: int, debug_mode: bool = False) -> Optional[SmetaPosition]:
        """Парсит одну строку сметы."""
        if self.df is None:
            return None
        
        number_val = self.df.iat[row_idx, self.cols.get('number', 0)]
        if pd.isna(number_val):
            return None
        try:
            number = int(str(number_val).strip())
            if number <= 0:
                return None
        except (ValueError, TypeError):
            return None
        
        name_val = self.df.iat[row_idx, self.cols.get('name', 2)] if 'name' in self.cols else None
        unit_val = self.df.iat[row_idx, self.cols.get('unit', 5)] if 'unit' in self.cols else None
        qty_val = self.df.iat[row_idx, self.cols.get('qty', 7)] if 'qty' in self.cols else None
        price_val = self.df.iat[row_idx, self.cols.get('price', 8)] if 'price' in self.cols else None
        cost_val = self.df.iat[row_idx, self.cols.get('cost', 9)] if 'cost' in self.cols else None
        justification_val = self.df.iat[row_idx, self.cols.get('justification', 1)] if 'justification' in self.cols else None
        
        if debug_mode:
            print(f"\n  _parse_row({row_idx}): number={number}")
            print(f"    name={str(name_val)[:50] if pd.notna(name_val) else 'None'}")
            print(f"    qty_val={qty_val}, price_val={price_val}, cost_val={cost_val}")
        
        # КРИТИЧНО: НЕ используем фильтр по длине наименования!
        # Методика не требует минимальной длины наименования.
        # Короткие наименования типа "Песок" (5 символов) должны обрабатываться.
        if pd.isna(name_val) or str(name_val).strip() == '':
            return None
        if pd.isna(unit_val) or str(unit_val).strip() == '':
            return None
        
        quantity = safe_float(qty_val)
        price = safe_float(price_val)
        cost = safe_float(cost_val)
        
        # КРИТИЧНО: НЕ используем ограничение на максимальное количество (qty < 1000000)
        # Методика не упоминает таких ограничений. Большие объемы должны обрабатываться.
        if quantity is None:
            return None
        if price is None and cost is not None and quantity:
            price = cost / quantity if quantity != 0 else None
        if cost is None and price is not None and quantity is not None:
            cost = price * quantity
        if price is None or cost is None:
            return None
        
        if debug_mode:
            print(f"    OK: qty={quantity:.2f}, price={price:.2f}, cost={cost:.2f}")
        
        justification = str(justification_val).strip() if pd.notna(justification_val) else ''
        
        return SmetaPosition(
            number=str(number),
            name=str(name_val).strip(),
            justification=justification,
            unit=str(unit_val).strip(),
            quantity=quantity,
            price=price,
            cost=cost
        )
    
    def _parse_float(self, value: Any) -> float:
        """Безопасное преобразование значения в float (обертка над safe_float)"""
        result = safe_float(value)
        return result if result is not None else 0.0
    
    def extract_number_from_cell(self, cell_value):
        """Универсальная функция для извлечения числа из значения ячейки"""
        if cell_value is None:
            return None
        
        # Если это уже число
        if isinstance(cell_value, (int, float)):
            try:
                return int(cell_value)
            except (TypeError, ValueError):
                pass
        
        # Если это строка или можно преобразовать в строку
        try:
            cell_str = str(cell_value).strip()
            
            # Сначала пробуем извлечь число с помощью регулярных выражений
            match = re.search(r'\d+', cell_str)
            if match:
                return int(match.group(0))
            
            # Если не получилось, пробуем очистить строку от всех нецифровых символов
            cleaned = re.sub(r'[^\d]', '', cell_str)
            if cleaned and cleaned.isdigit():
                return int(cleaned)
        except (TypeError, ValueError):
            pass
        
        return None
    
    def find_graph_numbers_row(self) -> Optional[int]:
        """
        Ищет строку с номерами граф от 1 до 14 в Excel файле.
        
        Алгоритм:
        1. Начинает с ячейки A1
        2. Ищет последовательность чисел 1, 2, 3, ..., 14 (игнорируя пустые ячейки)
        3. Учитывает объединенные ячейки (merged cells)
        4. Если последовательность 1-14 найдена - это строка с графами
        5. Если нет - переходит к A2 и повторяет проверку
        6. Диапазон поиска: A1 до A50
        7. Если не найдено в колонке A, начинает с B1 (смещение вправо)
        
        ВАЖНО: Между числами могут быть:
        - Пустые ячейки (скрытые колонки)
        - Объединенные ячейки (merged cells)
        
        Returns:
            Optional[int]: Номер строки (1-based), где найдена последовательность 1-14, или None
        """
        if self.wb is None or self.ws is None:
            return None
        
        # Получаем информацию об объединенных ячейках
        merged_ranges = list(self.ws.merged_cells.ranges)
        
        def is_cell_merged(row, col):
            """Проверяет, является ли ячейка частью объединенного диапазона"""
            for merged_range in merged_ranges:
                if (merged_range.min_row <= row <= merged_range.max_row and
                    merged_range.min_col <= col <= merged_range.max_col):
                    # Возвращаем True для всех ячеек кроме первой (где хранится значение)
                    if not (row == merged_range.min_row and col == merged_range.min_col):
                        return True
            return False
        
        # Максимальное количество строк для поиска
        max_rows = min(50, self.ws.max_row)
        # Максимальное количество столбцов для начала поиска (смещение)
        max_start_cols = min(6, self.ws.max_column)  # A, B, C, D, E, F
        # Максимальное количество столбцов для сканирования (учитываем возможные пропуски)
        max_scan_cols = min(30, self.ws.max_column)
        
        # Логируем информацию об объединенных ячейках
        if merged_ranges:
            logging.debug(f"Обнаружено {len(merged_ranges)} объединенных диапазонов в файле")
        
        # ОСНОВНОЙ АЛГОРИТМ: Перебираем начальные колонки (A, B, C, ...)
        for start_col in range(1, max_start_cols + 1):
            # Перебираем строки от 1 до 50 для текущей начальной колонки
            for row_idx in range(1, max_rows + 1):
                # Собираем все непустые числа в строке, начиная с start_col
                numbers_found = []  # список (номер_графы, индекс_колонки)
                merged_cells_skipped = 0  # счетчик пропущенных объединенных ячеек
                
                for col_idx in range(start_col, max_scan_cols + 1):
                    # Проверяем, является ли ячейка частью объединенного диапазона
                    if is_cell_merged(row_idx, col_idx):
                        merged_cells_skipped += 1
                        continue  # Пропускаем объединенную ячейку
                    
                    cell_value = self.ws.cell(row=row_idx, column=col_idx).value
                    num = self.extract_number_from_cell(cell_value)
                    
                    # Если нашли число, добавляем в список
                    if num is not None:
                        numbers_found.append((num, col_idx))
                
                # Проверяем, нашли ли мы последовательность 1-14
                if len(numbers_found) >= 14:
                    # Извлекаем только числа
                    found_numbers = [num for num, _ in numbers_found]
                    
                    # Проверяем, начинается ли с 1 и содержит ли все числа до 14
                    if found_numbers[:14] == list(range(1, 15)):
                        # НАЙДЕНО! Последовательность 1-14
                        logging.info(f"✓ Найдена строка с номерами граф (1-14) в строке {row_idx}, начиная с колонки {get_column_letter(start_col)}")
                        
                        if merged_cells_skipped > 0:
                            logging.info(f"  ℹ Обнаружено и пропущено {merged_cells_skipped} объединенных ячеек")
                        
                        # Сохраняем маппинг номеров граф к физическим колонкам
                        self.graph_to_column = {}
                        for i in range(14):
                            graph_num, col_idx = numbers_found[i]
                            self.graph_to_column[graph_num] = col_idx
                        
                        # Отладочный вывод маппинга
                        logging.debug(f"Маппинг граф к колонкам:")
                        for graph_num in range(1, 15):
                            col_idx = self.graph_to_column.get(graph_num)
                            if col_idx:
                                logging.debug(f"  Графа {graph_num} → колонка {get_column_letter(col_idx)}")
                        
                        # Выводим информацию об объединенных ячейках в этой строке
                        merged_in_row = [r for r in merged_ranges if r.min_row <= row_idx <= r.max_row]
                        if merged_in_row:
                            logging.debug(f"Объединенные ячейки в строке {row_idx}:")
                            for merged_range in merged_in_row:
                                logging.debug(f"  {merged_range}")
                        
                        # Возвращаем номер строки, где найдены графы
                        return row_idx
        
        # Если строгая последовательность 1-14 не найдена
        logging.error("✗ Строка с номерами граф (1-14) не найдена в файле")
        logging.error(f"  Проверено: {max_start_cols} начальных колонок × {max_rows} строк")
        logging.error(f"  Диапазон: {get_column_letter(1)}1 до {get_column_letter(max_start_cols)}{max_rows}")
        
        return None
    
    def determine_smeta_structure(self):
        """Определяет структуру сметы с поддержкой 14 граф"""
        # Загружаем файл через openpyxl для прямого доступа к ячейкам
        if self.wb is None:
            self.wb = openpyxl.load_workbook(self.filepath, data_only=True)
            self.ws = self.wb.active
        
        # Ищем строку с номерами граф
        graph_row = self.find_graph_numbers_row()
        
        if graph_row is None:
            logging.error("Не удалось найти строку с номерами граф. Невозможно определить структуру сметы.")
            return False
        
        # Определяем начало данных (обычно следующая строка после номеров граф)
        self.data_start_row = graph_row + 1
        
        # Проверяем следующую строку - часто там подзаголовки
        next_row_text = ''
        for col_idx in range(1, min(20, self.ws.max_column + 1)):
            cell = self.ws.cell(row=self.data_start_row, column=col_idx)
            if cell.value is not None:
                next_row_text += str(cell.value).strip().lower() + ' '
        
        # Если следующая строка содержит подобные слова, пропускаем и ее
        if any(keyword in next_row_text for keyword in ['наименование', 'работ', 'ед. изм', 'едизм', 
                                                        'количество', 'цена', 'стоимость', 'обоснование']):
            self.data_start_row += 1
            logging.info(f"Обнаружены подзаголовки, данные начинаются со строки {self.data_start_row}")
        
        # Определяем маппинг для нормализованной таблицы
        # Основные графы для нормализованной сметы:
        # 1: № п/п
        # 2: Обоснование (расценка)
        # 3: Наименование работ
        # 4: Единица измерения
        # 6: Количество "всего"
        # 7: Цена "на ед."
        # 8: Стоимость "Всего"
        self.column_mapping = {
            'number': self.graph_to_column.get(1),
            'justification': self.graph_to_column.get(2),
            'name': self.graph_to_column.get(3),
            'unit': self.graph_to_column.get(4),
            'quantity': self.graph_to_column.get(6),
            'price': self.graph_to_column.get(7),
            'cost': self.graph_to_column.get(8),
        }
        
        logging.info(f"Структура сметы определена. Начало данных: строка {self.data_start_row}")
        logging.info(f"Маппинг колонок: {self.column_mapping}")
        
        # Отладочный вывод первых нескольких строк
        logging.debug("Отладка первых 3 строк данных:")
        for row_idx in range(self.data_start_row, min(self.data_start_row + 3, self.ws.max_row + 1)):
            row_data = []
            for col_idx in range(1, min(15, self.ws.max_column + 1)):
                cell = self.ws.cell(row=row_idx, column=col_idx)
                if cell.value is not None:
                    row_data.append(f"{get_column_letter(col_idx)}{row_idx}: {str(cell.value)[:30]}")
            logging.debug(f"Строка {row_idx}: {' | '.join(row_data)}")
        
        return True
    
    def normalize_smeta_table(self) -> List[Dict[str, Any]]:
        """Нормализует смету с поддержкой до 14 граф
        
        Если встречаются два пункта строки с одинаковыми числами в графе 1,
        то ячейки в которых текст курсивный необходимо игнорировать и не учитывать эту строку для сверки.
        """
        if not self.column_mapping or self.data_start_row is None:
            logging.error("Структура сметы не определена. Сначала вызовите determine_smeta_structure()")
            return []
        
        # Для проверки курсива нужно загрузить файл с сохранением стилей
        # Загружаем файл без data_only для доступа к стилям
        wb_with_styles = None
        ws_with_styles = None
        try:
            wb_with_styles = openpyxl.load_workbook(self.filepath, data_only=False)
            ws_with_styles = wb_with_styles.active
        except Exception as e:
            logging.warning(f"Не удалось загрузить файл с сохранением стилей для проверки курсива: {e}")
            # Продолжаем без проверки курсива
        
        # Сначала собираем все строки с их данными и информацией о курсиве
        all_rows = []
        max_rows = min(self.data_start_row + 300, self.ws.max_row + 1)  # Ограничиваем до 300 строк
        
        for row_idx in range(self.data_start_row, max_rows):
            # Пропускаем пустые строки
            is_row_empty = True
            for col_idx in range(1, min(10, self.ws.max_column + 1)):
                cell = self.ws.cell(row=row_idx, column=col_idx)
                if cell.value is not None and str(cell.value).strip() != '':
                    is_row_empty = False
                    break
            
            if is_row_empty:
                continue
            
            # Пропускаем технические строки (итоги, разделы)
            first_cell = self.ws.cell(row=row_idx, column=1).value
            if first_cell is not None:
                first_cell_str = str(first_cell).strip().lower()
                if any(marker in first_cell_str for marker in [
                    'итого', 'всего', 'раздел', 'глава', 'в том числе', 
                    'затраты труда', 'накладные расходы', 'сметная прибыль',
                    'всего по смете', 'итого по разделу'
                ]):
                    continue
            
            try:
                # Извлекаем номер позиции
                number_col = self.column_mapping['number']
                if number_col is None:
                    continue
                    
                number_cell = self.ws.cell(row=row_idx, column=number_col)
                number_val = number_cell.value
                
                if number_val is None or str(number_val).strip() == '':
                    continue
                
                # Извлекаем число из номера
                number_num = self.extract_number_from_cell(number_val)
                if number_num is None:
                    continue
                    
                number_str = str(number_val).strip()
                clean_number = re.sub(r'[^\d.]', '', number_str)
                
                if not clean_number or not re.match(r'^[\d.]+$', clean_number):
                    continue
                
                # Проверяем, является ли текст в ячейке графы 1 курсивным
                is_italic = False
                if ws_with_styles:
                    try:
                        style_cell = ws_with_styles.cell(row=row_idx, column=number_col)
                        if style_cell.font and style_cell.font.italic:
                            is_italic = True
                    except Exception as e:
                        logging.debug(f"Не удалось проверить курсив для строки {row_idx}: {e}")
                
                # Извлекаем остальные данные
                position = {
                    'row_idx': row_idx,
                    'number': number_str,
                    'clean_number': clean_number,
                    'is_italic': is_italic,
                    'justification': '',
                    'name': '',
                    'unit': '',
                    'quantity': 0.0,
                    'price': 0.0,
                    'cost': 0.0
                }
                
                # Заполняем данные в соответствии с маппингом
                for field, col_idx in self.column_mapping.items():
                    if col_idx is None or field == 'number':
                        continue
                    
                    # Извлекаем значение из ячейки
                    cell_value = self.ws.cell(row=row_idx, column=col_idx).value
                    
                    if field == 'justification':
                        position['justification'] = str(cell_value).strip() if cell_value else ''
                    elif field == 'name':
                        position['name'] = str(cell_value).strip() if cell_value else ''
                        if not position['name']:  # Пропускаем пустые наименования
                            break
                    elif field == 'unit':
                        position['unit'] = str(cell_value).strip() if cell_value else ''
                        if not position['unit']:  # Пропускаем пустые единицы измерения
                            break
                    elif field == 'quantity':
                        position['quantity'] = self._parse_float(cell_value)
                    elif field == 'price':
                        position['price'] = self._parse_float(cell_value)
                    elif field == 'cost':
                        position['cost'] = self._parse_float(cell_value)
                
                # Если успешно извлекли основные данные, добавляем в список
                if position['name'] and position['unit']:
                    # Автоматический расчет стоимости при необходимости
                    if position['cost'] == 0 and position['quantity'] > 0 and position['price'] > 0:
                        position['cost'] = position['quantity'] * position['price']
                    
                    all_rows.append(position)
                
            except Exception as e:
                logging.error(f"Ошибка при обработке строки {row_idx}: {str(e)}")
                continue
        
        # Теперь обрабатываем дубликаты: группируем по clean_number
        rows_by_number = {}
        for row in all_rows:
            clean_num = row['clean_number']
            if clean_num not in rows_by_number:
                rows_by_number[clean_num] = []
            rows_by_number[clean_num].append(row)
        for row in all_rows:
            rows_by_number[row['clean_number']].append(row)
        
        # Фильтруем дубликаты: если есть несколько строк с одинаковым номером,
        # игнорируем те, где текст курсивный
        normalized_data = []
        seen_numbers = set()
        
        for clean_number, rows in rows_by_number.items():
            if len(rows) > 1:
                # Есть дубликаты - проверяем курсив
                non_italic_rows = [r for r in rows if not r['is_italic']]
                italic_rows = [r for r in rows if r['is_italic']]
                
                if non_italic_rows:
                    # Есть строки без курсива - используем их, игнорируем курсивные
                    if italic_rows:
                        logging.debug(f"Найдены дубликаты для номера {clean_number}: {len(non_italic_rows)} без курсива, {len(italic_rows)} с курсивом. Игнорируем курсивные.")
                    for row in non_italic_rows:
                        if clean_number not in seen_numbers:
                            # Удаляем служебные поля перед добавлением
                            row_data = {k: v for k, v in row.items() if k not in ['row_idx', 'clean_number', 'is_italic']}
                            normalized_data.append(row_data)
                            seen_numbers.add(clean_number)
                else:
                    # Все строки курсивные - используем первую
                    if clean_number not in seen_numbers:
                        row_data = {k: v for k, v in rows[0].items() if k not in ['row_idx', 'clean_number', 'is_italic']}
                        normalized_data.append(row_data)
                        seen_numbers.add(clean_number)
            else:
                # Нет дубликатов - просто добавляем
                if clean_number not in seen_numbers:
                    row_data = {k: v for k, v in rows[0].items() if k not in ['row_idx', 'clean_number', 'is_italic']}
                    normalized_data.append(row_data)
                    seen_numbers.add(clean_number)
        
        # Закрываем файл со стилями, если он был открыт
        if wb_with_styles:
            try:
                wb_with_styles.close()
            except:
                pass
        
        logging.info(f"Нормализованная таблица сформирована. Всего позиций: {len(normalized_data)}")
        return normalized_data

class KS2Parser(Parser):
    """Парсер для чтения КС-2"""
    
    def __init__(self, filepath: str):
        super().__init__(filepath)
        self.positions: List[KS2Position] = []
        self.header: Optional[KS2Header] = None
    
    def parse(self) -> List[KS2Position]:
        """Парсит КС-2 и возвращает список позиций"""
        self.load()
        self.positions = []
        self.cols = {}
        # ИСПРАВЛЕНИЕ: Используем более точные ключевые слова для поиска заголовка
        # 'наименование' и 'единиц' более надежно определяют строку заголовков
        self.header_row = self.find_header_row(['наименование', 'единиц']) or 31
        if self.header_row is None:
            raise ValueError("Не найдена строка с заголовками граф в КС-2")
        
        self._find_columns()
        
        # КРИТИЧНО: Найти строку с номерами граф (1, 2, 3, 4, 5, 6, 7, 8) и использовать её для точного определения колонок
        graph_numbers_row = None
        cols_map = {}  # Словарь: номер графы → индекс колонки
        
        for check_idx in range(self.header_row, min(self.header_row + 10, len(self.df))):
            row_check = self.df.iloc[check_idx]
            
            # Проверяем первые 8 колонок
            first_eight = []
            for col in range(min(8, len(row_check))):
                val = row_check.iloc[col]
                if pd.notna(val):
                    val_str = str(val).strip()
                    # Проверяем, что это число от 1 до 8
                    try:
                        num = int(val_str)
                        if 1 <= num <= 8:
                            first_eight.append(num)
                    except (ValueError, TypeError):
                        pass
            
            # Если нашли последовательность 1, 2, 3, 4, 5, 6, 7, 8
            if len(first_eight) >= 8 and first_eight[:8] == [1, 2, 3, 4, 5, 6, 7, 8]:
                graph_numbers_row = check_idx
                
                # Создаем маппинг: номер графы → индекс колонки
                for col_idx in range(min(8, len(row_check))):
                    val = row_check.iloc[col_idx]
                    if pd.notna(val):
                        try:
                            graph_num = int(str(val).strip())
                            if 1 <= graph_num <= 8:
                                cols_map[graph_num] = col_idx
                        except (ValueError, TypeError):
                            pass
                
                print(f"[КС-2] Найдена строка с номерами граф: {check_idx}")
                print(f"[КС-2] Маппинг граф: {cols_map}")
                break
        
        # Если нашли строку с номерами граф, обновляем self.cols точными значениями
        if cols_map:
            if 1 in cols_map:
                self.cols['number'] = cols_map[1]  # Графа 1: № п/п
            if 2 in cols_map:
                self.cols['smeta_pos'] = cols_map[2]  # Графа 2: по смете
            if 3 in cols_map:
                self.cols['name'] = cols_map[3]  # Графа 3: Наименование работ
            if 4 in cols_map:
                self.cols['justification'] = cols_map[4]  # Графа 4: Номер единичной расценки
            if 5 in cols_map:
                self.cols['unit'] = cols_map[5]  # Графа 5: Единица измерения
            if 6 in cols_map:
                self.cols['qty'] = cols_map[6]  # Графа 6: Выполнено работ / количество
            if 7 in cols_map:
                self.cols['price'] = cols_map[7]  # Графа 7: Выполнено работ / цена за единицу
            if 8 in cols_map:
                self.cols['cost'] = cols_map[8]  # Графа 8: Выполнено работ / стоимость
            print(f"[КС-2] Обновленные колонки по номерам граф: {self.cols}")
        
        self.skip_graph_numbers_row()
        if self.data_start is None:
            self.data_start = self.header_row + 2
        
        # ОТЛАДОЧНЫЙ ВЫВОД: проверяем первые 5-10 строк после data_start
        # Убеждаемся, что мы читаем реальные данные, а не технические строки
        import sys
        debug_mode = '--debug-ks2' in sys.argv
        if debug_mode and self.df is not None:
            print(f"\n[DEBUG KS2Parser] header_row={self.header_row}, data_start={self.data_start}")
            print(f"[DEBUG KS2Parser] Проверка первых 10 строк после data_start:")
            for i in range(self.data_start, min(self.data_start + 10, len(self.df))):
                row = self.df.iloc[i]
                first_cols = []
                for col_idx in range(min(10, len(row))):
                    val = row.iloc[col_idx] if col_idx < len(row) else None
                    if pd.notna(val):
                        first_cols.append(str(val).strip()[:20])  # Первые 20 символов
                    else:
                        first_cols.append("")
                print(f"  Строка {i}: {' | '.join(first_cols)}")
            print()
        
        seen_numbers = set()
        for row_idx in range(self.data_start, len(self.df)):
            pos = self._parse_row(row_idx, seen_numbers)
            if pos:
                self.positions.append(pos)
        
        # КРИТИЧНО: Фильтруем результат, оставляя только объекты KS2Position
        # Это защищает от попадания строк, None или других не-объектов в список
        self.positions = [pos for pos in self.positions if isinstance(pos, KS2Position)]
        
        # Всегда выводим сводку (даже без debug_mode)
        print(f"\n[КС-2] Найдено позиций: {len(self.positions)}")
        if self.positions:
            first_pos = self.positions[0]
            print(f"[КС-2] Первая позиция: №{first_pos.number}, по смете №{first_pos.smeta_number}, name='{first_pos.name[:50]}'")
            if len(self.positions) > 1:
                second_pos = self.positions[1]
                print(f"[КС-2] Вторая позиция: №{second_pos.number}, по смете №{second_pos.smeta_number}, name='{second_pos.name[:50]}'")
        else:
            print(f"[КС-2] ВНИМАНИЕ: Позиции не найдены! Проверьте:")
            print(f"[КС-2]   - header_row={self.header_row}")
            print(f"[КС-2]   - data_start={self.data_start}")
            print(f"[КС-2]   - cols={self.cols}")
        
        return self.positions
    
    def normalize(self) -> List[NormalizedPosition]:
        """Приводит позиции КС-2 к нормализованному виду.
        
        Возвращает список нормализованных позиций, содержащих только необходимые графы:
        - Графа 1: № п/п
        - Графа 2: № по смете
        - Графа 3: Наименование работ
        - Графа 4: Номер единичной расценки (обоснование)
        - Графа 5: Единица измерения
        - Графа 6: Выполнено работ / количество
        - Графа 7: Выполнено работ / цена за единицу
        - Графа 8: Выполнено работ / стоимость
        """
        normalized = []
        for pos in self.positions:
            normalized.append(NormalizedPosition(
                number=pos.number,
                justification_or_smeta_number=pos.smeta_number,  # Для КС-2 это номер по смете
                name=pos.name,
                unit=pos.unit,
                quantity=pos.quantity,
                price=pos.price,
                cost=pos.cost,
                source="ks2"
            ))
        return normalized
    
    def _find_columns(self):
        if self.df is None or self.header_row is None:
            return
        
        for col_idx in range(self.df.shape[1]):
            val0 = self.df.iat[self.header_row - 1, col_idx] if (self.header_row - 1) >= 0 else None
            val1 = self.df.iat[self.header_row, col_idx]
            val2 = self.df.iat[self.header_row + 1, col_idx] if (self.header_row + 1) < len(self.df) else None
            val_combined = f"{str(val0).lower() if pd.notna(val0) else ''} {str(val1).lower() if pd.notna(val1) else ''} {str(val2).lower() if pd.notna(val2) else ''}"
            
            # Графа 1 КС-2: № п/п (порядковый номер)
            if ('п/п' in val_combined or 'номер' in val_combined) and 'number' not in self.cols:
                # Убеждаемся, что это не "по смете"
                if 'по смете' not in val_combined:
                    self.cols['number'] = col_idx
            # Графа 2 КС-2: по смете (номер позиции в смете)
            elif 'по смете' in val_combined and 'smeta_pos' not in self.cols:
                self.cols['smeta_pos'] = col_idx
            # Графа 3 КС-2: Наименование работ
            elif 'наименование' in val_combined and 'name' not in self.cols:
                if 'работ' in val_combined or 'name' not in self.cols:
                    self.cols['name'] = col_idx
            # Графа 4 КС-2: Номер единичной расценки
            elif ('расцен' in val_combined or 'обосн' in val_combined or 'номер' in val_combined) and 'justification' not in self.cols:
                # Проверяем, что это графа с расценками, а не общий номер
                if 'единичн' in val_combined or 'расцен' in val_combined or self.cols.get('number') is not None:
                    self.cols['justification'] = col_idx
            # Графа 5 КС-2: Единица измерения
            elif 'единиц' in val_combined and 'измерен' in val_combined and 'unit' not in self.cols:
                self.cols['unit'] = col_idx
            # Графа 6 КС-2: Выполнено работ / количество
            elif ('количество' in val_combined or 'выполнено' in val_combined) and 'qty' not in self.cols:
                self.cols['qty'] = col_idx
            # Графа 7 КС-2: Выполнено работ / цена за единицу, руб.
            elif 'цена' in val_combined and 'price' not in self.cols:
                if 'единиц' in val_combined or 'price' not in self.cols:
                    self.cols['price'] = col_idx
            # Графа 8 КС-2: Выполнено работ / стоимость, руб.
            elif 'стоимость' in val_combined and 'cost' not in self.cols:
                self.cols['cost'] = col_idx
    
    def _parse_row(self, row_idx: int, seen_numbers: set) -> Optional[KS2Position]:
        if self.df is None:
            return None
        
        # КРИТИЧНО: Пропускаем технические строки (разделители глав)
        # Методика: "могут попадаться ячейки с объединенными наименованиями и текстом
        # где могут встречаться по тексту числа. Такие строки должны игнорироваться.
        # Они обычно разделяют КС-2 по главам."
        row_check = self.df.iloc[row_idx]
        row_str = ' '.join([str(x).lower() for x in row_check if pd.notna(x)])
        
        # Пропускаем разделители глав и итоговые строки
        technical_markers = ['раздел', 'глава', 'итого по разделу', 'итого', 'всего по смете', 'всего']
        for marker in technical_markers:
            if marker in row_str:
                return None
        
        # КРИТИЧНО: Дополнительная проверка на строку с номерами граф
        # Проверяем первые 8 колонок - если они содержат последовательные числа 1, 2, 3, 4, 5, 6, 7, 8
        # это строка с номерами граф, а не реальные данные
        first_eight = []
        for col_idx in range(min(8, len(row_check))):
            val = row_check.iloc[col_idx]
            if pd.notna(val):
                val_str = str(val).strip()
                first_eight.append(val_str)
        
        # Проверяем, является ли это строкой с номерами граф
        if len(first_eight) >= 8:
            try:
                nums = [int(x) for x in first_eight[:8]]
                if nums == [1, 2, 3, 4, 5, 6, 7, 8]:
                    return None  # Это строка с номерами граф, пропускаем
            except (ValueError, TypeError):
                pass
        
        # Альтернативная проверка: первые 3 значения = "1", "2", "3" и все они - числа
        if len(first_eight) >= 3:
            try:
                if first_eight[0] == "1" and first_eight[1] == "2" and first_eight[2] == "3":
                    # Проверяем, что остальные значения тоже числа (4, 5, 6, 7, 8...)
                    all_numbers = True
                    for i, val_str in enumerate(first_eight[3:min(8, len(first_eight))]):
                        try:
                            expected_num = i + 4  # Ожидаем 4, 5, 6, 7, 8
                            if int(val_str) != expected_num:
                                all_numbers = False
                                break
                        except (ValueError, TypeError):
                            all_numbers = False
                            break
                    if all_numbers:
                        return None  # Это строка с номерами граф, пропускаем
            except (ValueError, TypeError):
                pass
        
        number_val = self.df.iat[row_idx, self.cols.get('number', 0)]
        smeta_val = self.df.iat[row_idx, self.cols.get('smeta_pos', 1)] if 'smeta_pos' in self.cols else None
        name_val = self.df.iat[row_idx, self.cols.get('name', 2)] if 'name' in self.cols else None
        qty_val = self.df.iat[row_idx, self.cols.get('qty', 5)] if 'qty' in self.cols else None
        
        if pd.isna(number_val) or pd.isna(smeta_val) or pd.isna(name_val):
            return None
        
        try:
            number = int(str(number_val).strip())
            smeta_number = int(str(smeta_val).strip())
            if number <= 0 or smeta_number <= 0:
                return None
        except (ValueError, TypeError):
            return None
        
        number_str = str(number)
        if number_str in seen_numbers:
            return None
        seen_numbers.add(number_str)
        
        quantity = safe_float(qty_val)
        if quantity is None or quantity <= 0:
            return None
        
        justification_val = self.df.iat[row_idx, self.cols.get('justification', 3)] if 'justification' in self.cols else None
        unit_val = self.df.iat[row_idx, self.cols.get('unit', 4)] if 'unit' in self.cols else None
        if pd.isna(unit_val) or str(unit_val).strip() == '':
            return None
        price_val = self.df.iat[row_idx, self.cols.get('price', 6)] if 'price' in self.cols else None
        cost_val = self.df.iat[row_idx, self.cols.get('cost', 7)] if 'cost' in self.cols else None
        
        price = safe_float(price_val)
        cost = safe_float(cost_val)
        if price is None and cost is not None and quantity:
            price = cost / quantity if quantity != 0 else None
        if cost is None and price is not None and quantity is not None:
            cost = price * quantity
        if price is None or cost is None:
            return None
        
        return KS2Position(
            number=number_str,
            smeta_number=str(smeta_number),
            name=str(name_val).strip(),
            justification=str(justification_val).strip() if pd.notna(justification_val) else '',
            unit=str(unit_val).strip() if pd.notna(unit_val) else '',
            quantity=quantity,
            price=price,
            cost=cost
        )
    
    def parse_header(self) -> KS2Header:
        """Парсит реквизиты из шапки КС-2"""
        df = pd.read_excel(self.filepath, sheet_name=0, header=None)
        header = KS2Header()
        
        for idx in range(min(30, len(df))):
            row = df.iloc[idx]
            row_str = ' '.join([str(x) for x in row if pd.notna(x)])
            # Далее оставляем существующую логику из предыдущей версии
            if 'Заказчик' in row_str:
                for val in row:
                    if pd.notna(val):
                        val_str = str(val).strip()
                        if 'Заказчик' in val_str and len(val_str) > 10:
                            header.customer = val_str.replace('Заказчик:', '').replace('Заказчик', '').strip()
                        if re.match(r'^\d{8,10}$', val_str) and not header.customer_okpo:
                            header.customer_okpo = val_str
            if 'Подрядчик' in row_str:
                for val in row:
                    if pd.notna(val):
                        val_str = str(val).strip()
                        if 'Подрядчик' in val_str and len(val_str) > 10:
                            header.contractor = val_str.replace('Подрядчик:', '').replace('Подрядчик', '').strip()
                        if re.match(r'^\d{8,10}$', val_str) and not header.contractor_okpo:
                            header.contractor_okpo = val_str
            if 'Объект' in row_str:
                for val in row:
                    if pd.notna(val):
                        val_str = str(val).strip()
                        if 'Объект' in val_str and len(val_str) > 10:
                            header.object_name = val_str.replace('Объект:', '').replace('Объект', '').strip()
            if 'Договор' in row_str or 'контракт' in row_str.lower():
                for col_idx in range(len(row)):
                    val = row.iloc[col_idx]
                    if pd.notna(val):
                        val_str = str(val).strip()
                        if 'номер' in val_str.lower() and col_idx + 1 < len(row):
                            next_val = row.iloc[col_idx + 1]
                            if pd.notna(next_val):
                                header.contract_number = str(next_val).strip()
                        if 'дата' in val_str.lower() and col_idx + 1 < len(row):
                            next_val = row.iloc[col_idx + 1]
                            if pd.notna(next_val):
                                header.contract_date = str(next_val).strip()
        
        self.header = header
        return header

class KS2Validator:
    """Валидатор для проверки КС-2 против сметы
    
    Проверяет, что физические и стоимостные показатели в КС-2
    не превышают соответствующие показатели в смете.
    """
    
    def __init__(self, ks2_positions: List[KS2Position], smeta_positions: Dict[str, SmetaPosition], ks2_header: Optional[KS2Header] = None):
        # КРИТИЧНО: Фильтруем ks2_positions, оставляя только объекты KS2Position
        # Это защищает от попадания строк, None или других не-объектов
        self.ks2_positions = [pos for pos in ks2_positions if isinstance(pos, KS2Position)]
        self.smeta_positions = smeta_positions
        self.ks2_header = ks2_header
        self.results: List[ValidationResult] = []
        self.header_results: List[ValidationResult] = []
        
        # Нормализованные данные для сравнения
        self.normalized_ks2: List[NormalizedPosition] = []
        self.normalized_smeta: Dict[str, NormalizedPosition] = {}
        
        # Объединенная таблица для сверки (третья таблица)
        self.comparison_table: List[ComparisonRow] = []
        
    def normalize_data(self):
        """Приводит обе таблицы (КС-2 и Смета) к единому нормализованному виду.
        
        Это первый этап проверки - нормализация данных перед сравнением.
        Обе таблицы приводятся к единому формату NormalizedPosition,
        содержащему только необходимые графы для сверки.
        """
        # Нормализуем позиции сметы
        print(f"[НОРМАЛИЗАЦИЯ] Начало нормализации сметы. Всего позиций в словаре: {len(self.smeta_positions)}")
        for number, pos in self.smeta_positions.items():
            # КРИТИЧНО: Ключ в normalized_smeta - это number (номер позиции из сметы)
            # Это будет использоваться для поиска по smeta_number из КС-2
            self.normalized_smeta[number] = NormalizedPosition(
                number=str(pos.number) if pos.number else "",
                justification_or_smeta_number=str(pos.justification) if pos.justification else "",
                name=str(pos.name) if pos.name else "",
                unit=str(pos.unit) if pos.unit else "",
                quantity=float(pos.quantity) if pos.quantity is not None else 0.0,
                price=float(pos.price) if pos.price is not None else 0.0,
                cost=float(pos.cost) if pos.cost is not None else 0.0,
                source="smeta"
            )
            # Отладка первых 3 позиций
            if len(self.normalized_smeta) <= 3:
                print(f"[НОРМАЛИЗАЦИЯ СМЕТА] Позиция {number}: number='{pos.number}', name='{pos.name[:50] if pos.name else ''}', qty={pos.quantity}, price={pos.price}, cost={pos.cost}")
        
        # Нормализуем позиции КС-2
        print(f"[НОРМАЛИЗАЦИЯ] Начало нормализации КС-2. Всего позиций: {len(self.ks2_positions)}")
        for pos in self.ks2_positions:
            self.normalized_ks2.append(NormalizedPosition(
                number=str(pos.number) if pos.number else "",
                justification_or_smeta_number=str(pos.smeta_number) if pos.smeta_number else "",  # Для КС-2 это номер по смете
                name=str(pos.name) if pos.name else "",
                unit=str(pos.unit) if pos.unit else "",
                quantity=float(pos.quantity) if pos.quantity is not None else 0.0,
                price=float(pos.price) if pos.price is not None else 0.0,
                cost=float(pos.cost) if pos.cost is not None else 0.0,
                source="ks2"
            ))
            # Отладка первых 3 позиций
            if len(self.normalized_ks2) <= 3:
                print(f"[НОРМАЛИЗАЦИЯ КС-2] Позиция {pos.number}: smeta_number='{pos.smeta_number}', name='{pos.name[:50] if pos.name else ''}', qty={pos.quantity}, price={pos.price}, cost={pos.cost}")
        
        print(f"[НОРМАЛИЗАЦИЯ] КС-2: {len(self.normalized_ks2)} позиций")
        print(f"[НОРМАЛИЗАЦИЯ] Смета: {len(self.normalized_smeta)} позиций")
        print(f"[НОРМАЛИЗАЦИЯ] Ключи в normalized_smeta (первые 5): {list(self.normalized_smeta.keys())[:5]}")
    
    def build_comparison_table(self) -> List[ComparisonRow]:
        """ШАГ 2: Формирует третью таблицу с перенесенными данными из двух таблиц.
        
        Объединяет данные из нормализованных таблиц КС-2 и Сметы в одну таблицу
        для удобного сравнения. Каждая строка содержит данные из обеих таблиц.
        
        Returns:
            List[ComparisonRow]: Список строк объединенной таблицы
        """
        comparison_table = []
        
        # Сортируем позиции КС-2 по номеру
        sorted_ks2 = sorted(self.normalized_ks2, key=lambda x: int(x.number) if x.number.isdigit() else 9999)
        
        for idx, ks2_norm in enumerate(sorted_ks2):
            # Находим соответствующие данные из сметы
            smeta_number = ks2_norm.justification_or_smeta_number  # Для КС-2 это номер по смете
            
            # КРИТИЧНО: Пробуем найти по точному совпадению, затем по числовому
            smeta_norm = self.normalized_smeta.get(smeta_number)
            
            # Если не найдено, пробуем найти по числовому совпадению
            if smeta_norm is None and smeta_number:
                try:
                    smeta_num_int = int(float(smeta_number))
                    for key, norm_pos in self.normalized_smeta.items():
                        try:
                            if int(float(key)) == smeta_num_int:
                                smeta_norm = norm_pos
                                if idx < 3:
                                    print(f"[ОБЪЕДИНЕННАЯ ТАБЛИЦА] Найдено по числовому совпадению: '{smeta_number}' -> '{key}'")
                                break
                        except (ValueError, TypeError):
                            continue
                except (ValueError, TypeError):
                    pass
            
            # Отладка для первых 3 позиций
            if idx < 3:
                print(f"[ОБЪЕДИНЕННАЯ ТАБЛИЦА] КС-2 позиция {ks2_norm.number}: smeta_number='{smeta_number}', найдено: {smeta_norm is not None}")
                if smeta_norm:
                    print(f"  Данные из сметы: name='{smeta_norm.name[:50] if smeta_norm.name else ''}', qty={smeta_norm.quantity}, price={smeta_norm.price}, cost={smeta_norm.cost}")
                else:
                    print(f"  Доступные ключи в normalized_smeta (первые 10): {list(self.normalized_smeta.keys())[:10]}")
            
            # Находим оригинальные объекты
            ks2_orig = next((p for p in self.ks2_positions if p.number == ks2_norm.number), None)
            smeta_orig = self.smeta_positions.get(smeta_number) if smeta_number else None
            
            # Если smeta_orig не найден, пробуем найти по числовому совпадению
            if smeta_orig is None and smeta_number:
                try:
                    smeta_num_int = int(float(smeta_number))
                    for key, pos in self.smeta_positions.items():
                        try:
                            if int(float(key)) == smeta_num_int:
                                smeta_orig = pos
                                break
                        except (ValueError, TypeError):
                            continue
                except (ValueError, TypeError):
                    pass
            
            # Если smeta_orig не найден, но smeta_norm есть, создаем из нормализованных данных
            if smeta_orig is None and smeta_norm:
                smeta_orig = SmetaPosition(
                    number=str(smeta_norm.number),
                    name=str(smeta_norm.name),
                    justification=str(smeta_norm.justification_or_smeta_number),
                    unit=str(smeta_norm.unit),
                    quantity=float(smeta_norm.quantity) if smeta_norm.quantity is not None else 0.0,
                    price=float(smeta_norm.price) if smeta_norm.price is not None else 0.0,
                    cost=float(smeta_norm.cost) if smeta_norm.cost is not None else 0.0
                )
            
            # Формируем строку объединенной таблицы
            if smeta_norm:
                # Есть данные из обеих таблиц
                row = ComparisonRow(
                    ks2_number=ks2_norm.number,
                    smeta_number=smeta_number or "",
                    ks2_name=ks2_norm.name or "",
                    ks2_justification=ks2_orig.justification if ks2_orig and ks2_orig.justification else "",
                    ks2_unit=ks2_norm.unit or "",
                    ks2_quantity=float(ks2_norm.quantity) if ks2_norm.quantity is not None else 0.0,
                    ks2_price=float(ks2_norm.price) if ks2_norm.price is not None else 0.0,
                    ks2_cost=float(ks2_norm.cost) if ks2_norm.cost is not None else 0.0,
                    smeta_name=smeta_norm.name or "",
                    smeta_justification=smeta_norm.justification_or_smeta_number or "",
                    smeta_unit=smeta_norm.unit or "",
                    smeta_quantity=float(smeta_norm.quantity) if smeta_norm.quantity is not None else 0.0,
                    smeta_price=float(smeta_norm.price) if smeta_norm.price is not None else 0.0,
                    smeta_cost=float(smeta_norm.cost) if smeta_norm.cost is not None else 0.0,
                    has_ks2=True,
                    has_smeta=True,
                    ks2_orig=ks2_orig,
                    smeta_orig=smeta_orig
                )
            else:
                # Нет данных из сметы
                row = ComparisonRow(
                    ks2_number=ks2_norm.number,
                    smeta_number=smeta_number or "",
                    ks2_name=ks2_norm.name or "",
                    ks2_justification=ks2_orig.justification if ks2_orig and ks2_orig.justification else "",
                    ks2_unit=ks2_norm.unit or "",
                    ks2_quantity=float(ks2_norm.quantity) if ks2_norm.quantity is not None else 0.0,
                    ks2_price=float(ks2_norm.price) if ks2_norm.price is not None else 0.0,
                    ks2_cost=float(ks2_norm.cost) if ks2_norm.cost is not None else 0.0,
                    smeta_name="",
                    smeta_justification="",
                    smeta_unit="",
                    smeta_quantity=0.0,
                    smeta_price=0.0,
                    smeta_cost=0.0,
                    has_ks2=True,
                    has_smeta=False,
                    ks2_orig=ks2_orig,
                    smeta_orig=None
                )
            
            comparison_table.append(row)
        
        self.comparison_table = comparison_table
        print(f"[ОБЪЕДИНЕННАЯ ТАБЛИЦА] Сформировано {len(comparison_table)} строк")
        return comparison_table
    
    def validate_comparison_table(self) -> List[ValidationResult]:
        """ШАГ 3: Производит сверку и выявляет ошибки на основе объединенной таблицы.
        
        Проверяет каждую строку объединенной таблицы и выявляет несоответствия.
        
        Returns:
            List[ValidationResult]: Список результатов проверки
        """
        results = []
        
        # Проверка дублирования работ
        self._check_duplicates_normalized()
        
        # Проверка реквизитов шапки
        if self.ks2_header:
            self._validate_header()
        
        # Проверяем каждую строку объединенной таблицы
        for row in self.comparison_table:
            if not row.has_ks2:
                continue  # Пропускаем строки без данных КС-2
            
            if not row.has_smeta:
                # Позиция не найдена в смете
                results.append(ValidationResult(
                    ks2_pos=row.ks2_orig,
                    smeta_pos=None,
                    issue_type=IssueType.NOT_IN_SMETA,
                    message=f"В КС-2 позиция №{row.ks2_number} ссылается на позицию сметы №{row.smeta_number}, которая отсутствует в смете"
                ))
                continue
            
            # Сверяем данные из обеих таблиц
            result = self._compare_comparison_row(row)
            results.append(result)
        
        # Сохраняем результаты для использования в get_summary()
        self.results = results
        
        return results + self.header_results
    
    def _compare_comparison_row(self, row: ComparisonRow) -> ValidationResult:
        """Сравнивает данные из строки объединенной таблицы.
        
        Args:
            row: Строка объединенной таблицы с данными из КС-2 и Сметы
            
        Returns:
            ValidationResult: Результат проверки
        """
        # Графа 3: Наименование (должны совпадать)
        if not self._names_match(row.smeta_name, row.ks2_name):
            return ValidationResult(
                ks2_pos=row.ks2_orig,
                smeta_pos=row.smeta_orig,
                issue_type=IssueType.NAME_MISMATCH,
                message=f"В КС-2 позиция №{row.ks2_number}: графа 3 (наименование) '{row.ks2_name}' не совпадает с графой 3 сметы '{row.smeta_name}'"
            )
        
        # Графа 4: Единица измерения (должны совпадать)
        if not self._units_match(row.smeta_unit, row.ks2_unit):
            return ValidationResult(
                ks2_pos=row.ks2_orig,
                smeta_pos=row.smeta_orig,
                issue_type=IssueType.UNIT_MISMATCH,
                message=f"В КС-2 позиция №{row.ks2_number}: графа 4 (единица измерения) '{row.ks2_unit}' не совпадает с графой 4 сметы '{row.smeta_unit}'"
            )
        
        # Графа 2 КС-2 (№ по смете) должна совпадать с Графой 2 сметы (Обоснование)
        if row.ks2_orig and row.smeta_orig:
            if row.smeta_orig.justification:
                if not row.ks2_orig.justification or not self._names_match(row.smeta_orig.justification, row.ks2_orig.justification):
                    return ValidationResult(
                        ks2_pos=row.ks2_orig,
                        smeta_pos=row.smeta_orig,
                        issue_type=IssueType.JUSTIFICATION_MISMATCH,
                        message=f"В КС-2 позиция №{row.ks2_number}: графа 4 (обоснование) '{row.ks2_orig.justification or 'не указано'}' не совпадает с графой 2 сметы '{row.smeta_orig.justification}'"
                    )
        
        # Графы 6, 7, 8: Количество, Цена, Стоимость (КС-2 не должна превышать Смету)
        tolerance_qty = 0.01
        tolerance_price = 0.01
        tolerance_cost = 0.01
        
        quantity_diff = row.ks2_quantity - row.smeta_quantity
        price_diff = row.ks2_price - row.smeta_price
        cost_diff = row.ks2_cost - row.smeta_cost
        
        # Графа 6: Количество
        if quantity_diff > tolerance_qty:
            exceed_percent = (quantity_diff / row.smeta_quantity * 100) if row.smeta_quantity > 0 else 0
            return ValidationResult(
                ks2_pos=row.ks2_orig,
                smeta_pos=row.smeta_orig,
                issue_type=IssueType.EXCEED_QUANTITY,
                message=f"В КС-2 позиция №{row.ks2_number}: графа 6 (количество) {row.ks2_quantity:.4f} превышает графу 6 сметы {row.smeta_quantity:.4f} на {quantity_diff:.4f} ({exceed_percent:.1f}%)",
                quantity_diff=quantity_diff,
                cost_diff=cost_diff,
                severity="error"
            )
        
        # Графа 7: Цена
        if price_diff > tolerance_price:
            exceed_percent = (price_diff / row.smeta_price * 100) if row.smeta_price > 0 else 0
            return ValidationResult(
                ks2_pos=row.ks2_orig,
                smeta_pos=row.smeta_orig,
                issue_type=IssueType.EXCEED_PRICE,
                message=f"В КС-2 позиция №{row.ks2_number}: графа 7 (цена) {row.ks2_price:.2f} руб. превышает графу 7 сметы {row.smeta_price:.2f} руб. на {price_diff:.2f} руб. ({exceed_percent:.1f}%)",
                price_diff=price_diff,
                quantity_diff=quantity_diff,
                cost_diff=cost_diff,
                severity="error"
            )
        
        # Графа 8: Стоимость
        if cost_diff > tolerance_cost:
            exceed_percent = (cost_diff / row.smeta_cost * 100) if row.smeta_cost > 0 else 0
            return ValidationResult(
                ks2_pos=row.ks2_orig,
                smeta_pos=row.smeta_orig,
                issue_type=IssueType.EXCEED_COST,
                message=f"В КС-2 позиция №{row.ks2_number}: графа 8 (стоимость) {row.ks2_cost:.2f} руб. превышает графу 8 сметы {row.smeta_cost:.2f} руб. на {cost_diff:.2f} руб. ({exceed_percent:.1f}%)",
                quantity_diff=quantity_diff,
                cost_diff=cost_diff,
                price_diff=price_diff,
                severity="error"
            )
        
        # Все проверки пройдены
        return ValidationResult(
            ks2_pos=row.ks2_orig,
            smeta_pos=row.smeta_orig,
            issue_type=IssueType.OK,
            message="Ошибки не выявлены",
            quantity_diff=quantity_diff,
            price_diff=price_diff,
            cost_diff=cost_diff
        )
    
    def validate(self) -> List[ValidationResult]:
        """Выполняет полную проверку КС-2 против сметы
        
        НОВЫЙ ПОДХОД (4 шага):
        1. Приводим таблицы КС-2 и Сметы к единому виду (нормализация)
        2. Формируем третью таблицу с перенесенными данными из двух таблиц
        3. Производим сверку и выявляем ошибки
        4. Все результаты проверки переносятся в интерфейс
        
        Алгоритм проверки:
        1. Система проверяет каждую строку КС-2 по порядку (по первой графе - порядковый номер)
        2. Для каждой строки КС-2 система смотрит на вторую графу - позицию по смете
        3. Находит соответствующую позицию в смете
        4. Сравнивает по установленным критериям
        5. Результаты упорядочены по номеру позиции КС-2
        
        Проверяет:
        1. Объёмы работ в столбцах 6 и 8 КС-2 совпадают с данными сметы
        2. Наименования работ в столбце 3 КС-2 идентичны позициям в смете
        3. Единицы измерения в столбце 5 КС-2 соответствуют смете
        4. Расценки и стоимость в столбцах 4 и 7 КС-2 корректны
        5. В КС-2 нет работ, не предусмотренных сметой
        6. Реквизиты в шапке формы совпадают
        7. Нет дублирования работ
        """
        self.results = []
        self.header_results = []
        self.comparison_table = []
        
        # КРИТИЧНО: Дополнительная проверка типов перед валидацией
        # Фильтруем все не-объекты KS2Position (на случай, если они попали после инициализации)
        self.ks2_positions = [pos for pos in self.ks2_positions if isinstance(pos, KS2Position)]
        
        if not self.ks2_positions:
            # Если после фильтрации не осталось позиций, возвращаем пустой результат
            return []
        
        # ШАГ 1: Нормализация данных - приводим обе таблицы к единому виду
        print("[ШАГ 1] Нормализация данных...")
        self.normalize_data()
        
        # ШАГ 2: Формируем третью таблицу с перенесенными данными из двух таблиц
        print("[ШАГ 2] Формирование объединенной таблицы...")
        self.build_comparison_table()
        
        # ШАГ 3: Производим сверку и выявляем ошибки
        print("[ШАГ 3] Сверка данных и выявление ошибок...")
        results = self.validate_comparison_table()
        
        # ШАГ 4: Результаты уже готовы для переноса в интерфейс
        print(f"[ШАГ 4] Результаты готовы: {len(results)} записей")
        return results
    
    def normalize_positions_unified(self) -> List[UnifiedNormalizedPosition]:
        """НОВЫЙ ПОДХОД: Создает нормализованные позиции для единого представления КС-2 и Сметы
        
        Сразу объединяет данные из КС-2 и Сметы в одну структуру.
        Это более простой и эффективный подход.
        
        Returns:
            Список нормализованных позиций
        """
        normalized = []
        
        print(f"[НОРМАЛИЗАЦИЯ УНИФИЦИРОВАННАЯ] Начало. КС-2: {len(self.ks2_positions)}, Смета: {len(self.smeta_positions)}")
        
        for ks2_pos in self.ks2_positions:
            # Ключ для поиска в смете - номер позиции из графы 2 КС-2
            smeta_key = str(ks2_pos.smeta_number).strip() if ks2_pos.smeta_number else ""
            
            # Пробуем найти по точному совпадению
            smeta_pos = self.smeta_positions.get(smeta_key)
            
            # Если не найдено, пробуем найти по числовому совпадению
            if smeta_pos is None and smeta_key:
                try:
                    smeta_num_int = int(float(smeta_key))
                    for key, pos in self.smeta_positions.items():
                        try:
                            if int(float(key)) == smeta_num_int:
                                smeta_pos = pos
                                break
                        except (ValueError, TypeError):
                            continue
                except (ValueError, TypeError):
                    pass
            
            # Создаем нормализованную позицию с данными КС-2
            norm_pos = UnifiedNormalizedPosition(
                # Идентификаторы
                ks2_number=str(ks2_pos.number) if ks2_pos.number else "",
                smeta_ref=str(ks2_pos.smeta_number) if ks2_pos.smeta_number else "",
                
                # Данные КС-2
                ks2_name=str(ks2_pos.name) if ks2_pos.name else "",
                ks2_justification=str(ks2_pos.justification) if ks2_pos.justification else "",
                ks2_unit=str(ks2_pos.unit) if ks2_pos.unit else "",
                ks2_quantity=float(ks2_pos.quantity) if ks2_pos.quantity is not None else 0.0,
                ks2_price=float(ks2_pos.price) if ks2_pos.price is not None else 0.0,
                ks2_cost=float(ks2_pos.cost) if ks2_pos.cost is not None else 0.0,
            )
            
            # Если нашли соответствующую позицию в смете
            if smeta_pos:
                # Заполняем данные сметы
                norm_pos.smeta_number = str(smeta_pos.number) if smeta_pos.number else None
                norm_pos.smeta_name = str(smeta_pos.name) if smeta_pos.name else None
                norm_pos.smeta_justification = str(smeta_pos.justification) if smeta_pos.justification else None
                norm_pos.smeta_unit = str(smeta_pos.unit) if smeta_pos.unit else None
                norm_pos.smeta_quantity = float(smeta_pos.quantity) if smeta_pos.quantity is not None else 0.0
                norm_pos.smeta_price = float(smeta_pos.price) if smeta_pos.price is not None else 0.0
                norm_pos.smeta_cost = float(smeta_pos.cost) if smeta_pos.cost is not None else 0.0
                
                # Вычисляем остатки (Смета - КС-2)
                norm_pos.quantity_remainder = norm_pos.smeta_quantity - norm_pos.ks2_quantity
                norm_pos.price_remainder = norm_pos.smeta_price - norm_pos.ks2_price
                norm_pos.cost_remainder = norm_pos.smeta_cost - norm_pos.ks2_cost
                
                # Проверяем соответствие
                issues = []
                
                # Проверка наименования
                if norm_pos.ks2_name.strip() != norm_pos.smeta_name.strip():
                    issues.append(f"Наименование не совпадает")
                
                # Проверка единицы измерения
                if norm_pos.ks2_unit.strip() != norm_pos.smeta_unit.strip():
                    issues.append(f"Единица измерения не совпадает")
                
                # Проверка превышения количества
                if norm_pos.quantity_remainder < -0.01:  # КС-2 > Смета
                    issues.append(f"Превышение количества на {abs(norm_pos.quantity_remainder):.4f}")
                
                # Проверка превышения цены
                if norm_pos.price_remainder < -0.01:  # КС-2 > Смета
                    issues.append(f"Превышение цены на {abs(norm_pos.price_remainder):.2f} руб.")
                
                # Проверка превышения стоимости
                if norm_pos.cost_remainder < -0.01:  # КС-2 > Смета
                    issues.append(f"Превышение стоимости на {abs(norm_pos.cost_remainder):.2f} руб.")
                
                norm_pos.issues = issues
                norm_pos.status = "Есть проблемы" if issues else "Соответствует"
            else:
                # Позиция не найдена в смете
                norm_pos.status = "Не найдено в смете"
                norm_pos.issues = [f"Позиция №{smeta_key} отсутствует в смете"]
            
            normalized.append(norm_pos)
        
        print(f"[НОРМАЛИЗАЦИЯ УНИФИЦИРОВАННАЯ] Создано {len(normalized)} позиций")
        return normalized
    
    def get_test_tables(self) -> Dict:
        """Возвращает три таблицы для тестирования:
        1. Нормализованная таблица КС-2
        2. Нормализованная таблица Сметы
        3. Объединенная таблица (НОВЫЙ ПОДХОД - унифицированная нормализация)
        
        Returns:
            Dict с тремя таблицами в формате списков словарей
        """
        print(f"[ТЕСТОВЫЕ ТАБЛИЦЫ] Начало формирования. Позиций КС-2: {len(self.ks2_positions)}, позиций Сметы: {len(self.smeta_positions)}")
        
        # НОВЫЙ ПОДХОД: Используем унифицированную нормализацию
        unified_normalized = self.normalize_positions_unified()
        
        # Формируем три таблицы из унифицированных данных
        # Таблица 1: Нормализованная КС-2 (извлекаем только данные КС-2)
        normalized_ks2_table = []
        for pos in unified_normalized:
            normalized_ks2_table.append({
                'number': pos.ks2_number,
                'justification_or_smeta_number': pos.smeta_ref,
                'name': pos.ks2_name,
                'unit': pos.ks2_unit,
                'quantity': float(pos.ks2_quantity),
                'price': float(pos.ks2_price),
                'cost': float(pos.ks2_cost),
                'source': 'ks2'
            })
        
        # Таблица 2: Нормализованная Смета (извлекаем только данные Сметы, где они есть)
        normalized_smeta_table = []
        seen_smeta_numbers = set()
        for pos in unified_normalized:
            if pos.smeta_number and pos.smeta_number not in seen_smeta_numbers:
                normalized_smeta_table.append({
                    'number': pos.smeta_number,
                    'justification_or_smeta_number': pos.smeta_justification or '',
                    'name': pos.smeta_name or '',
                    'unit': pos.smeta_unit or '',
                    'quantity': float(pos.smeta_quantity),
                    'price': float(pos.smeta_price),
                    'cost': float(pos.smeta_cost),
                    'source': 'smeta'
                })
                seen_smeta_numbers.add(pos.smeta_number)
        
        # Также добавляем все позиции из сметы, которые не были сопоставлены
        for number, smeta_pos in self.smeta_positions.items():
            if number not in seen_smeta_numbers:
                normalized_smeta_table.append({
                    'number': str(smeta_pos.number) if smeta_pos.number else '',
                    'justification_or_smeta_number': str(smeta_pos.justification) if smeta_pos.justification else '',
                    'name': str(smeta_pos.name) if smeta_pos.name else '',
                    'unit': str(smeta_pos.unit) if smeta_pos.unit else '',
                    'quantity': float(smeta_pos.quantity) if smeta_pos.quantity is not None else 0.0,
                    'price': float(smeta_pos.price) if smeta_pos.price is not None else 0.0,
                    'cost': float(smeta_pos.cost) if smeta_pos.cost is not None else 0.0,
                    'source': 'smeta'
                })
        
        # Сортируем таблицу сметы по номеру
        normalized_smeta_table.sort(key=lambda x: int(x['number']) if x['number'].isdigit() else 9999)
        
        # Таблица 3: Объединенная таблица (полные данные из унифицированных позиций)
        comparison_table = []
        print(f"[ТЕСТОВЫЕ ТАБЛИЦЫ] Формирование объединенной таблицы из {len(unified_normalized)} унифицированных позиций...")
        for idx, pos in enumerate(unified_normalized):
            row_data = {
                'ks2_number': pos.ks2_number,
                'smeta_number': pos.smeta_ref,
                'ks2_name': pos.ks2_name,
                'ks2_justification': pos.ks2_justification,
                'ks2_unit': pos.ks2_unit,
                'ks2_quantity': float(pos.ks2_quantity),
                'ks2_price': float(pos.ks2_price),
                'ks2_cost': float(pos.ks2_cost),
                'smeta_name': pos.smeta_name or '',
                'smeta_justification': pos.smeta_justification or '',
                'smeta_unit': pos.smeta_unit or '',
                'smeta_quantity': float(pos.smeta_quantity),
                'smeta_price': float(pos.smeta_price),
                'smeta_cost': float(pos.smeta_cost),
                'has_ks2': True,
                'has_smeta': pos.smeta_number is not None
            }
            comparison_table.append(row_data)
            # Отладка первых 3 записей
            if idx < 3:
                print(f"[ТЕСТОВЫЕ ТАБЛИЦЫ] Объединенная таблица строка {idx + 1}: has_smeta={row_data['has_smeta']}, smeta_name='{row_data['smeta_name'][:50] if row_data['smeta_name'] else ''}', smeta_quantity={row_data['smeta_quantity']}, smeta_price={row_data['smeta_price']}, smeta_cost={row_data['smeta_cost']}")
        
        print(f"[ТЕСТОВЫЕ ТАБЛИЦЫ] Все таблицы сформированы:")
        print(f"  - Нормализованная КС-2: {len(normalized_ks2_table)} записей")
        print(f"  - Нормализованная Смета: {len(normalized_smeta_table)} записей")
        print(f"  - Объединенная таблица: {len(comparison_table)} записей")
        
        # Проверяем, что данные есть
        if len(normalized_smeta_table) == 0:
            print(f"[ОШИБКА] Таблица нормализованной Сметы пустая!")
            print(f"  - self.smeta_positions содержит: {len(self.smeta_positions)} записей")
            if len(self.smeta_positions) > 0:
                print(f"  - Первые 3 ключа в smeta_positions: {list(self.smeta_positions.keys())[:3]}")
        
        return {
            'normalized_ks2': normalized_ks2_table,
            'normalized_smeta': normalized_smeta_table,
            'comparison_table': comparison_table
        }
        
        # Таблица 1: Нормализованная КС-2
        normalized_ks2_table = []
        for norm_pos in sorted(self.normalized_ks2, key=lambda x: int(x.number) if x.number.isdigit() else 9999):
            normalized_ks2_table.append({
                'number': norm_pos.number,
                'justification_or_smeta_number': norm_pos.justification_or_smeta_number,
                'name': norm_pos.name,
                'unit': norm_pos.unit,
                'quantity': float(norm_pos.quantity) if norm_pos.quantity is not None else 0.0,
                'price': float(norm_pos.price) if norm_pos.price is not None else 0.0,
                'cost': float(norm_pos.cost) if norm_pos.cost is not None else 0.0,
                'source': norm_pos.source
            })
        
        # Таблица 2: Нормализованная Смета
        normalized_smeta_table = []
        print(f"[ТЕСТОВЫЕ ТАБЛИЦЫ] Формирование таблицы нормализованной Сметы.")
        print(f"  - self.normalized_smeta содержит: {len(self.normalized_smeta)} записей")
        print(f"  - Тип self.normalized_smeta: {type(self.normalized_smeta)}")
        
        if len(self.normalized_smeta) == 0:
            print(f"[ОШИБКА] self.normalized_smeta пустой! Проверяем self.smeta_positions...")
            print(f"  - self.smeta_positions содержит: {len(self.smeta_positions)} записей")
            if len(self.smeta_positions) > 0:
                print(f"  - Первые 3 ключа в smeta_positions: {list(self.smeta_positions.keys())[:3]}")
                first_key = list(self.smeta_positions.keys())[0]
                first_pos = self.smeta_positions[first_key]
                print(f"  - Первая позиция (ключ '{first_key}'): number='{first_pos.number}', name='{first_pos.name[:50] if first_pos.name else ''}'")
        
        # Сортируем и формируем таблицу
        try:
            sorted_items = sorted(self.normalized_smeta.items(), key=lambda x: int(x[0]) if str(x[0]).isdigit() else 9999)
        except Exception as e:
            print(f"[ОШИБКА] Ошибка при сортировке: {e}")
            sorted_items = list(self.normalized_smeta.items())
        
        print(f"[ТЕСТОВЫЕ ТАБЛИЦЫ] Обработка {len(sorted_items)} записей...")
        for idx, (number, norm_pos) in enumerate(sorted_items):
            try:
                row_data = {
                    'number': str(norm_pos.number) if norm_pos.number is not None else '',
                    'justification_or_smeta_number': str(norm_pos.justification_or_smeta_number) if norm_pos.justification_or_smeta_number is not None else '',
                    'name': str(norm_pos.name) if norm_pos.name is not None else '',
                    'unit': str(norm_pos.unit) if norm_pos.unit is not None else '',
                    'quantity': float(norm_pos.quantity) if norm_pos.quantity is not None and norm_pos.quantity != '' else 0.0,
                    'price': float(norm_pos.price) if norm_pos.price is not None and norm_pos.price != '' else 0.0,
                    'cost': float(norm_pos.cost) if norm_pos.cost is not None and norm_pos.cost != '' else 0.0,
                    'source': str(norm_pos.source) if norm_pos.source else 'smeta'
                }
                normalized_smeta_table.append(row_data)
                # Отладка первых 3 записей
                if idx < 3:
                    print(f"[ТЕСТОВЫЕ ТАБЛИЦЫ] Смета запись {idx + 1}: number='{row_data['number']}', name='{row_data['name'][:50] if row_data['name'] else ''}', qty={row_data['quantity']}, price={row_data['price']}, cost={row_data['cost']}")
            except Exception as e:
                print(f"[ОШИБКА] Ошибка при обработке записи {number}: {e}")
                print(f"  - norm_pos: {norm_pos}")
                continue
        
        print(f"[ТЕСТОВЫЕ ТАБЛИЦЫ] Таблица нормализованной Сметы сформирована. Записей: {len(normalized_smeta_table)}")
        
        # Таблица 3: Объединенная таблица
        comparison_table = []
        print(f"[ТЕСТОВЫЕ ТАБЛИЦЫ] Формирование объединенной таблицы. Всего строк: {len(self.comparison_table)}")
        for idx, row in enumerate(self.comparison_table):
            row_data = {
                'ks2_number': str(row.ks2_number) if row.ks2_number else '',
                'smeta_number': str(row.smeta_number) if row.smeta_number else '',
                'ks2_name': str(row.ks2_name) if row.ks2_name else '',
                'ks2_justification': str(row.ks2_justification) if row.ks2_justification else '',
                'ks2_unit': str(row.ks2_unit) if row.ks2_unit else '',
                'ks2_quantity': float(row.ks2_quantity),
                'ks2_price': float(row.ks2_price),
                'ks2_cost': float(row.ks2_cost),
                'smeta_name': str(row.smeta_name) if row.smeta_name else '',
                'smeta_justification': str(row.smeta_justification) if row.smeta_justification else '',
                'smeta_unit': str(row.smeta_unit) if row.smeta_unit else '',
                'smeta_quantity': float(row.smeta_quantity),
                'smeta_price': float(row.smeta_price),
                'smeta_cost': float(row.smeta_cost),
                'has_ks2': row.has_ks2,
                'has_smeta': row.has_smeta
            }
            comparison_table.append(row_data)
            # Отладка первых 3 записей
            if idx < 3:
                print(f"[ТЕСТОВЫЕ ТАБЛИЦЫ] Объединенная таблица строка {idx + 1}: has_smeta={row.has_smeta}, smeta_name='{row.smeta_name}', smeta_quantity={row.smeta_quantity}, smeta_price={row.smeta_price}, smeta_cost={row.smeta_cost}")
        
        print(f"[ТЕСТОВЫЕ ТАБЛИЦЫ] Все таблицы сформированы:")
        print(f"  - Нормализованная КС-2: {len(normalized_ks2_table)} записей")
        print(f"  - Нормализованная Смета: {len(normalized_smeta_table)} записей")
        print(f"  - Объединенная таблица: {len(comparison_table)} записей")
        
        # Проверяем, что данные есть
        if len(normalized_smeta_table) == 0:
            print(f"[ОШИБКА] Таблица нормализованной Сметы пустая!")
            print(f"  - self.normalized_smeta содержит: {len(self.normalized_smeta)} записей")
            print(f"  - self.smeta_positions содержит: {len(self.smeta_positions)} записей")
            if len(self.normalized_smeta) > 0:
                print(f"  - Первые 3 ключа в normalized_smeta: {list(self.normalized_smeta.keys())[:3]}")
                first_item = list(self.normalized_smeta.items())[0]
                print(f"  - Первая запись (ключ '{first_item[0]}'): {first_item[1]}")
        
        return {
            'normalized_ks2': normalized_ks2_table,
            'normalized_smeta': normalized_smeta_table,
            'comparison_table': comparison_table
        }
    
    def _validate_normalized(self) -> List[ValidationResult]:
        """Выполняет сверку нормализованных данных КС-2 и Сметы.
        
        Обе таблицы уже приведены к единому виду (NormalizedPosition),
        содержат только необходимые графы для сравнения.
        """
        results = []
        
        # Проверка дублирования работ (на основе нормализованных данных)
        self._check_duplicates_normalized()
        
        # Проверка реквизитов шапки (если доступны)
        if self.ks2_header:
            self._validate_header()
        
        # Сортируем нормализованные позиции КС-2 по номеру
        sorted_ks2 = sorted(self.normalized_ks2, key=lambda x: int(x.number) if x.number.isdigit() else 9999)
        
        # Проверяем каждую позицию КС-2
        for ks2_norm in sorted_ks2:
            # Находим соответствующую позицию в смете по номеру (графа 2 КС-2 = графа 1 сметы)
            smeta_number = ks2_norm.justification_or_smeta_number  # Для КС-2 это номер по смете
            smeta_norm = self.normalized_smeta.get(smeta_number)
            
            # Находим оригинальные позиции для создания ValidationResult
            ks2_orig = next((p for p in self.ks2_positions if p.number == ks2_norm.number), None)
            
            # Пробуем найти smeta_orig в словаре по номеру
            smeta_orig = None
            if smeta_number:
                # Пробуем найти по точному совпадению
                smeta_orig = self.smeta_positions.get(smeta_number)
                # Если не найдено, пробуем найти по числовому совпадению (на случай "1" vs "1.0")
                if smeta_orig is None:
                    try:
                        smeta_num_int = int(float(smeta_number))
                        for key, pos in self.smeta_positions.items():
                            try:
                                if int(float(key)) == smeta_num_int:
                                    smeta_orig = pos
                                    break
                            except (ValueError, TypeError):
                                continue
                    except (ValueError, TypeError):
                        pass
            
            # Отладка для первых 3 позиций
            if len(results) < 3:
                print(f"[ОТЛАДКА ВАЛИДАЦИЯ] Позиция КС-2 №{ks2_norm.number}:")
                print(f"  smeta_number (из КС-2): '{smeta_number}' (type: {type(smeta_number)})")
                print(f"  smeta_norm найден: {smeta_norm is not None}")
                print(f"  smeta_orig найден: {smeta_orig is not None}")
                if smeta_orig:
                    print(f"  smeta_orig.number: '{smeta_orig.number}'")
                    print(f"  smeta_orig.quantity: {smeta_orig.quantity}")
                    print(f"  smeta_orig.price: {smeta_orig.price}")
                    print(f"  smeta_orig.cost: {smeta_orig.cost}")
                elif smeta_norm:
                    print(f"  smeta_norm.number: '{smeta_norm.number}'")
                    print(f"  smeta_norm.quantity: {smeta_norm.quantity}")
                    print(f"  smeta_norm.price: {smeta_norm.price}")
                    print(f"  smeta_norm.cost: {smeta_norm.cost}")
                print(f"  Доступные ключи в smeta_positions (первые 5): {list(self.smeta_positions.keys())[:5]}")
            
            if smeta_norm is None:
                # Позиция не найдена в смете
                results.append(ValidationResult(
                    ks2_pos=ks2_orig,
                    smeta_pos=None,
                    issue_type=IssueType.NOT_IN_SMETA,
                    message=f"В КС-2 позиция №{ks2_norm.number} ссылается на позицию сметы №{smeta_number}, которая отсутствует в смете"
                ))
                continue
            
            # КРИТИЧНО: Если smeta_orig не найден в словаре, но smeta_norm есть, создаем SmetaPosition из нормализованных данных
            # Это гарантирует, что данные из сметы всегда будут доступны для отображения
            if smeta_orig is None and smeta_norm:
                smeta_orig = SmetaPosition(
                    number=str(smeta_norm.number),
                    name=str(smeta_norm.name),
                    justification=str(smeta_norm.justification_or_smeta_number),
                    unit=str(smeta_norm.unit),
                    quantity=float(smeta_norm.quantity) if smeta_norm.quantity is not None else 0.0,
                    price=float(smeta_norm.price) if smeta_norm.price is not None else 0.0,
                    cost=float(smeta_norm.cost) if smeta_norm.cost is not None else 0.0
                )
                if len(results) < 3:
                    print(f"  [СОЗДАН ИЗ НОРМАЛИЗОВАННЫХ] smeta_orig.quantity: {smeta_orig.quantity}, price: {smeta_orig.price}, cost: {smeta_orig.cost}")
            
            # КРИТИЧНО: Убеждаемся, что smeta_orig всегда существует, если smeta_norm найден
            if smeta_orig is None and smeta_norm:
                # Это не должно произойти, но на всякий случай создаем из нормализованных данных
                smeta_orig = SmetaPosition(
                    number=str(smeta_norm.number),
                    name=str(smeta_norm.name),
                    justification=str(smeta_norm.justification_or_smeta_number),
                    unit=str(smeta_norm.unit),
                    quantity=float(smeta_norm.quantity) if smeta_norm.quantity is not None else 0.0,
                    price=float(smeta_norm.price) if smeta_norm.price is not None else 0.0,
                    cost=float(smeta_norm.cost) if smeta_norm.cost is not None else 0.0
                )
            
            # Сверка нормализованных данных
            result = self._compare_normalized(ks2_norm, smeta_norm, ks2_orig, smeta_orig)
            results.append(result)
        
        # КРИТИЧНО: Сохраняем результаты в self.results для использования в get_summary()
        self.results = results
        
        # Возвращаем результаты валидации + результаты проверки заголовка
        return results + self.header_results
    
    def _check_duplicates_normalized(self):
        """Проверка дублирования работ на основе нормализованных данных."""
        seen_positions = {}
        for ks2_norm in self.normalized_ks2:
            # Ключ: номер по смете + начало наименования
            key = (ks2_norm.justification_or_smeta_number, ks2_norm.name[:50])
            if key in seen_positions:
                # Найдено дублирование
                first_pos = seen_positions[key]
                ks2_orig = next((p for p in self.ks2_positions if p.number == ks2_norm.number), None)
                first_orig = next((p for p in self.ks2_positions if p.number == first_pos.number), None)
                
                result = ValidationResult(
                    ks2_pos=ks2_orig,
                    smeta_pos=None,
                    issue_type=IssueType.DUPLICATE,
                    message=f"В КС-2 позиция №{ks2_norm.number} дублирует позицию №{first_pos.number}: обе ссылаются на смету №{ks2_norm.justification_or_smeta_number} '{ks2_norm.name[:50]}...'"
                )
                self.header_results.append(result)
            else:
                seen_positions[key] = ks2_norm
    
    def _compare_normalized(self, ks2_norm: NormalizedPosition, smeta_norm: NormalizedPosition,
                           ks2_orig: Optional[KS2Position], smeta_orig: Optional[SmetaPosition]) -> ValidationResult:
        """Сравнивает две нормализованные позиции.
        
        Обе позиции уже в едином формате, содержат только необходимые графы.
        """
        # Графа 3: Наименование (должны совпадать)
        if not self._names_match(smeta_norm.name, ks2_norm.name):
            return ValidationResult(
                ks2_pos=ks2_orig,
                smeta_pos=smeta_orig,
                issue_type=IssueType.NAME_MISMATCH,
                message=f"В КС-2 позиция №{ks2_norm.number}: графа 3 (наименование) '{ks2_norm.name}' не совпадает с графой 3 сметы '{smeta_norm.name}'"
            )
        
        # Графа 4: Единица измерения (должны совпадать)
        if not self._units_match(smeta_norm.unit, ks2_norm.unit):
            return ValidationResult(
                ks2_pos=ks2_orig,
                smeta_pos=smeta_orig,
                issue_type=IssueType.UNIT_MISMATCH,
                message=f"В КС-2 позиция №{ks2_norm.number}: графа 4 (единица измерения) '{ks2_norm.unit}' не совпадает с графой 4 сметы '{smeta_norm.unit}'"
            )
        
        # Графа 2 КС-2 (№ по смете) должна совпадать с Графой 2 сметы (Обоснование)
        # Для КС-2 justification_or_smeta_number = номер по смете
        # Для сметы justification_or_smeta_number = обоснование
        # Проверяем, что обоснование в КС-2 совпадает с обоснованием в смете
        if ks2_orig and smeta_orig:
            if smeta_orig.justification:
                if not ks2_orig.justification or not self._names_match(smeta_orig.justification, ks2_orig.justification):
                    return ValidationResult(
                        ks2_pos=ks2_orig,
                        smeta_pos=smeta_orig,
                        issue_type=IssueType.JUSTIFICATION_MISMATCH,
                        message=f"В КС-2 позиция №{ks2_norm.number}: графа 4 (обоснование) '{ks2_orig.justification or 'не указано'}' не совпадает с графой 2 сметы '{smeta_orig.justification}'"
                    )
        
        # Графы 6, 7, 8: Количество, Цена, Стоимость (КС-2 не должна превышать Смету)
        tolerance_qty = 0.01
        tolerance_price = 0.01
        tolerance_cost = 0.01
        
        quantity_diff = ks2_norm.quantity - smeta_norm.quantity
        price_diff = ks2_norm.price - smeta_norm.price
        cost_diff = ks2_norm.cost - smeta_norm.cost
        
        # Графа 6: Количество
        if quantity_diff > tolerance_qty:
            exceed_percent = (quantity_diff / smeta_norm.quantity * 100) if smeta_norm.quantity > 0 else 0
            return ValidationResult(
                ks2_pos=ks2_orig,
                smeta_pos=smeta_orig,
                issue_type=IssueType.EXCEED_QUANTITY,
                message=f"В КС-2 позиция №{ks2_norm.number}: графа 6 (количество) {ks2_norm.quantity:.4f} превышает графу 6 сметы {smeta_norm.quantity:.4f} на {quantity_diff:.4f} ({exceed_percent:.1f}%)",
                quantity_diff=quantity_diff,
                cost_diff=cost_diff,
                severity="error"
            )
        
        # Графа 7: Цена
        if price_diff > tolerance_price:
            exceed_percent = (price_diff / smeta_norm.price * 100) if smeta_norm.price > 0 else 0
            return ValidationResult(
                ks2_pos=ks2_orig,
                smeta_pos=smeta_orig,
                issue_type=IssueType.EXCEED_PRICE,
                message=f"В КС-2 позиция №{ks2_norm.number}: графа 7 (цена) {ks2_norm.price:.2f} руб. превышает графу 7 сметы {smeta_norm.price:.2f} руб. на {price_diff:.2f} руб. ({exceed_percent:.1f}%)",
                price_diff=price_diff,
                quantity_diff=quantity_diff,
                cost_diff=cost_diff,
                severity="error"
            )
        
        # Графа 8: Стоимость
        if cost_diff > tolerance_cost:
            exceed_percent = (cost_diff / smeta_norm.cost * 100) if smeta_norm.cost > 0 else 0
            return ValidationResult(
                ks2_pos=ks2_orig,
                smeta_pos=smeta_orig,
                issue_type=IssueType.EXCEED_COST,
                message=f"В КС-2 позиция №{ks2_norm.number}: графа 8 (стоимость) {ks2_norm.cost:.2f} руб. превышает графу 8 сметы {smeta_norm.cost:.2f} руб. на {cost_diff:.2f} руб. ({exceed_percent:.1f}%)",
                quantity_diff=quantity_diff,
                cost_diff=cost_diff,
                price_diff=price_diff,
                severity="error"
            )
        
        # Все проверки пройдены
        return ValidationResult(
            ks2_pos=ks2_orig,
            smeta_pos=smeta_orig,
            issue_type=IssueType.OK,
            message="Ошибки не выявлены",
            quantity_diff=quantity_diff,
            price_diff=price_diff,
            cost_diff=cost_diff
        )
    
    def _validate_position(self, ks2_pos: KS2Position) -> ValidationResult:
        """
        Проверяет одну позицию КС-2 согласно методике.
        Смета является эталоном (нормативом), с ней сверяются значения КС-2.
        
        Правила проверки (согласно методике):
        - Графа 2 КС-2 ("по смете") ↔ Графа 1 сметы ("№ п/п") - поиск позиции
        - Графа 3 КС-2 ("Наименование работ") ↔ Графа 3 сметы ("Наименование работ и затрат")
        - Графа 4 КС-2 ("Номер единичной расценки") ↔ Графа 2 сметы ("Обоснование")
        - Графа 5 КС-2 ("Единица измерения") ↔ Графа 4 сметы ("Единица измерения")
        - Графа 6 КС-2 ("Выполнено работ / количество") НЕ ДОЛЖНА ПРЕВЫШАТЬ Графу 6 сметы ("Количество / всего")
        - Графа 7 КС-2 ("цена за единицу, руб.") НЕ ДОЛЖНА ПРЕВЫШАТЬ Графу 7 сметы ("на ед.")
        - Графа 8 КС-2 ("стоимость, руб.") НЕ ДОЛЖНА ПРЕВЫШАТЬ Графу 8 сметы ("Всего")
        """
        smeta_pos = self.smeta_positions.get(ks2_pos.smeta_number)
        if smeta_pos is None:
            return ValidationResult(
                ks2_pos=ks2_pos,
                smeta_pos=None,
                issue_type=IssueType.NOT_IN_SMETA,
                message=f"В КС-2 позиция №{ks2_pos.number} ссылается на позицию сметы №{ks2_pos.smeta_number}, которая отсутствует в смете (работа не предусмотрена сметой)"
            )
        
        # Графа 3 КС-2 должна совпадать с графой 3 сметы (наименование)
        if not self._names_match(smeta_pos.name, ks2_pos.name):
            return ValidationResult(
                ks2_pos=ks2_pos,
                smeta_pos=smeta_pos,
                issue_type=IssueType.NAME_MISMATCH,
                message=f"В КС-2 позиция №{ks2_pos.number}: графа 3 (наименование) '{ks2_pos.name}' не совпадает с графой 3 сметы '{smeta_pos.name}'"
            )
        
        # Графа 4 КС-2 должна совпадать с графой 2 сметы (обоснование)
        if smeta_pos.justification:
            if not ks2_pos.justification or not self._names_match(smeta_pos.justification, ks2_pos.justification):
                return ValidationResult(
                    ks2_pos=ks2_pos,
                    smeta_pos=smeta_pos,
                    issue_type=IssueType.JUSTIFICATION_MISMATCH,
                    message=f"В КС-2 позиция №{ks2_pos.number}: графа 4 (обоснование) '{ks2_pos.justification or 'не указано'}' не совпадает с графой 2 сметы '{smeta_pos.justification}'"
                )
        
        # Графа 5 КС-2 должна совпадать с графой 4 сметы (единица измерения)
        if not self._units_match(smeta_pos.unit, ks2_pos.unit):
            return ValidationResult(
                ks2_pos=ks2_pos,
                smeta_pos=smeta_pos,
                issue_type=IssueType.UNIT_MISMATCH,
                message=f"В КС-2 позиция №{ks2_pos.number}: графа 5 (единица измерения) '{ks2_pos.unit}' не совпадает с графой 4 сметы '{smeta_pos.unit}'"
            )
        
        # Графы 6, 7, 8 КС-2 не должны превышать графы 6, 7, 8 сметы (количество, цена, стоимость)
        # Методика: "Число указанное КС-2 не должна превышать число в Смете"
        # Смета является эталоном, значения КС-2 не должны превышать значения сметы
        tolerance_qty = 0.01
        tolerance_price = 0.01
        tolerance_cost = 0.01
        quantity_diff = ks2_pos.quantity - smeta_pos.quantity
        price_diff = ks2_pos.price - smeta_pos.price
        cost_diff = ks2_pos.cost - smeta_pos.cost
        warnings: List[str] = []
        
        # Графа 6 КС-2 ("Выполнено работ / количество") не должна превышать Графу 6 сметы ("Количество / всего")
        if quantity_diff > tolerance_qty:
            exceed_percent = (quantity_diff / smeta_pos.quantity * 100) if smeta_pos.quantity > 0 else 0
            return ValidationResult(
                ks2_pos=ks2_pos,
                smeta_pos=smeta_pos,
                issue_type=IssueType.EXCEED_QUANTITY,
                message=f"В КС-2 позиция №{ks2_pos.number}: графа 6 (количество) {ks2_pos.quantity:.4f} превышает графу 6 сметы {smeta_pos.quantity:.4f} на {quantity_diff:.4f} ({exceed_percent:.1f}%)",
                quantity_diff=quantity_diff,
                cost_diff=cost_diff,
                severity="error"
            )
        
        # Графа 7 КС-2 ("цена за единицу, руб.") не должна превышать Графу 7 сметы ("на ед.")
        if price_diff > tolerance_price:
            exceed_percent = (price_diff / smeta_pos.price * 100) if smeta_pos.price > 0 else 0
            return ValidationResult(
                ks2_pos=ks2_pos,
                smeta_pos=smeta_pos,
                issue_type=IssueType.EXCEED_PRICE,
                message=f"В КС-2 позиция №{ks2_pos.number}: графа 7 (цена) {ks2_pos.price:.2f} руб. превышает графу 7 сметы {smeta_pos.price:.2f} руб. на {price_diff:.2f} руб. ({exceed_percent:.1f}%)",
                price_diff=price_diff,
                quantity_diff=quantity_diff,
                cost_diff=cost_diff,
                severity="error"
            )
        
        # Графа 8 КС-2 ("стоимость, руб.") не должна превышать Графу 8 сметы ("Всего")
        if cost_diff > tolerance_cost:
            exceed_percent = (cost_diff / smeta_pos.cost * 100) if smeta_pos.cost > 0 else 0
            return ValidationResult(
                ks2_pos=ks2_pos,
                smeta_pos=smeta_pos,
                issue_type=IssueType.EXCEED_COST,
                message=f"В КС-2 позиция №{ks2_pos.number}: графа 8 (стоимость) {ks2_pos.cost:.2f} руб. превышает графу 8 сметы {smeta_pos.cost:.2f} руб. на {cost_diff:.2f} руб. ({exceed_percent:.1f}%)",
                quantity_diff=quantity_diff,
                cost_diff=cost_diff,
                severity="error"
            )
        
        # Экономия (предупреждения)
        if quantity_diff < -tolerance_qty:
            warnings.append(
                f"Графа 6: экономия по количеству {abs(quantity_diff):.2f} ({abs(quantity_diff / smeta_pos.quantity) * 100:.1f}%)"
            )
        if price_diff < -tolerance_price:
            warnings.append(
                f"Графа 7: экономия по цене {abs(price_diff):.2f} руб. ({abs(price_diff / smeta_pos.price) * 100:.1f}%)"
            )
        if cost_diff < -tolerance_cost:
            warnings.append(
                f"Графа 8: экономия по стоимости {abs(cost_diff):.2f} руб. ({abs(cost_diff / smeta_pos.cost) * 100:.1f}%)"
            )
        
        ok_message = "Позиция в пределах нормы"
        if warnings:
            ok_message += f" (экономия: {', '.join(warnings)})"
        
        return ValidationResult(
            ks2_pos=ks2_pos,
            smeta_pos=smeta_pos,
            issue_type=IssueType.OK,
            message=ok_message,
            quantity_diff=quantity_diff,
            cost_diff=cost_diff,
            price_diff=price_diff,
            severity="info" if warnings else "ok",
            confidence=1.0,
            warnings=warnings
        )

    def _check_duplicates(self):
        """Проверка дублирования работ (пункт 7)"""
        seen_positions = {}
        for ks2_pos in self.ks2_positions:
            key = (ks2_pos.smeta_number, ks2_pos.name[:50])  # Номер сметы + начало наименования
            if key in seen_positions:
                # Найдено дублирование
                first_pos = seen_positions[key]
                result = ValidationResult(
                    ks2_pos=ks2_pos,
                    smeta_pos=None,
                    issue_type=IssueType.DUPLICATE,
                    message=f"В КС-2 позиция №{ks2_pos.number} дублирует позицию №{first_pos.number}: обе ссылаются на смету №{ks2_pos.smeta_number} '{ks2_pos.name[:50]}...'"
                )
                self.header_results.append(result)
            else:
                seen_positions[key] = ks2_pos
    
    def _validate_header(self):
        """Проверка реквизитов в шапке формы (пункт 6)"""
        # Базовая проверка наличия реквизитов
        if not self.ks2_header:
            return
        
        issues = []
        if not self.ks2_header.customer:
            issues.append("Отсутствует наименование заказчика")
        if not self.ks2_header.contractor:
            issues.append("Отсутствует наименование подрядчика")
        if not self.ks2_header.contract_number:
            issues.append("Отсутствует номер договора")
        
        if issues:
            result = ValidationResult(
                ks2_pos=None,  # Для шапки нет позиции
                smeta_pos=None,
                issue_type=IssueType.HEADER_MISMATCH,
                message=f"Несоответствие реквизитов в шапке: {', '.join(issues)}"
            )
            self.header_results.append(result)
    
    def _names_match(self, name1: str, name2: str) -> bool:
        """Проверяет строгое совпадение наименований"""
        if not name1 or not name2:
            return False
        n1 = ' '.join(name1.split()).lower()
        n2 = ' '.join(name2.split()).lower()
        return n1 == n2

    def _units_match(self, unit1: str, unit2: str) -> bool:
        """Проверяет соответствие единиц измерения (строгое сравнение)"""
        unit1_norm = unit1.strip().lower() if unit1 else ''
        unit2_norm = unit2.strip().lower() if unit2 else ''
        return unit1_norm == unit2_norm
    
    def get_summary(self) -> Dict:
        """Возвращает сводку по проверке
        
        Считает только результаты проверки позиций КС-2 (без ошибок шапки)
        """
        # Считаем только результаты проверки позиций КС-2
        position_results = [r for r in self.results if r.ks2_pos is not None]
        header_errors = [r for r in self.header_results if r.issue_type != IssueType.OK]
        
        total = len(position_results)  # Количество проверенных позиций КС-2
        ok = sum(1 for r in position_results if r.issue_type == IssueType.OK)
        warnings = sum(1 for r in position_results if r.issue_type == IssueType.WARNING)
        errors = sum(1 for r in position_results if r.issue_type not in (IssueType.OK, IssueType.WARNING))
        
        # Добавляем ошибки шапки к общему количеству ошибок
        total_errors = errors + len(header_errors)
        
        total_cost_diff = sum(r.cost_diff for r in position_results if r.cost_diff > 0)
        
        return {
            'total': total,  # Всего проверенных позиций КС-2
            'ok': ok,  # Соответствует смете
            'warnings': warnings,
            'errors': total_errors,  # Всего ошибок (включая ошибки шапки)
            'position_errors': errors,  # Ошибки в позициях
            'header_errors': len(header_errors),  # Ошибки в шапке
            'total_cost_diff': total_cost_diff,
            'exceed_quantity': sum(1 for r in position_results if r.issue_type == IssueType.EXCEED_QUANTITY),
            'exceed_cost': sum(1 for r in position_results if r.issue_type == IssueType.EXCEED_COST),
            'exceed_price': sum(1 for r in position_results if r.issue_type == IssueType.EXCEED_PRICE),
            'not_in_smeta': sum(1 for r in position_results if r.issue_type == IssueType.NOT_IN_SMETA),
            'unit_mismatch': sum(1 for r in position_results if r.issue_type == IssueType.UNIT_MISMATCH),
            'name_mismatch': sum(1 for r in position_results if r.issue_type == IssueType.NAME_MISMATCH),
            'justification_mismatch': sum(1 for r in position_results if r.issue_type == IssueType.JUSTIFICATION_MISMATCH),
            'quantity_mismatch': sum(1 for r in position_results if r.issue_type == IssueType.QUANTITY_MISMATCH),
            'cost_mismatch': sum(1 for r in position_results if r.issue_type == IssueType.COST_MISMATCH),
            'price_mismatch': sum(1 for r in position_results if r.issue_type == IssueType.PRICE_MISMATCH),
            'duplicate': sum(1 for r in self.header_results if r.issue_type == IssueType.DUPLICATE),
            'header_mismatch': sum(1 for r in header_errors if r.issue_type == IssueType.HEADER_MISMATCH),
            'all_ok': (errors == 0 and len(header_errors) == 0),  # Флаг: все проверки пройдены
        }


def _parse_float_safe(value: Any) -> float:
    """Безопасное преобразование значения в float с обработкой различных форматов"""
    if value is None:
        return 0.0
    
    if isinstance(value, (int, float)):
        return float(value)
    
    try:
        # Очищаем значение от пробелов, запятых и других символов
        cleaned = str(value).replace(' ', '').replace('\xa0', '').replace(',', '.')
        # Удаляем все после первого нечислового символа (кроме точки и минуса)
        cleaned = re.sub(r'[^0-9\.\-].*$', '', cleaned)
        return float(cleaned) if cleaned else 0.0
    except (ValueError, TypeError):
        return 0.0


def _find_columns_by_content(ws: Any, max_rows: int = 15) -> Optional[Dict[str, int]]:
    """Альтернативный метод определения колонок по содержимому заголовков"""
    logger = logging.getLogger(__name__)
    logger.debug("Используем альтернативный метод определения колонок по содержимому")
    
    # Ищем заголовки в первых строках
    header_candidates = []
    for row_idx in range(1, max_rows + 1):
        row_text = ' '.join(str(ws.cell(row_idx, col_idx).value or '').strip() 
                           for col_idx in range(1, 15))
        if any(keyword in row_text.lower() for keyword in ['наименование', 'обоснование', 'единица измерения']):
            header_candidates.append((row_idx, row_text))
    
    if not header_candidates:
        logger.warning("Не удалось найти заголовки в смете")
        return None
    
    # Берем первую подходящую строку как заголовок
    header_row = header_candidates[0][0]
    logger.debug(f"Заголовок найден на строке {header_row}")
    
    cols_map = {}
    # Определяем колонки по ключевым словам
    for col_idx in range(1, 15):
        cell_value = str(ws.cell(header_row, col_idx).value or '').strip().lower()
        
        if '№' in cell_value or 'п/п' in cell_value:
            cols_map['number'] = col_idx - 1  # 0-based
        elif 'обоснование' in cell_value or 'шифр' in cell_value or 'расценка' in cell_value:
            cols_map['justification'] = col_idx - 1
        elif 'наименование' in cell_value and 'работ' in cell_value:
            cols_map['name'] = col_idx - 1
        elif 'единиц' in cell_value and 'измер' in cell_value:
            cols_map['unit'] = col_idx - 1
        elif 'количеств' in cell_value or 'объем' in cell_value:
            cols_map['qty'] = col_idx - 1
        elif ('цена' in cell_value or 'на ед' in cell_value) and 'всего' not in cell_value:
            cols_map['price'] = col_idx - 1
        elif 'стоимость' in cell_value and 'всего' in cell_value:
            cols_map['cost'] = col_idx - 1
    
    if cols_map:
        logger.debug(f"Маппинг колонок по ключевым словам: {cols_map}")
        return cols_map
    
    return None


def auto_normalize_smeta(path: str, header_row: int = 2, data_start_row: int = 3,
                         cols: Optional[Dict[str, int]] = None) -> List[dict]:
    """
    Автоматическая нормализация сметы с использованием openpyxl.
    Улучшенная версия с гибким поиском структуры и умной фильтрацией.
    
    Структура файла:
    - Строка header_row (обычно 2): номера граф (1, 2, 3, 4, ..., 14)
    - Строка data_start_row (обычно 3): начало реальных данных
    
    Args:
        path: Путь к Excel-файлу сметы
        header_row: Номер строки с номерами граф (1-based, обычно 2)
        data_start_row: Номер строки начала данных (1-based, обычно 3)
        cols: Словарь с индексами столбцов (0-based). Если None, будет автопоиск.
              Формат: {'number': 0, 'justification': 1, 'name': 2, 'unit': 5, 'qty': 6, 'price': 7, 'cost': 8}
    
    Returns:
        Список словарей с нормализованными данными из сметы
    """
    logger = logging.getLogger(__name__)
    
    try:
        import openpyxl
    except ImportError:
        raise ImportError("Для использования auto_normalize_smeta требуется openpyxl. Установите: pip install openpyxl")
    
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        logger.info(f"Файл {path} успешно загружен с помощью openpyxl")
    except Exception as e:
        logger.error(f"Ошибка загрузки файла {path} с помощью openpyxl: {str(e)}")
        raise
    
    rows = []
    seen_numbers = set()
    
    # Если индексы столбцов не заданы, ищем строку с номерами граф для автопоиска
    if cols is None:
        cols = {}
        # Улучшенный поиск строки с номерами граф от 1 до 14 с гибким форматированием
        graph_row = None
        # Увеличиваем количество проверяемых строк - строка с графами может быть дальше
        max_rows_to_check = min(50, ws.max_row)  # Проверяем первые 50 строк (включая строку 26)
        max_cols_to_check = min(20, ws.max_column)  # Проверяем до 20 столбцов
        
        logger.info(f"Поиск строки с номерами граф (1-14) в строках 1-{max_rows_to_check}, проверяем до {max_cols_to_check} столбцов")
        
        for row_idx in range(1, max_rows_to_check + 1):
            row_values = []  # Список кортежей (col_idx, num, cell_str)
            graph_positions = {}  # Словарь: номер графы -> позиция колонки (0-based)
            
            # Собираем значения из всех ячеек строки с гибким извлечением чисел
            for col_idx in range(1, max_cols_to_check + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value is not None:
                    # Сначала пробуем получить число напрямую (если это число в Excel)
                    num = None
                    cell_str = None
                    
                    # Проверяем, является ли значение числом
                    if isinstance(cell.value, (int, float)):
                        try:
                            num_val = int(cell.value)
                            if 1 <= num_val <= 14:
                                num = num_val
                                cell_str = str(cell.value)
                        except (ValueError, TypeError):
                            pass
                    
                    # Если не получилось как число, пробуем извлечь из строки
                    if num is None:
                        # Очищаем значение от лишних символов и пробелов
                        cell_str = str(cell.value).strip().replace('\n', ' ')
                        
                        # Пробуем извлечь число из строки (даже если есть дополнительные символы)
                        # Например: "1.", "2)", "Графа 3", "№4", "1 ", " 2"
                        numbers_in_cell = re.findall(r'\d+', cell_str)
                        if numbers_in_cell:
                            try:
                                num = int(numbers_in_cell[0])
                                if not (1 <= num <= 14):
                                    num = None
                            except (ValueError, TypeError):
                                num = None
                    
                    # Если нашли число от 1 до 14, добавляем его
                    if num is not None:
                        row_values.append((col_idx, num, cell_str))
                        graph_positions[num] = col_idx - 1  # 0-based индекс
            
                # Анализируем собранную информацию
            # Универсальная проверка: должно быть минимум 3 номера граф
            if len(row_values) >= 3:
                # Сортируем по позиции в строке
                row_values.sort(key=lambda x: x[0])
                
                # Извлекаем только номера для анализа
                numbers = [num for _, num, _ in row_values]
                found_graphs = sorted(set(numbers))
                
                # Детальное логирование для строк с большим количеством найденных граф
                if len(found_graphs) >= 8:
                    logger.info(f"[ОТЛАДКА] Строка {row_idx}: найдено {len(found_graphs)} уникальных граф: {found_graphs}")
                    logger.info(f"[ОТЛАДКА] Строка {row_idx}: первые 15 значений ячеек: {[(col, num, str(text)[:30]) for col, num, text in row_values[:15]]}")
                
                # Отладочный вывод для первых 10 строк
                elif row_idx <= 10:
                    logger.debug(f"Строка {row_idx}: найдено {len(found_graphs)} уникальных граф: {found_graphs}")
                
                # Проверяем, есть ли в строке последовательность начиная с 1
                has_sequence_start = any(num == 1 for num in numbers[:10])  # 1 должна быть в первых 10 ячейках
                
                # Проверяем последовательность: ищем начало и считаем длину
                sequential_count = 1
                current_num = 1
                start_idx = -1
                
                # Ищем начало последовательности (графа 1)
                for i, num in enumerate(numbers):
                    if num == 1:
                        start_idx = i
                        break
                
                # Также проверяем, есть ли все графы от 1 до 14 (или большинство из них)
                has_all_graphs = len(found_graphs) >= 12  # Если найдено 12+ уникальных граф, это точно строка с графами
                has_majority_graphs = len(found_graphs) >= 10 and all(g in found_graphs for g in [1, 2, 3, 4, 5, 6, 7, 8])
                
                if start_idx != -1:
                    # Считаем длину последовательности
                    for i in range(start_idx + 1, len(numbers)):
                        if numbers[i] == current_num + 1:
                            sequential_count += 1
                            current_num += 1
                        elif numbers[i] > current_num + 1:
                            # Пропускаем пропуски, но продолжаем считать
                            if numbers[i] <= current_num + 3:  # Допускаем пропуск 1-2 граф
                                sequential_count += 1
                                current_num = numbers[i]
                            else:
                                # Если уже набрали достаточно последовательных граф, не прерываем
                                if sequential_count >= 5:
                                    continue
                                break  # Большой разрыв в последовательности
                
                # Дополнительная проверка: если найдено много граф (10+), проверяем их распределение
                if len(found_graphs) >= 10:
                    # Проверяем, что графы покрывают диапазон от 1 до 14
                    min_graph = min(found_graphs)
                    max_graph = max(found_graphs)
                    graph_range = max_graph - min_graph + 1
                    # Если диапазон большой (10+), это хороший признак
                    if min_graph == 1 and max_graph >= 10:
                        has_all_graphs = True
                
                # Логика принятия решения
                is_graph_numbers_row = False
                
                # ПРИОРИТЕТНАЯ ПРОВЕРКА: Если числа идут последовательно в первых ячейках (A, B, C, D...)
                # Это самый надежный признак строки с номерами граф - работает для любой строки
                if len(row_values) >= 5:
                    # Проверяем, что первые найденные числа идут последовательно (1, 2, 3, 4, 5...)
                    first_numbers = [num for _, num, _ in row_values[:14]]  # Первые 14 найденных чисел
                    if first_numbers[0] == 1:  # Начинается с 1
                        sequential_in_order = True
                        consecutive_count = 1
                        for i in range(1, min(14, len(first_numbers))):  # Проверяем первые 14
                            if first_numbers[i] == first_numbers[i-1] + 1:
                                consecutive_count += 1
                            elif first_numbers[i] > first_numbers[i-1] + 1:
                                # Допускаем небольшие пропуски (например, пропуск графы 5)
                                if first_numbers[i] <= first_numbers[i-1] + 3:
                                    consecutive_count += 1
                                else:
                                    break
                        
                        # Если первые числа идут последовательно 1,2,3,4,5,6,7,8... (минимум 5 подряд) - это строка с графами
                        if consecutive_count >= 5:
                            is_graph_numbers_row = True
                            logger.info(f"Строка {row_idx}: ✓ НАЙДЕНА по последовательности в первых ячейках: "
                                      f"{first_numbers[:consecutive_count]}... (всего {len(found_graphs)} граф)")
                
                # Вариант 1 (приоритетный): Найдено 12+ уникальных граф от 1 до 14 - это точно строка с графами
                if not is_graph_numbers_row and has_all_graphs:
                    is_graph_numbers_row = True
                    logger.info(f"Строка {row_idx}: ✓ НАЙДЕНА строка с номерами граф! Найдено {len(found_graphs)} уникальных граф: {found_graphs}")
                
                # Вариант 2: Найдена последовательность начиная с 1 (минимум 5 чисел подряд)
                elif sequential_count >= 5 and start_idx >= 0:
                    is_graph_numbers_row = True
                    logger.info(f"Строка {row_idx}: ✓ НАЙДЕНА последовательность номеров граф длиной {sequential_count}: "
                              f"{numbers[start_idx:start_idx+sequential_count] if start_idx >= 0 else []}")
                
                # Вариант 3: Найдено большинство основных граф (1-8) + дополнительные
                elif has_majority_graphs:
                    is_graph_numbers_row = True
                    logger.info(f"Строка {row_idx}: ✓ НАЙДЕНА строка с основными графами (1-8) и дополнительными: "
                              f"{len(found_graphs)} граф: {found_graphs}")
                
                # Вариант 4: Строка содержит много номеров (8+) и начинается с 1
                elif has_sequence_start and len(found_graphs) >= 8:
                    valid_count = sum(1 for num in found_graphs if 1 <= num <= 14)
                    if valid_count >= 8:
                        is_graph_numbers_row = True
                        logger.info(f"Строка {row_idx}: ✓ НАЙДЕНА строка с номерами граф по количеству: "
                                  f"{valid_count} из {len(found_graphs)} номеров в диапазоне 1-14")
                
                # Вариант 5: Строка содержит ключевые слова "графа", "колонка" и достаточно цифр
                if not is_graph_numbers_row:
                    row_text = ' '.join(cell_str for _, _, cell_str in row_values).lower()
                    if any(keyword in row_text for keyword in ['графа', 'колонка', '№', 'номер']) and len(found_graphs) >= 4:
                        is_graph_numbers_row = True
                        logger.info(f"Строка {row_idx}: ✓ НАЙДЕНА строка с номерами граф по ключевым словам: "
                                  f"{row_text[:50]}...")
                
                # Вариант 6: Если найдено 10+ граф и они покрывают диапазон от 1 до 10+, это тоже строка с графами
                if not is_graph_numbers_row and len(found_graphs) >= 10:
                    min_graph = min(found_graphs)
                    max_graph = max(found_graphs)
                    if min_graph == 1 and max_graph >= 10:
                        is_graph_numbers_row = True
                        logger.info(f"Строка {row_idx}: ✓ НАЙДЕНА строка с графами по диапазону: "
                                  f"от {min_graph} до {max_graph}, всего {len(found_graphs)} граф: {found_graphs}")
                
                if is_graph_numbers_row:
                    graph_row = row_idx
                    logger.info(f"✓ Строка с номерами граф найдена на позиции {row_idx}")
                    logger.debug(f"Значения в строке: {[(col, num, text[:20]) for col, num, text in row_values[:10]]}")
                    
                    # Создаем маппинг: номер графы → индекс колонки (0-based)
                    # Используем все найденные графы для точного маппинга
                    for graph_num in range(1, 15):
                        if graph_num in graph_positions:
                            col_idx = graph_positions[graph_num]
                            # Маппинг основных граф для нормализованной таблицы
                            if graph_num == 1:
                                cols['number'] = col_idx
                            elif graph_num == 2:
                                cols['justification'] = col_idx
                            elif graph_num == 3:
                                cols['name'] = col_idx
                            elif graph_num == 4:
                                cols['unit'] = col_idx
                            elif graph_num == 6:
                                cols['qty'] = col_idx
                            elif graph_num == 7:
                                cols['price'] = col_idx
                            elif graph_num == 8:
                                cols['cost'] = col_idx
                            # Остальные графы (9-14) сохраняем для информации, но не используем в нормализованной таблице
                    
                    logger.info(f"Найдено граф: {sorted(graph_positions.keys())}")
                    logger.debug(f"Маппинг граф определен: {cols}")
                    
                    # Обновляем data_start_row - обычно это следующая строка после номеров граф
                    data_start_row = graph_row + 1
                    
                    # Проверяем следующую строку - часто там подзаголовки
                    next_row_text = ''
                    for col_idx in range(1, min(10, max_cols_to_check + 1)):
                        cell = ws.cell(row=data_start_row, column=col_idx)
                        if cell.value is not None:
                            next_row_text += str(cell.value).strip().lower() + ' '
                    
                    # Если следующая строка содержит подзаголовки, пропускаем и ее
                    if any(keyword in next_row_text for keyword in ['наименование', 'ед. изм', 'едизм', 
                                                                     'количество', 'цена', 'стоимость', 'обоснование']):
                        data_start_row += 1
                        logger.info(f"Обнаружены подзаголовки, данные начинаются со строки {data_start_row}")
                    
                    break
        
        if graph_row is None:
            logger.warning("Строка с номерами граф не найдена стандартными методами")
        
        # Если не нашли по номерам граф, пробуем альтернативный метод
        if not cols:
            logger.warning("Строка с номерами граф не найдена. Используем альтернативный метод определения колонок")
            alt_cols = _find_columns_by_content(ws)
            if alt_cols:
                cols = alt_cols
                data_start_row = header_row + 1
            else:
                logger.warning("Альтернативный метод тоже не сработал, используем значения по умолчанию")
                cols = {'number': 0, 'justification': 1, 'name': 2, 'unit': 5, 'qty': 6, 'price': 7, 'cost': 8}
        else:
            logger.info(f"Маппинг столбцов определен: {cols}")
    
    logger.info(f"Начало чтения данных со строки {data_start_row}")
    logger.debug(f"Используемые столбцы: {cols}")
    
    # Читаем данные начиная с data_start_row
    max_data_rows = min(ws.max_row + 1, data_start_row + 500)  # Ограничиваем до 500 строк для производительности
    for row_idx in range(data_start_row, max_data_rows):
        row = ws[row_idx]
        
        # Пропускаем пустые строки
        if all(cell.value is None for cell in row[:10]):
            continue
        
        # Умная фильтрация: пропускаем строки с вспомогательными данными
        row_text = ' '.join(str(cell.value or '').strip() for cell in row[:10]).lower()
        if any(marker in row_text for marker in ['затраты труда машинистов', 'накладные расходы', 
                                               'сметная прибыль', 'итого', 'всего', 'в том числе',
                                               'итого по разделу', 'раздел', 'подраздел']):
            continue
        
        # Проверяем, что строка содержит числовой номер позиции
        if cols.get('number') is None or len(row) <= cols['number']:
            continue
        
        number_cell = row[cols['number']]
        number_val = number_cell.value
        if number_val is None:
            continue
        
        # Проверяем, что номер позиции - это число (с учетом различных форматов)
        try:
            number_str = str(number_val).strip()
            # Извлекаем числа из строки (даже если есть дополнительные символы типа "1.", "2)", "№3")
            numbers_in_cell = re.findall(r'\d+', number_str)
            if not numbers_in_cell:
                continue
            
            # Берем первое число
            clean_number = numbers_in_cell[0]
            # Проверяем, что это не просто номер графы (1-14)
            if clean_number in [str(i) for i in range(1, 15)]:
                continue
            
            # Убираем все кроме цифр и точки для сохранения
            number_str = re.sub(r'[^\d.]', '', number_str)
            if not number_str or not re.match(r'^[\d.]+$', number_str):
                continue
        except (ValueError, TypeError):
            continue
        
        # Извлекаем данные из строки
        try:
            # Получаем все значения строки
            vals = [cell.value for cell in row]
            
            # Извлекаем наименование
            name_idx = cols.get('name', 2)
            name_val = str(vals[name_idx] or '').strip() if name_idx < len(vals) else ''
            
            # Пропускаем строки без наименования
            if not name_val:
                continue
            
            # Извлекаем единицу измерения
            unit_idx = cols.get('unit', 5)
            unit_val = str(vals[unit_idx] or '').strip() if unit_idx < len(vals) else ''
            
            # Пропускаем позиции без единицы измерения (согласно рекомендациям)
            if not unit_val:
                continue
            
            # Формируем позицию
            position = {
                'number': number_str,
                'justification': str(vals[cols.get('justification', 1)] or '').strip() if cols.get('justification', 1) < len(vals) else '',
                'name': name_val,
                'unit': unit_val,
                'quantity': _parse_float_safe(vals[cols.get('qty', 6)]) if cols.get('qty', 6) < len(vals) else 0.0,
                'price': _parse_float_safe(vals[cols.get('price', 7)]) if cols.get('price', 7) < len(vals) else 0.0,
                'cost': _parse_float_safe(vals[cols.get('cost', 8)]) if cols.get('cost', 8) < len(vals) else 0.0,
            }
            
            # Автоматический расчет стоимости, если она отсутствует
            if position['cost'] == 0 and position['quantity'] > 0 and position['price'] > 0:
                position['cost'] = position['quantity'] * position['price']
            
            # Пропускаем дубликаты по номеру
            if position['number'] in seen_numbers:
                logger.debug(f"Пропущен дубликат позиции {position['number']}")
                continue
            
            seen_numbers.add(position['number'])
            rows.append(position)
            
            # Отладка первых 3 записей
            if len(rows) <= 3:
                logger.debug(f"Позиция {len(rows)}: number='{position['number']}', name='{position['name'][:50]}', "
                           f"qty={position['quantity']}, price={position['price']:.2f}, cost={position['cost']:.2f}")
        except (ValueError, TypeError, IndexError, AttributeError) as e:
            logger.warning(f"Ошибка при обработке строки {row_idx}: {e}")
            continue
    
    logger.info(f"Нормализованная таблица сформирована. Всего позиций: {len(rows)}")
    
    if not rows:
        logger.warning("Нормализованная таблица пустая. Проверяем первые 10 строк для отладки:")
        for row_idx in range(1, 11):
            row_values = [str(ws.cell(row_idx, col_idx).value or '')[:20] 
                         for col_idx in range(1, 11)]
            logger.warning(f"Строка {row_idx}: {row_values}")
    
    return rows


def normalize_smeta_table(smeta_positions: Dict[str, SmetaPosition]) -> List[dict]:
    """
    Формирует нормализованную таблицу только по смете из уже распарсенных позиций.
    
    Args:
        smeta_positions: Словарь позиций из сметы (ключ = номер позиции)
    
    Returns:
        Список словарей с нормализованными данными из сметы
    """
    result = []
    
    print(f"[НОРМАЛИЗАЦИЯ СМЕТЫ] Начало формирования таблицы. Всего позиций: {len(smeta_positions)}")
    
    for idx, (number, pos) in enumerate(smeta_positions.items()):
        try:
            row_data = {
                'number': str(pos.number) if pos.number is not None else '',
                'justification': str(pos.justification) if pos.justification is not None else '',
                'name': str(pos.name) if pos.name is not None else '',
                'unit': str(pos.unit) if pos.unit is not None else '',
                'quantity': float(pos.quantity) if pos.quantity is not None else 0.0,
                'price': float(pos.price) if pos.price is not None else 0.0,
                'cost': float(pos.cost) if pos.cost is not None else 0.0,
            }
            result.append(row_data)
            
            # Отладка первых 3 записей
            if idx < 3:
                print(f"[НОРМАЛИЗАЦИЯ СМЕТЫ] Позиция {idx + 1}: number='{row_data['number']}', name='{row_data['name'][:50] if row_data['name'] else ''}', qty={row_data['quantity']}, price={row_data['price']}, cost={row_data['cost']}")
        except Exception as e:
            print(f"[ОШИБКА] Ошибка при обработке позиции {number}: {e}")
            print(f"  - pos: {pos}")
            continue
    
    print(f"[НОРМАЛИЗАЦИЯ СМЕТЫ] Таблица сформирована. Записей: {len(result)}")
    return result

