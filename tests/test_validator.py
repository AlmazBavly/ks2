#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Юнит-тесты для модуля проверки КС-2
"""
import unittest
from unittest.mock import Mock, patch
import sys
import os

# Добавляем путь к корневой директории проекта
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ks2_validator import (
    SmetaPosition, KS2Position, KS2Header,
    SmetaParser, KS2Parser, KS2Validator,
    IssueType
)


class TestSmetaPosition(unittest.TestCase):
    """Тесты для класса SmetaPosition"""
    
    def test_smeta_position_creation(self):
        """Тест создания позиции сметы"""
        pos = SmetaPosition(
            number="1",
            name="Тестовая работа",
            unit="м3",
            quantity=10.0,
            price=100.0,
            cost=1000.0,
            justification="ГЭСН01-01-001"
        )
        self.assertEqual(pos.number, "1")
        self.assertEqual(pos.name, "Тестовая работа")
        self.assertEqual(pos.unit, "м3")
        self.assertEqual(pos.quantity, 10.0)
        self.assertEqual(pos.price, 100.0)
        self.assertEqual(pos.cost, 1000.0)


class TestKS2Position(unittest.TestCase):
    """Тесты для класса KS2Position"""
    
    def test_ks2_position_creation(self):
        """Тест создания позиции КС-2"""
        pos = KS2Position(
            number="1",
            smeta_number="1",
            name="Тестовая работа",
            justification="ГЭСН01-01-001",
            unit="м3",
            quantity=10.0,
            price=100.0,
            cost=1000.0
        )
        self.assertEqual(pos.number, "1")
        self.assertEqual(pos.smeta_number, "1")
        self.assertEqual(pos.name, "Тестовая работа")
        self.assertEqual(pos.unit, "м3")


class TestKS2Validator(unittest.TestCase):
    """Тесты для класса KS2Validator"""
    
    def setUp(self):
        """Настройка тестовых данных"""
        self.smeta_positions = {
            "1": SmetaPosition(
                number="1",
                name="Тестовая работа",
                unit="м3",
                quantity=10.0,
                price=100.0,
                cost=1000.0,
                justification="ГЭСН01-01-001"
            ),
            "2": SmetaPosition(
                number="2",
                name="Другая работа",
                unit="м2",
                quantity=5.0,
                price=200.0,
                cost=1000.0,
                justification="ГЭСН01-01-002"
            )
        }
    
    def test_validate_ok(self):
        """Тест успешной проверки"""
        ks2_positions = [
            KS2Position(
                number="1",
                smeta_number="1",
                name="Тестовая работа",
                justification="ГЭСН01-01-001",
                unit="м3",
                quantity=10.0,
                price=100.0,
                cost=1000.0
            )
        ]
        
        validator = KS2Validator(ks2_positions, self.smeta_positions)
        results = validator.validate()
        
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].issue_type, IssueType.OK)
    
    def test_validate_exceed_quantity(self):
        """Тест превышения количества"""
        ks2_positions = [
            KS2Position(
                number="1",
                smeta_number="1",
                name="Тестовая работа",
                justification="ГЭСН01-01-001",
                unit="м3",
                quantity=15.0,  # Превышение
                price=100.0,
                cost=1500.0
            )
        ]
        
        validator = KS2Validator(ks2_positions, self.smeta_positions)
        results = validator.validate()
        
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].issue_type, IssueType.EXCEED_QUANTITY)
        self.assertGreater(results[0].quantity_diff, 0)
    
    def test_validate_exceed_cost(self):
        """Тест превышения стоимости"""
        ks2_positions = [
            KS2Position(
                number="1",
                smeta_number="1",
                name="Тестовая работа",
                justification="ГЭСН01-01-001",
                unit="м3",
                quantity=10.0,
                price=150.0,  # Выше цены
                cost=1500.0  # Превышение
            )
        ]
        
        validator = KS2Validator(ks2_positions, self.smeta_positions)
        results = validator.validate()
        
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].issue_type, IssueType.EXCEED_COST)
        self.assertGreater(results[0].cost_diff, 0)
    
    def test_validate_not_in_smeta(self):
        """Тест позиции, отсутствующей в смете"""
        ks2_positions = [
            KS2Position(
                number="1",
                smeta_number="999",  # Не существует в смете
                name="Несуществующая работа",
                justification="ГЭСН01-01-999",
                unit="м3",
                quantity=10.0,
                price=100.0,
                cost=1000.0
            )
        ]
        
        validator = KS2Validator(ks2_positions, self.smeta_positions)
        results = validator.validate()
        
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].issue_type, IssueType.NOT_IN_SMETA)
    
    def test_validate_unit_mismatch(self):
        """Тест несоответствия единиц измерения"""
        ks2_positions = [
            KS2Position(
                number="1",
                smeta_number="1",
                name="Тестовая работа",
                justification="ГЭСН01-01-001",
                unit="м2",  # Не соответствует м3
                quantity=10.0,
                price=100.0,
                cost=1000.0
            )
        ]
        
        validator = KS2Validator(ks2_positions, self.smeta_positions)
        results = validator.validate()
        
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].issue_type, IssueType.UNIT_MISMATCH)
    
    def test_validate_name_mismatch(self):
        """Тест несоответствия наименования"""
        ks2_positions = [
            KS2Position(
                number="1",
                smeta_number="1",
                name="Совсем другая работа",  # Не соответствует
                justification="ГЭСН01-01-001",
                unit="м3",
                quantity=10.0,
                price=100.0,
                cost=1000.0
            )
        ]
        
        validator = KS2Validator(ks2_positions, self.smeta_positions)
        results = validator.validate()
        
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].issue_type, IssueType.NAME_MISMATCH)
    
    def test_validate_duplicate(self):
        """Тест дублирования работ"""
        ks2_positions = [
            KS2Position(
                number="1",
                smeta_number="1",
                name="Тестовая работа",
                justification="ГЭСН01-01-001",
                unit="м3",
                quantity=10.0,
                price=100.0,
                cost=1000.0
            ),
            KS2Position(
                number="2",
                smeta_number="1",  # Дубликат
                name="Тестовая работа",
                justification="ГЭСН01-01-001",
                unit="м3",
                quantity=5.0,
                price=100.0,
                cost=500.0
            )
        ]
        
        validator = KS2Validator(ks2_positions, self.smeta_positions)
        results = validator.validate()
        
        # Должна быть найдена ошибка дублирования
        duplicate_errors = [r for r in results if r.issue_type == IssueType.DUPLICATE]
        self.assertGreater(len(duplicate_errors), 0)
    
    def test_get_summary(self):
        """Тест получения сводки"""
        ks2_positions = [
            KS2Position(
                number="1",
                smeta_number="1",
                name="Тестовая работа",
                justification="ГЭСН01-01-001",
                unit="м3",
                quantity=15.0,  # Превышение
                price=100.0,
                cost=1500.0
            )
        ]
        
        validator = KS2Validator(ks2_positions, self.smeta_positions)
        validator.validate()
        summary = validator.get_summary()
        
        self.assertEqual(summary['total'], 1)
        self.assertEqual(summary['errors'], 1)
        self.assertEqual(summary['exceed_quantity'], 1)
        self.assertEqual(summary['ok'], 0)


class TestUnitsMatch(unittest.TestCase):
    """Тесты для проверки соответствия единиц измерения"""
    
    def setUp(self):
        """Настройка валидатора"""
        self.validator = KS2Validator([], {})
    
    def test_units_match_identical(self):
        """Тест идентичных единиц"""
        self.assertTrue(self.validator._units_match("м3", "м3"))
        self.assertTrue(self.validator._units_match("м2", "м2"))
        self.assertTrue(self.validator._units_match("м", "м"))
    
    def test_units_match_synonyms(self):
        """Тест синонимов единиц"""
        self.assertTrue(self.validator._units_match("м3", "м³"))
        self.assertTrue(self.validator._units_match("м2", "м²"))
        self.assertTrue(self.validator._units_match("т", "тонн"))
    
    def test_units_match_different(self):
        """Тест разных единиц"""
        self.assertFalse(self.validator._units_match("м3", "м2"))
        self.assertFalse(self.validator._units_match("м", "кг"))


class TestNamesMatch(unittest.TestCase):
    """Тесты для проверки соответствия наименований"""
    
    def setUp(self):
        """Настройка валидатора"""
        self.validator = KS2Validator([], {})
    
    def test_names_match_identical(self):
        """Тест идентичных наименований"""
        self.assertTrue(self.validator._names_match(
            "Копание ям вручную",
            "Копание ям вручную"
        ))
    
    def test_names_match_similar(self):
        """Тест похожих наименований"""
        self.assertTrue(self.validator._names_match(
            "Копание ям вручную без креплений",
            "Копание ям вручную"
        ))
    
    def test_names_match_different(self):
        """Тест разных наименований"""
        self.assertFalse(self.validator._names_match(
            "Копание ям вручную",
            "Укладка труб"
        ))


if __name__ == '__main__':
    unittest.main()

