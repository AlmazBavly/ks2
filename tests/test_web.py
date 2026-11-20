#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Юнит-тесты для веб-приложения
"""
import unittest
from unittest.mock import Mock, patch, MagicMock
import sys
import os

# Добавляем путь к корневой директории проекта
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ks2_validator_web import app, allowed_file


class TestWebApp(unittest.TestCase):
    """Тесты для веб-приложения"""
    
    def setUp(self):
        """Настройка тестового клиента"""
        self.app = app.test_client()
        self.app.testing = True
    
    def test_index_page(self):
        """Тест главной страницы"""
        response = self.app.get('/')
        self.assertEqual(response.status_code, 200)
        # Проверяем наличие ключевых элементов на странице
        response_text = response.data.decode('utf-8')
        self.assertIn('Сверка', response_text)
        self.assertIn('КС-2', response_text)
    
    def test_allowed_file(self):
        """Тест проверки разрешенных файлов"""
        self.assertTrue(allowed_file('test.xlsx'))
        self.assertTrue(allowed_file('test.XLSX'))
        self.assertTrue(allowed_file('test.xls'))
        self.assertFalse(allowed_file('test.pdf'))
        self.assertFalse(allowed_file('test.txt'))
        self.assertFalse(allowed_file('test'))
    
    @patch('ks2_validator_web.SmetaParser')
    @patch('ks2_validator_web.KS2Parser')
    @patch('ks2_validator_web.KS2Validator')
    def test_validate_api_no_files(self, mock_validator, mock_ks2_parser, mock_smeta_parser):
        """Тест API валидации без файлов"""
        response = self.app.post('/api/validate')
        self.assertEqual(response.status_code, 400)
        data = response.get_json()
        self.assertIn('error', data)
    
    def test_validate_api_invalid_format(self):
        """Тест API валидации с неправильным форматом"""
        from io import BytesIO
        data = {
            'smeta_file': (BytesIO(b'fake pdf content'), 'test.pdf'),
            'ks2_file': (BytesIO(b'fake pdf content'), 'test.pdf')
        }
        response = self.app.post('/api/validate', data=data, content_type='multipart/form-data')
        self.assertEqual(response.status_code, 400)
        json_data = response.get_json()
        self.assertIn('error', json_data)


if __name__ == '__main__':
    unittest.main()

