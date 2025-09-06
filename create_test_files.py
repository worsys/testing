#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для создания тестовых Excel файлов для демонстрации функции объединения
"""

import os
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta
import random

def create_test_data():
    """Создать тестовые данные для Excel файлов"""
    
    # Создать папку "файлы", если она не существует
    test_folder = Path("файлы")
    test_folder.mkdir(exist_ok=True)
    
    print("📁 Создаю тестовые Excel файлы...")
    print("=" * 50)
    
    # Данные для первого файла - продажи за январь
    dates1 = [datetime(2024, 1, i) for i in range(1, 16)]
    data1 = {
        'Дата': dates1,
        'Товар': ['Товар_' + str(random.randint(1, 10)) for _ in dates1],
        'Количество': [random.randint(1, 50) for _ in dates1],
        'Цена': [round(random.uniform(10, 1000), 2) for _ in dates1],
    }
    data1['Сумма'] = [q * p for q, p in zip(data1['Количество'], data1['Цена'])]
    
    df1 = pd.DataFrame(data1)
    file1 = test_folder / "продажи_январь.xlsx"
    df1.to_excel(file1, index=False)
    print(f"✅ Создан файл: {file1} ({len(df1)} строк)")
    
    # Данные для второго файла - продажи за февраль
    dates2 = [datetime(2024, 2, i) for i in range(1, 21)]
    data2 = {
        'Дата': dates2,
        'Товар': ['Товар_' + str(random.randint(1, 10)) for _ in dates2],
        'Количество': [random.randint(1, 50) for _ in dates2],
        'Цена': [round(random.uniform(10, 1000), 2) for _ in dates2],
    }
    data2['Сумма'] = [q * p for q, p in zip(data2['Количество'], data2['Цена'])]
    
    df2 = pd.DataFrame(data2)
    file2 = test_folder / "продажи_февраль.xlsx"
    df2.to_excel(file2, index=False)
    print(f"✅ Создан файл: {file2} ({len(df2)} строк)")
    
    # Данные для третьего файла - продажи за март
    dates3 = [datetime(2024, 3, i) for i in range(1, 26)]
    data3 = {
        'Дата': dates3,
        'Товар': ['Товар_' + str(random.randint(1, 10)) for _ in dates3],
        'Количество': [random.randint(1, 50) for _ in dates3],
        'Цена': [round(random.uniform(10, 1000), 2) for _ in dates3],
    }
    data3['Сумма'] = [q * p for q, p in zip(data3['Количество'], data3['Цена'])]
    
    df3 = pd.DataFrame(data3)
    file3 = test_folder / "продажи_март.xlsx"
    df3.to_excel(file3, index=False)
    print(f"✅ Создан файл: {file3} ({len(df3)} строк)")
    
    print("=" * 50)
    print(f"🎉 Создано {3} тестовых Excel файлов в папке '{test_folder}'")
    print(f"📊 Общее количество строк: {len(df1) + len(df2) + len(df3)}")
    print()
    print("Теперь можно протестировать объединение:")
    print("python collect_excel_files.py -m 'объединенные_продажи.xlsx'")

if __name__ == "__main__":
    try:
        create_test_data()
    except ImportError:
        print("❌ Для создания тестовых файлов требуется pandas и openpyxl")
        print("   Выполните: pip install pandas openpyxl")
    except Exception as e:
        print(f"❌ Ошибка при создании тестовых файлов: {e}")