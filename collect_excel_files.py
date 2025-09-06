#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для сбора и объединения всех Excel файлов из папки "файлы"
Script to collect and merge all Excel files from "файлы" folder
"""

import os
import shutil
import glob
from pathlib import Path
from datetime import datetime
import argparse

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False
    print("⚠️  Для объединения файлов требуется установка pandas и openpyxl:")
    print("   pip install pandas openpyxl")
    print()


def get_file_info(file_path):
    """Получить информацию о файле"""
    stat = os.stat(file_path)
    size = stat.st_size
    modified = datetime.fromtimestamp(stat.st_mtime)
    
    # Конвертировать размер в читаемый формат
    if size < 1024:
        size_str = f"{size} B"
    elif size < 1024**2:
        size_str = f"{size/1024:.1f} KB"
    elif size < 1024**3:
        size_str = f"{size/(1024**2):.1f} MB"
    else:
        size_str = f"{size/(1024**3):.1f} GB"
    
    return {
        'size': size,
        'size_str': size_str,
        'modified': modified,
        'modified_str': modified.strftime('%Y-%m-%d %H:%M:%S')
    }


def find_excel_files(source_folder):
    """Найти все Excel файлы в указанной папке"""
    source_path = Path(source_folder)
    
    if not source_path.exists():
        print(f"❌ Папка '{source_folder}' не найдена!")
        return []
    
    if not source_path.is_dir():
        print(f"❌ '{source_folder}' не является папкой!")
        return []
    
    # Поиск Excel файлов
    excel_extensions = ['*.xlsx', '*.xls', '*.xlsm', '*.xlsb']
    excel_files = []
    
    for ext in excel_extensions:
        pattern = source_path / "**" / ext
        files = glob.glob(str(pattern), recursive=True)
        excel_files.extend(files)
    
    return sorted(excel_files)


def collect_excel_files(source_folder="файлы", destination_folder=None, copy_files=False, merge_files=False, merge_output=None, merge_sheet=None, add_source_column=True):
    """
    Собрать все Excel файлы из указанной папки
    
    Args:
        source_folder: Папка-источник (по умолчанию "файлы")
        destination_folder: Папка назначения (если нужно копировать файлы)
        copy_files: Копировать ли файлы в папку назначения
        merge_files: Объединять ли файлы в один
        merge_output: Имя выходного файла для объединения
        merge_sheet: Имя листа для чтения (если None, берется первый лист)
        add_source_column: Добавлять ли колонку с именем исходного файла при объединении
    """
    
    print(f"🔍 Поиск Excel файлов в папке: {source_folder}")
    print("=" * 60)
    
    excel_files = find_excel_files(source_folder)
    
    if not excel_files:
        print(f"📭 В папке '{source_folder}' не найдено Excel файлов")
        return excel_files
    
    print(f"📊 Найдено Excel файлов: {len(excel_files)}")
    print()
    
    total_size = 0
    
    # Показать информацию о каждом файле
    for i, file_path in enumerate(excel_files, 1):
        file_info = get_file_info(file_path)
        total_size += file_info['size']
        
        print(f"{i:2d}. {os.path.basename(file_path)}")
        print(f"    📁 Путь: {file_path}")
        print(f"    📏 Размер: {file_info['size_str']}")
        print(f"    📅 Изменен: {file_info['modified_str']}")
        print()
    
    # Показать общую статистику
    if total_size < 1024:
        total_size_str = f"{total_size} B"
    elif total_size < 1024**2:
        total_size_str = f"{total_size/1024:.1f} KB"
    elif total_size < 1024**3:
        total_size_str = f"{total_size/(1024**2):.1f} MB"
    else:
        total_size_str = f"{total_size/(1024**3):.1f} GB"
    
    print("=" * 60)
    print(f"📈 Общий размер: {total_size_str}")
    print()
    
    # Копирование файлов, если указана папка назначения
    if copy_files and destination_folder:
        copy_excel_files(excel_files, destination_folder)
        print()
    
    # Объединение файлов, если нужно
    if merge_files and merge_output:
        merge_excel_files(excel_files, merge_output, merge_sheet, add_source_column)
    
    return excel_files


def copy_excel_files(excel_files, destination_folder):
    """Копировать Excel файлы в папку назначения"""
    dest_path = Path(destination_folder)
    
    # Создать папку назначения, если она не существует
    dest_path.mkdir(parents=True, exist_ok=True)
    
    print(f"📦 Копирование файлов в: {destination_folder}")
    print("-" * 40)
    
    copied_count = 0
    
    for file_path in excel_files:
        source_file = Path(file_path)
        dest_file = dest_path / source_file.name
        
        try:
            # Если файл уже существует, добавить номер
            counter = 1
            original_dest = dest_file
            while dest_file.exists():
                stem = original_dest.stem
                suffix = original_dest.suffix
                dest_file = dest_path / f"{stem}_{counter}{suffix}"
                counter += 1
            
            shutil.copy2(file_path, dest_file)
            print(f"✅ {source_file.name} → {dest_file.name}")
            copied_count += 1
            
        except Exception as e:
            print(f"❌ Ошибка при копировании {source_file.name}: {e}")
    
    print("-" * 40)
    print(f"📋 Скопировано файлов: {copied_count} из {len(excel_files)}")


def merge_excel_files(excel_files, output_file, sheet_name=None, add_source_column=True):
    """
    Объединить все Excel файлы в один файл
    
    Args:
        excel_files: Список путей к Excel файлам
        output_file: Путь к выходному файлу
        sheet_name: Имя листа для чтения (если None, берется первый лист)
        add_source_column: Добавлять ли колонку с именем исходного файла
    """
    if not PANDAS_AVAILABLE:
        print("❌ Для объединения файлов требуется установка pandas и openpyxl")
        print("   Выполните: pip install pandas openpyxl")
        return False
    
    if not excel_files:
        print("❌ Нет файлов для объединения")
        return False
    
    print(f"🔗 Объединение {len(excel_files)} Excel файлов...")
    print(f"📁 Выходной файл: {output_file}")
    print("-" * 60)
    
    all_dataframes = []
    successful_files = []
    failed_files = []
    
    for i, file_path in enumerate(excel_files, 1):
        try:
            print(f"📖 Читаю файл {i}/{len(excel_files)}: {os.path.basename(file_path)}")
            
            # Читаем Excel файл
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                # Читаем первый лист
                df = pd.read_excel(file_path, sheet_name=0)
            
            if df.empty:
                print(f"   ⚠️  Файл пустой, пропускаем")
                continue
            
            # Добавляем колонку с именем исходного файла, если нужно
            if add_source_column:
                df['Исходный_файл'] = os.path.basename(file_path)
            
            all_dataframes.append(df)
            successful_files.append(file_path)
            print(f"   ✅ Прочитано строк: {len(df)}")
            
        except Exception as e:
            print(f"   ❌ Ошибка при чтении файла: {e}")
            failed_files.append(file_path)
    
    if not all_dataframes:
        print("❌ Не удалось прочитать ни одного файла")
        return False
    
    print("-" * 60)
    print(f"📊 Успешно прочитано файлов: {len(successful_files)}")
    if failed_files:
        print(f"❌ Файлов с ошибками: {len(failed_files)}")
    
    try:
        # Объединяем все DataFrame'ы
        print("🔗 Объединяю данные...")
        merged_df = pd.concat(all_dataframes, ignore_index=True, sort=False)
        
        # Создаем папку для выходного файла, если нужно
        output_path = Path(output_file)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Сохраняем объединенный файл
        print(f"💾 Сохраняю объединенный файл...")
        merged_df.to_excel(output_file, index=False, engine='openpyxl')
        
        print("-" * 60)
        print(f"✅ Объединение завершено успешно!")
        print(f"📄 Итоговый файл: {output_file}")
        print(f"📊 Общее количество строк: {len(merged_df)}")
        print(f"📈 Количество колонок: {len(merged_df.columns)}")
        
        # Показываем информацию о колонках
        if len(merged_df.columns) <= 10:
            print(f"📋 Колонки: {', '.join(merged_df.columns.tolist())}")
        else:
            print(f"📋 Первые колонки: {', '.join(merged_df.columns.tolist()[:10])}...")
        
        return True
        
    except Exception as e:
        print(f"❌ Ошибка при сохранении объединенного файла: {e}")
        return False


def main():
    """Главная функция"""
    parser = argparse.ArgumentParser(
        description="Сбор и объединение Excel файлов из папки",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры использования:
  python collect_excel_files.py                                    # Показать файлы из папки "файлы"
  python collect_excel_files.py -s "мои_документы"                 # Указать другую папку-источник
  python collect_excel_files.py -d "собранные_excel"               # Скопировать в папку "собранные_excel"
  python collect_excel_files.py -m "объединенный.xlsx"             # Объединить все файлы в один
  python collect_excel_files.py -s "данные" -m "результат.xlsx"    # Объединить файлы из "данные"
  python collect_excel_files.py -m "итог.xlsx" --sheet "Лист1"     # Объединить конкретный лист
  python collect_excel_files.py -m "итог.xlsx" --no-source-column  # Без колонки с именем файла
        """
    )
    
    parser.add_argument(
        '-s', '--source',
        default='файлы',
        help='Папка-источник (по умолчанию: "файлы")'
    )
    
    parser.add_argument(
        '-d', '--destination',
        help='Папка назначения для копирования файлов'
    )
    
    parser.add_argument(
        '-m', '--merge',
        help='Объединить все файлы в один (указать имя выходного файла)'
    )
    
    parser.add_argument(
        '--sheet',
        help='Имя листа для чтения при объединении (по умолчанию: первый лист)'
    )
    
    parser.add_argument(
        '--no-source-column',
        action='store_true',
        help='Не добавлять колонку с именем исходного файла при объединении'
    )
    
    args = parser.parse_args()
    
    # Определить параметры операций
    copy_files = args.destination is not None
    merge_files = args.merge is not None
    add_source_column = not args.no_source_column
    
    collect_excel_files(
        source_folder=args.source,
        destination_folder=args.destination,
        copy_files=copy_files,
        merge_files=merge_files,
        merge_output=args.merge,
        merge_sheet=args.sheet,
        add_source_column=add_source_column
    )


if __name__ == "__main__":
    main()