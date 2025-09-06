#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для сбора всех Excel файлов из папки "файлы"
Script to collect all Excel files from "файлы" folder
"""

import os
import shutil
import glob
from pathlib import Path
from datetime import datetime
import argparse


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


def collect_excel_files(source_folder="файлы", destination_folder=None, copy_files=False):
    """
    Собрать все Excel файлы из указанной папки
    
    Args:
        source_folder: Папка-источник (по умолчанию "файлы")
        destination_folder: Папка назначения (если нужно копировать файлы)
        copy_files: Копировать ли файлы в папку назначения
    """
    
    print(f"🔍 Поиск Excel файлов в папке: {source_folder}")
    print("=" * 60)
    
    excel_files = find_excel_files(source_folder)
    
    if not excel_files:
        print(f"📭 В папке '{source_folder}' не найдено Excel файлов")
        return
    
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
    
    # Копирование файлов, если указана папка назначения
    if copy_files and destination_folder:
        copy_excel_files(excel_files, destination_folder)


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


def main():
    """Главная функция"""
    parser = argparse.ArgumentParser(
        description="Сбор Excel файлов из папки",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры использования:
  python collect_excel_files.py                           # Показать файлы из папки "файлы"
  python collect_excel_files.py -s "мои_документы"        # Указать другую папку-источник
  python collect_excel_files.py -d "собранные_excel"      # Скопировать в папку "собранные_excel"
  python collect_excel_files.py -s "данные" -d "результат" # Копировать из "данные" в "результат"
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
    
    args = parser.parse_args()
    
    # Определить, нужно ли копировать файлы
    copy_files = args.destination is not None
    
    collect_excel_files(
        source_folder=args.source,
        destination_folder=args.destination,
        copy_files=copy_files
    )


if __name__ == "__main__":
    main()