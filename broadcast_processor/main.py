#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Главный файл системы анализа файлов
Автоматически запускает анализ и предлагает просмотр сводки
"""

import os
from modules.folder_processor import read_folder_path, process_folder
from modules.database_viewer import get_points_summary

def main():
    try:
        
        folder_path = read_folder_path()
        print(f"📂 Папка для анализа: {folder_path}")
        print("=" * 40)
        
        print("🔄 Выполняется анализ...")
        db_manager = process_folder()
        print("=" * 40)
        
        print("📊 Общая сводка:")
        get_points_summary()
                
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        print("👋 Программа завершена")

if __name__ == "__main__":
    main() 