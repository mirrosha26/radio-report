#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль для работы с SQLite базой данных
Создает таблицы и сохраняет данные о пунктах
"""

import sqlite3
import os
from pathlib import Path
from datetime import datetime

class DatabaseManager:
    def __init__(self, db_path="points_database.db"):
        """Инициализация менеджера базы данных"""
        self.db_path = db_path
        self.init_database()
    
    def init_database(self):
        """Создает таблицы в базе данных"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                # Создаем таблицу для файлов
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS files (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        filename TEXT NOT NULL,
                        file_path TEXT NOT NULL,
                        file_type TEXT,
                        encoding TEXT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
                
                # Создаем таблицу для пунктов
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS points (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        file_id INTEGER,
                        point_number INTEGER,
                        tag TEXT,
                        seconds INTEGER,
                        content TEXT,
                        short_content TEXT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY (file_id) REFERENCES files (id)
                    )
                ''')
                

                
                conn.commit()
                
        except sqlite3.Error as e:
            print(f"❌ Ошибка создания базы данных: {e}")
    
    def save_file(self, filename, file_path, file_type, encoding):
        """Сохраняет информацию о файле и возвращает его ID"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                # Проверяем, существует ли уже такой файл
                cursor.execute('''
                    SELECT id FROM files WHERE file_path = ?
                ''', (str(file_path),))
                
                existing_file = cursor.fetchone()
                
                if existing_file:
                    # Обновляем существующий файл
                    cursor.execute('''
                        UPDATE files 
                        SET filename = ?, file_type = ?, encoding = ?, created_at = CURRENT_TIMESTAMP
                        WHERE id = ?
                    ''', (filename, file_type, encoding, existing_file[0]))
                    file_id = existing_file[0]
                else:
                    # Создаем новый файл
                    cursor.execute('''
                        INSERT INTO files (filename, file_path, file_type, encoding)
                        VALUES (?, ?, ?, ?)
                    ''', (filename, str(file_path), file_type, encoding))
                    file_id = cursor.lastrowid
                
                conn.commit()
                return file_id
                
        except sqlite3.Error as e:
            print(f"❌ Ошибка сохранения файла: {e}")
            return None
    
    def save_points(self, file_id, points):
        """Сохраняет пункты для указанного файла"""
        if not points:
            return
        
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                # Удаляем старые пункты для этого файла
                cursor.execute('DELETE FROM points WHERE file_id = ?', (file_id,))
                
                # Сохраняем новые пункты
                for i, point in enumerate(points, 1):
                    cursor.execute('''
                        INSERT INTO points (file_id, point_number, tag, seconds, content, short_content)
                        VALUES (?, ?, ?, ?, ?, ?)
                    ''', (
                        file_id,
                        i,
                        point.get('tag', ''),
                        point.get('seconds', 0),
                        point.get('value', ''),
                        point.get('short_content', '')  # Добавляем короткое описание
                    ))
                
                conn.commit()
                
        except sqlite3.Error as e:
            print(f"❌ Ошибка сохранения пунктов: {e}")
    
    def update_point_short_content(self, point_id, short_content):
        """Обновляет короткое описание пункта"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                cursor.execute('''
                    UPDATE points 
                    SET short_content = ? 
                    WHERE id = ?
                ''', (short_content, point_id))
                
                conn.commit()
                return True
                
        except sqlite3.Error as e:
            print(f"❌ Ошибка обновления короткого описания: {e}")
            return False
    
    def get_all_points(self):
        """Получает все пункты из базы данных"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                cursor.execute('''
                    SELECT 
                        f.filename,
                        p.point_number,
                        p.tag,
                        p.seconds,
                        p.content,
                        p.short_content,
                        p.created_at,
                        p.id
                    FROM points p
                    JOIN files f ON p.file_id = f.id
                    ORDER BY f.filename, p.point_number
                ''')
                
                return cursor.fetchall()
                
        except sqlite3.Error as e:
            print(f"❌ Ошибка получения пунктов: {e}")
            return []
    
 