#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль для обработки папок и файлов
Извлекает пункты и сохраняет в SQLite базу данных
"""

import os
import re
import yaml
from pathlib import Path
from modules.database import DatabaseManager

def read_folder_path():
    """Читает путь к папке из config.yaml"""
    try:
        with open('config.yaml', 'r', encoding='utf-8') as file:
            config = yaml.safe_load(file)
            return config.get('folder_path', 'ИЮЛЬ')
    except FileNotFoundError:
        print("❌ Файл config.yaml не найден")
        return 'ИЮЛЬ'
    except yaml.YAMLError as e:
        print(f"❌ Ошибка чтения config.yaml: {e}")
        return 'ИЮЛЬ'

def parse_text_into_points(text_content):
    """Разбивает текст на пункты и создает список словарей"""
    if not text_content or not isinstance(text_content, str):
        return []
    
    # Ищем пункты по форматам: число.(тег-время) текст или число) (тег-время) текст
    # Паттерны: 1.(губер-57) текст, 2.(губер-63) текст, 3) (губер-49) текст
    # Также поддерживаем формат с пробелом: 1.(губер 57) текст
    # Учитываем различные варианты пробелов: перед скобками, за скобками, внутри
    patterns = [
        # Универсальные паттерны с гибкой обработкой пробелов
        # Формат с точкой: 1. (тег-время) или 1.(тег время)
        r'(\d+)\.\s*\(\s*([^-()]+?)\s*-\s*(\d+)\s*\)\s*(.*?)(?=\d+[.)]\s*\(|$)',  # тег-время
        r'(\d+)\.\s*\(\s*([^-()0-9\s]+(?:\s+[^-()0-9\s]+)*)\s+(\d+)\s*\)\s*(.*?)(?=\d+[.)]\s*\(|$)',  # тег время
        
        # Формат со скобкой: 1) (тег-время) или 1)(тег время)  
        r'(\d+)\)\s*\(\s*([^-()]+?)\s*-\s*(\d+)\s*\)\s*(.*?)(?=\d+[.)]\s*\(|$)',  # тег-время
        r'(\d+)\)\s*\(\s*([^-()0-9\s]+(?:\s+[^-()0-9\s]+)*)\s+(\d+)\s*\)\s*(.*?)(?=\d+[.)]\s*\(|$)',  # тег время
        
        # Запасные паттерны для последнего пункта (без lookahead)
        r'(\d+)\.\s*\(\s*([^-()]+?)\s*-\s*(\d+)\s*\)\s*(.*?)$',
        r'(\d+)\.\s*\(\s*([^-()0-9\s]+(?:\s+[^-()0-9\s]+)*)\s+(\d+)\s*\)\s*(.*?)$',
        r'(\d+)\)\s*\(\s*([^-()]+?)\s*-\s*(\d+)\s*\)\s*(.*?)$',
        r'(\d+)\)\s*\(\s*([^-()0-9\s]+(?:\s+[^-()0-9\s]+)*)\s+(\d+)\s*\)\s*(.*?)$'
    ]
    
    points = []
    found_positions = set()  # Отслеживаем уже найденные позиции, чтобы избежать дубликатов
    
    for i, pattern in enumerate(patterns):
        matches = re.finditer(pattern, text_content, re.DOTALL)
        
        for match in matches:
            start_pos = match.start()
            
            # Проверяем, что эту позицию мы еще не обрабатывали
            if start_pos not in found_positions:
                found_positions.add(start_pos)
                
                point_number = int(match.group(1))
                tag = match.group(2).strip()  # тег (губер, админ и т.д.)
                seconds = int(match.group(3))  # время из скобок
                point_text = match.group(4).strip()  # текст пункта
                
                # Проверяем, что пункт содержит достаточную длину
                if len(point_text) >= 10:
                    points.append({
                        'value': point_text,
                        'seconds': seconds,
                        'tag': tag,
                        'point_number': point_number
                    })
    
    # Сортируем по номеру пункта
    points.sort(key=lambda x: x.get('point_number', 0))
    
    return points

def read_odt_content_simple(file_path):
    """Читает .odt файл как ZIP архив и извлекает текст из content.xml"""
    try:
        import zipfile
        with zipfile.ZipFile(file_path, 'r') as zip_file:
            if 'content.xml' in zip_file.namelist():
                content = zip_file.read('content.xml').decode('utf-8')
                # Убираем XML теги, оставляем только текст
                import re
                content = re.sub(r'<[^>]+>', '', content)
                content = re.sub(r'\s+', ' ', content).strip()
                return content, "odt_xml"
            else:
                return f"❌ В .odt файле не найден content.xml", "odt_error"
    except Exception as e:
        return f"❌ Ошибка чтения .odt файла: {e}", "odt_error"

def read_odt_content_advanced(file_path):
    """Читает .odt файл используя odfpy для чистого извлечения текста"""
    try:
        from odf import text, teletype
        from odf.opendocument import load
        doc = load(str(file_path))
        content = teletype.extractText(doc)
        return content, "odt_clean"
    except ImportError:
        # Если odfpy не установлен, используем простой метод
        return read_odt_content_simple(file_path)
    except Exception as e:
        # Если что-то пошло не так, пробуем простой метод
        return read_odt_content_simple(file_path)

def read_docx_content(file_path):
    """Читает .docx файл используя python-docx"""
    try:
        from docx import Document
        doc = Document(file_path)
        content = []
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                content.append(paragraph.text)
        
        if content:
            return '\n'.join(content), "docx_clean"
        else:
            return "", "docx_empty"
    except ImportError:
        return f"❌ Библиотека python-docx не установлена", "docx_error"
    except Exception as e:
        return f"❌ Ошибка чтения .docx файла: {e}", "docx_error"

def read_doc_content(file_path):
    """Читает .doc файл используя mammoth или docx2txt"""
    try:
        import mammoth
        with open(file_path, "rb") as docx_file:
            result = mammoth.extract_raw_text(docx_file)
            if result.value:
                return result.value, "doc_clean"
            else:
                return "", "doc_empty"
    except ImportError:
        try:
            import docx2txt
            content = docx2txt.process(str(file_path))
            if content:
                return content, "doc_clean"
            else:
                return "", "doc_empty"
        except ImportError:
            return f"❌ Библиотеки mammoth и docx2txt не установлены", "doc_error"
        except Exception as e:
            return f"❌ Ошибка чтения .doc файла через docx2txt: {e}", "doc_error"
    except Exception as e:
        return f"❌ Ошибка чтения .doc файла через mammoth: {e}", "doc_error"

def read_file_content(file_path):
    """Читает содержимое файла в зависимости от его типа"""
    file_ext = file_path.suffix.lower()
    
    # Для .odt файлов используем специальную обработку
    if file_ext == '.odt':
        return read_odt_content_advanced(file_path)
    
    # Для .docx файлов используем специальную обработку
    elif file_ext == '.docx':
        return read_docx_content(file_path)
    
    # Для .doc файлов используем специальную обработку
    elif file_ext == '.doc':
        return read_doc_content(file_path)
    
    # Для остальных файлов пробуем обычное чтение
    try:
        # Пробуем разные кодировки
        encodings = ['utf-8', 'cp1251', 'windows-1251', 'latin-1', 'iso-8859-1']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as file:
                    content = file.read()
                    return content, encoding
            except UnicodeDecodeError:
                continue
        
        # Если не удалось прочитать как текст, пробуем как бинарный
        try:
            with open(file_path, 'rb') as file:
                content = file.read()
                return f"[Бинарный файл, размер: {len(content)} байт]", "binary"
        except:
            return f"❌ Не удалось прочитать файл", "error"
            
    except Exception as e:
        return f"❌ Ошибка чтения файла: {e}", "error"

def read_files_from_folder(folder_path, db_manager):
    """Читает файлы из указанной папки и анализирует их содержимое"""
    folder = Path(folder_path)
    
    if not folder.exists():
        print(f"❌ Ошибка: папка '{folder_path}' не найдена")
        return
    
    if not folder.is_dir():
        print(f"❌ Ошибка: '{folder_path}' не является папкой")
        return
    
    total_files = 0
    total_points = 0
    folder_results = []
    
    # Проходим по всем элементам в папке
    for item in folder.iterdir():
        if item.is_dir():
            print(f"📁 Папка: {item.name}")
            
            # Получаем список файлов в папке
            files = list(item.glob('*'))
            files = [f for f in files if f.is_file()]  # Только файлы, не папки
            
            if files:
                folder_files = 0
                folder_points = 0
                
                # Обрабатываем каждый файл
                for file in files:
                    # Читаем содержимое файла
                    content, encoding = read_file_content(file)
                    
                    # Сохраняем файл в базу данных
                    file_id = db_manager.save_file(
                        filename=file.name,
                        file_path=file,
                        file_type=file.suffix.lower(),
                        encoding=encoding
                    )
                    
                    if file_id:
                        # Разбиваем текст на пункты
                        points = parse_text_into_points(content)
                        if points:
                            # Сохраняем пункты в базу данных
                            db_manager.save_points(file_id, points)
                            folder_points += len(points)
                            total_points += len(points)
                            print(f"  ✅ {file.name}: {len(points)} пунктов")
                        else:
                            print(f"  ⚠️ {file.name}: пункты не найдены")
                        
                        folder_files += 1
                        total_files += 1
                
                # Результаты по папке
                folder_results.append({
                    'name': item.name,
                    'files': folder_files,
                    'points': folder_points
                })
                
                print(f"  📊 Итого: {folder_files} файлов, {folder_points} пунктов")
                
            else:
                print("  📭 Папка пуста")
        
        elif item.is_file():
            # Читаем содержимое файла
            content, encoding = read_file_content(item)
            
            # Сохраняем файл в базу данных
            file_id = db_manager.save_file(
                filename=item.name,
                file_path=item,
                file_type=item.suffix.lower(),
                encoding=encoding
            )
            
            if file_id:
                # Разбиваем текст на пункты
                points = parse_text_into_points(content)
                if points:
                    # Сохраняем пункты в базу данных
                    db_manager.save_points(file_id, points)
                    total_points += len(points)
                    print(f"   ✅ {item.name} → {len(points)} пунктов")
                else:
                    print(f"   ⚠️  {item.name} → пункты не найдены")
                
                total_files += 1
    
    # Итоговая сводка
    print("=" * 40)
    print(f"📁 Папок: {len([item for item in folder.iterdir() if item.is_dir()])}")
    print(f"📄 Файлов: {total_files}")
    print(f"📋 Пунктов: {total_points}")
    
    if folder_results:
        print("📂 Детализация:")
        for result in folder_results:
            if result['points'] > 0:
                print(f"  🗂️ {result['name']}: {result['files']} файлов, {result['points']} пунктов")
            else:
                print(f"  🗂️ {result['name']}: {result['files']} файлов, пункты не найдены")
    
    return total_files, total_points

def process_folder():
    """Основная функция для обработки папки"""
    # Удаляем существующую базу данных
    db_path = "points_database.db"
    if os.path.exists(db_path):
        try:
            os.remove(db_path)
        except Exception as e:
            print(f"❌ Ошибка удаления БД: {e}")
    
    # Инициализируем новую базу данных
    db_manager = DatabaseManager()
    
    # Читаем путь к папке из конфига
    folder_path = read_folder_path()
    
    # Анализируем файлы
    total_files, total_points = read_files_from_folder(folder_path, db_manager)
    
    # Проверяем данные в БД
    points = db_manager.get_all_points()
    if points:
        if len(points) != total_points:
            print(f"⚠️  Несоответствие: найдено {total_points}, сохранено {len(points)}")
    else:
        print("❌ В базе данных нет пунктов")
    
    return db_manager 