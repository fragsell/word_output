import zipfile
import os

# with zipfile.ZipFile('Гайка шлицевая молочная, резьбовая, DIN AISI 316.docx', 'r') as zip_ref:
#     for file in zip_ref.namelist():
#         if file.startswith('word/media/'):
#             zip_ref.extract(file, 'output_folder')

import json
import re
from docx import Document

def parse_docx_by_section_titles(file_path):
    """Парсит документ по названиям разделов"""
    doc = Document(file_path)
    
    # Собираем все параграфы
    all_paragraphs = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            all_paragraphs.append(text)
    
    # Заголовок документа
    title = all_paragraphs[0] if all_paragraphs else ""
    
    # Список названий разделов в правильном порядке
    section_titles = [
        "Область применения",
        "Соответствие стандартам и спецификации", 
        "Основные параметры",
        "Материалы исполнения",
        "Рабочие параметры",
        "Конструктивные особенности",
        "Маркировка",
        "Контроль качества",
        "Упаковка и хранение",
        "Гарантия",
        "Сведения о поставке"
    ]
    
    # Создаем структуру для хранения данных
    sections = {}
    current_section = None
    section_content = []
    
    # Проходим по всем параграфам, начиная с "ТЕХНИЧЕСКИЙ ПАСПОРТ"
    start_index = 0
    for i, p in enumerate(all_paragraphs):
        if "ТЕХНИЧЕСКИЙ ПАСПОРТ" in p:
            start_index = i + 1
            break
    
    # Собираем данные по разделам
    for i in range(start_index, len(all_paragraphs)):
        p = all_paragraphs[i]
        
        # Проверяем, является ли это названием раздела
        is_section_title = False
        section_name = None
        
        for title in section_titles:
            if p == title:
                is_section_title = True
                section_name = title
                break
        
        if is_section_title:
            # Сохраняем предыдущий раздел
            if current_section is not None and section_content:
                section_text = "\n".join(section_content).strip()
                sections[current_section] = {
                    "title": current_section,
                    "content": section_text
                }
            
            # Начинаем новый раздел
            current_section = section_name
            section_content = []
        elif current_section is not None:
            # Добавляем текст к текущему разделу
            section_content.append(p)
    
    # Сохраняем последний раздел
    if current_section is not None and section_content:
        section_text = "\n".join(section_content).strip()
        sections[current_section] = {
            "title": current_section,
            "content": section_text
        }
    
    # Извлекаем таблицу
    table_data = extract_table_data(doc)
    
    # Особый случай для раздела "Основные параметры" - добавляем обозначения отдельно
    if "Основные параметры" in sections:
        # Разделяем описание и обозначения
        content = sections["Основные параметры"]["content"]
        if "Обозначения:" in content:
            parts = content.split("Обозначения:", 1)
            description = parts[0].strip()
            notations = "Обозначения:" + parts[1] if len(parts) > 1 else ""
            
            sections["Основные параметры"]["content"] = description
            sections["Основные параметры"]["notations"] = notations
    
    return {
        "document_title": all_paragraphs[0] if all_paragraphs else "",
        "sections": sections,
        "table_data": table_data
    }

def extract_table_data(doc):
    """Извлекает данные из таблиц"""
    for table in doc.tables:
        rows = []
        headers = []
        
        for i, row in enumerate(table.rows):
            cells = [cell.text.strip() for cell in row.cells]
            
            if i == 0:
                headers = cells
            else:
                row_dict = {}
                for j, cell_text in enumerate(cells):
                    if j < len(headers) and headers[j]:
                        # Очищаем названия заголовков
                        header = headers[j].replace('\n', ' ').strip()
                        row_dict[header] = cell_text
                
                if row_dict and any(row_dict.values()):
                    rows.append(row_dict)
        
        if rows and any("DN" in key for key in (headers + list(rows[0].keys()) if rows else [])):
            return rows
    
    return []

def clean_text(text):
    """Очищает текст"""
    if not text:
        return ""
    
    # Заменяем переносы строк в середине предложений
    text = re.sub(r'([а-яa-z])\n([а-яa-z])', r'\1 \2', text, flags=re.IGNORECASE)
    
    # Удаляем лишние пробелы
    text = re.sub(r'\s+', ' ', text)
    
    return text.strip()

def save_to_json(data, output_file):
    """Сохраняет данные в JSON"""
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def main():
    input_file = "Гайка шлицевая молочная, резьбовая, DIN AISI 316.docx"
    output_file = "гайка_полные_данные.json"
    
    print("=" * 80)
    print("ИЗВЛЕЧЕНИЕ ДАННЫХ ИЗ ДОКУМЕНТА")
    print("=" * 80)
    
    try:
        # Парсим документ
        data = parse_docx_by_section_titles(input_file)
        
        # Выводим результаты
        print(f"\nЗаголовок документа: {data['document_title']}")
        print(f"\nНайдено разделов: {len(data['sections'])}")
        
        # Список разделов в правильном порядке для вывода
        ordered_sections = [
            "Область применения",
            "Соответствие стандартам и спецификации", 
            "Основные параметры",
            "Материалы исполнения",
            "Рабочие параметры",
            "Конструктивные особенности",
            "Маркировка",
            "Контроль качества",
            "Упаковка и хранение",
            "Гарантия",
            "Сведения о поставке"
        ]
        
        print("\nСОДЕРЖАНИЕ РАЗДЕЛОВ:")
        print("-" * 80)
        
        for section_title in ordered_sections:
            if section_title in data["sections"]:
                section = data["sections"][section_title]
                content = section.get("content", "")
                
                print(f"\n{section_title}:")
                if content:
                    # Очищаем и форматируем текст
                    clean_content = clean_text(content)
                    char_count = len(clean_content)
                    
                    print(f"  Длина: {char_count} символов")
                    
                    # Показываем превью
                    preview_lines = clean_content.split('. ')
                    if len(preview_lines) > 0:
                        print(f"  Превью: {preview_lines[0][:150]}...")
                else:
                    print("  [Нет содержания]")
                
                # Если есть обозначения
                if "notations" in section:
                    notations = section["notations"]
                    print(f"  Обозначения: {len(notations)} символов")
            else:
                print(f"\n{section_title}: [Раздел не найден]")
        
        # Выводим таблицу
        table_data = data.get("table_data", [])
        if table_data:
            print(f"\n" + "=" * 80)
            print(f"ТАБЛИЦА ПАРАМЕТРОВ: {len(table_data)} строк")
            print("-" * 80)
            
            # Красиво выводим таблицу
            if table_data:
                headers = list(table_data[0].keys())
                print(" | ".join(headers))
                print("-" * 60)
                
                for row in table_data[:5]:  # Первые 5 строк
                    values = [str(row.get(h, "")) for h in headers]
                    print(" | ".join(values))
                
                if len(table_data) > 5:
                    print(f"... и еще {len(table_data) - 5} строк")
        
        # Сохраняем в JSON
        save_to_json(data, output_file)
        
        print(f"\n" + "=" * 80)
        print(f"✓ Данные сохранены в файл: {output_file}")
        
        # Дополнительная информация
        print(f"\nДОПОЛНИТЕЛЬНАЯ ИНФОРМАЦИЯ:")
        print(f"- Заголовок документа: {data['document_title']}")
        print(f"- Всего разделов: {len(data['sections'])}")
        print(f"- Строк в таблице: {len(table_data)}")
        
        # Проверяем, все ли разделы найдены
        missing_sections = [s for s in ordered_sections if s not in data["sections"]]
        if missing_sections:
            print(f"- Пропущенные разделы: {', '.join(missing_sections)}")
        
    except Exception as e:
        print(f"✗ Ошибка: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()