import zipfile
import os

# with zipfile.ZipFile('Гайка шлицевая молочная, резьбовая, DIN AISI 316.docx', 'r') as zip_ref:
#     for file in zip_ref.namelist():
#         if file.startswith('word/media/'):
#             zip_ref.extract(file, 'output_folder')

import json
import re
from docx import Document


def parse_milk_nut_docx(file_path):
    # Загрузка документа
    doc = Document(file_path)

    # Словарь для хранения данных
    data = {
        "title": "",
        "description": "",
        "application_area": "",
        "standards": [],
        "table_data": [],
        "materials": [],
        "working_params": {},
        "marking_example": "",
        "quality_control": [],
        "packaging": [],
        "warranty": "",
        "supply_info": {}
    }

    # Сбор текстовых данных
    current_section = None
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # Заголовок
        if "Гайка шлицевая молочная" in text and not data["title"]:
            data["title"] = text
        # Раздел 1: Область применения
        elif text.startswith("1.") or "Область применения" in text:
            current_section = "application_area"
            data["application_area"] = text
        elif current_section == "application_area" and not text.startswith("2."):
            data["application_area"] += " " + text
        # Раздел 2: Соответствие стандартам
        elif text.startswith("2.") or "Соответствие стандартам" in text:
            current_section = "standards"
            data["standards"].append(text)
        elif current_section == "standards" and not text.startswith("3."):
            data["standards"].append(text)
        # Раздел 3: Основные параметры (таблица будет обработана отдельно)
        elif text.startswith("3.") or "Основные параметры" in text:
            current_section = "table_section"
        # Раздел 4: Материалы исполнения
        elif text.startswith("4.") or "Материалы исполнения" in text:
            current_section = "materials"
            data["materials"].append(text)
        elif current_section == "materials" and not text.startswith("5."):
            data["materials"].append(text)
        # Раздел 5: Рабочие параметры
        elif text.startswith("5.") or "Рабочие параметры" in text:
            current_section = "working_params"
            data["working_params"]["title"] = text
        elif current_section == "working_params" and not text.startswith("6."):
            if "description" not in data["working_params"]:
                data["working_params"]["description"] = text
            else:
                data["working_params"]["description"] += " " + text
        # Раздел 7: Маркировка
        elif text.startswith("7.") or "Маркировка" in text:
            current_section = "marking"
            data["marking_example"] = text
        elif current_section == "marking" and not text.startswith("8."):
            data["marking_example"] += " " + text
        # Раздел 8: Контроль качества
        elif text.startswith("8.") or "Контроль качества" in text:
            current_section = "quality_control"
            data["quality_control"].append(text)
        elif current_section == "quality_control" and not text.startswith("9."):
            data["quality_control"].append(text)
        # Раздел 9: Упаковка и хранение
        elif text.startswith("9.") or "Упаковка и хранение" in text:
            current_section = "packaging"
            data["packaging"].append(text)
        elif current_section == "packaging" and not text.startswith("10."):
            data["packaging"].append(text)
        # Раздел 10: Гарантия
        elif text.startswith("10.") or "Гарантия" in text:
            current_section = "warranty"
            data["warranty"] = text
        elif current_section == "warranty" and not text.startswith("11."):
            data["warranty"] += " " + text
        # Раздел 11: Сведения о поставке
        elif text.startswith("11.") or "Сведения о поставке" in text:
            current_section = "supply_info"

    # Парсинг таблицы
    for table in doc.tables:
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        if "DN" in headers and "Резьба G" in headers:
            for row in table.rows[1:]:  # пропускаем заголовок
                row_data = [cell.text.strip() for cell in row.cells]
                if len(row_data) == len(headers):
                    table_entry = {headers[i]: row_data[i]
                                   for i in range(len(headers))}
                    data["table_data"].append(table_entry)

    # Удаление дубликатов из списков
    for key in ["standards", "materials", "quality_control", "packaging"]:
        if key in data and isinstance(data[key], list):
            # Убираем пустые строки и оставляем уникальные значения
            data[key] = [item for item in data[key] if item]
            data[key] = list(dict.fromkeys(data[key]))

    return data


def save_to_json(data, output_path):
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


if __name__ == "__main__":
    input_file = "Гайка шлицевая молочная, резьбовая, DIN AISI 316.docx"
    output_file = "гайка_шлицевая_молочная.json"

    try:
        parsed_data = parse_milk_nut_docx(input_file)
        save_to_json(parsed_data, output_file)
        print(f"Данные успешно сохранены в {output_file}")
    except Exception as e:
        print(f"Ошибка: {e}")
