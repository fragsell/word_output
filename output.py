import zipfile
import os

with zipfile.ZipFile('document.docx', 'r') as zip_ref:
    for file in zip_ref.namelist():
        if file.startswith('word/media/'):
            zip_ref.extract(file, 'output_folder')