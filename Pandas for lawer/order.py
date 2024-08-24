import pandas as pd
from docx import Document
import os

def create_document(row, column_names, template_path, output_folder, counter):
    doc = Document(template_path)

    for paragraph in doc.paragraphs:
        text = paragraph.text  # читаем абзац 1 раз
        for k, v in column_names.items():  # и в цикле
            text = text.replace(k, row[v])  # заменяем маркеры на значения, сопоставляя из словаря
        if text != paragraph.text:  # проверяем, если в исходном тексте абзаца что-то изменлось
            paragraph.text = text  # то записываем результат в документ

    fio = row['[ФИО должника]']
    output_path = os.path.join(output_folder, f"Заявление_{fio}.docx")
    doc.save(output_path)
    print(f"Документ для {fio} успешно создан.")
    counter[0] += 1  # приращиваем счетчик успешно сохраненных документов


column_names = {
    '[ФИО_должника]': '[ФИО должника]', '[место_жительства]': '[место жительства]', '[дата_и_место рождения]': '[дата и место рождения]', '[данные_паспорта]': '[данные паспорта]',
    '[размер_задолженности]': '[размер задолженности]', '[пошлина]': '[пошлина]', '[свидетельство]': '[свидетельство]'
}  # словарь для сопоставления маркеров в тексте и названий полей в таблице

excel_file = r'Book2.xlsx'
template_path = r'Заявление.docx'
output_folder = r'output_docs'

os.makedirs(output_folder, exist_ok=True)

counter = [0]
pd.read_excel(excel_file).astype(str).apply(create_document, axis=1, args=(
column_names, template_path, output_folder, counter))  # читаем данные, сразу преобразуем все в текст и применяем к каждой строке фрейма функцию с доп аргументами
print(f"{counter[0]} документов успешно создано и сохранены в папку {output_folder}.")


