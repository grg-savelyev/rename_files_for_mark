import os
import pandas as pd
import shutil
import re


def create_excel_dict(folder_path: str, excel_name: str) -> dict | None:
    """
    Создание словаря на основе данных из файла Excel.

    Args:
        folder_path (str): Путь к папке с файлом Excel.
        excel_name (str): Название файла Excel.

    Returns:
        dict: Словарь, где ключ - название файла из Adobe Indesign, а значение - арт, наименование, цвет, размер.
    """
    try:
        excel_path = os.path.join(folder_path, f'{excel_name}.xlsx')
        df = pd.read_excel(excel_path)
    except FileNotFoundError:
        print(f'Файл Excel "{excel_name}" не найден.')
        return

    return dict(zip(df['id'], df['name'].astype(str).apply(replace_special_chars)))


def replace_special_chars(text: str) -> str:
    """
    Замена спец символов в значении словаря.

    Args:
        text (str): Значение name из Excel.

    Returns:
        str: Исправленное значение в словарь
    """
    return re.sub(r'[:\\?#]', '_', text)


def sort_by_folders(folder_path: str, excel_dict: dict) -> None:
    """
    Сортировка файлов в папки на основе словаря.

    Args:
        folder_path (str): Путь к папке с файлами PDF.
        excel_dict (dict): Словарь с данными из Excel.
    """
    all_files = os.listdir(folder_path)
    pdf_file_names = [file.split('.')[0] for file in all_files if file.endswith('.pdf')]

    for pdf_file_name in pdf_file_names:
        table_value = excel_dict.get(pdf_file_name)
        if table_value is None:
            print(f'Имя файла {pdf_file_name} не найдено в столбце "name" Excel файла')
        else:
            product_code_name = (', '.join(table_value.split(', ')[:2]))
            # Путь к папке, в которую нужно скопировать файл
            new_folder_path = os.path.join(folder_path, product_code_name)
            if not os.path.exists(new_folder_path):  # Если папки не существует, создаем ее
                os.makedirs(new_folder_path)
            shutil.copy(os.path.join(folder_path, f'{pdf_file_name}.pdf'), new_folder_path)
            renaming(new_folder_path, pdf_file_name, table_value)


def renaming(new_folder_path: str, pdf_file_name: str, table_value: str) -> None:
    """
    Переименовывание файла.

    Args:
        new_folder_path (str): Путь к новой папке с файлами.
        pdf_file_name (str): Имя файла.
        table_value (str): Новое имя файла.
    """
    current_file_path = os.path.join(new_folder_path, f'{pdf_file_name}.pdf')
    new_file_path = os.path.join(new_folder_path, f'{table_value}.pdf')
    os.rename(current_file_path, new_file_path)


def main():
    print('Укажите путь к папке для обработки:')
    folder_path = input().strip()
    excel_name = 'InDesign'
    print('Запуск обработки.\n')
    excel_dict = create_excel_dict(folder_path, excel_name)
    if excel_dict is not None:
        sort_by_folders(folder_path, excel_dict)
        print('Обработка завершена')
    else:
        print('Обработка завершена с ошибкой.')


if __name__ == "__main__":
    main()
