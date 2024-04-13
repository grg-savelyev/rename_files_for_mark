import gspread
import os
import pandas as pd
import shutil
from datetime import datetime


def create_labels_dict(df):
    labels_dict = dict(zip(df['barcode'],
                           zip(df['id_label'], df['name_label'],
                               df['id_barcode'], df['name_barcode'],
                               df['id_masterbox'], df['name_master box'],
                               df['name_folder'].str.replace(': ', '_'), df['chestny_znak'])))
    return dict_check(labels_dict)


def create_care_labels_dict(df):
    care_labels_dict = dict(zip(df['id_care_label'], zip(df['name'], df['name_folder'].str.replace(': ', '_'))))
    return dict_check(care_labels_dict)


def dict_check(data_dict):
    all_filled = all(all(value) for value in data_dict.values())
    if not all_filled:
        print(f'Словарь {data_dict} не заполнен или содержит пустые строки. Операция прервана.')
        return None
    return data_dict


def start_time():
    current_time = datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    return current_time


def sort_by_folders(start_folder_path: str, folder_path: str, data_dict: dict, label_type: str) -> None:
    if label_type == 'label':
        for key, value in data_dict.items():
            product_folder_name = value[6]
            # Путь к папке, в которую нужно скопировать файл
            new_folder_path = os.path.join(start_folder_path, product_folder_name)
            if not os.path.exists(new_folder_path):  # Если папки не существует, создаем ее
                os.makedirs(new_folder_path)
                shutil.copy(os.path.join(folder_path, 'label 1_45x100.pdf'), new_folder_path)

            #  навесные бирки
            shutil.copy(os.path.join(folder_path, f'{value[0]}.pdf'), new_folder_path)
            renaming(new_folder_path, f'{value[0]}', f'{value[1]}')

            #  мастер короба
            shutil.copy(os.path.join(folder_path, f'{value[4]}.pdf'), new_folder_path)
            renaming(new_folder_path, f'{value[4]}', f'{value[5]}')

            #  баркод
            if not value[7].lower() == 'нет':
                if not os.path.exists(os.path.join(new_folder_path, 'label 3_45x70.pdf')):
                    shutil.copy(os.path.join(folder_path, 'label 3_45x70.pdf'), new_folder_path)
            else:
                shutil.copy(os.path.join(folder_path, f'{value[2]}.pdf'), new_folder_path)
                renaming(new_folder_path, f'{value[2]}', f'{value[3]}')
        print('Навесные отработали')

    if label_type == 'care_label':
        for key, value in data_dict.items():
            product_folder_name = value[1]
            # Путь к папке, в которую нужно скопировать файл
            new_folder_path = os.path.join(start_folder_path, product_folder_name)
            if not os.path.exists(new_folder_path):  # Если папки не существует, создаем ее
                print(f'Папка {product_folder_name} не найдена и создана новая. '
                      f'Проверьте имена папок в таблице на идентичность')
                os.makedirs(new_folder_path)
            shutil.copy(os.path.join(folder_path, f'{key}.pdf'), new_folder_path)
            renaming(new_folder_path, f'{key}', f'{value[0]}')


def renaming(new_folder_path: str, pdf_file_name: str, product_code_name: str) -> None:
    """Переименовывание файла."""
    current_file_path = os.path.join(new_folder_path, f'{pdf_file_name}.pdf')
    new_file_path = os.path.join(new_folder_path, f'{product_code_name}.pdf')
    os.rename(current_file_path, new_file_path)


def main():
    print('Укажите путь к папке для обработки:')
    folder_path = r'C:\Users\Lenovo\Desktop\fleece 2025 full pack\fleece 2025 full pack\#_PREPRINT'  # input().strip()
    print('Запуск обработки.\n')
    start_folder_path = os.path.join(folder_path, start_time())
    os.makedirs(start_folder_path)

    gc = gspread.service_account(filename='creds.json')

    # навесные бирки, баркоды, мастер короба
    worksheet_label = gc.open('InDesign').worksheet('label')
    df = pd.DataFrame(worksheet_label.get_all_records())
    labels_dict: dict = create_labels_dict(df)
    if labels_dict:
        sort_by_folders(start_folder_path, folder_path, labels_dict, 'label')

    # вшивные бирки
    worksheet_care_label = gc.open('InDesign').worksheet('care_label')
    df2 = pd.DataFrame(worksheet_care_label.get_all_records())
    care_labels_dict: dict = create_care_labels_dict(df2)
    if care_labels_dict:
        sort_by_folders(start_folder_path, folder_path, care_labels_dict, 'care_label')

    print('Обработка завершена')


if __name__ == "__main__":
    main()
