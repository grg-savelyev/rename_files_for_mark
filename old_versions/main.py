import gspread
import os
import pandas as pd
import shutil
from datetime import datetime


def create_dicts(df):
    labels_dict = dict(zip(df['id_label'], zip(df['name_label'], df['name_folder'].replace(':', '_'))))
    barcodes_dict = dict(
        zip(df['id_barcode'], zip(df['name_barcode'], df['honest_sign'], df['name_folder'].replace(':', '_'))))

    data_dicts = [labels_dict, barcodes_dict]

    for data_dict in data_dicts:
        all_filled = all(all(value) for value in data_dict.values())
        if not all_filled:
            dict_name = 'labels_dict' if 'labels_dict' in data_dict else 'barcodes_dict'
            print(f'Словарь {dict_name} не заполнен или содержит пустые строки. Операция прервана.')
            return None

    return data_dicts


def sort_by_folders(folder_path: str, data_dicts: list) -> None:
    """Сортировка файлов в папки на основе словаря."""
    pdf_file_names = [file.split('.')[0] for file in os.listdir(folder_path) if file.endswith('.pdf')]
    second_labels = [file for file in pdf_file_names if '_'.join(file.split('_')[:-1]) == 'second_label']
    barcodes = [file for file in pdf_file_names if file.split('_')[0] == 'barcode']

    start_folder_path = os.path.join(folder_path, start_time())
    os.makedirs(start_folder_path)

    for data_dict in data_dicts:
        dict_name = 'barcodes_dict' if all(
            key.split('_')[0] == 'barcode' for key in data_dict.keys()) else 'labels_dict'

        if dict_name == 'labels_dict':
            for second_label in second_labels:
                table_value = data_dict.get(second_label)
                if table_value is not None:
                    product_name = table_value[0]
                    product_folder_name = table_value[1].rstrip()

                # Путь к папке, в которую нужно скопировать файл
                    new_folder_path = os.path.join(start_folder_path, product_folder_name)
                    if not os.path.exists(new_folder_path):  # Если папки не существует, создаем ее
                        os.makedirs(new_folder_path)
                        shutil.copy(os.path.join(folder_path, 'first label.pdf'), new_folder_path)
                    shutil.copy(os.path.join(folder_path, f'{second_label}.pdf'), new_folder_path)
                    renaming(new_folder_path, second_label, product_name)

        else:
            for barcode in barcodes:
                table_value = data_dict.get(barcode)
                if table_value is not None:
                    product_name = table_value[0]
                    honest_sign = table_value[1]
                    product_folder_name = table_value[-1].rstrip()

                    if honest_sign == 'НЕТ':
                        new_folder_path = os.path.join(start_folder_path, product_folder_name)
                        if not os.path.exists(new_folder_path):  # Если папки не существует, выводим сообщение
                            print(f'Папки {product_folder_name} не существует. Проверьте исходники')
                        shutil.copy(os.path.join(folder_path, f'{barcode}.pdf'), new_folder_path)
                        renaming(new_folder_path, barcode, product_name)
                    else:
                        new_folder_path = os.path.join(start_folder_path, product_folder_name)
                        if not os.path.exists(os.path.join(new_folder_path, 'third label.pdf')):
                            shutil.copy(os.path.join(folder_path, 'third label.pdf'), new_folder_path)


def start_time():
    current_time = datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    return current_time


def renaming(new_folder_path: str, pdf_file_name: str, product_code_name: str) -> None:
    """Переименовывание файла."""
    current_file_path = os.path.join(new_folder_path, f'{pdf_file_name}.pdf')
    new_file_path = os.path.join(new_folder_path, f'{product_code_name}.pdf')
    os.rename(current_file_path, new_file_path)


def main():
    print('Укажите путь к папке для обработки:')
    folder_path = r'C:\Users\Lenovo\Desktop\Вязанка_март\#_PRINT'  # input().strip()
    print('Запуск обработки.\n')

    gc = gspread.service_account(filename='creds.json')
    wks = gc.open('МАТРИЦА ЛЕГПРОМ').worksheet('Data_mark')
    df = pd.DataFrame(wks.get_all_records())
    data_dicts: list = create_dicts(df)

    sort_by_folders(folder_path, data_dicts)

    print('Обработка завершена')


if __name__ == "__main__":
    main()
