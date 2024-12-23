import os#собирает все эксельки из папки в 1 документ со всеми листами
import pandas as pd
from tkinter import Tk, filedialog


def merge_excel_files(folder_path):
    # Инициализация пустого списка для хранения содержимого DataFrame
    frames = []

    # Перебор всех файлов Excel в указанной директории
    for file in os.listdir(folder_path):
        if file.endswith('.xlsx'):
            file_path = os.path.join(folder_path, file)
            xls = pd.ExcelFile(file_path)

            # Чтение каждого листа в файле
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                df['Имя файла'] = file  # Добавляем имя файла как столбец
                df['Имя листа'] = sheet_name  # Добавляем имя листа как столбец
                frames.append(df)

    # Объединяем все собранные DataFrame, игнорируя индексы и заголовки столбцов
    combined_df = pd.concat(frames, ignore_index=True)

    return combined_df


def select_folder_and_merge():
    # Интерфейс для выбора директории
    root = Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory()
    root.destroy()

    if folder_path:
        result_df = merge_excel_files(folder_path)
        result_df.to_excel('Объединенный_файл.xlsx', index=False)
        print("Файл успешно создан и сохранен как 'Объединенный_файл.xlsx'")
    else:
        print("Папка не выбрана.")


select_folder_and_merge()
