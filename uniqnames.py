import pandas as pd


def process_csv(file_path, output_excel):
    # Загружаем CSV в DataFrame
    df = pd.read_csv(file_path, encoding='utf-8', sep=";")

    # Получаем уникальные значения и их количество
    unique_counts = df["Наименование события КОД ОИВ"].value_counts().reset_index()
    unique_counts.columns = ["Наименование события КОД ОИВ", "Количество"]

    # Сохраняем в Excel
    unique_counts.to_excel(output_excel, index=False, engine='openpyxl')

    print(f"Результат сохранен в {output_excel}")


# Пример использования
file_path = "reports/Годовой 14.03.2025/tmp_files/Обработанный АИС.csv"  # Укажите путь к вашему CSV-файлу
output_excel = "unique_counts.xlsx"  # Имя выходного Excel-файла
process_csv(file_path, output_excel)
