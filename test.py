from main_full import make_main_full_presentation


#
# ais_file = "Отчет по событиям за 30.08.2025 06.30 - 30.08.2025 15.29.xlsx"
# previous_period = {
#         "ЦАО": 1,
#         "САО": 1,
#         "СВАО": 1,
#         "ВАО": 1,
#         "ЮВАО": 1,
#         "ЮАО": 1,
#         "ЮЗАО": 1,
#         "ЗАО": 1,
#         "СЗАО": 1,
#         "ЗелАО": 1,
#         "ТиНАО": 1,
#         'ГБУ "АВД"': 1,
#         'Иные': 1,
#     }
# make_main_full_presentation(ais_file, "None", previous_period, "02.01.2003", True)




import pandas as pd
import glob

# Путь к папке с Excel-файлами (замени на свой)
files = glob.glob("data/*.xlsx")

all_pairs = pd.DataFrame()

for file in files:
    df = pd.read_excel(file, dtype=str)  # читаем как строки, чтобы не потерять формат
    if {"Наименование события", "Наименование события КОД ОИВ"}.issubset(df.columns):
        pairs = df[["Наименование события", "Наименование события КОД ОИВ"]]
        all_pairs = pd.concat([all_pairs, pairs], ignore_index=True)

# Оставляем уникальные пары
unique_pairs = all_pairs.drop_duplicates()

# Сохраняем в Excel
unique_pairs.to_excel("mapping.xlsx", index=False)
