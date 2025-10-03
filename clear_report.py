import locale
import os
import numpy as np
import pandas as pd


def fint(x):
    locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')
    return locale._format('%d', x, grouping=True)





def update_ais_data(df_ais, addresses_file="adresses.xlsx"):
    try:
        required_cols = ["Адрес объекта", "Район", "Округ"]
        for col in required_cols:
            if col not in df_ais.columns:
                print(f"В файле нет колонки '{col}'!")
                return None
        if os.path.exists(addresses_file):
            df_exclude = pd.read_excel(addresses_file)
            if "Адрес объекта" in df_exclude.columns:
                excluded_addresses = set(df_exclude["Адрес объекта"].astype(str).str.lower())

                df_ais.loc[
                    df_ais["Адрес объекта"].astype(str).str.lower().isin(excluded_addresses), ["Район", "Округ"]] = [
                    "Щербинка", "ТиНАО"]
            else:
                print(f"В файле {addresses_file} нет колонки 'Адрес объекта'!")
        else:
            print(f"Файл {addresses_file} не найден, замены не выполняются.")
        print("Обновление данных завершено.")
        df_ais = fix_districts(df_ais)
        return df_ais
    except Exception as e:
        print(f"Ошибка при обработке файла: {e}")
        return None



def fix_districts(df_main):


    mask_autodor = df_main[
                       'Ответственный'] == 'Государственное Бюджетное Учреждение города Москвы "Автомобильные дороги"'
    df_main.loc[mask_autodor, 'Округ'] = 'ГБУ "АВД"'

    df_main['Округ'] = df_main['Округ'].replace(['', 'Общегородской'], 'Иные')
    df_main['Округ'] = df_main['Округ'].fillna('Иные')
    return df_main



def make_month_report(ais_file, date, spec, type="nedelniymji"):
    is_cod_oiv = True
    dist_path = ""
    date_text = f"{date}"
    path_os = f"reports/{type} {date_text}".replace(":", ".")
    if not os.path.exists(path_os):
        os.makedirs(path_os)
    tmp_files_path = os.path.join(dist_path, "reports", f"{type} {date_text}", "tmp_files")
    tmp_files_img_path = os.path.join(dist_path, "reports", f"{type} {date_text}", "tmp_files", "img")
    reports_path = os.path.join(dist_path, "reports")
    for path in [tmp_files_path, tmp_files_img_path, reports_path]:
        if not os.path.exists(path):
            print(f"Создание директории: {path}")
            os.makedirs(path, exist_ok=True)
    allert = []
    order = ["ЦАО", "САО", "СВАО", "ВАО", "ЮВАО", "ЮАО", "ЮЗАО", "ЗАО", "СЗАО", "ЗелАО", "ТиНАО", "Общий итог"]
    if "xlsx" in ais_file:
        df_ais = pd.read_excel(ais_file)
    else:
        df_ais = pd.read_csv(ais_file)
    df_ais = update_ais_data(df_ais, "adresses.xlsx")
    if "Наименование события КОД ОИВ" in df_ais.columns and is_cod_oiv:
        df_ais["Наименование события"] = df_ais["Наименование события КОД ОИВ"]
    ais_event_name = "Наименование события"
    df_ais = df_ais[df_ais[ais_event_name].str.strip().notna()]
    status_mapping = {"Новое": "В работе", "Отменено": "Закрыто"}
    df_ais["Округ"] = df_ais["Округ"].replace({"НАО": "ТиНАО", "ТАО": "ТиНАО"})
    df_ais["Статус во внешней системе"] = df_ais["Статус во внешней системе"].replace(status_mapping)
    df_ais.loc[df_ais["Район"].isin(["Ново-Переделкино", "Солнцево"]), "Округ"] = "ЗАО"
    df_ais.loc[df_ais["Район"] == "Внуково", "Округ"] = "ТиНАО"
    df_ais["Район"] = df_ais["Район"].str.split(",").str[0].str.strip()
    df_ais = df_ais[df_ais["Ответственный"].str.strip().notna()]
    df_ais["Округ"] = df_ais["Округ"].replace(np.nan, "")
    df_ais = df_ais[~df_ais["Округ"].str.strip().isin(["", "Общегородской"])]

    df_ais["Дата создания"] = pd.to_datetime(df_ais["Дата создания во внешней системе"],
                                             format="%d.%m.%Y, %H:%M:%S").dt.date

    df_ais["Дата (разделенная)"] = df_ais["Дата создания"]

    unique_dates = sorted(df_ais["Дата создания"].unique())

    mid_index = 30

    max_date = max(unique_dates)

    if len(unique_dates) >= 2:
        df_ais.loc[df_ais["Дата создания"].isin(unique_dates[:mid_index]), "Дата (разделенная)"] = max_date
        df_ais.loc[df_ais["Дата создания"].isin(unique_dates[mid_index:]), "Дата (разделенная)"] = max_date
    else:
        allert.append("Количество дат меньше 2!!!")
    date = f"{max_date}".replace(":", ".")
    df_ais.to_excel(f"{tmp_files_path}/Обработанный АИС {spec} {date}.xlsx", index=False)
    return


ais_file = "Отчет_по_событиям_за_15_09_2025_00_00_15_09_2025_23_59_1.xlsx"
make_month_report(ais_file, "29.09.2025", "mnt_mji")
