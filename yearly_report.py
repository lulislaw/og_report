import os
from datetime import datetime, timedelta
import shutil
import pandas as pd
import openpyxl
import numpy as np
from python_pptx_text_replacer import TextReplacer
import locale
import re
from pptx_functions import get_top_rows_with_ties, runs_from_pptx, update_diagramms, make_yearly_slides
from str_pptx import keys_weekly_widget, keys_main, keys_table_svod
from xlsx_functions import update_ais_data


def sanitize_sheet_name(name):
    return re.sub(r'[\\/*?:\[\]]', '_', name)[:21]


def better_or(value):
    numeric_value = int(value.replace("\xa0", "").replace(" ", ""))
    allper_value = value
    if numeric_value > 0:
        allper_value = f"↗ {value}"  # Положительное значение
    elif numeric_value < 0:
        allper_value = f"↘ {value}"  # Отрицательное значение
    return allper_value


def fint(x):
    locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')
    return locale._format('%d', x, grouping=True)


def make_yearly_report(ais_file, date, type="Годовой"):
    dist_path = ""
    date_text = f"{date}"
    path_os = f"reports/{type} {date_text}".replace(":", ".")
    if not os.path.exists(path_os):
        os.makedirs(path_os)
    tmp_files_path = os.path.join(dist_path, "reports", f"{type} {date_text}", "tmp_files")
    report_path = os.path.join(dist_path, "reports", f"{type} {date_text}")
    tmp_files_img_path = os.path.join(dist_path, "reports", f"{type} {date_text}", "tmp_files", "img")
    reports_path = os.path.join(dist_path, "reports")
    for path in [tmp_files_path, tmp_files_img_path, reports_path]:
        if not os.path.exists(path):
            print(f"Создание директории: {path}")
            os.makedirs(path, exist_ok=True)
    allert = []
    order = ["ЦАО", "САО", "СВАО", "ВАО", "ЮВАО", "ЮАО", "ЮЗАО", "ЗАО", "СЗАО", "ЗелАО", "ТиНАО", "Общий итог"]

    if ".xlsx" in ais_file:
        df_ais = pd.read_excel(ais_file)
    elif ".csv" in ais_file:
        df_ais = pd.read_csv(ais_file)
    else:
        print("Файл не правильный")
        return
    update_ais_data(df_ais)
    df_ais["Причина"] = df_ais["Наименование события"]
    if "Наименование события КОД ОИВ" in df_ais.columns:
        df_ais["Наименование события"] = df_ais["Наименование события КОД ОИВ"]
    ais_event_name = "Наименование события"
    df_ais = df_ais[df_ais["Наименование события"].str.strip().notna()]
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

    try:
        # Пробуем сохранить в Excel
        df_ais.to_excel(f"{tmp_files_path}/Обработанный АИС.xlsx", index=False)
        print("Файл успешно сохранен в Excel.")
    except ValueError as e:
        # Если ошибка связана с ограничением размера, сохраняем в CSV
        if "This sheet is too large" in str(e) or "DataFrame size" in str(e):
            print("Файл слишком большой для Excel, сохраняем в CSV...")
            df_ais.to_csv(f"{tmp_files_path}/Обработанный АИС.csv", index=False, sep=";")  # CSV с разделителем `;`
            print(f"Файл сохранен в CSV")
        else:
            raise  # Если ошибка другая, пробрасываем её дальше
    allert.extend(calculation_month(df_ais, report_path, date))
    return allert


def calculation_month(df_ais, report_path, date):
    allert = []
    tmp_files_path = f"{report_path}/tmp_files"
    presentation_maket = "makets/presentation/yearly_presentation.pptx"
    output_file = f"{tmp_files_path}/Результаты.xlsx"
    order = ["ЦАО", "САО", "СВАО", "ВАО", "ЮВАО", "ЮАО", "ЮЗАО", "ЗАО", "СЗАО", "ЗелАО", "ТиНАО"]
    population_moscow = {
        "ЦАО": 774430,
        "САО": 1217909,
        "СВАО": 1455811,
        "ВАО": 1508678,
        "ЮВАО": 1515787,
        "ЮАО": 1768752,
        "ЮЗАО": 1435550,
        "ЗАО": 1399932,
        "СЗАО": 1039596,
        "ЗелАО": 270527,
        "ТиНАО": 762831,
        "Общий итог": 13149803
    }

    df_filtered = df_ais[df_ais["Тип объекта"] == "МКД"]
    # df_filtered = df_ais[df_ais["Тип объекта"] != "МКД"] # для огх нужно еще изменить позицию слайда

    # Подсчет количества объектов для каждого события
    event_counts = df_filtered.groupby("Наименование события").size().reset_index(name="Общее количество объектов")

    # Сортируем события по количеству объектов (по убыванию)
    sorted_events = event_counts.sort_values(by="Общее количество объектов", ascending=False)["Наименование события"]

    # Список округов, который должен быть в каждом DataFrame
    order = ["ЦАО", "САО", "СВАО", "ВАО", "ЮВАО", "ЮАО", "ЮЗАО", "ЗАО", "СЗАО", "ЗелАО", "ТиНАО"]

    # Словарь для хранения отдельных DataFrame по каждому событию
    event_dfs = {}
    event_dfs_reason = {}
    event_dfs_regions = {}

    for event in sorted_events:
        # Фильтруем данные по текущему событию
        df_event = df_filtered[df_filtered["Наименование события"] == event]

        # Группируем по округу и считаем количество объектов
        df_grouped = df_event.groupby("Округ").size().reset_index(name="Количество объектов")

        # Приводим к единому списку округов (merge с полным списком округов)
        df_full = pd.DataFrame({"Округ": order}).merge(df_grouped, on="Округ", how="left").fillna(0)

        # Преобразуем количество объектов в int (убираем float от NaN)
        df_full["Количество объектов"] = df_full["Количество объектов"].astype(int)

        df_grouped_reason = df_event.groupby("Причина").size().reset_index(name="Количество объектов")
        df_grouped_reason["Процент"] = (df_grouped_reason["Количество объектов"] / df_grouped_reason[
            "Количество объектов"].sum()) * 100
        df_grouped_reason["Процент"] = df_grouped_reason["Процент"].fillna(0).round(0)/100  # Округляем процент до 2 знаков

        # Заполняем пропущенные значения 0 и приводим к int
        df_grouped_reason["Количество объектов"] = df_grouped_reason["Количество объектов"].fillna(0).astype(int)
        # Сохраняем результат

        df_regions = df_full.copy()


        # Считаем процент от общего числа обращений в каждом округе
        total_objects = df_regions["Количество объектов"].sum()
        df_regions["Процент от общего"] = ((df_regions["Количество объектов"] / total_objects) * 100) \
                                              .fillna(0).round(1).astype(str).str.replace('.',
                                                                                          ',') + "%"  # ✅ Заменяем точки на запятые

        # Считаем количество обращений на 1000 жителей
        df_regions["Обращений на 1000 жителей"] = df_regions["Округ"].map(lambda x:
                                                                          df_regions.loc[df_regions[
                                                                                             "Округ"] == x, "Количество объектов"].values[
                                                                              0] / population_moscow.get(x, 1) * 1000)
        df_regions["Обращений на 1000 жителей"] = df_regions["Обращений на 1000 жителей"].fillna(0).apply(
            lambda x: f"{x:.3f}".replace('.', ','))

        df_regions = df_regions.drop(df_regions.columns[1], axis=1)
        # Сохраняем результат
        event_dfs_reason[event] = df_grouped_reason
        event_dfs[event] = df_full
        event_dfs_regions[event] = df_regions

    pptx_file = f"{tmp_files_path}/Годовой ОГ.pptx"
    shutil.copy(presentation_maket, pptx_file)

    make_yearly_slides(pptx_file, event_dfs, event_dfs_reason, event_dfs_regions)
    count = 0
    with pd.ExcelWriter(output_file, mode="w", engine="openpyxl") as writer:
        for event, df in event_dfs.items():
            count += 1
            sheet_name = sanitize_sheet_name(event)  # Очищаем имя листа
            df.to_excel(writer, sheet_name=f"Округ_{count}{sheet_name}", index=False)
            sheet_name = sanitize_sheet_name(event)  # Очищаем имя листа
            event_dfs_reason[event].to_excel(writer, sheet_name=f"Причина_{count}{sheet_name}", index=False)
            event_dfs_regions[event].to_excel(writer, sheet_name=f"Reg_{count}{sheet_name}", index=False)


    try:
        locale.setlocale(locale.LC_ALL, 'C')
    except locale.Error:
        pass
    return allert
