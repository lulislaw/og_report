import os

import pandas as pd
import openpyxl
import numpy as np
from str_pptx import keys_table, keys_main
from python_pptx_text_replacer import TextReplacer
import locale
from pptx_functions import runs_from_pptx, get_top_rows_with_ties
from xlsx_functions import update_ais_data


def fint(x):
    locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')
    return locale._format('%d', x, grouping=True)


def make_daily_presentation(ais_file, previous_period, date, config):
    dist_path = ""
    date_text = f"{date}"
    path_os = f"reports/{date_text}".replace(":", ".")
    if not os.path.exists(path_os):
        os.makedirs(path_os)
    tmp_files_path = os.path.join(dist_path, "reports", f"{date_text}", "tmp_files")
    tmp_files_img_path = os.path.join(dist_path, "reports", f"{date_text}", "tmp_files", "img")
    reports_path = os.path.join(dist_path, "reports")
    for path in [tmp_files_path, tmp_files_img_path, reports_path]:
        if not os.path.exists(path):
            print(f"Создание директории: {path}")
            os.makedirs(path, exist_ok=True)

    # Население Москвы по округам
    allert = []
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
    previous_period["Общий итог"] = None
    # Заданный порядок округов
    order = ["ЦАО", "САО", "СВАО", "ВАО", "ЮВАО", "ЮАО", "ЮЗАО", "ЗАО", "СЗАО", "ЗелАО", "ТиНАО", "Общий итог"]

    # Данные для предыдущего периода

    presentation_maket = "makets/presentation/daily-presentation.pptx"
    replacer = TextReplacer(presentation_maket, slides='',
                            tables=True, charts=True, textframes=True)

    replacer.replace_text([("date", date_text)])
    # Загрузка данных AIS
    df_ais = pd.read_excel(ais_file)
    update_ais_data(df_ais)
    # Фильтрация и замены в df_ais
    if "Наименование события КОД ОИВ" in df_ais.columns:
        df_ais["Наименование события"] = df_ais["Наименование события КОД ОИВ"]
    ais_event_name = "Наименование события"
    df_ais = df_ais[df_ais["Наименование события"].str.strip().notna()]
    status_mapping = {"Новое": "В работе", "Отменено": "Закрыто"}
    df_ais["Статус во внешней системе"] = df_ais["Статус во внешней системе"].replace(status_mapping)
    df_ais["Округ"] = df_ais["Округ"].replace({"НАО": "ТиНАО", "ТАО": "ТиНАО"})
    # Присваиваем округа
    df_ais.loc[df_ais["Район"].isin(["Ново-Переделкино", "Солнцево"]), "Округ"] = "ЗАО"
    df_ais.loc[df_ais["Район"] == "Внуково", "Округ"] = "ТиНАО"
    df_ais["Район"] = df_ais["Район"].str.split(",").str[0].str.strip()

    # Удаляем пустых ответственных и округов
    df_ais = df_ais[df_ais["Ответственный"].str.strip().notna()]
    df_ais["Округ"] = df_ais["Округ"].replace(np.nan, "")
    df_ais = df_ais[~df_ais["Округ"].str.strip().isin(["", "Общегородской"])]

    # Сохраняем обработанный AIS
    if config:
        # Сохраняем обработанный AIS
        df_ais.to_excel(f"{tmp_files_path}/Обработанный АИС_до_очистки.xlsx", index=False)
        df_ais = drop_random_by_config(df_ais, "resource/config.yaml", 52)
    df_ais.to_excel(f"{tmp_files_path}/Обработанный АИС.xlsx", index=False)

    # Группировка данных AIS
    df_summary = df_ais.groupby("Округ").size().reset_index(name="Отчетный период")

    # Добавляем "Предыдущий период"
    df_summary["Предыдущий период"] = df_summary["Округ"].map(previous_period)

    # Считаем общую сумму и добавляем проценты
    total_sum = df_summary["Отчетный период"].sum()
    df_summary["%"] = (df_summary["Отчетный период"] / total_sum) * 100
    df_summary["*1000"] = df_summary.apply(
        lambda row: (row["Отчетный период"] * 1000) / population_moscow.get(row["Округ"], 1), axis=1
    )

    # Добавляем строку "Общий итог"
    total_row = pd.DataFrame({
        "Округ": ["Общий итог"],
        "Предыдущий период": [sum(value for value in previous_period.values() if value is not None)],
        "Отчетный период": [total_sum],
        "%": [100],
        "*1000": (total_sum * 1000) / population_moscow.get("Общий итог", 1)
    })
    df_summary = pd.concat([df_summary, total_row], ignore_index=True)

    # Приводим к заданному порядку
    df_summary["Округ"] = pd.Categorical(df_summary["Округ"], categories=order, ordered=True)
    df_summary = df_summary.sort_values("Округ").reset_index(drop=True)

    # Пересчёт процентов и итогов
    total_sum_updated = df_summary[df_summary["Округ"] != "Общий итог"]["Отчетный период"].sum()
    df_summary["%"] = (df_summary["Отчетный период"] / total_sum_updated) * 100
    df_summary["*1000"] = df_summary.apply(
        lambda row: (row["Отчетный период"] * 1000) / population_moscow.get(row["Округ"], 1), axis=1
    )

    # Форматирование процентов и *1000
    df_summary["%"] = df_summary["%"].apply(lambda x: "100%" if x == 100 else f"{x:.2f}%".replace('.', ','))
    df_summary["*1000"] = df_summary["*1000"].apply(lambda x: f"{x:.2f}".replace('.', ','))

    # Обновление строки "Общий итог"
    df_summary.loc[df_summary["Округ"] == "Общий итог", "Предыдущий период"] = sum(
        value for value in previous_period.values() if value is not None)
    df_summary.loc[df_summary["Округ"] == "Общий итог", "Отчетный период"] = total_sum_updated
    df_summary.loc[df_summary["Округ"] == "Общий итог", "%"] = "100%"
    udel_all = (total_sum_updated * 1000) / population_moscow["Общий итог"]
    df_summary.loc[df_summary["Округ"] == "Общий итог", "*1000"] = f"{udel_all:.2f}".replace('.', ',')

    change_value = total_sum_updated - sum(value for value in previous_period.values() if value is not None)
    previous_total = sum(value for value in previous_period.values() if value is not None)
    change_percentage = (change_value / previous_total) * 100

    # Округляем процент изменения и корректируем для случая 0
    if change_percentage == 0:
        rounded_change_percentage = change_percentage
    else:
        rounded_change_percentage = round(change_percentage)
    if rounded_change_percentage == 0:
        rounded_change_percentage = 1 if change_percentage > 0 else -1

    change_row = pd.DataFrame({
        "Округ": ["Изменение"],
        "Предыдущий период": ["-"],
        "Отчетный период": [f"{rounded_change_percentage}%".replace('.', ',')],
        "%": ["-"],
        "*1000": ["-"]  # Указываем "-" для этой колонки, так как она не применима
    })

    df_combined = pd.concat([df_summary, change_row], ignore_index=True)
    # Приводим к заданному порядку
    columns_order = ["Округ", "Предыдущий период", "Отчетный период", "%", "*1000"]
    df_combined = df_combined[columns_order]

    replacer_list = []
    for i, row in enumerate(keys_main):  # Перебираем ключи построчно
        value = ""
        for j, key in enumerate(row):
            if j == 0:  # "Предыдущий период"
                previous_period = df_combined.iloc[i]["Предыдущий период"]
                value = previous_period
            elif j == 1:  # "Отчетный период"
                report_period = df_combined.iloc[i]["Отчетный период"]
                value = report_period
            elif j == 2:  # "%"
                value = df_combined.iloc[i]["%"]
            elif j == 3:  # "*1000"
                value = df_combined.iloc[i]["*1000"]

            # Применяем функцию fint, если значение числовое
            if isinstance(value, (int, float)):
                value = fint(int(value))

            if j == 1 and isinstance(previous_period, (int, float)) and isinstance(report_period, (int, float)):
                if report_period < previous_period:
                    value = f"↘ {value}"
                elif report_period > previous_period:
                    value = f"↗ {value}"

            # Добавляем в список
            replacer_list.append((key, value))
    all_value = df_combined.iloc[-2]["Отчетный период"]  # Предпоследнее значение "Отчетный период"
    allper_value = df_combined.iloc[-1]["Отчетный период"]  # Последнее значение "Изменение"

    numeric_value = float(allper_value.replace('%', '').strip())

    # Определяем знак числа и добавляем стрелку
    if numeric_value > 0:
        allper_value = f"↗ {allper_value}"  # Положительное значение
    elif numeric_value < 0:
        allper_value = f"↘ {allper_value}"  # Отрицательное значение

    # Применяем форматирование, если числовые значения
    if isinstance(all_value, (int, float)):
        all_value = fint(int(all_value))

    # Добавляем в список replacer_list
    replacer_list.append(("*all*", all_value))
    replacer_list.append(("*allper*", allper_value))
    replacer.replace_text(replacer_list)

    # Сохраняем итоговые данные
    with pd.ExcelWriter(f"{tmp_files_path}/Результат АИС.xlsx", mode="w") as writer:
        df_combined.to_excel(writer, sheet_name="Обновленные Итоги", index=False)
    # Создаем сводную таблицу
    pivot_table = pd.pivot_table(
        df_ais,
        index=ais_event_name,
        columns="Округ",
        aggfunc="size",
        fill_value=0
    )

    # Добавляем итоговые суммы по строкам (справа)
    pivot_table["Итого по строке"] = pivot_table.sum(axis=1)

    # Приводим столбцы к заданному порядку округов
    ordered_columns = [col for col in order if col in pivot_table.columns] + ["Итого по строке"]
    pivot_table = pivot_table[ordered_columns]

    # Копируем оригинальный список для вычисления "Иные"
    original_pivot_table = pivot_table.copy()

    # Сортируем строки по убыванию итогов
    pivot_table_sorted = pivot_table.sort_values(by="Итого по строке", ascending=False)

    # Ограничиваем таблицу до топ-10 строк
    top_10 = pivot_table_sorted.iloc[:10]
    pivot_table_with_heating = pd.concat([top_10])
    pivot_table_with_heating = pivot_table_with_heating.sort_values(by="Итого по строке", ascending=False)

    # Вычисляем строку "Иные" из оригинального списка
    remaining_rows = original_pivot_table.drop(index=pivot_table_with_heating.index, errors="ignore")
    other_row = remaining_rows.sum(axis=0)  # Суммируем оставшиеся строки
    other_row.name = "Иные"
    pivot_table_with_heating = pd.concat([pivot_table_with_heating, pd.DataFrame([other_row])])

    # Добавляем строку "Итого по столбцу"
    column_sums = pivot_table_with_heating.sum(axis=0)
    column_sums.name = "Итого по столбцу"
    pivot_table_with_heating = pd.concat([pivot_table_with_heating, pd.DataFrame([column_sums])])

    # Приводим порядок столбцов к изначальному
    final_table = pivot_table_with_heating[ordered_columns]

    df_sorted = final_table.reset_index()
    top_themes = df_sorted.iloc[:3]

    # Формируем список пар (key, value)
    table_list = []
    for i, row in enumerate(top_themes.itertuples(), start=1):
        table_list.append((f"*theme{i}*", str(row[1])))  # Тема как строка
        table_list.append((f"*top{i}*", fint(int(row[-1]))))  # П

    # Выводим результат

    replacer.replace_text(table_list)
    df_sorted.columns.values[0] = "Тема"  # Задаем имя первой колонке
    top_disctricts_dict = {}
    # Пройдемся по топ-3 темам
    for i, row in df_sorted.iloc[:3].iterrows():  # Сортированный DataFrame с топ-3 темами
        theme = row["Тема"]  # Название темы

        # Преобразуем округа и их значения в список кортежей (округ, значение)
        district_values = [(district, row[district]) for district in row.index[1:-1]]

        # Определяем топ-3 округа с учетом равных значений
        top_districts = get_top_rows_with_ties(district_values, top_n=3)
        districts_string = ", ".join(top_districts)
        replacer.replace_text([(f"*ao{i + 1}*", districts_string)])
        top_lst = []
        for district in top_districts:
            top_lst.append((district, fint(int(row[district]))))
        top_disctricts_dict[f"ao{i + 1}"] = top_lst

    with pd.ExcelWriter(f"{tmp_files_path}/Результат АИС.xlsx", mode="a", engine="openpyxl",
                        if_sheet_exists="replace") as writer:
        final_table.to_excel(writer, sheet_name="Сводная таблица")

    final_table = final_table.reset_index()  # Сбрасываем индекс, чтобы "Наименование события КОД ОИВ" стало колонкой
    keys_short_table = keys_table.copy()
    keys_short_table.pop(-2)
    # Преобразуем ключи в одномерный список для всех ячеек
    flat_keys = [key for row in keys_short_table for key in row]

    # Преобразуем данные таблицы в одномерный список
    flat_values = final_table.to_numpy().flatten()

    # Создаем словарь
    table_dict = dict(zip(flat_keys, flat_values))
    table_list = [
        (
            key,
            "Пуск отопления**" if value == "Пуск отопления" else
            fint(int(value)) if isinstance(value, (int, float)) and value.is_integer() else str(value)
        )
        for key, value in table_dict.items()
    ]

    replacer.replace_text(table_list)

    df_ais["Статус во внешней системе"] = df_ais["Статус во внешней системе"].replace("Закрыта", "Закрыто")

    # Подсчет количества записей в df_ais по колонке "Статус во внешней системе"
    ais_status_count = df_ais["Статус во внешней системе"].value_counts().reset_index()
    ais_status_count.columns = ["Статус", "Количество (df_ais)"]

    # Подсчет количества записей в df_edc по колонке "Статус"

    # Проверяем наличие статусов, отличных от "Закрыто" и "В работе"
    valid_statuses = {"Закрыто", "В работе"}
    invalid_ais_statuses = set(ais_status_count["Статус"]) - valid_statuses

    if invalid_ais_statuses:
        print("Предупреждение: обнаружены статусы, отличные от 'Закрыто' и 'В работе'.")
        if invalid_ais_statuses:
            allert.append(f"Неизвестные статусы в df_ais: {invalid_ais_statuses}")

    # Преобразуем значения в целочисленные (где применимо)
    ais_status_count["Количество (df_ais)"] = ais_status_count["Количество (df_ais)"].astype(int)

    ais_status_count["Сумма по статусу"] = (ais_status_count["Количество (df_ais)"])
    closed_total = 0  # Для статуса "Закрыто"
    in_progress_total = 0  # Для статуса "В работе"
    for _, row in ais_status_count.iterrows():
        if row["Статус"] == "Закрыто":
            closed_total = row["Сумма по статусу"]
        elif row["Статус"] == "В работе":
            in_progress_total = row["Сумма по статусу"]
    replacer.replace_text([("*clos*", fint(closed_total)), ("*work*", fint(in_progress_total))])

    with pd.ExcelWriter(f"{tmp_files_path}/Результат АИС.xlsx", engine="openpyxl", mode="a",
                        if_sheet_exists="replace") as writer:
        ais_status_count.to_excel(writer, sheet_name="Статусы", index=False)

    replacer.replace_text([(" ", " ")])
    replacer.write_presentation_to_file(f"{tmp_files_path}/ОтчетБезЦвета.pptx")

    file_pptx = f"{tmp_files_path}/ОтчетБезЦвета.pptx"

    file_path_save = f"{path_os}/Суточный ОГ на {date}.pptx".replace(":", ".")
    runs_from_pptx(file_pptx, top_disctricts_dict, "daily").save(file_path_save)
    try:
        locale.setlocale(locale.LC_ALL, 'C')
    except locale.Error:
        pass
    return allert
