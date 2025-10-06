import os
from datetime import datetime, timedelta

import pandas as pd
import numpy as np
from python_pptx_text_replacer import TextReplacer
import locale

from pptx_functions import get_top_rows_with_ties, runs_from_pptx, update_diagramms
from str_pptx import keys_weekly_widget, keys_main, keys_table_svod, keys_table_weekly
from xlsx_functions import update_ais_data, population

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


def make_month_report(ais_file, date, mid_index, type="Недельный"):

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
    if ".xlsx" in ais_file:
        df_ais = pd.read_excel(ais_file)
    else:
        df_ais = pd.read_csv(ais_file)
    df_ais = update_ais_data(df_ais)
    print(df_ais.columns)
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

    # Создаем новую колонку "Дата (разделенная)" с теми же значениями, что и "Дата создания"
    df_ais["Дата (разделенная)"] = df_ais["Дата создания"]
    # Получаем список 14 уникальных дат

    unique_dates = sorted(df_ais["Дата создания"].unique())

    for i, unique_date in enumerate(unique_dates):
        print(i, unique_date)
    print(mid_index, "mid", unique_dates[mid_index])

    # Определяем максимальную дату
    max_date = max(unique_dates)

    # Приравниваем первую половину дат к средней дате, а вторую половину к максимальной
    if len(unique_dates) >= 2:
        mid_date = unique_dates[mid_index - 1]  # Берем дату середины (наименьшая из второй половины)

        df_ais.loc[df_ais["Дата создания"].isin(unique_dates[:mid_index]), "Дата (разделенная)"] = mid_date
        df_ais.loc[df_ais["Дата создания"].isin(unique_dates[mid_index:]), "Дата (разделенная)"] = max_date
    else:
        allert.append("Количество дат меньше 2!!!")
    df_ais.to_excel(f"{tmp_files_path}/Обработанный АИС.xlsx", index=False)
    # Группируем данные по "Дата (разделенная)" и "Статус во внешней системе"
    summary_df = df_ais.groupby(["Дата (разделенная)", "Статус во внешней системе"]).size().unstack(fill_value=0)

    # Добавляем колонку с общим количеством записей в каждом периоде
    summary_df["Всего записей"] = summary_df.sum(axis=1)

    # Вычисляем процент изменения между периодами
    summary_df["Процент изменения"] = ((summary_df.iloc[1] - summary_df.iloc[0]) / summary_df.iloc[1]) * 100
    allert.extend(calculation_month(df_ais, report_path, date))
    return allert


def calculation_month(df_ais, report_path, date):
    population_moscow = population()
    allert = []
    tmp_files_path = f"{report_path}/tmp_files"
    presentation_maket = "makets/presentation/weekly-full-presentation.pptx"
    replacer = TextReplacer(presentation_maket, slides='',
                            tables=True, charts=True, textframes=True)
    output_file = f"{tmp_files_path}/Результаты.xlsx"
    order = ["ЦАО", "САО", "СВАО", "ВАО", "ЮВАО", "ЮАО", "ЮЗАО", "ЗАО", "СЗАО", "ЗелАО", "ТиНАО", 'ГБУ "АВД"', "Иные",
             "Общий итог"]

    date_obj = datetime.strptime(date, "%d.%m.%Y")
    today_date_str = date_obj.strftime("%d.%m.%Y")
    date_obj -= timedelta(days=1)
    last_date_str = str(date_obj.strftime("%d.%m.%Y"))
    date_obj -= timedelta(days=7)
    middle_date_str = str(date_obj.strftime("%d.%m.%Y"))
    date_obj -= timedelta(days=6)
    earl_date_str = str(date_obj.strftime("%d.%m.%Y"))
    replacer.replace_text([("*allperioddate*", f"{earl_date_str} - {last_date_str}"),
                           ("*nowperioddate*", f"{middle_date_str} - {last_date_str}")])

    summary_df = df_ais.groupby(["Дата (разделенная)", "Статус во внешней системе"]).size().unstack(fill_value=0)

    # Добавляем колонку с общим количеством записей в каждом периоде
    summary_df["Всего записей"] = summary_df.sum(axis=1)

    # Сортируем по убыванию даты
    summary_df = summary_df.sort_index(ascending=False)

    # Вычисляем процент изменения между последними двумя периодами
    summary_df["Процент изменения"] = ((summary_df.iloc[0]["Всего записей"] - summary_df.iloc[1]["Всего записей"]) /
                                       summary_df.iloc[1]["Всего записей"]) * 100
    flat_keys = [key for row in keys_weekly_widget for key in row]

    # Преобразуем данные таблицы в одномерный список
    flat_values = summary_df.to_numpy().flatten()  # Пропускаем дату
    percent = flat_values[-1]
    # Создаем словарь
    table_dict = dict(zip(flat_keys, flat_values))
    # Формируем список с форматированием значений
    table_list = [
        (
            key,
            fint(int(value)) if isinstance(value, (int, float)) and float(value).is_integer() else
            f"{round(value)}%" if isinstance(value, (int, float)) else str(value)
        )
        for key, value in table_dict.items()
    ]
    if "-" in f"{percent}":
        percent = f"↘ {round(percent)}%"
    else:
        percent = f"↗ {round(percent)}%"
    table_list.append(("*pmap*", percent))
    replacer.replace_text(table_list)

    # Определяем последний период
    latest_period = df_ais["Дата (разделенная)"].max()
    latest_period_data = df_ais[df_ais["Дата (разделенная)"] == latest_period]
    earliest_period = df_ais["Дата (разделенная)"].min()
    earliest_period_data = df_ais[df_ais["Дата (разделенная)"] == earliest_period]
    # Создаем сводную таблицу
    latest_period_data.to_excel(f"{tmp_files_path}/последнийпериод.xlsx")
    earliest_period_data.to_excel(f"{tmp_files_path}/ранийпериод.xlsx")
    pivot_table = pd.pivot_table(
        latest_period_data,
        index="Наименование события",
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
    df_sorted.columns.values[0] = "Тема"  # Задаем имя первой колонке
    max_districts = max(
        len(get_top_rows_with_ties([(district, row[district]) for district in row.index[1:-1]], top_n=3)) for _, row in
        df_sorted.iloc[:3].iterrows())
    table_data = []
    table_list = []
    for i, row in enumerate(df_sorted.itertuples(), start=1, ):
        if i > 4: break
        theme = str(row[1])
        theme_count = int(row[-1])
        table_list.append((f"*theme{i}*", theme))
        table_list.append((f"*t{i}cnt*", fint(theme_count)))
        theme_count_earliest = (earliest_period_data["Наименование события"] == theme).sum()
        inwork_early = ((earliest_period_data["Наименование события"] == theme) &
                        (earliest_period_data["Статус во внешней системе"] == "В работе")).sum()
        closed_early = ((earliest_period_data["Наименование события"] == theme) &
                        (earliest_period_data["Статус во внешней системе"] == "Закрыто")).sum()
        inwork_late = ((latest_period_data["Наименование события"] == theme) &
                       (latest_period_data["Статус во внешней системе"] == "В работе")).sum()
        closed_late = ((latest_period_data["Наименование события"] == theme) &
                       (latest_period_data["Статус во внешней системе"] == "Закрыто")).sum()
        table_list.append((f"*t{i}cc*", fint(closed_early)))
        table_list.append((f"*t{i}cn*", fint(closed_late)))
        table_list.append((f"*t{i}wc*", fint(inwork_early)))
        table_list.append((f"*t{i}wn*", fint(inwork_late)))
        if theme_count_earliest != 0:
            percent = round(((theme_count - theme_count_earliest) / theme_count_earliest) * 100)
        else:
            percent = round(1 * 100)
        percent = f"↘ {percent}%" if percent < 0 else f"↗ {percent}%"
        table_list.append((f"*t{i}ce*", fint(theme_count_earliest)))
        table_list.append((f"*t{i}prc*", percent))
    replacer.replace_text(table_list)
    # Пройдемся по топ-3 темам
    for i, row in df_sorted.iloc[:3].iterrows():  # Сортированный DataFrame с топ-3 темами
        theme = row["Тема"]  # Название темы
        # Преобразуем округа и их значения в список кортежей (округ, значение)
        district_values = [(district, row[district]) for district in row.index[1:-1]]

        # Определяем топ-3 округа с учетом равных значений
        top_districts = get_top_rows_with_ties(district_values, top_n=3)
        while len(top_districts) < 3:
            top_districts.append("")

            # Формируем строку для таблицы
        table_data.append([theme] + top_districts)
        districts_string = ", ".join(top_districts)
        replacer.replace_text([(f"*t{i + 1}ao*", districts_string)])
    flat_keys = [key for row in keys_table_weekly for key in row]
    final_table = final_table.reset_index()
    # Преобразуем данные таблицы в одномерный список
    flat_values = final_table.to_numpy().flatten()

    # Создаем словарь
    table_dict = dict(zip(flat_keys, flat_values))
    table_list = [
        (
            key,
            fint(int(value)) if isinstance(value, (int, float)) and value.is_integer() else str(value)
        )
        for key, value in table_dict.items()
    ]
    replacer.replace_text(table_list)
    columns = ["Тема"] + [f"Округ {i + 1}" for i in range(max_districts)]
    df_table_corrected = pd.DataFrame(table_data, columns=columns)
    # Сохраняем результат в Excel на новый лист

    # Группировка данных AIS
    df_summary = latest_period_data.groupby("Округ").size().reset_index(name="Отчетный период")
    # Преобразуем данные предыдущего периода в словарь {Округ: Количество записей}
    previous_period = earliest_period_data.groupby("Округ").size().to_dict()
    print(previous_period)
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
    print(df_summary, "df_summary")
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
    print(change_row, "change_row")
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
    all_value = df_summary.iloc[-2]["Отчетный период"]  # Предпоследнее значение "Отчетный период"
    allper_value = df_summary.iloc[-1]["Отчетный период"]  # Последнее значение "Изменение"
    numeric_value = float(str(allper_value).replace('%', '').strip())

    # Определяем знак числа и добавляем стрелку
    if numeric_value > 0:
        allper_value = f"↗ {allper_value}"  # Положительное значение
    elif numeric_value < 0:
        allper_value = f"↘ {allper_value}"  # Отрицательное значение

    # Применяем форматирование, если числовые значения
    if isinstance(all_value, (int, float)):
        all_value = fint(int(all_value))

    # Добавляем в список replacer_list
    replacer.replace_text(replacer_list)

    event_tables = {}
    diagramms_data = {}
    # Проходим по строкам df_table_corrected (список топ-3 округов для каждой темы)
    counter = 0
    for _, row in df_table_corrected.iterrows():
        event_name = row["Тема"]  # Получаем название события
        districts = row[1:].dropna().str.strip().tolist()  # Убираем пробелы и NaN

        comparison_data = []

        for district in districts:
            # Определяем топ-3 ответственных за поздний период
            top_responsible_latest = (
                latest_period_data[
                    (latest_period_data["Округ"].str.strip() == district) &
                    (latest_period_data["Наименование события"] == event_name)
                    ]
                .groupby("Ответственный")["Наименование события"]
                .count()
                .nlargest(3)
            )

            # **Получаем те же топ-3 в раннем периоде**
            top_responsible_earliest = (
                earliest_period_data[
                    (earliest_period_data["Округ"].str.strip() == district) &
                    (earliest_period_data["Наименование события"] == event_name)
                    ]
                .groupby("Ответственный")["Наименование события"]
                .count()
            )

            # **Собираем данные**
            for responsible in top_responsible_latest.index:
                latest_count = top_responsible_latest.get(responsible, 0)
                earliest_count = top_responsible_earliest.get(responsible, 0)  # Берем данные за ранний период или 0
                difference = latest_count - earliest_count  # Вычисляем разницу

                comparison_data.append([
                    district, responsible, earliest_count, latest_count, difference
                ])

        # Создаем DataFrame с данными
        df_event_comparison = pd.DataFrame(
            comparison_data,
            columns=["Округ", "Ответственный", "Кол-во (ранний)", "Кол-во (поздний)", "Разница"]
        )

        # **Добавляем строку с названием темы (пустые колонки в остальных местах)**
        header_row = pd.DataFrame([[event_name, "", "", "", ""]], columns=df_event_comparison.columns)
        df_event_comparison["Округ"] = pd.Categorical(df_event_comparison["Округ"], categories=order,
                                                      ordered=True)
        df_event_comparison = df_event_comparison.sort_values("Округ").reset_index(drop=True)
        districts_lst = df_event_comparison["Округ"].unique().tolist()

        for i, district in enumerate(districts_lst, start=1):
            # Filter rows for the current district
            df_slice = df_event_comparison[df_event_comparison["Округ"] == district]

            # Convert the relevant columns to lists
            respons = df_slice["Ответственный"].tolist()
            early_c = df_slice["Кол-во (ранний)"].tolist()

            late_c = df_slice["Кол-во (поздний)"].tolist()
            diff_c = df_slice["Разница"].tolist()
            replacer.replace_text(
                [(str(f"*t{counter}ao{i}*"), str(district)), (str(f"*t{counter}as{i}*"), str(f"{fint(sum(late_c))}"))])

            # Put these into your diagramms_data dict
            diagramms_data[f"theme{counter}_okrug{i}_resp"] = respons
            diagramms_data[f"theme{counter}_okrug{i}_early"] = early_c
            diagramms_data[f"theme{counter}_okrug{i}_late"] = late_c

            # Replace text for differences
            data_diff_counts_replace = [
                (f"*t{counter}o{i}d{idx}*", better_or(fint(value)))
                for idx, value in enumerate(diff_c)
            ]
            replacer.replace_text(data_diff_counts_replace)

        counter += 1

        # **Добавляем строку с итоговой разницей по всему событию**
        total_difference = df_event_comparison["Разница"].sum()
        total_row = pd.DataFrame([["Итого", "", "", "", total_difference]], columns=df_event_comparison.columns)

        # **Объединяем заголовок, основную таблицу и итоговую строку**
        df_event_comparison = pd.concat([header_row, df_event_comparison, total_row], ignore_index=True)

        # Сохраняем DataFrame в словаре
        event_tables[event_name] = df_event_comparison

        event_district_tables = {}

        # Получаем уникальный список всех округов
        all_districts = sorted(
            set(earliest_period_data["Округ"].str.strip()) | set(latest_period_data["Округ"].str.strip()))

        # Проходим по всем темам
        for ev_num, event_name in enumerate(df_table_corrected["Тема"].unique()):
            comparison_data = []

            for district in all_districts:
                # Количество событий по округу за два периода
                earliest_count = earliest_period_data[
                    (earliest_period_data["Округ"].str.strip() == district) &
                    (earliest_period_data["Наименование события"] == event_name)
                    ]["Наименование события"].count()

                latest_count = latest_period_data[
                    (latest_period_data["Округ"].str.strip() == district) &
                    (latest_period_data["Наименование события"] == event_name)
                    ]["Наименование события"].count()

                difference = latest_count - earliest_count

                comparison_data.append([district, earliest_count, latest_count, difference])

            df_district_comparison = pd.DataFrame(
                comparison_data,
                columns=["Округ", "Кол-во (ранний)", "Кол-во (поздний)", "Разница"]
            )
            df_district_comparison["Округ"] = pd.Categorical(df_district_comparison["Округ"], categories=order,
                                                             ordered=True)
            df_district_comparison = df_district_comparison.sort_values("Округ").reset_index(drop=True)
            diagramms_data[f"theme{ev_num}_earliest"] = df_district_comparison["Кол-во (ранний)"].tolist()
            diagramms_data[f"theme{ev_num}_latest"] = df_district_comparison["Кол-во (поздний)"].tolist()
            data_diff_replace = [(f"*t{ev_num}d{d + 1}*", better_or(fint(row["Разница"]))) for d, row in
                                 df_district_comparison.iterrows()]
            replacer.replace_text(data_diff_replace)
            # Добавляем строку с названием темы
            header_row = pd.DataFrame([[event_name, "", "", ""]], columns=df_district_comparison.columns)

            # Добавляем итоговую строку
            total_row = pd.DataFrame([
                ["Итого", df_district_comparison["Кол-во (ранний)"].sum(),
                 df_district_comparison["Кол-во (поздний)"].sum(),
                 df_district_comparison["Разница"].sum()]
            ], columns=df_district_comparison.columns)

            # Объединяем заголовок, основную таблицу и итог
            df_district_comparison = pd.concat([header_row, df_district_comparison, total_row], ignore_index=True)

            # Сохраняем DataFrame в словаре
            event_district_tables[event_name] = df_district_comparison

    status_comparison_tables = {}

    # Получаем список всех возможных статусов
    all_statuses = sorted(set(earliest_period_data["Статус во внешней системе"].dropna().unique()) |
                          set(latest_period_data["Статус во внешней системе"].dropna().unique()))

    # Проходим по всем темам
    for i, event_name in enumerate(df_table_corrected["Тема"].unique()):
        comparison_data = []

        # **Считаем количество событий по статусам за два периода**
        status_counts_earliest = earliest_period_data[
            earliest_period_data["Наименование события"] == event_name
            ]["Статус во внешней системе"].value_counts()

        status_counts_latest = latest_period_data[
            latest_period_data["Наименование события"] == event_name
            ]["Статус во внешней системе"].value_counts()

        # Формируем строки для двух периодов
        earliest_row = ["Ранний период"] + [status_counts_earliest.get(status, 0) for status in all_statuses]
        latest_row = ["Поздний период"] + [status_counts_latest.get(status, 0) for status in all_statuses]
        difference_row = ["Разница"] + [(latest_row[i + 1] - earliest_row[i + 1]) for i in range(len(all_statuses))]

        # **Добавляем колонку "Всего" (сумма по статусам)**
        earliest_total = sum(earliest_row[1:])
        latest_total = sum(latest_row[1:])
        difference_total = sum(difference_row[1:])

        earliest_row.append(earliest_total)
        latest_row.append(latest_total)
        difference_row.append(difference_total)

        # **Добавляем строку с процентным изменением только для "Всего"**
        if earliest_total == 0:
            percent_change = "0%"  # Избегаем деления на ноль
        else:
            percent_change = f"{(difference_total / earliest_total) * 100:.2f}%"

        percent_change_row = ["% Δ"] + [""] * len(all_statuses) + [percent_change]

        # **Создаем DataFrame**
        df_status_comparison = pd.DataFrame(
            [earliest_row, latest_row, difference_row, percent_change_row],
            columns=["Период"] + all_statuses + ["Всего"]
        )

        # **Добавляем строку с названием темы**
        header_row = pd.DataFrame([[event_name] + [""] * (len(df_status_comparison.columns) - 1)],
                                  columns=df_status_comparison.columns)

        # **Объединяем заголовок и основную таблицу**
        df_status_comparison = pd.concat([header_row, df_status_comparison], ignore_index=True)

        # Сохраняем таблицу в словаре
        status_comparison_tables[event_name] = df_status_comparison
        df_date_data = df_ais[df_ais["Наименование события"] == event_name]["Дата создания"].value_counts()

        df_date_data = df_date_data.reset_index()
        df_date_data.columns = ["Дата", "Количество"]  # Явно задаем названия

        df_date_data = df_date_data.sort_values(by="Дата")
        dates = df_date_data["Дата"].tolist()
        values = df_date_data["Количество"].tolist()

        diagramms_data[f"theme{i}_dates"] = dates
        diagramms_data[f"theme{i}_values"] = values
    district_counts = latest_period_data["Район"].value_counts().reset_index()
    district_counts.columns = ["region", "record_count"]

    population_data = pd.read_excel("resource/region_citizen.xlsx")
    merged_data = district_counts.merge(population_data, on="region", how="left")
    merged_data["records_per_1000"] = (merged_data["record_count"] / merged_data["count"]) * 1000
    merged_data = merged_data.sort_values(by="records_per_1000", ascending=False).reset_index(drop=True)
    merged_data["records_per_1000"] = (merged_data["records_per_1000"]).round(2)
    # Сортируем данные по убыванию records_per_1000

    # Создаем массив пар, где первый элемент — строка с шаблоном, второй — значение из DataFrame
    data_array = [(f"*rt{i + 1}*", row["region"]) for i, row in merged_data.iterrows()]
    data_array += [(f"*rtk{i + 1}*", "{:.2f}".format(row["records_per_1000"]).replace(".", ",")) for i, row in
                   merged_data.iterrows()]

    print(data_array)
    replacer.replace_text(data_array)
    replacer.replace_text([(" ", " ")])
    replacer.write_presentation_to_file(f"{tmp_files_path}/ОтчетБезЦвета.pptx")
    runs_from_pptx(f"{tmp_files_path}/ОтчетБезЦвета.pptx", None, "week").save(
        f"{tmp_files_path}/ОтчетБезДиаграмм.pptx")
    update_diagramms(f"{tmp_files_path}/ОтчетБезДиаграмм.pptx", f"{report_path}/Недельный отчёт на {date}.pptx",
                     diagramms_data)
    with pd.ExcelWriter(f"{output_file}", mode="w", engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Виджет")
        df_table_corrected.to_excel(writer, sheet_name="Топ темы и округа")
        df_combined.to_excel(writer, sheet_name="Карта")
        final_table.to_excel(writer, sheet_name="Сводная таблица")
        merged_data.to_excel(writer, sheet_name="Удель")
        for i, event_name in enumerate(event_tables.keys()):
            status_comparison_tables[event_name].to_excel(writer, sheet_name=f"Тема {i + 1} Статусы", index=False,
                                                          header=True)
            event_district_tables[event_name].to_excel(writer, sheet_name=f"Тема {i + 1} Округ", index=False,
                                                       header=True)
            event_tables[event_name].to_excel(writer, sheet_name=f"Тема {i + 1} Ответственные", index=False,
                                              header=True)
    try:
        locale.setlocale(locale.LC_ALL, 'C')
    except locale.Error:
        pass
    return allert
