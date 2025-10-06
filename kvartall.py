import pandas as pd
import os

from xlsx_functions import update_ais_data, population


def make_kvartal_report_excel(file,date):
    population_moscow = population()

    date_text = f"{date}"
    dist_path = ""
    tmp_files_img_path = os.path.join(dist_path, "reports", f"{date_text}", "tmp_files", "img").replace(":", ".")
    tmp_files_path = os.path.join(dist_path, "reports", f"{date_text}", "tmp_files").replace(":", ".")
    reports_path = os.path.join(dist_path, "reports")

    for path in [tmp_files_path, tmp_files_img_path, reports_path]:
        if not os.path.exists(path):
            print(f"Создание директории: {path}")
            os.makedirs(path, exist_ok=True)
    excel_path = f"{reports_path}/kvartal.xlsx"
    order = ["ЦАО", "САО", "СВАО", "ВАО", "ЮВАО", "ЮАО", "ЮЗАО", "ЗАО", "СЗАО", "ЗелАО", "ТиНАО", 'ГБУ "АВД"', 'Иные']
    if ".xlsx" in file:
        df = pd.read_excel(file)
    else:
        df = pd.read_csv(file, low_memory=False)
    df = update_ais_data(df)
    df["Наименование события"] = df["Наименование события КОД ОИВ"]
    df["Округ"] = df["Округ"].astype(str).str.strip()
    df = df[~df["Округ"].isin(["", "nan", "None"])]
    status_mapping = {"Новое": "В работе", "Отменено": "Закрыто"}
    df["Округ"] = df["Округ"].replace({"НАО": "ТиНАО", "ТАО": "ТиНАО"})
    df["Статус во внешней системе"] = df["Статус во внешней системе"].replace(status_mapping)
    df.loc[df["Район"].isin(["Ново-Переделкино", "Солнцево"]), "Округ"] = "ЗАО"
    df.loc[df["Район"] == "Внуково", "Округ"] = "ТиНАО"
    df["Район"] = df["Район"].str.split(",").str[0].str.strip()
    print(df["Округ"].unique())
    count_table = (
        df["Округ"].value_counts()
        .rename_axis("Округ")
        .reset_index(name="Количество")
    )
    count_table["Сумма"] = count_table["Количество"]
    count_table = count_table[count_table["Округ"].isin(order)]
    count_table["Округ"] = pd.Categorical(count_table["Округ"], categories=order, ordered=True)
    count_table = count_table.sort_values("Округ").reset_index(drop=True)

    # Вторая таблица
    rate_table = pd.DataFrame()
    rate_table["Округ"] = count_table["Округ"].astype(str)
    rate_table["Количество"] = count_table["Количество"].astype(int)
    rate_table["Население"] = rate_table["Округ"].map(population_moscow)
    rate_table["Удельное количество на 100 тыс."] = (
                                                            rate_table["Количество"] / rate_table["Население"]
                                                    ) * 100000

    # Общая удельная величина
    total_count = rate_table["Количество"].sum()
    total_population = population_moscow["Общий итог"]
    overall_rate = (total_count / total_population) * 100000
    overall_rate_df = pd.DataFrame({
        "Показатель": ["Общая удельная величина"],
        "Значение на 100 тыс.": [overall_rate]
    })


    with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
        count_table.to_excel(writer, sheet_name="Количество по округам", index=False)
        rate_table.to_excel(writer, sheet_name="Удельные значения", index=False)
        overall_rate_df.to_excel(writer, sheet_name="Общий показатель", index=False)

    event_table = (
        df["Наименование события"]
        .value_counts()
        .rename_axis("Наименование события")
        .reset_index(name="Количество")
    )

    # Сохраняем в тот же Excel-файл, добавив лист
    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        event_table.to_excel(writer, sheet_name="События", index=False)
    print("Файл успешно сохранён:", excel_path)

    # Фильтры по типу объекта
    df_mkd = df[df["Тип объекта"] == "МКД"]
    df_other = df[df["Тип объекта"] != "МКД"]

    # Считаем количество событий в обеих группах
    event_mkd = df_mkd["Наименование события"].value_counts().rename("Количество (МКД)")
    event_other = df_other["Наименование события"].value_counts().rename("Количество (прочие)")

    # Объединяем по индексу (Наименование события)
    event_by_type = pd.concat([event_mkd, event_other], axis=1).fillna(0).astype(int)
    event_by_type = event_by_type.reset_index().rename(columns={"index": "Наименование события"})

    # Сохраняем в Excel
    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        event_by_type.to_excel(writer, sheet_name="События по типу", index=False)

    print("Таблица 'События по типу объекта' добавлена.")

    # Отфильтрованные датафреймы
    df_mkd = df[df["Тип объекта"] == "МКД"]
    df_other = df[df["Тип объекта"] != "МКД"]

    def top_events_with_top_districts(df_filtered, label):
        # Топ-8 событий
        top_events = df_filtered["Наименование события"].value_counts().head(8).index

        # Отбор только топ-событий
        df_top = df_filtered[df_filtered["Наименование события"].isin(top_events)]

        # Группировка по событию и округу
        grouped = (
            df_top.groupby(["Наименование события", "Округ"])
            .size()
            .reset_index(name="Количество")
        )

        # Для каждого события — топ-3 округа
        top3_per_event = (
            grouped.sort_values(["Наименование события", "Количество"], ascending=[True, False])
            .groupby("Наименование события")
            .head(3)
            .reset_index(drop=True)
        )

        return top3_per_event

    # Получение таблиц
    top_mkd = top_events_with_top_districts(df_mkd, "МКД")
    top_other = top_events_with_top_districts(df_other, "Прочие")

    # Сохранение в Excel
    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        top_mkd.to_excel(writer, sheet_name="Топ событий МКД", index=False)
        top_other.to_excel(writer, sheet_name="Топ событий Прочие", index=False)

    print("Топ-8 событий с округами (МКД и Прочие) добавлены.")

    # Шаг 1. Загрузка данных по населению районов
    population_by_region = pd.read_excel("resource/region_citizen.xlsx")
    population_by_region.columns = ["Район", "Население"]

    # Шаг 2. Подсчет количества событий по району
    events_by_region = df["Район"].value_counts().rename_axis("Район").reset_index(name="Количество")

    # Шаг 3. Объединение по району
    region_stats = events_by_region.merge(population_by_region, on="Район", how="left")

    # Шаг 4. Расчет удельного количества на 1 тыс. человек
    region_stats["Удельное на 1 тыс."] = (region_stats["Количество"] / region_stats["Население"]) * 1000

    # Шаг 5. Сохранение в Excel
    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        region_stats.to_excel(writer, sheet_name="Удельно по районам", index=False)

    print("Таблица с удельными значениями по районам добавлена.")


    # 1. Фильтрация: только МКД
    df_mkd = df[df["Тип объекта"] == "МКД"]

    # 2. Получаем топ-8 событий
    top_8_events = df_mkd["Наименование события"].value_counts().head(8).index.tolist()

    # 3. Только записи по этим событиям
    df_top_events = df_mkd[df_mkd["Наименование события"].isin(top_8_events)]

    # 4. Сводная таблица: округ × событие
    pivot_table = pd.pivot_table(
        df_top_events,
        index="Округ",
        columns="Наименование события",
        aggfunc="size",
        fill_value=0
    )

    # 5. Приводим строки к нужному порядку округов
    pivot_table = pivot_table.reindex(order).fillna(0)

    # 6. Сортируем события (столбцы) по убыванию суммарного количества
    event_totals = pivot_table.sum(axis=0).sort_values(ascending=False)
    sorted_events = event_totals.index.tolist()
    pivot_table = pivot_table[sorted_events]

    # 7. Добавляем колонку "Итого" — по всем событиям, а не только топ-8
    total_counts_all = df_mkd["Округ"].value_counts().reindex(order).fillna(0).astype(int)
    pivot_table["Итого"] = total_counts_all

    # 8. Добавляем строку "Итого"
    total_row = pivot_table.sum(numeric_only=True).to_frame().T
    total_row.index = ["Итого"]
    pivot_table = pd.concat([pivot_table, total_row])

    # 9. Сброс индекса перед сохранением
    pivot_table_reset = pivot_table.reset_index()

    # 10. Сохраняем в Excel
    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        pivot_table_reset.to_excel(writer, sheet_name="Округ × Топ события МКД", index=False)

    print("Сводная таблица по МКД обновлена: события отсортированы, учтены полные итоги.")

    # Фильтруем по МКД
    df_mkd = df[df["Тип объекта"] == "МКД"]

    # Определим топ-8 тем
    top_8_events = df_mkd["Наименование события"].value_counts().head(8).index.tolist()

    # ExcelWriter сразу откроем один раз
    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        for i, event in enumerate(top_8_events):
            # Фильтруем по конкретному событию
            df_event = df_mkd[df_mkd["Наименование события"] == event]

            # --- 1. Количество по округам ---
            event_by_district = (
                df_event["Округ"]
                .value_counts()
                .rename_axis("Округ")
                .reset_index(name="Количество")
            )

            # --- 2. Топ-3 округа ---
            top_3_districts = event_by_district.head(3)["Округ"].tolist()

            # --- 3. По каждому из топ-3 округов — топ-3 районов ---
            top_regions_data = []
            for district in top_3_districts:
                df_district = df_event[df_event["Округ"] == district]
                top_regions = (
                    df_district["Район"]
                    .value_counts()
                    .head(3)
                    .rename_axis("Район")
                    .reset_index(name="Количество")
                )
                top_regions["Округ"] = district
                top_regions["Событие"] = event
                top_regions_data.append(top_regions)

            # Объединяем районы
            event_top_regions = pd.concat(top_regions_data, ignore_index=True)[
                ["Событие", "Округ", "Район", "Количество"]]

            # Запись на отдельные листы
            sheet_base = f"Тема {i + 1}"  # обрезаем имя листа, чтобы не было слишком длинным
            event_by_district.to_excel(writer, sheet_name=f"{sheet_base} - округа", index=False)
            event_top_regions.to_excel(writer, sheet_name=f"{sheet_base} - районы", index=False)

    print("Для каждой из топ-8 тем по МКД созданы таблицы по округам и районам.")

    # Фильтруем по МКД
    df_mkd = df[df["Тип объекта"] == "МКД"]

    # Топ-8 событий
    top_8_events = df_mkd["Наименование события"].value_counts().head(8).index.tolist()

    # Открываем ExcelWriter
    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        for i, event in enumerate(top_8_events):
            # Фильтруем по событию
            df_event = df_mkd[df_mkd["Наименование события"] == event]

            # Группировка по адресу
            top_addresses = (
                df_event.groupby(["Адрес объекта", "Район", "Округ", "Наименование события"])
                .size()
                .reset_index(name="Количество")
                .sort_values("Количество", ascending=False)
                .head(10)
            )

            # Сохраняем на лист
            sheet_name = f"Тема {i + 1} - адреса"
            top_addresses.to_excel(writer, sheet_name=sheet_name, index=False)

    print("Топ-10 адресов для топ-8 событий по МКД сохранены.")
    # 1. Фильтрация: Прочие (не-МКД)
    df_other = df[df["Тип объекта"] != "МКД"]

    # 2. Топ-8 событий по Прочим
    top_8_events_other = df_other["Наименование события"].value_counts().head(8).index.tolist()

    # 3. Только записи по топ-8 событиям
    df_top_events_other = df_other[df_other["Наименование события"].isin(top_8_events_other)]

    # 4. Сводная таблица: Округ × Событие
    pivot_table_other = pd.pivot_table(
        df_top_events_other,
        index="Округ",
        columns="Наименование события",
        aggfunc="size",
        fill_value=0
    )

    # 5. Приводим порядок округов
    pivot_table_other = pivot_table_other.reindex(order).fillna(0)

    # 6. Сортировка событий по убыванию суммарного количества
    event_totals = pivot_table_other.sum(axis=0).sort_values(ascending=False)
    sorted_events = event_totals.index.tolist()
    pivot_table_other = pivot_table_other[sorted_events]

    # 7. Добавляем колонку "Итого" по строкам (сумма по всем событиям — не только топ-8)
    total_counts_all_other = df_other["Округ"].value_counts().reindex(order).fillna(0).astype(int)
    pivot_table_other["Итого"] = total_counts_all_other

    # 8. Добавляем строку "Итого" по колонкам
    total_row_other = pivot_table_other.sum(numeric_only=True).to_frame().T
    total_row_other.index = ["Итого"]
    pivot_table_other = pd.concat([pivot_table_other, total_row_other])

    # 9. Сброс индекса для сохранения
    pivot_table_other_reset = pivot_table_other.reset_index()

    # 10. Сохраняем в Excel
    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        pivot_table_other_reset.to_excel(writer, sheet_name="Округ × Топ события Прочие", index=False)

    print("Сводная таблица по Прочим обновлена, события отсортированы по убыванию.")

    # --- Топ-3 округа и топ-3 района по каждой теме ---

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        for i, event in enumerate(top_8_events_other):
            df_event = df_other[df_other["Наименование события"] == event]

            # Кол-во по округам
            event_by_district = (
                df_event["Округ"]
                .value_counts()
                .rename_axis("Округ")
                .reset_index(name="Количество")
            )

            top_3_districts = event_by_district.head(3)["Округ"].tolist()

            # Топ-3 района в каждом из топ-3 округов
            top_regions_data = []
            for district in top_3_districts:
                df_district = df_event[df_event["Округ"] == district]
                top_regions = (
                    df_district["Район"]
                    .value_counts()
                    .head(3)
                    .rename_axis("Район")
                    .reset_index(name="Количество")
                )
                top_regions["Округ"] = district
                top_regions["Событие"] = event
                top_regions_data.append(top_regions)

            event_top_regions = pd.concat(top_regions_data, ignore_index=True)[
                ["Событие", "Округ", "Район", "Количество"]]

            # Сохраняем
            sheet_base = f"Прочие {i + 1}"
            event_by_district.to_excel(writer, sheet_name=f"{sheet_base} - округа", index=False)
            event_top_regions.to_excel(writer, sheet_name=f"{sheet_base} - районы", index=False)

    print("Для топ-8 тем Прочего созданы таблицы по округам и районам.")

    # --- Топ-10 адресов для каждой темы ---

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        for i, event in enumerate(top_8_events_other):
            df_event = df_other[df_other["Наименование события"] == event]

            top_addresses = (
                df_event.groupby(["Адрес объекта", "Район", "Округ", "Наименование события"])
                .size()
                .reset_index(name="Количество")
                .sort_values("Количество", ascending=False)
                .head(10)
            )

            sheet_name = f"Прочие {i + 1} - адреса"
            top_addresses.to_excel(writer, sheet_name=sheet_name, index=False)

    print("Топ-10 адресов для топ-8 событий Прочего сохранены.")

    # Фильтрация
    df_mkd = df[df["Тип объекта"] == "МКД"]
    df_other = df[df["Тип объекта"] != "МКД"]

    # Подсчет количества уникальных адресов
    unique_addresses_mkd = df_mkd["Адрес объекта"].nunique()
    unique_addresses_other = df_other["Адрес объекта"].nunique()

    return [("Количество уникальных адресов (Прочие):", unique_addresses_other), ("Количество уникальных адресов (МКД):", unique_addresses_mkd)]