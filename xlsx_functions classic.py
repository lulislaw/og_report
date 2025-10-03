import os
import pandas as pd
import yaml
from pathlib import Path
from typing import Mapping
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import time
import numpy as np

def update_ais_data(df_ais, addresses_file="adresses.xlsx"):
    """
    Обновляет данные в df_ais:
    Если 'Адрес объекта' присутствует в adresses.xlsx, то меняем 'Район' на 'Щербинка', 'Округ' на 'ТиНАО'.
    """

    try:

        # Проверяем наличие нужных колонок
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
        # df_ais = fix_districts(df_ais, False)
        return df_ais

    except Exception as e:
        print(f"Ошибка при обработке файла: {e}")
        return None


def _load_config(config: str | Path | Mapping[str, float]) -> dict[str, float]:
    """
    Принимает путь/строку/Path к YAML‑файлу или уже готовый dict
    и возвращает словарь {район: frac}.
    """
    if isinstance(config, Mapping):
        return dict(config)

    config_path = Path(config)
    if not config_path.is_file():
        raise FileNotFoundError(f"Config not found: {config_path}")

    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def drop_random_by_config(
        df: pd.DataFrame,
        config: str | Path | Mapping[str, float],
        seed: int | None = None,
) -> pd.DataFrame:
    """
    Удаляет случайную долю строк для каждого района,
    согласно конфигурации {район: frac}.
    """
    cfg = _load_config(config)
    if cfg:
        to_drop = []
        for district, frac in cfg.items():
            if not (0 < frac <= 1):
                raise ValueError(f"Некорректная доля для «{district}»: {frac}")

            mask = df["Район"] == district
            if mask.any():
                to_drop.extend(
                    df[mask].sample(frac=frac, random_state=seed).index
                )

        return df.drop(index=to_drop)
    return df


def get_district_gorod(driver, number: str) -> str:
    try:
        driver.get('https://gorod.mos.ru/')
        time.sleep(3)
        search_button = driver.find_element(By.XPATH,
                                            '/html/body/div[2]/div/div/div[1]/header/div/nav/div[2]/ul/li[1]/a')
        search_button.click()
        search_placeholder = driver.find_element(By.XPATH,
                                                 "/html/body/div[2]/div/div/div[1]/header/div/nav/div[1]/div[2]/div/div[1]/form/div/input")

        search_placeholder.send_keys(number)
        time.sleep(2)
        search_find = driver.find_element(By.XPATH,
                                          "/html/body/div[2]/div/div/div[1]/header/div/nav/div[1]/div[2]/div/div[2]/div/a")
        if search_find:
            search_find.click()
        time.sleep(3)
        district_tag = driver.find_element(By.XPATH,
                                           value="/html/body/div[2]/div/div/header/div/div/div/div[1]/nav/ul/li[1]")
        if district_tag:
            return abbreviate_district(district_tag.text)
        else:
            return ''  # или None, если не найдено

    finally:
        print("NG")


def get_district_edc(driver, number: str) -> str:
    try:
        # URL для поиска на edc.mos.ru
        driver.get('https://edc.mos.ru/')
        time.sleep(5)
        search_input = driver.find_element(By.XPATH,
                                           "/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div[1]/span[1]/input")
        search_input.send_keys(number)
        time.sleep(1)
        search_button = driver.find_element(By.XPATH,
                                            "/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div[1]/button[1]")
        search_button.click()
        district_tag = driver.find_element(By.XPATH,
                                           "/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[2]/div/div/div/div/div[2]/table/tbody/tr[2]/td[2]/div/span/div/div[1]")
        print(district_tag.text)
        if district_tag:
            return district_tag.text
        else:
            return ''
    finally:
        print("edc")


def abbreviate_district(full_name: str) -> str:
    DISTRICT_ABBREVIATIONS = {
        "Центральный округ": "ЦАО",
        "Северный округ": "САО",
        "Северо-Восточный округ": "СВАО",
        "Северо-Западный округ": "СЗАО",
        "Восточный округ": "ВАО",
        "Юго-Восточный округ": "ЮВАО",
        "Южный округ": "ЮАО",
        "Юго-Западный округ": "ЮЗАО",
        "Западный округ": "ЗАО",
        "Зеленоградский округ": "ЗелАО",
        "Новомосковский округ": "ТиНАО",
        "Троицкий округ": "ТиНАО",
    }
    name = full_name.strip()
    return DISTRICT_ABBREVIATIONS.get(name, name)


def fix_districts(df_main,need_selenium, ):
    # 1) Путь к файлам (при желании поменяйте на свои названия/пути)
    MAPPING_FILE = 'resource/mapping.xlsx'  # файл со столбцами "район" и "округ"

    # 2) Читаем основной DataFrame и справочник
    df_mapping = pd.read_excel(MAPPING_FILE, dtype=str)  # читаем mapping.xlsx как строки

    mask_autodor = df_main[
                       'Ответственный'] == 'Государственное Бюджетное Учреждение города Москвы "Автомобильные дороги"'
    df_main.loc[mask_autodor, 'Округ'] = 'ГБУ "АВД"'


    df_main["Район"] = df_main["Район"].str.split(",").str[0].str.strip()
    df_mapping = df_mapping.dropna(subset=['Район', 'Округ'])
    mapping_dict = dict(zip(df_mapping['Район'], df_mapping['Округ']))

    def is_empty_or_obs(row_value):
        return pd.isna(row_value) or str(row_value).strip() == '' or str(row_value).strip() == 'Общегородской'

    import pymorphy3, re

    morph = pymorphy3.MorphAnalyzer()

    def normalize_words(text: str) -> list[str]:
        text = text.lower()
        text = re.sub(r'[^а-яё\s]', ' ', text)
        tokens = [tok for tok in text.split() if tok]
        lemmas = [morph.parse(tok)[0].normal_form for tok in tokens]
        return lemmas

    mask_need_mapping = df_main['Округ'].apply(is_empty_or_obs)

    df_main.loc[mask_need_mapping, 'Округ'] = df_main.loc[mask_need_mapping, 'Район'].map(mapping_dict)
    mask_still_empty = df_main['Округ'].apply(is_empty_or_obs)
    for idx in df_main[mask_still_empty].index:
        resp_text = str(df_main.at[idx, 'Ответственный'])
        # получаем перечень лемм из строки "Ответственный"
        resp_norm = normalize_words(resp_text)

        for rayon, okrug in mapping_dict.items():
            rayon_clean = rayon.lower().replace('район', '').strip()
            if not rayon_clean:
                continue

            rayon_norm = normalize_words(rayon_clean)

            if all(token in resp_norm for token in rayon_norm):
                df_main.at[idx, 'Округ'] = okrug
                break

    okrug_gen_dict = {
        'центрального административного округа': 'ЦАО',
        'северного административного округа': 'САО',
        'северо-восточного административного округа': 'СВАО',
        'северо-западного административного округа': 'СЗАО',
        'восточного административного округа': 'ВАО',
        'юго-восточного административного округа': 'ЮВАО',
        'южного административного округа': 'ЮАО',
        'юго-западного административного округа': 'ЮЗАО',
        'западного административного округа': 'ЗАО',
        'зеленоградского административного округа': 'ЗелАО',
        'новомосковского административного округа': 'ТиНАО',
        'троицкого административного округа': 'ТиНАО',
    }

    # Словарь для преобразования полного названия округа в сокращение

    mask_still_empty = df_main['Округ'].apply(is_empty_or_obs)

    for idx in df_main[mask_still_empty].index:
        resp_text = str(df_main.at[idx, 'Ответственный'])
        resp_lower = resp_text.lower()

        # 1) Сначала пытаем маппинг по району (ваш существующий код)
        rayon_found = False
        resp_norm = normalize_words(resp_text)
        for rayon, okrug in mapping_dict.items():
            rayon_clean = rayon.lower().replace('район', '').strip()
            if not rayon_clean:
                continue
            rayon_norm = normalize_words(rayon_clean)
            if all(token in resp_norm for token in rayon_norm):
                df_main.at[idx, 'Округ'] = okrug
                rayon_found = True
                break

        if rayon_found:
            continue

        for key_gen, okrug_name in okrug_gen_dict.items():
            if key_gen in resp_lower:
                df_main.at[idx, 'Округ'] = okrug_name
                break
        # если ничего не нашлось — остаётся пустым

    # 7) Сохраняем результат обратно в Excel

    df_main['Округ'] = df_main['Округ'].fillna('')  # заменим NaN на пустую строку
    df_main['Система'] = df_main['Система'].astype(str)
    df_main['№ во внешней системе'] = df_main['№ во внешней системе'].astype(str)
    mask_okrug_empty = (df_main['Округ'].str.strip() == '') | (df_main['Округ'].str.strip() == 'Общегородской')

    # Обработаем отдельно "НГ" и "ЕДЦ"
    mask_ng = mask_okrug_empty & (df_main['Система'] == 'НГ')
    mask_edc = mask_okrug_empty & (df_main['Система'] == 'ЕДЦ')
    if need_selenium:
        # Настройки Chrome
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)
        chrome_options.add_argument("user-data-dir=C:\\Profile")  # Использование профиля пользователя
        driver = webdriver.Chrome(options=chrome_options)
        for idx, row in df_main[mask_ng].iterrows():
            number = row['№ во внешней системе'].strip()
            if number:
                time.sleep(3)
                district = get_district_gorod(driver, number)
                df_main.at[idx, 'Округ'] = district
        for idx, row in df_main[mask_edc].iterrows():
            number = row['№ во внешней системе'].strip()
            if number:
                time.sleep(3)
                district = get_district_edc(driver, number)
                df_main.at[idx, 'Округ'] = district
        driver.quit()
    else:
        print("УДАЛЕНИЕ ПУСТЫХ ОКРУГОВ")
        df_main = df_main[df_main["Округ"].str.strip().notna()]

    return df_main
