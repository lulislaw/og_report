import os
import pandas as pd
import yaml
from pathlib import Path
from typing import Mapping

from remote import dwnl_cfg


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
        dwnl_cfg()
        cfg_path = Path("resource/config.yaml")
        if cfg_path.is_file():
            try:
                df_ais = drop_random_by_config(df_ais, str(cfg_path), 42)
            except Exception as e:
                print(f"[skip] drop_random_by_config: {e}")
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


def fix_districts(df_main):
    mask_autodor = df_main[
                       'Ответственный'] == 'ГБУ «Автомобильные дороги»'
    df_main.loc[mask_autodor, 'Округ'] = 'ГБУ "АВД"'

    mask_autodor = df_main[
                       'Ответственный'] == 'Государственное Бюджетное Учреждение города Москвы "Автомобильные дороги"'
    df_main.loc[mask_autodor, 'Округ'] = 'ГБУ "АВД"'

    df_main['Округ'] = df_main['Округ'].replace(['', 'Общегородской'], 'Иные')
    df_main['Округ'] = df_main['Округ'].fillna('Иные')
    return df_main


def pusk_otoplenia_list():
    return ["Включение/отключение центрального отопления",
            "Прорыв, сильная течь элементов системы отопления (квартира)",
            "Прорыв, сильная течь элементов системы отопления (подъезд)",
            "Слабая течь элементов системы отопления (квартира)",
            "Слабая течь элементов системы отопления (подъезд)",
            "Экстремально высокая температура элементов системы отопления",
            "Шум, гул, вибрация в системе отопления"]


import pandas as pd


def fill_event_codes(df: pd.DataFrame) -> pd.DataFrame:
    ref_df = pd.read_excel("resource/mapping.xlsx")
    mapping = (
        ref_df.drop_duplicates(subset=["Наименование события", "Наименование события КОД ОИВ"])
        .set_index("Наименование события")["Наименование события КОД ОИВ"]
        .to_dict()
    )
    df["Наименование события КОД ОИВ"] = df["Наименование события"].map(mapping)
    return df
