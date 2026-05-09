import customtkinter as ctk
from tkinter import filedialog, messagebox, Text
from tkcalendar import DateEntry
from tkinterdnd2 import DND_FILES, TkinterDnD
from datetime import datetime
import threading
import os
import sys
import subprocess
import textwrap
import pandas as pd
from tkinter import ttk

from kvartall import make_kvartal_report_excel
from main_full import make_main_full_presentation
from month_report import make_month_report
from weekly_report import make_weekly_report


# ==================================================
# Настройки
# ==================================================

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

REPORTS_FOLDER = "reports"
ADDRESSES_FILE = "adresses.xlsx"

DISTRICTS = [
    "ЦАО",
    "САО",
    "СВАО",
    "ВАО",
    "ЮВАО",
    "ЮАО",
    "ЮЗАО",
    "ЗАО",
    "СЗАО",
    "ЗелАО",
    "ТиНАО",
    'ГБУ "АВД"',
    "Иные",
    "Общий итог",
]

ADDRESS_PATTERNS = [
    "брусилова",
    "дачная",
    "захарьин",
    "савицкого",
    "сироткин",
    "типограф",
]

COLOR_BG = "#121212"              # общий фон
COLOR_PANEL = "#1a1a1a"           # основные панели
COLOR_CARD = "#232323"            # карточки
COLOR_CARD_LIGHT = "#2d2d2d"      # вторичные элементы / строки

COLOR_ACCENT = "#0077FF"          # основной акцент
COLOR_ACCENT_HOVER = "#006BE6"    # hover для акцентных кнопок
COLOR_ACCENT_SOFT = "#0F2A4A"     # мягкая синяя подложка

COLOR_MUTED = "#9a9a9a"           # вторичный текст
COLOR_SUCCESS = "#0077FF"         # статус тоже в одном акценте
COLOR_WARNING = "#0077FF"
COLOR_ERROR = "#d9534f"           # ошибку оставляем красной


# ==================================================
# Консоль
# ==================================================

class ConsoleText:
    """Безопасное перенаправление print в Text-поле Tkinter."""

    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.after(0, self._write, message)

    def _write(self, message):
        self.text_widget.insert("end", message)
        self.text_widget.see("end")

    def flush(self):
        pass


# ==================================================
# Общие функции
# ==================================================

def set_today():
    date_entry.set_date(datetime.today())


def check_time():
    now = datetime.now().time()
    target_time = datetime.strptime("15:30", "%H:%M").time()
    return now <= target_time


def only_digits(new_value: str) -> bool:
    return new_value.isdigit() or new_value == ""


def get_date_str():
    return date_entry.get_date().strftime("%d.%m.%Y")


def get_ais_file():
    return ais_file_entry.get().strip()


def get_previous_period_data():
    data = {}

    for district, var in entry_vars.items():
        raw_value = var.get().strip()

        if raw_value == "":
            data[district] = 0
            continue

        try:
            data[district] = int(raw_value)
        except ValueError:
            data[district] = 0

    return data


def calculate_total():
    total = 0

    for district in previous_period:
        if district == "Общий итог":
            continue

        try:
            total += int(entry_vars[district].get() or 0)
        except ValueError:
            pass

    entry_vars["Общий итог"].set(str(total))
    print(f"Общий итог пересчитан: {total}")


def open_reports_folder():
    if not os.path.exists(REPORTS_FOLDER):
        messagebox.showerror("Ошибка", "Папка reports не найдена!")
        return

    if sys.platform.startswith("win"):
        subprocess.run(["explorer", REPORTS_FOLDER])
    elif sys.platform == "darwin":
        subprocess.run(["open", REPORTS_FOLDER])
    else:
        subprocess.run(["xdg-open", REPORTS_FOLDER])


def switch_to_console():
    try:
        tabs.set("Консоль")
    except Exception:
        pass


def switch_to_addresses():
    try:
        tabs.set("Адреса на проверку")
    except Exception:
        pass


def clear_console():
    report_output.delete("1.0", "end")
    open_folder_button.pack_forget()


def show_open_folder_button():
    open_folder_button.pack(side="left", padx=(8, 0))


def validate_ais_file():
    ais_file = get_ais_file()

    if not ais_file:
        messagebox.showwarning("Внимание", "Выберите файл АИС.")
        return False

    if not os.path.isfile(ais_file):
        messagebox.showerror("Ошибка", "Файл АИС не найден.")
        return False

    return True


def set_status(text, color=COLOR_SUCCESS):
    status_badge.configure(text=text, text_color=color)


def run_report_in_thread(report_func):
    if not validate_ais_file():
        return

    switch_to_console()
    clear_console()
    save_tree_to_excel(only_checked=True)

    set_status("● Формирование отчета", COLOR_WARNING)
    print("Процесс запущен...")

    def report_task():
        try:
            report_lines = report_func()

            if report_lines:
                for line in report_lines:
                    print(line)

            print("Конец" + "\n" * 5)

            root.after(0, show_open_folder_button)
            root.after(0, lambda: set_status("●", COLOR_SUCCESS))

        except Exception as e:
            print(f"Ошибка при формировании отчета: {e}")
            root.after(0, lambda: set_status("●", "#ef4444"))

    threading.Thread(target=report_task, daemon=True).start()


# ==================================================
# Генерация отчетов
# ==================================================

def generate_report_halfdaily():
    def task():
        ais_file = get_ais_file()
        date_str = get_date_str()
        fix_oiv = fix_oiv_var.get()
        previous_period_data = get_previous_period_data()

        return make_main_full_presentation(
            ais_file,
            previous_period_data,
            date_str,
            True,       # morning специально всегда True
            fix_oiv,
        )

    run_report_in_thread(task)


def generate_report_weekly():
    def task():
        ais_file = get_ais_file()
        date_str = get_date_str()

        return make_weekly_report(
            ais_file,
            date_str,
            False,
        )

    run_report_in_thread(task)


def apply_mid_index(event=None):
    raw = mid_index_var.get().strip()

    if raw == "":
        messagebox.showwarning("Внимание", "Введите число для mid_index.")
        mid_index_entry.focus_set()
        return None

    try:
        val = int(raw)
    except ValueError:
        messagebox.showerror("Ошибка", "mid_index должен быть целым числом.")
        return None

    if val < 0 or val > 100:
        messagebox.showwarning("Внимание", "mid_index должен быть в диапазоне 0..100.")
        return None

    return val


def generate_report_month():
    mid_index = apply_mid_index()

    if mid_index is None:
        return

    def task():
        ais_file = get_ais_file()
        date_str = get_date_str()

        print(f"mid_index: {mid_index}")

        return make_month_report(
            ais_file,
            date_str,
            mid_index=mid_index,
        )

    run_report_in_thread(task)


def generate_report_kvartal():
    def task():
        ais_file = get_ais_file()
        date_str = get_date_str()

        return make_kvartal_report_excel(
            ais_file,
            date_str,
        )

    run_report_in_thread(task)


# ==================================================
# Переключение типа отчета
# ==================================================

def update_buttons_visibility():
    report_type = report_type_var.get()

    for btn in [halfdaily_btn, weekly_btn, month_btn, kvartal_btn]:
        btn.pack_forget()

    morning_report_frame.pack_forget()
    fix_oiv_frame.pack_forget()
    midind_frame.pack_forget()

    if report_type == "Полусуточный":
        morning_report_frame.pack(side="left", padx=(0, 10))
        fix_oiv_frame.pack(side="left", padx=(0, 10))
        halfdaily_btn.pack(fill="x", padx=12, pady=(0, 12))
        morning_report_frame.pack_forget()


    elif report_type == "Недельный":
        weekly_btn.pack(fill="x", padx=12, pady=(0, 12))

    elif report_type == "Месячный":
        midind_frame.pack(side="left", padx=(0, 10))
        month_btn.pack(fill="x", padx=12, pady=(0, 12))

    elif report_type == "Квартальный":
        kvartal_btn.pack(fill="x", padx=12, pady=(0, 12))


# ==================================================
# Работа с таблицей адресов
# ==================================================

check_states = {}
original_data = {}


def on_tree_click(event):
    region = tree.identify_region(event.x, event.y)

    if region != "cell":
        return

    column = tree.identify_column(event.x)
    row_id = tree.identify_row(event.y)

    if row_id and column == "#1":
        check_states[row_id] = not check_states.get(row_id, False)
        tree.set(row_id, "check", "✓" if check_states[row_id] else "")


def select_file(entry):
    filename = filedialog.askopenfilename(
        filetypes=[
            ("Excel/CSV files", "*.xlsx;*.xls;*.csv"),
            ("All files", "*.*"),
        ]
    )

    if not filename:
        return

    entry.delete(0, "end")
    entry.insert(0, filename)

    if entry == ais_file_entry:
        load_ais_data()


def drop_file(event, entry):
    file_path = event.data.strip().replace("{", "").replace("}", "")

    entry.delete(0, "end")
    entry.insert(0, file_path)

    if entry == ais_file_entry:
        load_ais_data()


def read_ais_dataframe(ais_path):
    file_ext = os.path.splitext(ais_path)[1].lower()

    if file_ext in [".xlsx", ".xls"]:
        return pd.read_excel(ais_path)

    if file_ext == ".csv":
        return pd.read_csv(ais_path)

    raise ValueError("Неподдерживаемый формат файла. Нужен .xlsx, .xls или .csv")


def load_ais_data():
    ais_path = get_ais_file()

    if not os.path.isfile(ais_path):
        print("AIS-файл не указан или не существует!")
        return

    try:
        switch_to_addresses()
        table_frame.pack(fill="both", expand=True, padx=4, pady=(4, 4))

        df = read_ais_dataframe(ais_path)

        required_cols = ["Адрес объекта", "Район", "Округ"]

        for col in required_cols:
            if col not in df.columns:
                print(f"В файле нет колонки '{col}'!")
                return

        df = df.sort_values(by="Адрес объекта", ascending=False)

        address_series = df["Адрес объекта"].astype(str).str.lower()

        mask = address_series.apply(
            lambda addr: any(pattern in addr for pattern in ADDRESS_PATTERNS)
        )

        df_filtered = df[mask].drop_duplicates(subset=["Адрес объекта"])

        if os.path.exists(ADDRESSES_FILE):
            df_exclude = pd.read_excel(ADDRESSES_FILE)

            if "Адрес объекта" in df_exclude.columns:
                excluded_addresses = set(
                    df_exclude["Адрес объекта"].astype(str).str.lower()
                )

                df_filtered = df_filtered[
                    ~df_filtered["Адрес объекта"].astype(str).str.lower().isin(excluded_addresses)
                ]
            else:
                print(f"В файле '{ADDRESSES_FILE}' нет колонки 'Адрес объекта'!")
        else:
            print(f"Файл '{ADDRESSES_FILE}' не найден, исключения не применяются.")

        for row_id in tree.get_children():
            tree.delete(row_id)

        check_states.clear()
        original_data.clear()

        wrap_width = 52

        for _, row_data in df_filtered.iterrows():
            address = str(row_data["Адрес объекта"])
            rayon = str(row_data["Район"])
            okrug = str(row_data["Округ"])

            wrapped_address = textwrap.fill(address, width=wrap_width)

            item_id = tree.insert(
                "",
                "end",
                values=("", wrapped_address, rayon, okrug),
            )

            check_states[item_id] = False
            original_data[item_id] = (address, rayon, okrug)

        address_counter_label.configure(
            text=f"Найдено адресов на проверку: {len(df_filtered)}"
        )

        print(
            f"Загружено {len(df_filtered)} строк из {ais_path}, "
            f"исключены адреса из {ADDRESSES_FILE}."
        )

    except Exception as e:
        print(f"Ошибка при загрузке файла AIS: {e}")


def save_tree_to_excel(only_checked=False):
    data = []

    for row_id in tree.get_children():
        if only_checked and not check_states.get(row_id, False):
            continue

        if row_id not in original_data:
            continue

        address, rayon, okrug = original_data[row_id]
        data.append([address, rayon, okrug])

    if not data:
        print("Нет данных для сохранения, возможно, ничего не отмечено.")
        table_frame.pack_forget()
        return

    df_new = pd.DataFrame(
        data,
        columns=["Адрес объекта", "Район", "Округ"],
    )

    if os.path.exists(ADDRESSES_FILE):
        df_existing = pd.read_excel(ADDRESSES_FILE)
        df_combined = (
            pd.concat([df_existing, df_new])
            .drop_duplicates()
            .reset_index(drop=True)
        )
    else:
        df_combined = df_new

    df_combined.to_excel(ADDRESSES_FILE, index=False)

    print(
        f"Данные сохранены в файл: {ADDRESSES_FILE}, "
        f"количество записей: {len(df_combined)}"
    )

    table_frame.pack_forget()


# ==================================================
# Окно
# ==================================================

root = TkinterDnD.Tk()
root.title("ОГ")
root.geometry("1360x800")
root.minsize(1366, 768)
root.configure(background=COLOR_BG)


# ==================================================
# ttk стили
# ==================================================

style = ttk.Style()
style.theme_use("default")

style.configure(
    "Treeview",
    rowheight=56,
    background="#1f1f1f",
    foreground="#f2f2f2",
    fieldbackground="#1f1f1f",
    borderwidth=0,
    font=("Arial", 10),
)

style.configure(
    "Treeview.Heading",
    background="#1f538d",
    foreground="white",
    font=("Arial", 10, "bold"),
    relief="flat",
)

style.map(
    "Treeview",
    background=[("selected", "#2563eb")],
    foreground=[("selected", "white")],
)


# ==================================================
# Основной контейнер
# ==================================================

app_frame = ctk.CTkFrame(
    root,
    fg_color=COLOR_BG,
    corner_radius=0,
)
app_frame.pack(fill="both", expand=True)


# ==================================================
# Верхняя компактная шапка
# ==================================================

header_frame = ctk.CTkFrame(
    app_frame,
    fg_color=COLOR_PANEL,
    corner_radius=0,
    height=48,
)
header_frame.pack(fill="x", side="top")
header_frame.pack_propagate(False)

header_title = ctk.CTkLabel(
    header_frame,
    text="ОГ",
    font=("Arial", 18, "bold"),
    text_color="white",
)
header_title.pack(side="left", padx=18)

status_badge = ctk.CTkLabel(
    header_frame,
    text="●",
    font=("Arial", 12, "bold"),
    text_color=COLOR_SUCCESS,
)
status_badge.pack(side="right", padx=18)


# ==================================================
# Рабочая область
# ==================================================

content_frame = ctk.CTkFrame(
    app_frame,
    fg_color=COLOR_BG,
    corner_radius=0,
)
content_frame.pack(fill="both", expand=True, padx=14, pady=14)


# ==================================================
# Левая панель
# ==================================================

left_panel = ctk.CTkFrame(
    content_frame,
    width=480,
    fg_color=COLOR_PANEL,
    corner_radius=16,
)
left_panel.pack(side="left", fill="y", expand=False, padx=(0, 10))
left_panel.pack_propagate(False)


# ==================================================
# Тип отчета + файл
# ==================================================

top_card = ctk.CTkFrame(
    left_panel,
    fg_color=COLOR_CARD,
    corner_radius=14,
)
top_card.pack(fill="x", padx=12, pady=(12, 8))

top_row = ctk.CTkFrame(top_card, fg_color="transparent")
top_row.pack(fill="x", padx=12, pady=(12, 8))

ctk.CTkLabel(
    top_row,
    text="Тип отчета",
    font=("Arial", 13, "bold"),
).pack(side="left", padx=(0, 10))

report_type_var = ctk.StringVar(value="Полусуточный")

report_type_combobox = ctk.CTkComboBox(
    top_row,
    values=[
        "Полусуточный",
        "Недельный",
        "Месячный",
        "Квартальный",
    ],
    variable=report_type_var,
    command=lambda choice: update_buttons_visibility(),
    width=210,
    height=34,
    button_color=COLOR_ACCENT,
    button_hover_color=COLOR_ACCENT_HOVER,
    dropdown_fg_color="#2a2a2a",
    dropdown_hover_color="#333333",
)
report_type_combobox.pack(side="left")

file_label = ctk.CTkLabel(
    top_card,
    text="Файл АИС",
    font=("Arial", 13, "bold"),
)
file_label.pack(anchor="w", padx=12, pady=(4, 4))

ais_file_row = ctk.CTkFrame(top_card, fg_color="transparent")
ais_file_row.pack(fill="x", padx=12, pady=(0, 12))

ais_file_entry = ctk.CTkEntry(
    ais_file_row,
    height=36,
    placeholder_text="Выберите или перетащите .xlsx / .xls / .csv",
)
ais_file_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

ais_file_entry.drop_target_register(DND_FILES)
ais_file_entry.dnd_bind("<<Drop>>", lambda e: drop_file(e, ais_file_entry))

ais_file_button = ctk.CTkButton(
    ais_file_row,
    text="...",
    width=44,
    height=36,
    command=lambda: select_file(ais_file_entry),
    fg_color=COLOR_ACCENT,
    hover_color=COLOR_ACCENT_HOVER,
)
ais_file_button.pack(side="left")


# ==================================================
# Дата и параметры
# ==================================================

settings_card = ctk.CTkFrame(
    left_panel,
    fg_color=COLOR_CARD,
    corner_radius=14,
)
settings_card.pack(fill="x", padx=12, pady=8)

settings_row = ctk.CTkFrame(settings_card, fg_color="transparent")
settings_row.pack(fill="x", padx=12, pady=12)

today_button = ctk.CTkButton(
    settings_row,
    text="Сегодня",
    command=set_today,
    width=90,
    height=34,
    fg_color="#0077FF",
    hover_color="#0066EE",
)
today_button.pack(side="left", padx=(0, 8))

date_entry = DateEntry(
    settings_row,
    width=13,
    background="darkblue",
    foreground="white",
    borderwidth=2,
    date_pattern="dd.mm.yyyy",
)
date_entry.pack(side="left", padx=(0, 14))
date_entry.set_date(datetime.today())

morning_var = ctk.BooleanVar(value=check_time())
morning_report_frame = ctk.CTkFrame(settings_row, fg_color="transparent")

morning_checkbox = ctk.CTkCheckBox(
    morning_report_frame,
    text="Утро",
    variable=morning_var,
)
morning_checkbox.pack(side="left")

fix_oiv_var = ctk.BooleanVar(value=True)
fix_oiv_frame = ctk.CTkFrame(settings_row, fg_color="transparent")

fix_oiv_checkbox = ctk.CTkCheckBox(
    fix_oiv_frame,
    text="Исправить КОД ОИВ",
    variable=fix_oiv_var,
)
fix_oiv_checkbox.pack(side="left")

midind_frame = ctk.CTkFrame(settings_row, fg_color="transparent")

ctk.CTkLabel(
    midind_frame,
    text="mid_index:",
    font=("Arial", 12),
).pack(side="left", padx=(0, 6))

mid_index_var = ctk.StringVar(value="0")
vcmd = (root.register(only_digits), "%P")

mid_index_entry = ctk.CTkEntry(
    midind_frame,
    width=80,
    height=32,
    textvariable=mid_index_var,
    validate="key",
    validatecommand=vcmd,
    placeholder_text="0..100",
)
mid_index_entry.pack(side="left")


# ==================================================
# Предыдущий период — компактная сетка
# ==================================================

previous_period_card = ctk.CTkFrame(
    left_panel,
    fg_color=COLOR_CARD,
    corner_radius=14,
)
previous_period_card.pack(fill="both", expand=True, padx=12, pady=8)

period_header = ctk.CTkFrame(previous_period_card, fg_color="transparent")
period_header.pack(fill="x", padx=12, pady=(8, 6))

ctk.CTkLabel(
    period_header,
    text="Предыдущий период",
    font=("Arial", 14, "bold"),
).pack(side="left")

summary_btn = ctk.CTkButton(
    period_header,
    text="Пересчитать общий итог",
    width=180,
    height=30,
    command=calculate_total,
    fg_color="#0077FF",
    hover_color="#0066EE",
)
summary_btn.pack(side="right")

previous_period = {district: 0 for district in DISTRICTS}
entry_vars = {}

# Контейнер без скролла, компактно в 2 колонки
period_grid = ctk.CTkFrame(
    previous_period_card,
    fg_color="transparent",
)
period_grid.pack(fill="both", expand=True, padx=10, pady=(0, 10))

# 2 колонки
left_col = ctk.CTkFrame(period_grid, fg_color="transparent")
right_col = ctk.CTkFrame(period_grid, fg_color="transparent")

left_col.pack(side="left", fill="both", expand=True, padx=(0, 5))
right_col.pack(side="left", fill="both", expand=True, padx=(5, 0))

district_items = list(previous_period.keys())

# Делим список пополам
split_index = (len(district_items) + 1) // 2
left_items = district_items[:split_index]
right_items = district_items[split_index:]


def create_period_row(parent, district):
    row_frame = ctk.CTkFrame(
        parent,
        fg_color="#1f1f1f",
        corner_radius=8,
        height=34,
    )
    row_frame.pack(fill="x", pady=2)
    row_frame.pack_propagate(False)

    label = ctk.CTkLabel(
        row_frame,
        text=district,
        anchor="w",
        font=("Arial", 12),
    )
    label.pack(side="left", fill="x", expand=True, padx=(8, 4))

    entry_var = ctk.StringVar(value="0")
    entry_vars[district] = entry_var

    entry_field = ctk.CTkEntry(
        row_frame,
        textvariable=entry_var,
        width=64,
        height=26,
        justify="center",
    )
    entry_field.pack(side="right", padx=(4, 8))

    return row_frame


for district in left_items:
    create_period_row(left_col, district)

for district in right_items:
    create_period_row(right_col, district)


# ==================================================
# Запуск отчета
# ==================================================

run_card = ctk.CTkFrame(
    left_panel,
    fg_color=COLOR_CARD,
    corner_radius=14,
)
run_card.pack(fill="x", padx=12, pady=(8, 12))

button_frame = ctk.CTkFrame(run_card, fg_color="transparent")
button_frame.pack(fill="x", padx=12, pady=(12, 12))

halfdaily_btn = ctk.CTkButton(
    button_frame,
    text="Сформировать полусуточный отчет",
    height=42,
    command=generate_report_halfdaily,
    fg_color=COLOR_ACCENT,
    hover_color=COLOR_ACCENT_HOVER,
)

weekly_btn = ctk.CTkButton(
    button_frame,
    text="Сформировать недельный отчет",
    height=42,
    command=generate_report_weekly,
    fg_color=COLOR_ACCENT,
    hover_color=COLOR_ACCENT_HOVER,
)

month_btn = ctk.CTkButton(
    button_frame,
    text="Сформировать месячный отчет",
    height=42,
    command=generate_report_month,
    fg_color=COLOR_ACCENT,
    hover_color=COLOR_ACCENT_HOVER,
)

kvartal_btn = ctk.CTkButton(
    button_frame,
    text="Сформировать квартальный отчет",
    height=42,
    command=generate_report_kvartal,
    fg_color=COLOR_ACCENT,
    hover_color=COLOR_ACCENT_HOVER,
)


# ==================================================
# Правая панель
# ==================================================

right_panel = ctk.CTkFrame(
    content_frame,
    fg_color=COLOR_PANEL,
    corner_radius=16,
)
right_panel.pack(side="right", fill="both", expand=True, padx=(10, 0))

tabs = ctk.CTkTabview(
    right_panel,
    fg_color=COLOR_PANEL,
    segmented_button_fg_color=COLOR_CARD,
    segmented_button_selected_color=COLOR_ACCENT,
    segmented_button_selected_hover_color=COLOR_ACCENT_HOVER,
    segmented_button_unselected_color=COLOR_CARD,
    segmented_button_unselected_hover_color="#333333",
)
tabs.pack(fill="both", expand=True, padx=12, pady=12)

console_tab = tabs.add("Консоль")
addresses_tab = tabs.add("Адреса на проверку")


# ==================================================
# Консоль
# ==================================================

console_header = ctk.CTkFrame(console_tab, fg_color="transparent")
console_header.pack(fill="x", padx=4, pady=(2, 6))

ctk.CTkLabel(
    console_header,
    text="Лог выполнения",
    font=("Arial", 15, "bold"),
).pack(side="left")

console_actions = ctk.CTkFrame(console_header, fg_color="transparent")
console_actions.pack(side="right")

clear_console_btn = ctk.CTkButton(
    console_actions,
    text="Очистить",
    width=90,
    height=30,
    fg_color="#0077FF",
    hover_color="#0066ee",
    command=lambda: report_output.delete("1.0", "end"),
)
clear_console_btn.pack(side="left")

open_folder_button = ctk.CTkButton(
    console_actions,
    text="Открыть reports",
    width=130,
    height=30,
    command=open_reports_folder,
    fg_color="#334155",
    hover_color="#475569",
)
open_folder_button.pack_forget()

report_output = Text(
    console_tab,
    wrap="word",
    height=15,
    bg="#101010",
    fg="#f5f5f5",
    insertbackground="white",
    relief="flat",
    font=("Consolas", 10),
    padx=12,
    pady=12,
)
report_output.pack(fill="both", expand=True, padx=4, pady=(0, 4))

sys.stdout = ConsoleText(report_output)
sys.stderr = ConsoleText(report_output)


# ==================================================
# Адреса на проверку
# ==================================================

addresses_header = ctk.CTkFrame(addresses_tab, fg_color="transparent")
addresses_header.pack(fill="x", padx=4, pady=(2, 6))

ctk.CTkLabel(
    addresses_header,
    text="Адреса на проверку",
    font=("Arial", 15, "bold"),
).pack(side="left")

address_counter_label = ctk.CTkLabel(
    addresses_header,
    text="Файл еще не загружен",
    font=("Arial", 12),
    text_color=COLOR_MUTED,
)
address_counter_label.pack(side="right")

table_frame = ctk.CTkFrame(
    addresses_tab,
    fg_color=COLOR_CARD,
    corner_radius=14,
)

tree_scrollbar = ttk.Scrollbar(table_frame, orient="vertical")
tree_scrollbar.pack(side="right", fill="y", padx=(0, 4), pady=8)

tree_columns = (
    "check",
    "Адрес объекта",
    "Район",
    "Округ",
)

tree = ttk.Treeview(
    table_frame,
    columns=tree_columns,
    show="headings",
    yscrollcommand=tree_scrollbar.set,
)

tree.heading("check", text="✓")
tree.column("check", width=42, anchor="center", stretch=False)

tree.heading("Адрес объекта", text="Адрес объекта")
tree.column("Адрес объекта", width=500, anchor="w")

tree.heading("Район", text="Район")
tree.column("Район", width=170, anchor="w")

tree.heading("Округ", text="Округ")
tree.column("Округ", width=100, anchor="center")

tree.pack(side="left", fill="both", expand=True, padx=8, pady=8)
tree_scrollbar.config(command=tree.yview)

tree.bind("<Button-1>", on_tree_click)


# ==================================================
# Финальная инициализация
# ==================================================

update_buttons_visibility()
tabs.set("Консоль")

root.mainloop()