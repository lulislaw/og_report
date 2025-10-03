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

from daily_report_full import make_daily_full_presentation
from main_full import make_main_full_presentation
from month_report import make_month_report
from weekly_report import make_weekly_report

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

REPORTS_FOLDER = "reports"


class ConsoleText:
    """Перенаправляем print в текстовое поле Tkinter"""

    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.insert("end", message)
        self.text_widget.see("end")  # Автопрокрутка

    def flush(self):
        pass


def set_today():
    date_entry.set_date(datetime.today())


def calculate_total():
    total = sum(int(entry_vars[district].get() or 0) for district in previous_period if district != "Общий итог")
    entry_vars["Общий итог"].set(str(total))


def open_reports_folder():
    if os.path.exists(REPORTS_FOLDER):
        if os.name == 'nt':  # Windows
            subprocess.run(['explorer', REPORTS_FOLDER])
        else:  # macOS, Linux
            subprocess.run(['xdg-open', REPORTS_FOLDER])
    else:
        messagebox.showerror("Ошибка", "Папка не найдена!")


def generate_report_halfdaily():
    report_output.delete("1.0", "end")
    save_tree_to_excel(True)
    print("Процесс запущен...")

    def report_task():
        ais_file = ais_file_entry.get()
        date_str = date_entry.get_date().strftime("%d.%m.%Y")
        morning = morning_var.get()
        fix_oiv = fix_oiv_var.get()
        summer = True
        # summer = summer_var.get()
        previous_period_data = {district: int(entry_vars[district].get() or 0) for district in entry_vars}
        if not summer:
            edc_file = None
        report_lines = make_main_full_presentation(ais_file, previous_period_data, date_str, morning, fix_oiv)
        for line in report_lines:
            print(line)

        open_folder_button.pack(pady=10)
        print("Конец" + "\n" * 5)

    threading.Thread(target=report_task, daemon=True).start()


def generate_report_daily_full():
    report_output.delete("1.0", "end")
    save_tree_to_excel(True)
    print("Процесс запущен...")

    def report_task():
        ais_file = ais_file_entry.get()
        date_str = date_entry.get_date().strftime("%d.%m.%Y")
        previous_period_data = {district: int(entry_vars[district].get() or 0) for district in entry_vars}
        report_lines = make_daily_full_presentation(ais_file, previous_period_data, date_str)
        for line in report_lines:
            print(line)
        open_folder_button.pack(pady=10)
        print("Конец" + "\n" * 5)

    threading.Thread(target=report_task, daemon=True).start()


def generate_report_weekly():
    report_output.delete("1.0", "end")
    save_tree_to_excel(True)
    print("Процесс запущен...")

    def report_task():
        ais_file = ais_file_entry.get()
        date_str = date_entry.get_date().strftime("%d.%m.%Y")
        report_lines = make_weekly_report(ais_file, date_str, False)
        for line in report_lines:
            print(line)
        open_folder_button.pack(pady=10)
        print("Конец" + "\n" * 5)

    threading.Thread(target=report_task, daemon=True).start()


def generate_report_month():
    report_output.delete("1.0", "end")
    save_tree_to_excel(True)
    print("Процесс запущен...")

    def report_task():
        ais_file = ais_file_entry.get()
        date_str = date_entry.get_date().strftime("%d.%m.%Y")
        report_lines = make_month_report(ais_file, date_str, False)
        for line in report_lines:
            print(line)
        open_folder_button.pack(pady=10)
        print("Конец" + "\n" * 5)

    threading.Thread(target=report_task, daemon=True).start()


def update_buttons_visibility():
    report_type = report_type_var.get()
    all_buttons = [
        halfdaily_btn, weekly_btn, month_btn
    ]
    for btn in all_buttons:
        btn.pack_forget()
    if report_type == "Полусуточный":
        halfdaily_btn.pack(side="right", padx=20)
        morning_report_frame.pack()
    elif report_type == "Недельный":
        weekly_btn.pack(side="right", padx=20)
        morning_report_frame.pack_forget()
    elif report_type == "Месячный":
        month_btn.pack(side="right", padx=20)
        morning_report_frame.pack_forget()


# ========= ЛОГИКА ДЛЯ TREEVIEW С ЧЕКБОКСАМИ =========

# Храним состояние чекбоксов здесь:
check_states = {}  # словарь: row_id -> bool
original_data = {}


def on_tree_click(event):
    region = tree.identify_region(event.x, event.y)
    if region == "cell":
        column = tree.identify_column(event.x)  # "#1", "#2", "#3", ...
        row_id = tree.identify_row(event.y)
        if row_id and column == "#1":
            # мы настроили "check" как первый столбец -> column == "#1"
            check_states[row_id] = not check_states.get(row_id, False)
            # Обновляем символ в ячейке
            tree.set(row_id, "check", "✓" if check_states[row_id] else "")


def select_file(entry):
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls;*.csv")])
    if filename:
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


def load_ais_data():
    ais_path = ais_file_entry.get()
    addresses_file = "adresses.xlsx"

    if not os.path.isfile(ais_path):
        print("AIS-файл не указан или не существует!")
        return

    try:
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        if ".xlsx" in ais_path:
            df = pd.read_excel(ais_path)
        elif ".csv" in ais_path:
            df = pd.read_csv(ais_path)
        else:
            print("Файл не правильный")
            return
        df.sort_values(by="Адрес объекта", ascending=False)
        required_cols = ["Адрес объекта", "Район", "Округ"]
        for col in required_cols:
            if col not in df.columns:
                print(f"В файле нет колонки '{col}'!")
                return

        patterns = ["брусилова", "дачная", "захарьин", "савицкого", "сироткин", "типограф"]
        mask = df["Адрес объекта"].str.lower().apply(
            lambda addr: any(pat in addr for pat in patterns if isinstance(addr, str))
        )
        df_filtered = df[mask].drop_duplicates(subset=["Адрес объекта"])

        if os.path.exists(addresses_file):
            df_exclude = pd.read_excel(addresses_file)
            if "Адрес объекта" in df_exclude.columns:
                excluded_addresses = set(df_exclude["Адрес объекта"].astype(str).str.lower())
                df_filtered = df_filtered[~df_filtered["Адрес объекта"].str.lower().isin(excluded_addresses)]
            else:
                print(f"В файле '{addresses_file}' нет колонки 'Адрес объекта'!")
        else:
            print(f"Файл '{addresses_file}' не найден, исключения не применяются.")

        for row_id in tree.get_children():
            tree.delete(row_id)
        check_states.clear()
        original_data.clear()

        WRAP_WIDTH = 40

        for i, row_data in df_filtered.iterrows():
            address = str(row_data["Адрес объекта"])
            rayon = str(row_data["Район"])
            okrug = str(row_data["Округ"])

            wrapped_address = textwrap.fill(address, width=WRAP_WIDTH)

            item_id = tree.insert("", "end", values=("", wrapped_address, rayon, okrug))
            check_states[item_id] = False
            original_data[item_id] = (address, rayon, okrug)

        print(f"Загружено {len(df_filtered)} строк из {ais_path}, исключены адреса из {addresses_file}.")
    except Exception as e:
        print(f"Ошибка при загрузке файла AIS: {e}")


def save_tree_to_excel(only_checked=False):
    data = []
    for row_id in tree.get_children():
        if only_checked and not check_states.get(row_id, False):
            continue

        address, rayon, okrug = original_data[row_id]
        data.append([address, rayon, okrug])

    if not data:
        print("Нет данных для сохранения (возможно, ничего не отмечено).")
        table_frame.pack_forget()
        return

    # Создаем DataFrame из новых данных
    df_new = pd.DataFrame(data, columns=["Адрес объекта", "Район", "Округ"])

    # Проверяем, существует ли файл
    if os.path.exists("adresses.xlsx"):
        # Загружаем существующие данные
        df_existing = pd.read_excel("adresses.xlsx")

        # Объединяем и убираем дубликаты
        df_combined = pd.concat([df_existing, df_new]).drop_duplicates().reset_index(drop=True)
    else:
        df_combined = df_new

    # Сохраняем обновленные данные
    df_combined.to_excel("adresses.xlsx", index=False)
    print(f"Данные сохранены в файл: {"adresses.xlsx"}, количество записей: {len(df_combined)}")
    table_frame.pack_forget()


# ============ ОСНОВНОЕ ОКНО ============

root = TkinterDnD.Tk()
style = ttk.Style()
style.configure("Treeview", rowheight=50)  # Например, 50 пикселей
root.title("ОГ")
root.geometry("1280x720")
root.configure(background="#2f2f2f")

main_frame = ctk.CTkFrame(root)
main_frame.pack(fill="both", expand=True, padx=10, pady=10)

# Левая панель
left_panel = ctk.CTkFrame(main_frame, width=500)
left_panel.pack(side="left", fill="both", expand=False, padx=10, pady=10)

title_label = ctk.CTkLabel(left_panel, text="Настройка", font=("Arial", 24, "bold"))
title_label.pack(pady=10)

report_type_frame = ctk.CTkFrame(left_panel)
report_type_frame.pack(pady=5)

report_type_label = ctk.CTkLabel(report_type_frame, text="Тип отчета:", font=("Arial", 14, "bold"))
report_type_label.pack(side="left", padx=5)

report_type_var = ctk.StringVar(value="Полусуточный")
report_type_combobox = ctk.CTkComboBox(
    report_type_frame,
    values=["Полусуточный", "Суточный", "Недельный", "Месячный"],
    # , "Недельный", "Месячный"],
    variable=report_type_var,
    command=lambda choice: update_buttons_visibility()
)
report_type_combobox.pack(side="left", padx=5)

# Выбор файлов (AIS, EDC)
file_selection_frame = ctk.CTkFrame(left_panel)
file_selection_frame.pack(pady=10)

ais_file_frame = ctk.CTkFrame(file_selection_frame)
ais_file_frame.pack(pady=5)

ais_file_label = ctk.CTkLabel(ais_file_frame, text="Файл АИС:", font=("Arial", 12))
ais_file_label.pack(side="left", padx=5)

ais_file_entry = ctk.CTkEntry(ais_file_frame, width=310)
ais_file_entry.pack(side="left", padx=5)
ais_file_entry.drop_target_register(DND_FILES)
ais_file_entry.dnd_bind("<<Drop>>", lambda e: drop_file(e, ais_file_entry))

ais_file_button = ctk.CTkButton(ais_file_frame, text="...", width=30, command=lambda: select_file(ais_file_entry))
ais_file_button.pack(side="left", padx=5)



# Дата + флажок "Утренний отчет"
date_frame = ctk.CTkFrame(left_panel)
date_frame.pack()
ctk.CTkButton(date_frame, text="Сегодня", command=set_today, width=30).pack(side="left", padx=10)

date_entry = DateEntry(date_frame, width=15, background='darkblue', foreground='white', borderwidth=2,
                       date_pattern="dd.mm.yyyy")
date_entry.pack(side="left", padx=10)
date_entry.set_date(datetime.today())

morning_report_frame = ctk.CTkFrame(date_frame)
morning_report_frame.pack(pady=5, padx=15)
fix_oiv_frame = ctk.CTkFrame(date_frame)
fix_oiv_frame.pack(pady=5, padx=15)

def check_time():
    now = datetime.now().time()
    target_time = datetime.strptime("15:30", "%H:%M").time()
    return now <= target_time


morning_var = ctk.BooleanVar(value=check_time())
morning_checkbox = ctk.CTkCheckBox(morning_report_frame, text="Утро", variable=morning_var)


fix_oiv_var = ctk.BooleanVar(value=check_time())
fix_oiv_checkbox = ctk.CTkCheckBox(fix_oiv_frame, text="Исправить КОД ОИВ", variable=fix_oiv_var)
# summer_report_frame = ctk.CTkFrame(date_frame)
# summer_report_frame.pack(pady=5, padx=15)
# summer_var = ctk.BooleanVar(value=True)
# summer_checkbox = ctk.CTkCheckBox(summer_report_frame, text="Лето", variable=summer_var)

morning_checkbox.pack()
fix_oiv_checkbox.pack()
# summer_checkbox.pack()

# Предыдущий период (округа)
previous_period_frame = ctk.CTkFrame(left_panel)
previous_period_frame.pack(pady=5, fill="both", expand=True)

previous_period_label = ctk.CTkLabel(previous_period_frame, text="Предыдущий период:", font=("Arial", 14, "bold"))
previous_period_label.pack(pady=5)

scroll_frame = ctk.CTkScrollableFrame(previous_period_frame, height=150)
scroll_frame.pack(fill="both", expand=True, padx=10, pady=5)

previous_period = {district: 0 for district in
                   ["ЦАО", "САО", "СВАО", "ВАО", "ЮВАО", "ЮАО", "ЮЗАО", "ЗАО", "СЗАО", "ЗелАО", "ТиНАО",
                    'ГБУ "АВД"', "Иные", "Общий итог"]}
#      'ГБУ "АВД"',
entry_vars = {}
max_label_width = max(len(d) for d in previous_period.keys())

for district in previous_period.keys():
    district_frame = ctk.CTkFrame(scroll_frame)
    district_frame.pack(fill="x", pady=2, padx=5)

    label = ctk.CTkLabel(district_frame, text=f"{district}:", width=max_label_width * 10, anchor="w")
    label.grid(row=0, column=0, sticky="w", padx=10)

    entry_var = ctk.StringVar(value="1")
    entry_vars[district] = entry_var

    entry_field = ctk.CTkEntry(district_frame, textvariable=entry_var, width=70)
    entry_field.grid(row=0, column=1, padx=10, sticky="w")

    district_frame.columnconfigure(0, weight=1)
    district_frame.columnconfigure(1, weight=1)

# Кнопки
button_frame = ctk.CTkFrame(left_panel)
button_frame.pack(fill="x", pady=10)

summary_btn = ctk.CTkButton(button_frame, text="Пересчитать итог", command=calculate_total)
summary_btn.pack(side="left", padx=20)

halfdaily_btn = ctk.CTkButton(button_frame, text="Полусуточный отчет", command=generate_report_halfdaily)
halfdaily_btn.pack(side="right", padx=20)

weekly_btn = ctk.CTkButton(button_frame, text="Недельный отчет", command=generate_report_weekly)
weekly_btn.pack(side="right", padx=20)

month_btn = ctk.CTkButton(button_frame, text="Месячный отчет", command=generate_report_month)
weekly_btn.pack(side="right", padx=20)

# Правая панель
right_panel = ctk.CTkFrame(main_frame, width=600)
right_panel.pack(side="right", fill="both", expand=True, padx=10, pady=10)

ctk.CTkLabel(right_panel, text="Консоль:", font=("Arial", 20, "bold")).pack(pady=10)
report_output = Text(right_panel, wrap="word", height=15)
report_output.pack(fill="both", expand=True, padx=10, pady=5)

sys.stdout = ConsoleText(report_output)
sys.stderr = ConsoleText(report_output)

open_folder_button = ctk.CTkButton(right_panel, text="Открыть папку", command=open_reports_folder)
open_folder_button.pack_forget()

table_frame = ctk.CTkFrame(right_panel)

tree_scrollbar = ttk.Scrollbar(table_frame, orient="vertical")
tree_scrollbar.pack(side="right", fill="y")

tree_columns = ("check", "Адрес объекта", "Район", "Округ")
tree = ttk.Treeview(table_frame, columns=tree_columns, show="headings", yscrollcommand=tree_scrollbar.set)

tree.heading("check", text="✓")
tree.column("check", width=30, anchor="center")

tree.heading("Адрес объекта", text="Адрес объекта")
tree.column("Адрес объекта", width=220)

tree.heading("Район", text="Район")
tree.column("Район", width=100)

tree.heading("Округ", text="Округ")
tree.column("Округ", width=80)

tree.pack(side="left", fill="both", expand=True)
tree_scrollbar.config(command=tree.yview)

tree.bind("<Button-1>", on_tree_click)

update_buttons_visibility()

root.mainloop()
