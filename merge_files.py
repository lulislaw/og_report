import pandas as pd
import os
folder_path = 'files_for_merge/files_01'
csv = False
output_path = 'files_for_merge/files_01_final'
files = [
    os.path.join(folder_path, f)
    for f in os.listdir(folder_path)
    if f.endswith('.xlsx')
]
print(files)
df_list = []
for f in files:
    df = pd.read_excel(f)
    df_list.append(df)
df_merged = pd.concat(df_list, ignore_index=True)
filter = True

if filter:
    selected_columns = ['Наименование события', 'Наименование события КОД ОИВ', 'Статус во внешней системе',
                        'Ответственный', 'Адрес объекта', 'Район', 'Округ', 'Тип объекта',
                        'Дата создания во внешней системе']

    df_filtered = df_merged[selected_columns]

    if csv:
        df_filtered.to_csv(f"{output_path}.csv", index=False)
    else:
        df_filtered.to_excel(f"{output_path}.xlsx", index=False)
else:
    if csv:
        df_merged.to_csv(f"{output_path}.csv", index=False)
    else:
        df_merged.to_excel(f"{output_path}.xlsx", index=False)

print(f"Файлы успешно объединены и сохранены в {output_path}")

