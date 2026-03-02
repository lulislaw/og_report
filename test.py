#
# from main_full import make_main_full_presentation
#
# ais_file = "Отчет_по_событиям_за_03_02_2026_15_30_04_02_2026_06_29.xlsx"
# previous_period = {
#         "ЦАО": 1,
#         "САО": 1,
#         "СВАО": 1,
#         "ВАО": 1,
#         "ЮВАО": 1,
#         "ЮАО": 1,
#         "ЮЗАО": 1,
#         "ЗАО": 1,
#         "СЗАО": 1,
#         "ЗелАО": 1,
#         "ТиНАО": 1,
#         'ГБУ "АВД"': 1,
#         'Иные': 1,
#     }
# make_main_full_presentation(ais_file, previous_period, "03.02.2026", False, True)
#


from svod import make_svod_presentation

make_svod_presentation("reports/03.02.2026 на 17.00/tmp_files/Обработанный АИС.xlsx",
                       "reports/03.02.2026 на 17.00/tmp_files/Обработанный АИС(Пуск отопления).xlsx", "03.02.2026",
                       False, False)


