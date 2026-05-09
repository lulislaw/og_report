def make_keys_table(
    data_rows: int,
    include_ob: bool = False,
    include_ot: bool = False,
) -> list[list[str]]:
    """
    Генерирует таблицу ключей вида:
    *t1t*, *t1c*, ..., *t1su*
    ...
    empty, *tsuc*, ..., *tsusu*

    data_rows — количество строк данных до итоговой строки.
    include_ob — добавить колонку *tNob*
    include_ot — добавить колонку *tNot*
    """

    district_suffixes = [
        "c",
        "s",
        "sv",
        "v",
        "yv",
        "y",
        "yz",
        "z",
        "sz",
        "ze",
        "ti",
    ]

    rows = []

    for i in range(1, data_rows + 1):
        row = [f"*t{i}t*"]

        for suffix in district_suffixes:
            row.append(f"*t{i}{suffix}*")

        if include_ob:
            row.append(f"*t{i}ob*")

        if include_ot:
            row.append(f"*t{i}ot*")

        row.append(f"*t{i}su*")

        rows.append(row)

    total_row = ["empty"]

    for suffix in district_suffixes:
        total_row.append(f"*tsu{suffix}*")

    if include_ob:
        total_row.append("*tsuob*")

    if include_ot:
        total_row.append("*tsuot*")

    total_row.append("*tsusu*")

    rows.append(total_row)

    return rows


def make_main_keys() -> list[list[str]]:
    return [
        ["*cpr*", "*ccu*", "*cpe*", "*cci*"],
        ["*spr*", "*scu*", "*spe*", "*sci*"],
        ["*svpr*", "*svcu*", "*svpe*", "*svci*"],
        ["*vpr*", "*vcu*", "*vpe*", "*vci*"],
        ["*yvpr*", "*yvcu*", "*yvpe*", "*yvci*"],
        ["*ypr*", "*ycu*", "*ype*", "*yci*"],
        ["*yzpr*", "*yzcu*", "*yzpe*", "*yzci*"],
        ["*zpr*", "*zcu*", "*zpe*", "*zci*"],
        ["*szpr*", "*szcu*", "*szpe*", "*szci*"],
        ["*zepr*", "*zecu*", "*zepe*", "*zeci*"],
        ["*tipr*", "*ticu*", "*tipe*", "*tici*"],
        ["*obr*", "*obcu*", "*obpe*", "*obci*"],
        ["*otr*", "*otcu*", "*otpe*", "*otci*"],
        ["*supr*", "*sucu*", "*supe*", "*suci*"],
    ]


def generate_table(letter: str, row_length: int, num_rows: int) -> list[list[str]]:
    """
    Генерирует таблицу с заданной буквой.

    Пример:
    generate_table("c", 3, 4)

    Вернет ключи:
    @c1@, @c4@, @cs1@
    @c2@, @c5@, @cs2@
    @c3@, @c6@, @cs3@
    @ci1@, @ci2@, @cis@
    """

    table = [["" for _ in range(row_length)] for _ in range(num_rows)]

    counter = 1

    for i in range(row_length - 1):
        for j in range(num_rows - 1):
            table[j][i] = f"@{letter}{counter}@"
            counter += 1

    for j in range(num_rows - 1):
        table[j][-1] = f"@{letter}s{j + 1}@"

    for i in range(row_length - 1):
        table[-1][i] = f"@{letter}i{i + 1}@"

    table[-1][-1] = f"@{letter}is@"

    return table

keys_table = make_keys_table(
    data_rows=12,
    include_ob=False,
    include_ot=False,
)

keys_table_sum = make_keys_table(
    data_rows=11,
    include_ob=False,
    include_ot=False,
)

keys_table_full = make_keys_table(
    data_rows=12,
    include_ob=True,
    include_ot=False,
)

keys_table_win_full_half = make_keys_table(
    data_rows=12,
    include_ob=True,
    include_ot=True,
)

keys_table_sum_full_half = make_keys_table(
    data_rows=11,
    include_ob=True,
    include_ot=True,
)

keys_table_sum_full_FULL = keys_table_win_full_half

keys_table_svod = keys_table_win_full_half

keys_table_svod_summer = keys_table_sum_full_half

keys_table_weekly = keys_table_sum_full_half


keys_main = make_main_keys()
keys_main_full = keys_main


keys_weekly_widget = [
    ["*worn*", "*cln*", "*alln*", "*p*"],
    ["*worp*", "*clp*", "*allp*", "*p_empty*"],
]