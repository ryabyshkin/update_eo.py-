"""
update_eo.py — Автоматическое обновление файла ЕО из исходных отчётов 1С.

Источники (папка data/):
  - vypusk_shi.xlsx       → колонка "Поступление цех, шт"
  - opt.xlsx              → колонки "Заказ ОПТ, шт" и "Отгружено ОПТ, шт"
  - ostatok.xlsx          → колонка "Остатки 01.04 шт" (переименована в "Остатки")
  - prodazhi_nedelya.xlsx → новая колонка продаж шт + "Продажи 02.04.-08.04. руб"
  - rezerv_lamoda.xlsx    → колонка "Резерв LA 01.04 шт"
  - rezervy_obsh.xlsx     → колонка "Резерв  01.04 шт"

Шаблон: EO_template.xlsx  →  Результат: output/EO_updated.xlsx

Ключ связи: "Номенклатура+характеристика"
Формула ключа: Номенклатура + " (" + Характеристика_до_точкизапятой + ")"

ABC-анализ пересчитывается по каждому листу (дропу) по "Продажи ИТОГО, шт".
"""

import re
import shutil
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

warnings.filterwarnings("ignore")

# ── Пути ──────────────────────────────────────────────────────────────────────
DATA_DIR = Path("data")
TEMPLATE  = DATA_DIR / "EO_template.xlsx"
OUTPUT    = Path("output") / "EO_updated.xlsx"

# Листы ЕО, которые обновляем
TARGET_SHEETS = ["SS26", "FW25", "БАЗА", "Предыдущие дропы", "Сопутка и прочее"]

# Колонка-ключ в ЕО
KEY_COL = "Номенклатура+характеристика"

# Название новой колонки продаж (меняй каждую неделю здесь)
NEW_SALES_WEEK_LABEL = "Продажи 09.04.-15.04. шт"

# ── Вспомогательные функции ───────────────────────────────────────────────────

def normalize_key(s: str) -> str:
    """Приводим ключ к нижнему регистру и убираем лишние пробелы."""
    return str(s).strip().lower() if pd.notna(s) else ""


def make_key(nom: pd.Series, char: pd.Series) -> pd.Series:
    """
    Формирует ключ по формуле: =СЦЕПИТЬ(ном;" ";"(";хар;")")
    Характеристика обрезается до первого ";".
    """
    char_clean = char.astype(str).str.split(";").str[0].str.strip()
    nom_clean  = nom.astype(str).str.strip()
    key = nom_clean + " (" + char_clean + ")"
    return key.str.strip().str.lower()


def read_1c_file(path: Path, skiprows: int = 4) -> pd.DataFrame:
    """
    Читает файл 1С с двойной шапкой (строки 0-3 — заголовок отчёта).
    skiprows=4 → строка 4 становится заголовком.
    Колонки дублируются из-за merged cells — берём первое непустое значение.
    """
    df = pd.read_excel(path, header=None, skiprows=skiprows)
    return df


# ── 1. Выпуск ШИ ──────────────────────────────────────────────────────────────

def process_vypusk(path: Path) -> pd.Series:
    """
    Возвращает Series: key → сумма количества (поступление цех).
    """
    raw = pd.read_excel(path, header=None)
    # row3 = ["Номенклатура", nan, "Характеристика", nan, "Итого"]
    # row4 = [nan, nan, nan, nan, "Количество"]
    # данные с row5
    df = raw.iloc[5:].copy()
    df.columns = range(len(df.columns))
    df = df.rename(columns={0: "ном", 2: "хар", 4: "кол"})
    df = df[["ном", "хар", "кол"]].dropna(subset=["кол"])
    df["ном"] = df["ном"].astype(str).str.strip()
    df["хар"] = df["хар"].astype(str).str.strip()
    # убираем строки-итоги (номенклатура пустая или nan)
    df = df[df["ном"].str.lower() != "nan"]
    df["key"] = make_key(df["ном"], df["хар"])
    df["кол"] = pd.to_numeric(df["кол"], errors="coerce").fillna(0)
    return df.groupby("key")["кол"].sum()


# ── 2. ОПТ ────────────────────────────────────────────────────────────────────

def process_opt(path: Path) -> pd.DataFrame:
    """
    Возвращает DataFrame с колонками: key, заказано, отгружено.
    Удаляем ненужные денежные колонки, оставляем Заказано(col7) и Отгружено(col8).
    """
    raw = pd.read_excel(path, header=None)
    df = raw.iloc[6:].copy()       # данные с 7-й строки (0-based row6)
    df.columns = range(len(df.columns))
    # col0=Номенклатура, col2=Характеристика, col7=Заказано, col8=Отгружено
    df = df.rename(columns={0: "ном", 2: "хар", 7: "заказано", 8: "отгружено"})
    df = df[["ном", "хар", "заказано", "отгружено"]]
    df["ном"] = df["ном"].astype(str).str.strip()
    df = df[df["ном"].str.lower() != "nan"]
    df["key"] = make_key(df["ном"], df["хар"])
    df["заказано"]  = pd.to_numeric(df["заказано"],  errors="coerce").fillna(0)
    df["отгружено"] = pd.to_numeric(df["отгружено"], errors="coerce").fillna(0)
    grp = df.groupby("key")[["заказано", "отгружено"]].sum().reset_index()
    return grp


# ── 3. Остатки ────────────────────────────────────────────────────────────────

STOCK_COLS = [
    "В пути в контент отдел Москва Бережковская (В пути)",
    "В пути в магазин Красноярск (В пути)",
    "В пути в магазин Москва Авиапарк (В пути)",
    "В пути в магазин Новосибирск (В пути)",
    "В пути в магазин С-Петербург (В пути)",
    "В пути в магазин С. Вражек (В пути)",
    "В пути в магазин Томск (В пути)",
    "В пути на ГП Москва (Бережковская)",
    "Готовая продукция Москва (Бережковская)",
    "Готовая продукция Томск МАКСПРО",
    "Контент отдел Москва Бережковская",
    "Магазин Красноярск",
    "Магазин Москва Авиапарк",
    "Магазин Москва С. Вражек",
    "Магазин Новосибирск",
    "Магазин С-Петербург",
    "Магазин Томск",
]

def process_ostatok(path: Path) -> pd.Series:
    """
    row3 = названия складов, row4 = "Свободный остаток".
    Номенклатура в col3, Характеристика в col5, склады с col6.
    Возвращает Series: key → итого (сумма по всем складам).
    """
    raw = pd.read_excel(path, header=None)
    # Строим заголовок из двух строк
    h3 = raw.iloc[3].fillna("").astype(str).str.strip()
    h4 = raw.iloc[4].fillna("").astype(str).str.strip()
    # Для складских колонок заголовок = h3 (название склада)
    headers = []
    for i, (a, b) in enumerate(zip(h3, h4)):
        if i in (0, 1, 2, 3, 4, 5):   # служебные
            headers.append(a if a else f"col{i}")
        else:
            headers.append(a if a else f"col{i}")
    df = raw.iloc[5:].copy()
    df.columns = headers[:len(df.columns)] + [f"col{i}" for i in range(len(headers), len(df.columns))]
    df = df.rename(columns={"Номенклатура": "ном", "Характеристика": "хар"})
    df["ном"] = df["ном"].astype(str).str.strip()
    df = df[df["ном"].str.lower() != "nan"]
    # суммируем склады
    present_stock_cols = [c for c in STOCK_COLS if c in df.columns]
    for c in present_stock_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["итого"] = df[present_stock_cols].sum(axis=1) if present_stock_cols else 0
    df["key"] = make_key(df["ном"], df["хар"])
    return df.groupby("key")["итого"].sum()


# ── 4. Продажи за неделю ──────────────────────────────────────────────────────

def process_prodazhi(path: Path) -> pd.DataFrame:
    """
    Формула: продажи_шт = V - AC,  выручка = W - AD.
    V=col21(Кол. тов. за текущий период), W=col22(Сумма со скидкой за текущий период)
    AC=col28(Кол. возвратов), AD=col29(Сумма возвратов)
    Строки с данными начинаются после row20 (строки-итоги/источники заканчиваются).
    """
    raw = pd.read_excel(path, header=None)
    # Берём только строки где col0=Номенклатура, col1=Характеристика
    # Итоговые строки (Итого, источники) — col1 = NaN и col0 не пустое
    # Строки с товарами: col1 не пустой
    df = raw.iloc[1:].copy()   # пропускаем row0 (даты)
    df.columns = range(len(df.columns))
    df = df.rename(columns={0: "ном", 1: "хар", 21: "V", 22: "W", 28: "AC", 29: "AD"})
    # Оставляем только строки с непустой характеристикой (товарные строки)
    df["хар"] = df["хар"].astype(str).str.strip()
    df = df[df["хар"].str.lower() != "nan"]
    df["ном"] = df["ном"].astype(str).str.strip()
    df = df[df["ном"].str.lower() != "nan"]
    for c in ["V", "W", "AC", "AD"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["продажи_шт"] = df["V"] - df["AC"]
    df["выручка"]    = df["W"] - df["AD"]
    df["key"] = make_key(df["ном"], df["хар"])
    grp = df.groupby("key")[["продажи_шт", "выручка"]].sum().reset_index()
    return grp


# ── 5. Резерв Ламода ──────────────────────────────────────────────────────────

def process_rezerv_lamoda(path: Path) -> pd.Series:
    """
    Структура: Место хранения | Характеристика | Количество | Резерв | Свободный остаток
    Строка 0 = реальный заголовок ["Номенклатура", "Характеристика", ...]
    Данные с row1. Номенклатура в col0 (Место хранения), col1 = Характеристика.
    Резерв берём из колонки "Резерв" (col3) — это количество зарезервированного.
    Знак минус в резерве → берём abs().
    """
    df = pd.read_excel(path)
    # row0 уже стал заголовком при чтении, но pandas взял ["Место хранения","Unnamed:1","Количество","Резерв","Свободный остаток"]
    # Первая строка данных — row0 = ["Номенклатура","Характеристика",nan,nan,nan] — пропускаем
    df = df.iloc[1:].copy()
    df.columns = ["ном", "хар", "кол", "резерв", "своб"]
    # Строки-итоги (например "КУПИШУЗ ООО") — характеристика NaN
    df["хар"] = df["хар"].astype(str).str.strip()
    df = df[df["хар"].str.lower() != "nan"]
    df["ном"] = df["ном"].astype(str).str.strip()
    df = df[df["ном"].str.lower() != "nan"]
    df["резерв"] = pd.to_numeric(df["резерв"], errors="coerce").fillna(0).abs()
    df["key"] = make_key(df["ном"], df["хар"])
    return df.groupby("key")["резерв"].sum()


# ── 6. Резервы общие ──────────────────────────────────────────────────────────

def process_rezervy_obsh(path: Path) -> pd.Series:
    """
    Та же структура что и ОПТ.
    col0=Номенклатура, col2=Характеристика, col10=Зарезервировано на складе.
    """
    raw = pd.read_excel(path, header=None)
    df = raw.iloc[6:].copy()
    df.columns = range(len(df.columns))
    df = df.rename(columns={0: "ном", 2: "хар", 10: "зарез"})
    df["ном"] = df["ном"].astype(str).str.strip()
    df = df[df["ном"].str.lower() != "nan"]
    df["key"] = make_key(df["ном"], df["хар"])
    df["зарез"] = pd.to_numeric(df["зарез"], errors="coerce").fillna(0)
    return df.groupby("key")["зарез"].sum()


# ── ABC-анализ ────────────────────────────────────────────────────────────────

def compute_abc(series: pd.Series) -> pd.Series:
    """
    Входит Series с продажами (индекс = порядковый по строкам).
    Возвращает Series с A/B/C.
    """
    total = series.sum()
    if total == 0:
        return pd.Series([""] * len(series), index=series.index)
    sorted_idx = series.sort_values(ascending=False).index
    cumsum = series[sorted_idx].cumsum() / total
    abc = pd.Series(index=series.index, dtype=str)
    abc[sorted_idx[cumsum <= 0.8]]                              = "A"
    abc[sorted_idx[(cumsum > 0.8) & (cumsum <= 0.95)]]         = "B"
    abc[sorted_idx[cumsum > 0.95]]                              = "C"
    abc = abc.fillna("")
    return abc


# ── Главная функция обновления ─────────────────────────────────────────────────

def update_eo():
    print("=== Загрузка исходных данных ===")

    vypusk   = process_vypusk(DATA_DIR / "vypusk_shi.xlsx")
    print(f"  Выпуск ШИ:       {len(vypusk)} позиций")

    opt_df   = process_opt(DATA_DIR / "opt.xlsx")
    print(f"  ОПТ:             {len(opt_df)} позиций")

    ostatok  = process_ostatok(DATA_DIR / "ostatok.xlsx")
    print(f"  Остатки:         {len(ostatok)} позиций")

    prodazhi = process_prodazhi(DATA_DIR / "prodazhi_nedelya.xlsx")
    print(f"  Продажи (неделя):{len(prodazhi)} позиций")

    rez_la   = process_rezerv_lamoda(DATA_DIR / "rezerv_lamoda.xlsx")
    print(f"  Резерв Ламода:   {len(rez_la)} позиций")

    rez_obsh = process_rezervy_obsh(DATA_DIR / "rezervy_obsh.xlsx")
    print(f"  Резервы общие:   {len(rez_obsh)} позиций")

    # Копируем шаблон → output
    OUTPUT.parent.mkdir(exist_ok=True)
    shutil.copy(TEMPLATE, OUTPUT)

    print("\n=== Обновление листов ЕО ===")
    wb = load_workbook(OUTPUT)

    for sheet_name in TARGET_SHEETS:
        if sheet_name not in wb.sheetnames:
            print(f"  [ПРОПУСК] Лист '{sheet_name}' не найден")
            continue

        ws = wb[sheet_name]
        print(f"\n  Лист: {sheet_name}")

        # Читаем лист через pandas для получения данных
        df_sheet = pd.read_excel(OUTPUT, sheet_name=sheet_name)
        if KEY_COL not in df_sheet.columns:
            print(f"    [ОШИБКА] Колонка '{KEY_COL}' не найдена")
            continue

        n_rows = len(df_sheet)

        # Находим индексы нужных колонок
        cols = list(df_sheet.columns)

        def col_idx(name):
            """0-based индекс колонки → 1-based для openpyxl."""
            try:
                return cols.index(name) + 1
            except ValueError:
                return None

        # Индексы целевых колонок в Excel (1-based)
        ci_key          = col_idx(KEY_COL)
        ci_vypusk       = col_idx("Поступление цех, шт")
        ci_opt_zak      = col_idx("Заказ ОПТ, шт")
        ci_opt_otg      = col_idx("Отгружено ОПТ, шт")
        ci_ostatok      = col_idx("Остатки 01.04 шт")
        ci_rez_obsh     = col_idx("Резерв  01.04 шт")
        ci_rez_la       = col_idx("Резерв LA 01.04 шт")
        ci_prodazhi_rub = col_idx("Продажи 02.04.-08.04. руб")

        # Ищем последнюю существующую колонку продаж шт, чтобы вставить новую после
        sales_cols = [c for c in cols if re.match(r"Продажи \d{2}\.\d{2}", c) and "шт" in c and "ИТОГО" not in c]
        itogo_col  = col_idx("Продажи ИТОГО, шт")
        # Новая колонка продаж — вставляем перед "Продажи ИТОГО, шт" если её нет
        if NEW_SALES_WEEK_LABEL not in cols and itogo_col:
            # Вставляем колонку в openpyxl перед ИТОГО
            ws.insert_cols(itogo_col)
            ws.cell(row=1, column=itogo_col).value = NEW_SALES_WEEK_LABEL
            # Сдвигаем все ci после вставки
            def shift(ci):
                return ci + 1 if ci and ci >= itogo_col else ci
            ci_opt_zak      = shift(ci_opt_zak)
            ci_opt_otg      = shift(ci_opt_otg)
            ci_ostatok      = shift(ci_ostatok)
            ci_rez_obsh     = shift(ci_rez_obsh)
            ci_rez_la       = shift(ci_rez_la)
            ci_prodazhi_rub = shift(ci_prodazhi_rub)
            ci_new_sales    = itogo_col      # теперь наша новая колонка
            itogo_col      += 1
            print(f"    Добавлена колонка: {NEW_SALES_WEEK_LABEL}")
        else:
            ci_new_sales = col_idx(NEW_SALES_WEEK_LABEL) if NEW_SALES_WEEK_LABEL in cols else None

        # Строим lookup-словарь key → row_number (2-based, строка 1 = заголовок)
        # Перечитываем ws напрямую
        opt_lookup = opt_df.set_index("key")

        for row_num in range(2, n_rows + 2):   # строки Excel (2..n+1)
            cell_key = ws.cell(row=row_num, column=ci_key).value
            if not cell_key:
                continue
            key = normalize_key(str(cell_key))

            # ── Выпуск ШИ → Поступление цех, шт
            if ci_vypusk:
                val = vypusk.get(key, None)
                if val is not None:
                    ws.cell(row=row_num, column=ci_vypusk).value = int(val)
                elif ws.cell(row=row_num, column=ci_vypusk).value == "нет данных":
                    ws.cell(row=row_num, column=ci_vypusk).value = None

            # ── ОПТ → Заказ ОПТ, Отгружено ОПТ
            if key in opt_lookup.index:
                row_opt = opt_lookup.loc[key]
                if ci_opt_zak:
                    v = row_opt["заказано"]
                    ws.cell(row=row_num, column=ci_opt_zak).value = int(v) if v else None
                if ci_opt_otg:
                    v = row_opt["отгружено"]
                    ws.cell(row=row_num, column=ci_opt_otg).value = int(v) if v else None
            else:
                if ci_opt_zak and ws.cell(row=row_num, column=ci_opt_zak).value == "нет данных":
                    ws.cell(row=row_num, column=ci_opt_zak).value = None
                if ci_opt_otg and ws.cell(row=row_num, column=ci_opt_otg).value == "нет данных":
                    ws.cell(row=row_num, column=ci_opt_otg).value = None

            # ── Остатки → Остатки 01.04 шт
            if ci_ostatok:
                val = ostatok.get(key, None)
                ws.cell(row=row_num, column=ci_ostatok).value = int(val) if val is not None else None

            # ── Резерв общий → Резерв  01.04 шт
            if ci_rez_obsh:
                val = rez_obsh.get(key, None)
                ws.cell(row=row_num, column=ci_rez_obsh).value = int(val) if val is not None else None

            # ── Резерв Ламода → Резерв LA 01.04 шт
            if ci_rez_la:
                val = rez_la.get(key, None)
                ws.cell(row=row_num, column=ci_rez_la).value = int(val) if val is not None else None

            # ── Продажи (новая неделя) шт и руб
            prod_lookup = prodazhi.set_index("key")
            if key in prod_lookup.index:
                row_pr = prod_lookup.loc[key]
                if ci_new_sales:
                    ws.cell(row=row_num, column=ci_new_sales).value = int(row_pr["продажи_шт"])
                if ci_prodazhi_rub:
                    ws.cell(row=row_num, column=ci_prodazhi_rub).value = round(float(row_pr["выручка"]), 2)

        # ── Обновляем формулу Продажи ИТОГО (добавляем новую колонку в SUM)
        if itogo_col and ci_new_sales:
            # Находим диапазон всех колонок продаж шт (от первой до новой включительно)
            from openpyxl.utils import get_column_letter
            first_sales_ci = None
            for i, c in enumerate(cols, 1):
                if re.match(r"Продажи \d{2}\.\d{2}", c) and "шт" in c and "ИТОГО" not in c:
                    first_sales_ci = i
                    break
            if first_sales_ci:
                col_from = get_column_letter(first_sales_ci)
                col_to   = get_column_letter(ci_new_sales)
                for row_num in range(2, n_rows + 2):
                    ws.cell(row=row_num, column=itogo_col).value = (
                        f"=SUM({col_from}{row_num}:{col_to}{row_num})"
                    )

        # ── ABC-анализ
        ci_abc      = col_idx("ABC-анализ")
        ci_itogo_sh = col_idx("Продажи ИТОГО, шт")
        if ci_abc and ci_itogo_sh:
            sales_values = []
            for row_num in range(2, n_rows + 2):
                v = ws.cell(row=row_num, column=ci_itogo_sh).value
                try:
                    sales_values.append(float(v) if v is not None else 0.0)
                except (TypeError, ValueError):
                    sales_values.append(0.0)
            sales_series = pd.Series(sales_values)
            abc_series   = compute_abc(sales_series)
            fill_a = PatternFill("solid", fgColor="C6EFCE")
            fill_b = PatternFill("solid", fgColor="FFEB9C")
            fill_c = PatternFill("solid", fgColor="FFC7CE")
            for i, row_num in enumerate(range(2, n_rows + 2)):
                abc_val = abc_series.iloc[i]
                cell = ws.cell(row=row_num, column=ci_abc)
                cell.value = abc_val
                if abc_val == "A":
                    cell.fill = fill_a
                elif abc_val == "B":
                    cell.fill = fill_b
                elif abc_val == "C":
                    cell.fill = fill_c

        print(f"    Обновлено: {n_rows} строк")

    wb.save(OUTPUT)
    print(f"\n✅ Готово! Файл сохранён: {OUTPUT}")


if __name__ == "__main__":
    update_eo()
