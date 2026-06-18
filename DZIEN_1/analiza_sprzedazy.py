from pathlib import Path
import csv
import re
import statistics
from collections import defaultdict, Counter
from datetime import datetime

from artifact_tool import Workbook, SpreadsheetFile


# ============================================================
# KONFIGURACJA
# ============================================================

INPUT_CSV = Path("sprzedaz_poprawne.csv")
OUTPUT_XLSX = Path("analiza_sprzedazy.xlsx")

CSV_SEPARATOR = ";"
CSV_ENCODING = "utf-8-sig"
DECIMAL_SEPARATOR = ","

# Reguła biznesowa:
# sprzedaż liczona jest tylko dla zamówień zrealizowanych.
STATUS_SPRZEDAZY = "Zrealizowane"


# ============================================================
# FUNKCJE POMOCNICZE
# ============================================================

def parse_decimal_pl(value):
    """
    Zamienia liczbę z polskim przecinkiem dziesiętnym na float.

    Przykład:
    "2384,05" -> 2384.05
    """
    if value is None:
        return None

    text = str(value).strip()

    if text == "":
        return None

    text = text.replace(" ", "")
    text = text.replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return None


def parse_int(value):
    """
    Zamienia tekst na liczbę całkowitą.
    """
    if value is None:
        return None

    text = str(value).strip()

    if re.fullmatch(r"-?\d+", text):
        return int(text)

    return None


def parse_date(value):
    """
    Parsuje datę w formacie YYYY-MM-DD.
    """
    if value is None:
        return None

    text = str(value).strip()

    try:
        return datetime.strptime(text, "%Y-%m-%d").date()
    except ValueError:
        return None


def month_key(date_value):
    """
    Zwraca miesiąc w formacie YYYY-MM.
    """
    return f"{date_value.year:04d}-{date_value.month:02d}"


def safe_avg(values):
    """
    Liczy średnią, jeśli lista nie jest pusta.
    """
    return sum(values) / len(values) if values else 0


def write_table(sheet, start_cell, headers, rows, table_name=None):
    """
    Wpisuje tabelę do arkusza.
    """
    matrix = [headers] + rows
    sheet.get_range(start_cell).write_values(matrix)

    row_count = len(matrix)
    col_count = len(headers)

    # Zakładamy, że start_cell to A1.
    # W tym skrypcie używamy wyłącznie A1.
    end_col_letter = chr(ord("A") + col_count - 1)
    table_range = f"A1:{end_col_letter}{row_count}"

    if table_name:
        sheet.tables.add(table_range, True, table_name)

    return table_range, row_count, col_count


def style_table(sheet, used_range, money_cols=None, number_cols=None, percent_cols=None):
    """
    Proste formatowanie arkusza.
    """
    money_cols = money_cols or []
    number_cols = number_cols or []
    percent_cols = percent_cols or []

    header = sheet.get_range("A1:Z1")
    header.format = {
        "fill": "#0F172A",
        "font": {"bold": True, "color": "#FFFFFF"},
        "horizontal_alignment": "center",
        "vertical_alignment": "center",
        "wrap_text": True,
    }

    sheet.freeze_panes.freeze_rows(1)

    # Autofit i rozsądne szerokości.
    sheet.get_range(used_range).format.autofit_columns()
    sheet.get_range(used_range).format.autofit_rows()

    for col in money_cols:
        sheet.get_range(f"{col}:{col}").format.number_format = '# ##0,00 zł'

    for col in number_cols:
        sheet.get_range(f"{col}:{col}").format.number_format = '0'

    for col in percent_cols:
        sheet.get_range(f"{col}:{col}").format.number_format = '0,00'


# ============================================================
# WCZYTANIE DANYCH
# ============================================================

def read_sales_data(input_csv):
    """
    Wczytuje plik CSV:
    - separator: średnik
    - kodowanie: UTF-8
    - przecinek dziesiętny
    """
    records = []

    with input_csv.open("r", encoding=CSV_ENCODING, newline="") as file:
        reader = csv.DictReader(file, delimiter=CSV_SEPARATOR)

        for row_number, row in enumerate(reader, start=2):
            order_date = parse_date(row.get("data_zamowienia"))
            quantity = parse_int(row.get("ilosc"))
            unit_price = parse_decimal_pl(row.get("cena_jedn"))
            discount = parse_decimal_pl(row.get("rabat_proc"))
            value = parse_decimal_pl(row.get("wartosc"))

            records.append({
                "row_number": row_number,
                "nr_zamowienia": row.get("nr_zamowienia", "").strip(),
                "data_zamowienia": order_date,
                "klient": row.get("klient", "").strip(),
                "miasto": row.get("miasto", "").strip(),
                "wojewodztwo": row.get("wojewodztwo", "").strip(),
                "kategoria": row.get("kategoria", "").strip(),
                "produkt": row.get("produkt", "").strip(),
                "ilosc": quantity,
                "cena_jedn": unit_price,
                "rabat_proc": discount,
                "wartosc": value,
                "status": row.get("status", "").strip(),
            })

    return records


# ============================================================
# ANALIZY
# ============================================================

def aggregate_sales(records):
    """
    Tworzy analizy:
    - sprzedaż według miesiąca,
    - sprzedaż według kategorii,
    - sprzedaż według miasta,
    - top 10 produktów,
    - średni rabat.
    """
    sales_records = [
        r for r in records
        if r["status"] == STATUS_SPRZEDAZY
        and r["wartosc"] is not None
        and r["data_zamowienia"] is not None
    ]

    by_month = defaultdict(lambda: {"sprzedaz": 0.0, "liczba_zamowien": 0})
    by_category = defaultdict(lambda: {"sprzedaz": 0.0, "liczba_zamowien": 0})
    by_city = defaultdict(lambda: {"sprzedaz": 0.0, "liczba_zamowien": 0})
    by_product = defaultdict(lambda: {"sprzedaz": 0.0, "liczba_zamowien": 0, "ilosc": 0})

    discount_all = []
    discount_by_category = defaultdict(list)

    for r in sales_records:
        value = r["wartosc"]
        month = month_key(r["data_zamowienia"])
        category = r["kategoria"]
        city = r["miasto"]
        product = r["produkt"]

        by_month[month]["sprzedaz"] += value
        by_month[month]["liczba_zamowien"] += 1

        by_category[category]["sprzedaz"] += value
        by_category[category]["liczba_zamowien"] += 1

        by_city[city]["sprzedaz"] += value
        by_city[city]["liczba_zamowien"] += 1

        by_product[product]["sprzedaz"] += value
        by_product[product]["liczba_zamowien"] += 1
        by_product[product]["ilosc"] += r["ilosc"] or 0

        if r["rabat_proc"] is not None:
            discount_all.append(r["rabat_proc"])
            discount_by_category[category].append(r["rabat_proc"])

    monthly_rows = []
    for month, data in sorted(by_month.items()):
        avg_order = data["sprzedaz"] / data["liczba_zamowien"] if data["liczba_zamowien"] else 0
        monthly_rows.append([
            month,
            round(data["sprzedaz"], 2),
            data["liczba_zamowien"],
            round(avg_order, 2),
        ])

    category_rows = []
    for category, data in sorted(by_category.items(), key=lambda x: x[1]["sprzedaz"], reverse=True):
        avg_order = data["sprzedaz"] / data["liczba_zamowien"] if data["liczba_zamowien"] else 0
        category_rows.append([
            category,
            round(data["sprzedaz"], 2),
            data["liczba_zamowien"],
            round(avg_order, 2),
        ])

    city_rows = []
    for city, data in sorted(by_city.items(), key=lambda x: x[1]["sprzedaz"], reverse=True):
        avg_order = data["sprzedaz"] / data["liczba_zamowien"] if data["liczba_zamowien"] else 0
        city_rows.append([
            city,
            round(data["sprzedaz"], 2),
            data["liczba_zamowien"],
            round(avg_order, 2),
        ])

    product_rows = []
    for product, data in sorted(by_product.items(), key=lambda x: x[1]["sprzedaz"], reverse=True)[:10]:
        avg_order = data["sprzedaz"] / data["liczba_zamowien"] if data["liczba_zamowien"] else 0
        product_rows.append([
            product,
            round(data["sprzedaz"], 2),
            data["liczba_zamowien"],
            data["ilosc"],
            round(avg_order, 2),
        ])

    avg_discount = safe_avg(discount_all)

    discount_rows = [
        ["Średni rabat ogółem", round(avg_discount, 2), len(discount_all)]
    ]

    for category, discounts in sorted(discount_by_category.items()):
        discount_rows.append([
            f"Średni rabat: {category}",
            round(safe_avg(discounts), 2),
            len(discounts),
        ])

    summary = {
        "liczba_rekordow": len(records),
        "liczba_zamowien_zrealizowanych": len(sales_records),
        "sprzedaz_razem": round(sum(r["wartosc"] for r in sales_records), 2),
        "sredni_rabat": round(avg_discount, 2),
        "zakres_dat_od": min((r["data_zamowienia"] for r in sales_records), default=None),
        "zakres_dat_do": max((r["data_zamowienia"] for r in sales_records), default=None),
    }

    return {
        "summary": summary,
        "monthly_rows": monthly_rows,
        "category_rows": category_rows,
        "city_rows": city_rows,
        "product_rows": product_rows,
        "discount_rows": discount_rows,
    }


# ============================================================
# RAPORT XLSX
# ============================================================

def build_workbook(analysis, output_xlsx):
    """
    Buduje plik XLSX z osobnymi arkuszami i wykresami.
    """
    wb = Workbook.create()

    # --------------------------------------------------------
    # README
    # --------------------------------------------------------
    readme = wb.worksheets.add("README")
    summary = analysis["summary"]

    readme_data = [
        ["Raport analizy sprzedaży", ""],
        ["Plik źródłowy", str(INPUT_CSV)],
        ["Separator CSV", CSV_SEPARATOR],
        ["Kodowanie", CSV_ENCODING],
        ["Separator dziesiętny", DECIMAL_SEPARATOR],
        ["Reguła sprzedaży", f"Tylko status: {STATUS_SPRZEDAZY}"],
        ["Liczba rekordów w pliku", summary["liczba_rekordow"]],
        ["Liczba zamówień zrealizowanych", summary["liczba_zamowien_zrealizowanych"]],
        ["Sprzedaż razem", summary["sprzedaz_razem"]],
        ["Średni rabat", summary["sredni_rabat"]],
        ["Zakres dat od", str(summary["zakres_dat_od"])],
        ["Zakres dat do", str(summary["zakres_dat_do"])],
    ]
    readme.get_range("A1:B12").values = readme_data
    readme.get_range("A1:B1").format = {
        "fill": "#0F172A",
        "font": {"bold": True, "color": "#FFFFFF", "size": 14},
    }
    readme.get_range("A2:A12").format = {"font": {"bold": True}}
    readme.get_range("B9:B10").format.number_format = '# ##0,00'
    readme.get_range("A1:B12").format.autofit_columns()

    # --------------------------------------------------------
    # Sprzedaż według miesiąca
    # --------------------------------------------------------
    s_month = wb.worksheets.add("Sprzedaz_miesiac")
    headers = ["Miesiąc", "Sprzedaż", "Liczba zamówień", "Średnia wartość zamówienia"]
    rng, rows, cols = write_table(
        s_month,
        "A1",
        headers,
        analysis["monthly_rows"],
        "TabelaSprzedazMiesiac",
    )
    style_table(s_month, rng, money_cols=["B", "D"], number_cols=["C"])

    chart = s_month.charts.add("line", s_month.get_range(f"A1:B{rows}"))
    chart.title_text = "Sprzedaż według miesiąca"
    chart.set_position("F2", "N20")

    # --------------------------------------------------------
    # Sprzedaż według kategorii
    # --------------------------------------------------------
    s_cat = wb.worksheets.add("Sprzedaz_kategoria")
    headers = ["Kategoria", "Sprzedaż", "Liczba zamówień", "Średnia wartość zamówienia"]
    rng, rows, cols = write_table(
        s_cat,
        "A1",
        headers,
        analysis["category_rows"],
        "TabelaSprzedazKategoria",
    )
    style_table(s_cat, rng, money_cols=["B", "D"], number_cols=["C"])

    chart = s_cat.charts.add("bar", s_cat.get_range(f"A1:B{rows}"))
    chart.title_text = "Sprzedaż według kategorii"
    chart.set_position("F2", "N20")

    # --------------------------------------------------------
    # Sprzedaż według miasta
    # --------------------------------------------------------
    s_city = wb.worksheets.add("Sprzedaz_miasto")
    headers = ["Miasto", "Sprzedaż", "Liczba zamówień", "Średnia wartość zamówienia"]
    rng, rows, cols = write_table(
        s_city,
        "A1",
        headers,
        analysis["city_rows"],
        "TabelaSprzedazMiasto",
    )
    style_table(s_city, rng, money_cols=["B", "D"], number_cols=["C"])

    # --------------------------------------------------------
    # Top 10 produktów
    # --------------------------------------------------------
    s_prod = wb.worksheets.add("Top10_produkty")
    headers = ["Produkt", "Sprzedaż", "Liczba zamówień", "Ilość sprzedana", "Średnia wartość zamówienia"]
    rng, rows, cols = write_table(
        s_prod,
        "A1",
        headers,
        analysis["product_rows"],
        "TabelaTop10Produkty",
    )
    style_table(s_prod, rng, money_cols=["B", "E"], number_cols=["C", "D"])

    chart = s_prod.charts.add("bar", s_prod.get_range(f"A1:B{rows}"))
    chart.title_text = "Top 10 produktów według sprzedaży"
    chart.set_position("G2", "O20")

    # --------------------------------------------------------
    # Średni rabat
    # --------------------------------------------------------
    s_disc = wb.worksheets.add("Sredni_rabat")
    headers = ["Analiza", "Średni rabat [%]", "Liczba zamówień"]
    rng, rows, cols = write_table(
        s_disc,
        "A1",
        headers,
        analysis["discount_rows"],
        "TabelaSredniRabat",
    )
    style_table(s_disc, rng, percent_cols=["B"], number_cols=["C"])

    # --------------------------------------------------------
    # Eksport
    # --------------------------------------------------------
    SpreadsheetFile.export_xlsx(wb).save(str(output_xlsx))

    return wb


def main():
    records = read_sales_data(INPUT_CSV)
    analysis = aggregate_sales(records)
    build_workbook(analysis, OUTPUT_XLSX)

    print("Analiza zakończona.")
    print(f"Plik XLSX zapisany jako: {OUTPUT_XLSX}")
    print(f"Liczba rekordów: {analysis['summary']['liczba_rekordow']}")
    print(f"Liczba zamówień zrealizowanych: {analysis['summary']['liczba_zamowien_zrealizowanych']}")
    print(f"Sprzedaż razem: {analysis['summary']['sprzedaz_razem']}")
    print(f"Średni rabat: {analysis['summary']['sredni_rabat']}%")


if __name__ == "__main__":
    main()
