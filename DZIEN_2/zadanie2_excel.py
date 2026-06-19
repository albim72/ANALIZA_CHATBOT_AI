# -*- coding: utf-8 -*-

"""
Rozwiązanie części 2 zadania szkoleniowego.

Cel:
1. Wczytać plik CSV z fikcyjnymi danymi osobowymi.
2. Sprawdzić jakość danych.
3. Wykryć:
   - braki danych,
   - błędne adresy e-mail,
   - błędne telefony,
   - błędne numery PESEL,
   - błędne daty urodzenia,
   - niezgodność daty urodzenia z PESEL,
   - duplikaty.
4. Zapisać wyniki do pliku Excel.

Plik wejściowy:
kursanci_dane_osobowe_bledy_10000.csv

Plik wynikowy:
wyniki_analizy_danych_osobowych.xlsx
"""

import re
from datetime import date
from pathlib import Path

import pandas as pd


INPUT_CSV = "kursanci_dane_osobowe_bledy_10000.csv"
OUTPUT_XLSX = "wyniki_analizy_danych_osobowych.xlsx"

EXPECTED_COLUMNS = [
    "id_klienta",
    "imie",
    "nazwisko",
    "email",
    "telefon",
    "pesel",
    "data_urodzenia",
]


def clean_text(value):
    """Zamienia wartość na tekst i usuwa zbędne spacje."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def is_missing(value):
    """Sprawdza, czy pole jest puste."""
    return clean_text(value) == ""


def is_valid_email(email):
    """Prosta walidacja adresu e-mail."""
    email = clean_text(email)

    if not email:
        return False

    pattern = r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$"
    return bool(re.match(pattern, email))


def normalize_phone(phone):
    """Usuwa spacje, myślniki i nawiasy z numeru telefonu."""
    phone = clean_text(phone)
    phone = phone.replace(" ", "")
    phone = phone.replace("-", "")
    phone = phone.replace("(", "")
    phone = phone.replace(")", "")
    return phone


def is_valid_polish_phone(phone):
    """
    Sprawdza uproszczony format polskiego numeru telefonu.

    Akceptowane formaty:
    - 501222333
    - +48501222333
    - 48501222333
    """
    phone = normalize_phone(phone)

    if not phone:
        return False

    return bool(re.match(r"^(\+48\d{9}|48\d{9}|\d{9})$", phone))


def pesel_birth_date(pesel):
    """
    Odczytuje datę urodzenia z numeru PESEL.
    Zwraca obiekt date albo None.
    """
    pesel = clean_text(pesel)

    if not re.match(r"^\d{11}$", pesel):
        return None

    year = int(pesel[0:2])
    month = int(pesel[2:4])
    day = int(pesel[4:6])

    if 1 <= month <= 12:
        century = 1900
    elif 21 <= month <= 32:
        century = 2000
        month -= 20
    elif 41 <= month <= 52:
        century = 2100
        month -= 40
    elif 61 <= month <= 72:
        century = 2200
        month -= 60
    elif 81 <= month <= 92:
        century = 1800
        month -= 80
    else:
        return None

    try:
        return date(century + year, month, day)
    except ValueError:
        return None


def is_valid_pesel_checksum(pesel):
    """Sprawdza sumę kontrolną PESEL."""
    pesel = clean_text(pesel)

    if not re.match(r"^\d{11}$", pesel):
        return False

    weights = [1, 3, 7, 9, 1, 3, 7, 9, 1, 3]
    checksum_sum = sum(int(pesel[i]) * weights[i] for i in range(10))
    control_digit = (10 - checksum_sum % 10) % 10

    return control_digit == int(pesel[10])


def is_valid_pesel(pesel):
    """Sprawdza format, datę i sumę kontrolną PESEL."""
    pesel = clean_text(pesel)

    return (
        bool(re.match(r"^\d{11}$", pesel))
        and pesel_birth_date(pesel) is not None
        and is_valid_pesel_checksum(pesel)
    )


def parse_birth_date(value):
    """Parsuje datę w formacie RRRR-MM-DD."""
    value = clean_text(value)

    if not value:
        return pd.NaT

    return pd.to_datetime(value, errors="coerce", format="%Y-%m-%d")


def join_error_names(row, flag_columns):
    """Tworzy opis błędów dla jednego rekordu."""
    errors = []

    for column in flag_columns:
        if bool(row[column]):
            errors.append(column)

    return ", ".join(errors)


# ------------------------------------------------------------
# 1. Wczytanie danych
# ------------------------------------------------------------

input_path = Path(INPUT_CSV)

if not input_path.exists():
    raise FileNotFoundError(
        f"Nie znaleziono pliku: {INPUT_CSV}. "
        "Umieść plik CSV w tym samym katalogu co skrypt."
    )

df = pd.read_csv(
    INPUT_CSV,
    sep=";",
    dtype=str,
    keep_default_na=False,
    encoding="utf-8"
)

df.columns = [clean_text(column) for column in df.columns]

missing_columns = [
    column for column in EXPECTED_COLUMNS
    if column not in df.columns
]

if missing_columns:
    raise ValueError(f"Brakuje kolumn: {missing_columns}")

df = df[EXPECTED_COLUMNS].copy()

for column in EXPECTED_COLUMNS:
    df[column] = df[column].map(clean_text)


# ------------------------------------------------------------
# 2. Braki danych
# ------------------------------------------------------------

for column in EXPECTED_COLUMNS:
    df[f"brak_{column}"] = df[column].map(is_missing)


# ------------------------------------------------------------
# 3. Walidacja formatów
# ------------------------------------------------------------

df["bledny_email"] = ~df["email"].map(is_valid_email)
df["bledny_telefon"] = ~df["telefon"].map(is_valid_polish_phone)
df["bledny_pesel"] = ~df["pesel"].map(is_valid_pesel)

df["data_urodzenia_parsed"] = df["data_urodzenia"].map(parse_birth_date)
df["bledna_data_urodzenia"] = df["data_urodzenia_parsed"].isna()


# ------------------------------------------------------------
# 4. Zgodność daty urodzenia z PESEL
# ------------------------------------------------------------

df["data_z_pesel"] = df["pesel"].map(pesel_birth_date)
df["data_z_pesel"] = pd.to_datetime(df["data_z_pesel"], errors="coerce")

df["niezgodnosc_pesel_data"] = (
    df["data_z_pesel"].notna()
    & df["data_urodzenia_parsed"].notna()
    & (df["data_z_pesel"].dt.date != df["data_urodzenia_parsed"].dt.date)
)


# ------------------------------------------------------------
# 5. Duplikaty
# ------------------------------------------------------------

for column in ["id_klienta", "email", "pesel"]:
    non_empty = df[column] != ""

    df[f"duplikat_{column}"] = False
    df.loc[non_empty, f"duplikat_{column}"] = (
        df.loc[non_empty, column].duplicated(keep=False)
    )


# ------------------------------------------------------------
# 6. Podsumowanie błędów w rekordach
# ------------------------------------------------------------

FLAG_COLUMNS = [
    column for column in df.columns
    if column.startswith("brak_")
    or column.startswith("bledny_")
    or column.startswith("bledna_")
    or column.startswith("duplikat_")
    or column == "niezgodnosc_pesel_data"
]

df["liczba_bledow"] = df[FLAG_COLUMNS].sum(axis=1)

df["opis_bledow"] = df.apply(
    lambda row: join_error_names(row, FLAG_COLUMNS),
    axis=1
)

df["czy_rekord_problemowy"] = df["liczba_bledow"] > 0


# ------------------------------------------------------------
# 7. Tabele wynikowe
# ------------------------------------------------------------

summary = pd.DataFrame(
    [
        ["Liczba rekordów", len(df)],
        ["Liczba kolumn źródłowych", len(EXPECTED_COLUMNS)],
        ["Rekordy bez wykrytych błędów", int((df["liczba_bledow"] == 0).sum())],
        ["Rekordy z co najmniej jednym błędem", int(df["czy_rekord_problemowy"].sum())],
        ["Odsetek rekordów problemowych", round(df["czy_rekord_problemowy"].mean() * 100, 2)],
        ["Błędne adresy e-mail", int(df["bledny_email"].sum())],
        ["Błędne telefony", int(df["bledny_telefon"].sum())],
        ["Błędne numery PESEL", int(df["bledny_pesel"].sum())],
        ["Błędne daty urodzenia", int(df["bledna_data_urodzenia"].sum())],
        ["Niezgodność PESEL vs data urodzenia", int(df["niezgodnosc_pesel_data"].sum())],
        ["Duplikaty id_klienta", int(df["duplikat_id_klienta"].sum())],
        ["Duplikaty email", int(df["duplikat_email"].sum())],
        ["Duplikaty PESEL", int(df["duplikat_pesel"].sum())],
    ],
    columns=["metryka", "wartosc"]
)

missing_summary = pd.DataFrame(
    {
        "kolumna": EXPECTED_COLUMNS,
        "liczba_brakow": [
            int(df[f"brak_{column}"].sum())
            for column in EXPECTED_COLUMNS
        ],
    }
)

missing_summary["odsetek_brakow"] = (
    missing_summary["liczba_brakow"] / len(df) * 100
).round(2)

errors_by_type = pd.DataFrame(
    {
        "typ_bledu": FLAG_COLUMNS,
        "liczba_rekordow": [
            int(df[column].sum())
            for column in FLAG_COLUMNS
        ],
    }
)

errors_by_type["odsetek_rekordow"] = (
    errors_by_type["liczba_rekordow"] / len(df) * 100
).round(2)

errors_by_type = errors_by_type.sort_values(
    "liczba_rekordow",
    ascending=False
)

problem_records = df[df["czy_rekord_problemowy"]].copy()

problem_records = problem_records[
    EXPECTED_COLUMNS + ["liczba_bledow", "opis_bledow"]
].sort_values("liczba_bledow", ascending=False)

clean_records = df[df["liczba_bledow"] == 0][EXPECTED_COLUMNS].copy()

invalid_email = df[df["bledny_email"]][
    EXPECTED_COLUMNS + ["opis_bledow"]
]

invalid_phone = df[df["bledny_telefon"]][
    EXPECTED_COLUMNS + ["opis_bledow"]
]

invalid_pesel = df[df["bledny_pesel"]][
    EXPECTED_COLUMNS + ["opis_bledow"]
]

invalid_date = df[df["bledna_data_urodzenia"]][
    EXPECTED_COLUMNS + ["opis_bledow"]
]

duplicates = df[
    df["duplikat_id_klienta"]
    | df["duplikat_email"]
    | df["duplikat_pesel"]
][
    EXPECTED_COLUMNS + ["opis_bledow"]
]

recommendations = pd.DataFrame(
    [
        [1, "Ustalić obowiązkowe pola: id_klienta, imie, nazwisko, email, telefon, pesel, data_urodzenia."],
        [2, "Dodać walidację formatu e-mail i telefonu na etapie wprowadzania danych."],
        [3, "Sprawdzać PESEL: długość, cyfry, suma kontrolna i zgodność daty urodzenia."],
        [4, "Blokować duplikaty po id_klienta, PESEL oraz, zależnie od procesu, po e-mailu."],
        [5, "Oddzielić analizę braków danych od analizy błędnych formatów, bo to inne klasy problemów."],
    ],
    columns=["lp", "rekomendacja"]
)


# ------------------------------------------------------------
# 8. Zapis do Excela
# ------------------------------------------------------------

with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
    summary.to_excel(writer, sheet_name="01_podsumowanie", index=False)
    missing_summary.to_excel(writer, sheet_name="02_braki_danych", index=False)
    errors_by_type.to_excel(writer, sheet_name="03_bledy_wg_typu", index=False)
    problem_records.to_excel(writer, sheet_name="04_rekordy_problemowe", index=False)
    clean_records.to_excel(writer, sheet_name="05_rekordy_poprawne", index=False)
    invalid_email.to_excel(writer, sheet_name="06_bledne_email", index=False)
    invalid_phone.to_excel(writer, sheet_name="07_bledne_telefony", index=False)
    invalid_pesel.to_excel(writer, sheet_name="08_bledne_pesele", index=False)
    invalid_date.to_excel(writer, sheet_name="09_bledne_daty", index=False)
    duplicates.to_excel(writer, sheet_name="10_duplikaty", index=False)
    recommendations.to_excel(writer, sheet_name="11_rekomendacje", index=False)

    workbook = writer.book

    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"

        for column_cells in worksheet.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter

            for cell in column_cells[:200]:
                value = "" if cell.value is None else str(cell.value)
                max_length = max(max_length, len(value))

            worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)

        for cell in worksheet[1]:
            cell.font = cell.font.copy(bold=True)


print(f"Gotowe. Zapisano plik Excel: {OUTPUT_XLSX}")
print(f"Liczba rekordów: {len(df)}")
print(f"Rekordy problemowe: {int(df['czy_rekord_problemowy'].sum())}")
print(f"Rekordy poprawne: {int((df['liczba_bledow'] == 0).sum())}")
