import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os


def load_data():
    try:
        df_raw = pd.read_excel("rohdaten.xlsx")
        print(
            f"Datei erfolgreich geladen. {df_raw.shape[0]} Zeilen und {df_raw.shape[1]} Spalten."
        )
    except Exception as e:
        print(f"Fehler beim Laden der Datei: {e}")

    return df_raw


def sort_data(data):
    # Alle eindeutigen matnr ermitteln
    alle_matnr = data["matnr"].unique()
    print(f"Verarbeite {len(alle_matnr)} verschiedene Materialnummern...")

    # Liste für alle Ergebnisse
    alle_ergebnisse = []

    for matnr in alle_matnr:
        # Filter für spezifische Materialnummer
        df_filtered = data[data["matnr"] == matnr].copy()

        # Gruppierung und Summierung der wavor_bstlmg pro matnr, Baumarkt, Monat
        result = (
            df_filtered.groupby(["matnr", "Baumarkt", "bedmo"])
            .agg(
                {
                    "wavor_bstlmg": "sum",
                    "progmo": "first",
                    "prog_mg1": "sum",
                    "progmo2": "first",
                    "prog_mg2": "sum",
                }
            )
            .reset_index()
        )

        # Zu Gesamtergebnis hinzufügen
        alle_ergebnisse.append(result)

    # Alle Ergebnisse zusammenfügen
    gesamt_result = pd.concat(alle_ergebnisse, ignore_index=True)

    # Excel Export
    os.makedirs("./output", exist_ok=True)
    gesamt_result.to_excel("./output/alle_matnr_bestellungen.xlsx", index=False)

    print(
        f"Datei erstellt: alle_matnr_bestellungen.xlsx mit {len(gesamt_result)} Datensätzen"
    )
    return gesamt_result

def sum_art_monthly_by_baumarkt(data):
    # Alle eindeutigen matnr ermitteln
    alle_matnr = data["bedmo"].unique()
    print(f"Verarbeite {len(alle_matnr)} verschiedene Materialnummern...")

    # Liste für alle Ergebnisse
    alle_ergebnisse = []

    for matnr in alle_matnr:
        # Filter für spezifische Baumarkt
        df_filtered = data[data["bedmo"] == matnr].copy()

        # Gruppierung und Summierung der wavor_bstlmg pro matnr, Baumarkt, Monat
        result = (
            df_filtered.groupby(["Baumarkt", "bedmo"])
            .agg(
                {
                    "wavor_bstlmg": "sum",
                    "progmo": "first",
                    "prog_mg1": "sum",
                    "progmo2": "first",
                    "prog_mg2": "sum",
                }
            )
            .reset_index()
        )

        # Zu Gesamtergebnis hinzufügen
        alle_ergebnisse.append(result)

    # Alle Ergebnisse zusammenfügen
    gesamt_result = pd.concat(alle_ergebnisse, ignore_index=True)

    # Excel Export
    os.makedirs("./output", exist_ok=True)
    gesamt_result.to_excel("./output/sum_art_monthly_by_baumarkt.xlsx", index=False)

    print(
        f"Datei erstellt: sum_art_monthly_by_baumarkt.xlsx mit {len(gesamt_result)} Datensätzen"
    )
    return gesamt_result

def sort_BaumartArtikel(data):
    # Alle eindeutigen Baumärkte ermitteln
    alle_baumärkte = data["Baumarktartikel"].unique()
    print(f"Verarbeite {len(alle_baumärkte)} verschiedene Baumarktartikel...")

    # Liste für alle Ergebnisse
    alle_ergebnisse = []

    for baumarkt in alle_baumärkte:
        # Filter für spezifische Baumarkt
        df_filtered = data[data["Baumarktartikel"] == baumarkt].copy()

        # Gruppierung und Summierung der wavor_bstlmg pro Baumarkt, Artikel
        result = (
            df_filtered.groupby(["Baumarktartikel", "bedmo"])
            .agg(
                {
                    "wavor_bstlmg": "sum",
                    "progmo": "first",
                    "prog_mg1": "sum",
                    "progmo2": "first",
                    "prog_mg2": "sum",
                }
            )
            .reset_index()
        )

        # Zu Gesamtergebnis hinzufügen
        alle_ergebnisse.append(result)

    # Alle Ergebnisse zusammenfügen
    gesamt_result = pd.concat(alle_ergebnisse, ignore_index=True)
    gesamt_result.to_excel("./output/sorted_Baumarktartikel_bestellungen.xlsx", index=False)

    return gesamt_result

def main():
    # 1. Teilaufgabe
    data = load_data()

    # # Sort data by bedmo
    # sortedBedmo = sort_data(data)
    # # Count Artkel monthly by Baumarkt
    # sum_art_monthly_by_baumarkt(sortedBedmo)
    # # ----> muss nur noch geploted werden

    # Sort data by Baumarktartikel / Monat / Baumarkt
    sortedBArtikel = sort_BaumartArtikel(data)


if __name__ == "__main__":
    main()
