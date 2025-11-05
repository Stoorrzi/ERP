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


def structure_data(df_raw):
    id_spalten = [
        "matnr",  # -> Artikel
        "Baumarktartikel",  # -> Teilegruppe
        "Baumarkt",  # -> Kunde
        "kundnr",  # -> Kunden-ID
    ]

    # --- Prognose 2026 verarbeiten ---
    df_2026 = df_raw[id_spalten + ["progmo", "prog_mg1"]].copy()
    df_2026 = df_2026.rename(
        columns={"progmo": "Prognose_Monat_Code", "prog_mg1": "Prognose_Menge"}
    )
    # Nur Zeilen behalten, die Prognosedaten enthalten
    df_2026 = df_2026.dropna(subset=["Prognose_Monat_Code", "Prognose_Menge"])

    # --- Prognose 2027 verarbeiten ---
    df_2027 = df_raw[id_spalten + ["progmo2", "prog_mg2"]].copy()
    df_2027 = df_2027.rename(
        columns={"progmo2": "Prognose_Monat_Code", "prog_mg2": "Prognose_Menge"}
    )
    df_2027 = df_2027.dropna(subset=["Prognose_Monat_Code", "Prognose_Menge"])

    # --- Beide Jahre zusammenführen ---
    df_prognose_lang = pd.concat([df_2026, df_2027], ignore_index=True)

    df_prognose_lang["Datum"] = pd.to_datetime(
        df_prognose_lang["Prognose_Monat_Code"].astype(int).astype(str), format="%Y%m"
    )
    # Alte Spalte entfernen
    df_prognose_lang = df_prognose_lang.drop(columns=["Prognose_Monat_Code"])

    print("\n--- Strukturierte 'lange' Daten (Auszug) ---")
    print(df_prognose_lang.head())

    # Output-Ordner erstellen falls er nicht existiert
    os.makedirs("./output", exist_ok=True)

    # Korrekte Verwendung von to_excel
    df_prognose_lang.to_excel("./output/prognose_lang.xlsx", index=False)

    return df_prognose_lang


def bottom_up_sum(df_prognose_lang):
    # Aggregation auf Ebene "Kunde" und "Datum"
    df_agg_kunde = (
        df_prognose_lang.groupby(["Baumarkt", "Datum"])
        .agg(Prognose_Original_Gesamt=("Prognose_Menge", "sum"))
        .reset_index()
    )

    print("\n--- Aggregiert auf Kunde & Monat (Bottom-Up-Summe) ---")
    print(df_agg_kunde.head())

    # Aggregation auf "Gesamt"-Ebene
    df_agg_gesamt = (
        df_prognose_lang.groupby("Datum")
        .agg(Prognose_Original_Gesamt=("Prognose_Menge", "sum"))
        .reset_index()
    )

    print("\n--- Aggregiert auf Gesamt & Monat ---")
    print(df_agg_gesamt.head())

    # Korrekte Verwendung von to_excel
    df_agg_gesamt.to_excel("./output/bottom_up_gesamt.xlsx", index=False)

    return df_agg_gesamt


def plot_trends(df_agg_gesamt):
    plt.figure(figsize=(15, 6))
    plt.plot(
        df_agg_gesamt["Datum"],
        df_agg_gesamt["Prognose_Original_Gesamt"],
        label="Bottom-Up Prognose (Gesamt)",
    )

    # Fügen Sie einen gleitenden Durchschnitt hinzu, um den Trend besser zu sehen
    df_agg_gesamt["Trend_12M"] = (
        df_agg_gesamt["Prognose_Original_Gesamt"]
        .rolling(window=12, center=True, min_periods=6)
        .mean()
    )
    plt.plot(
        df_agg_gesamt["Datum"],
        df_agg_gesamt["Trend_12M"],
        label="12-Monats-Trend",
        color="red",
        linestyle="--",
    )

    plt.title("Trendanalyse der Bottom-Up-Gesamtprognose (2026-2027)")
    plt.ylabel("Prognosemenge (Stück)")
    plt.xlabel("Datum")
    plt.legend()
    plt.grid(True)
    plt.show()

    # Erstellen einer Monats-Spalte für die Gruppierung
    df_agg_gesamt["Monat_Nr"] = df_agg_gesamt["Datum"].dt.month

    plt.figure(figsize=(12, 6))
    sns.boxplot(data=df_agg_gesamt, x="Monat_Nr", y="Prognose_Original_Gesamt")
    plt.title("Saisonale Analyse: Sind bestimmte Monate stärker?")
    plt.xlabel("Monat (1=Jan, 12=Dez)")
    plt.ylabel("Prognosemenge (Stück)")
    plt.show()


def analyse_umsatz_pro_baumarkt(df_prognose_lang):
    """
    Summiert alle umgesetzten Artikel pro Baumarkt und gibt die Ergebnisse in der Konsole aus.
    Zusätzlich wird eine monatliche Aufschlüsselung pro Baumarkt angezeigt.
    """
    # Aggregation der Prognosemenge pro Baumarkt (Gesamt)
    df_umsatz_baumarkt = (
        df_prognose_lang.groupby("Baumarkt")
        .agg(
            Gesamt_Prognosemenge=("Prognose_Menge", "sum"),
            Anzahl_Artikel=("matnr", "nunique"),
            Anzahl_Datenpunkte=("matnr", "count"),
        )
        .reset_index()
    )

    # Sortierung nach Gesamt_Prognosemenge (absteigend)
    df_umsatz_baumarkt = df_umsatz_baumarkt.sort_values(
        "Gesamt_Prognosemenge", ascending=False
    )

    print("\n" + "=" * 70)
    print("UMSATZANALYSE PRO BAUMARKT")
    print("=" * 70)

    # Gesamtübersicht
    gesamt_prognose = df_umsatz_baumarkt["Gesamt_Prognosemenge"].sum()
    anzahl_baumaerkte = len(df_umsatz_baumarkt)

    print(f"Anzahl Baumärkte: {anzahl_baumaerkte}")
    print(f"Gesamte Prognosemenge: {gesamt_prognose:,.0f} Stück")
    print(f"Durchschnitt pro Baumarkt: {gesamt_prognose/anzahl_baumaerkte:,.0f} Stück")
    print("-" * 70)

    # Detaillierte Ausgabe pro Baumarkt
    for idx, row in df_umsatz_baumarkt.iterrows():
        anteil = (row["Gesamt_Prognosemenge"] / gesamt_prognose) * 100
        print(
            f"{row['Baumarkt']:30} | "
            f"Menge: {row['Gesamt_Prognosemenge']:>10,.0f} | "
            f"Anteil: {anteil:>5.1f}% | "
            f"Artikel: {row['Anzahl_Artikel']:>4.0f}"
        )

    print("-" * 70)

    # Top 3 und Bottom 3 Baumärkte
    print("\nTOP 3 BAUMÄRKTE:")
    for idx, row in df_umsatz_baumarkt.head(3).iterrows():
        print(f"  {idx+1}. {row['Baumarkt']}: {row['Gesamt_Prognosemenge']:,.0f} Stück")

    print("\nSCHWÄCHSTE 3 BAUMÄRKTE:")
    for idx, row in df_umsatz_baumarkt.tail(3).iterrows():
        print(f"  {row['Baumarkt']}: {row['Gesamt_Prognosemenge']:,.0f} Stück")

    # NEUE FUNKTION: Monatliche Aufschlüsselung
    analyse_monatlicher_umsatz(df_prognose_lang, df_umsatz_baumarkt)

    return df_umsatz_baumarkt


def analyse_monatlicher_umsatz(df_prognose_lang, df_umsatz_baumarkt):
    """
    Analysiert den Umsatz pro Monat für jeden Baumarkt.
    """
    # Monat und Jahr aus Datum extrahieren
    df_prognose_lang["Jahr_Monat"] = df_prognose_lang["Datum"].dt.to_period("M")
    df_prognose_lang["Monat_Name"] = df_prognose_lang["Datum"].dt.strftime("%Y-%m")

    # Aggregation pro Baumarkt und Monat
    df_monatlich = (
        df_prognose_lang.groupby(["Baumarkt", "Jahr_Monat", "Monat_Name"])
        .agg(
            Monatliche_Menge=("Prognose_Menge", "sum"),
            Anzahl_Artikel_Monat=("matnr", "nunique"),
        )
        .reset_index()
    )

    print("\n" + "=" * 90)
    print("MONATLICHE UMSATZANALYSE PRO BAUMARKT")
    print("=" * 90)

    # Für die Top 5 Baumärkte detaillierte monatliche Aufschlüsselung
    top_5_baumaerkte = df_umsatz_baumarkt.head(5)["Baumarkt"].tolist()

    for baumarkt in top_5_baumaerkte:
        baumarkt_daten = df_monatlich[df_monatlich["Baumarkt"] == baumarkt].sort_values(
            "Jahr_Monat"
        )

        print(f"\n🏢 {baumarkt}")
        print("-" * 70)
        print("Monat      | Menge      | Artikel | Anteil am Baumarkt-Gesamt")
        print("-" * 70)

        baumarkt_gesamt = df_umsatz_baumarkt[
            df_umsatz_baumarkt["Baumarkt"] == baumarkt
        ]["Gesamt_Prognosemenge"].iloc[0]

        for _, row in baumarkt_daten.iterrows():
            anteil_monat = (row["Monatliche_Menge"] / baumarkt_gesamt) * 100
            print(
                f"{row['Monat_Name']} | "
                f"{row['Monatliche_Menge']:>10,.0f} | "
                f"{row['Anzahl_Artikel_Monat']:>7.0f} | "
                f"{anteil_monat:>6.1f}%"
            )

    # Gesamtanalyse aller Monate (alle Baumärkte zusammen)
    df_gesamt_monatlich = (
        df_prognose_lang.groupby(["Jahr_Monat", "Monat_Name"])
        .agg(
            Gesamt_Monatliche_Menge=("Prognose_Menge", "sum"),
            Gesamt_Artikel_Monat=("matnr", "nunique"),
            Anzahl_Baumaerkte=("Baumarkt", "nunique"),
        )
        .reset_index()
        .sort_values("Jahr_Monat")
    )

    print("\n" + "=" * 90)
    print("GESAMTÜBERSICHT ALLER MONATE (ALLE BAUMÄRKTE)")
    print("=" * 90)
    print("Monat      | Gesamt-Menge | Artikel | Baumärkte | Durchschn./Baumarkt")
    print("-" * 80)

    for _, row in df_gesamt_monatlich.iterrows():
        durchschnitt = row["Gesamt_Monatliche_Menge"] / row["Anzahl_Baumaerkte"]
        print(
            f"{row['Monat_Name']} | "
            f"{row['Gesamt_Monatliche_Menge']:>12,.0f} | "
            f"{row['Gesamt_Artikel_Monat']:>7.0f} | "
            f"{row['Anzahl_Baumaerkte']:>9.0f} | "
            f"{durchschnitt:>10,.0f}"
        )

    # Stärkste und schwächste Monate identifizieren
    staerkster_monat = df_gesamt_monatlich.loc[
        df_gesamt_monatlich["Gesamt_Monatliche_Menge"].idxmax()
    ]
    schwaechster_monat = df_gesamt_monatlich.loc[
        df_gesamt_monatlich["Gesamt_Monatliche_Menge"].idxmin()
    ]

    print(
        f"\n📈 STÄRKSTER MONAT: {staerkster_monat['Monat_Name']} mit {staerkster_monat['Gesamt_Monatliche_Menge']:,.0f} Stück"
    )
    print(
        f"📉 SCHWÄCHSTER MONAT: {schwaechster_monat['Monat_Name']} mit {schwaechster_monat['Gesamt_Monatliche_Menge']:,.0f} Stück"
    )

    # Excel-Export für weitere Analyse
    os.makedirs("./output", exist_ok=True)
    df_monatlich.to_excel("./output/monatlicher_umsatz_pro_baumarkt.xlsx", index=False)
    df_gesamt_monatlich.to_excel("./output/gesamt_monatlicher_umsatz.xlsx", index=False)

    print(f"\n📊 Monatliche Daten wurden in './output/' exportiert.")


def analysiere_hierarchieebenen(df_prognose_lang):
    """
    Identifiziert und analysiert die relevanten Hierarchieebenen
    (Artikel → Teilegruppe → Kundengruppe → Kunde → Gesamt)
    """
    print("\n" + "=" * 80)
    print("ANALYSE DER HIERARCHIEEBENEN")
    print("=" * 80)

    # 1. ARTIKEL-EBENE
    artikel_count = df_prognose_lang["matnr"].nunique()
    artikel_gesamt_menge = df_prognose_lang["Prognose_Menge"].sum()

    print(f"\n🔹 EBENE 1: ARTIKEL")
    print(f"   Anzahl eindeutige Artikel: {artikel_count:,}")
    print(f"   Gesamtmenge aller Artikel: {artikel_gesamt_menge:,.0f} Stück")

    # Top 5 Artikel
    top_artikel = (
        df_prognose_lang.groupby("matnr")["Prognose_Menge"]
        .sum()
        .sort_values(ascending=False)
        .head(5)
    )
    print(f"   Top 5 Artikel:")
    for artikel, menge in top_artikel.items():
        anteil = (menge / artikel_gesamt_menge) * 100
        print(f"     {artikel}: {menge:,.0f} Stück ({anteil:.1f}%)")

    # 2. TEILEGRUPPEN-EBENE
    teilegruppen_count = df_prognose_lang["Baumarktartikel"].nunique()

    print(f"\n🔹 EBENE 2: TEILEGRUPPEN")
    print(f"   Anzahl eindeutige Teilegruppen: {teilegruppen_count:,}")

    teilegruppen_agg = (
        df_prognose_lang.groupby("Baumarktartikel")
        .agg({"Prognose_Menge": "sum", "matnr": "nunique"})
        .sort_values("Prognose_Menge", ascending=False)
    )

    print(f"   Top 5 Teilegruppen:")
    for idx, (teilegruppe, row) in enumerate(teilegruppen_agg.head(5).iterrows()):
        anteil = (row["Prognose_Menge"] / artikel_gesamt_menge) * 100
        print(
            f"     {teilegruppe}: {row['Prognose_Menge']:,.0f} Stück ({anteil:.1f}%) - {row['matnr']} Artikel"
        )

    # 3. KUNDEN-EBENE
    kunden_count = df_prognose_lang["Baumarkt"].nunique()

    print(f"\n🔹 EBENE 3: KUNDEN (BAUMÄRKTE)")
    print(f"   Anzahl eindeutige Kunden: {kunden_count:,}")

    kunden_agg = (
        df_prognose_lang.groupby("Baumarkt")
        .agg(
            {"Prognose_Menge": "sum", "matnr": "nunique", "Baumarktartikel": "nunique"}
        )
        .sort_values("Prognose_Menge", ascending=False)
    )

    print(f"   Top 5 Kunden:")
    for idx, (kunde, row) in enumerate(kunden_agg.head(5).iterrows()):
        anteil = (row["Prognose_Menge"] / artikel_gesamt_menge) * 100
        print(
            f"     {kunde}: {row['Prognose_Menge']:,.0f} Stück ({anteil:.1f}%) - {row['matnr']} Artikel, {row['Baumarktartikel']} Teilegruppen"
        )

    # 4. GESAMT-EBENE
    print(f"\n🔹 EBENE 4: GESAMT")
    print(f"   Gesamtmenge über alle Ebenen: {artikel_gesamt_menge:,.0f} Stück")
    print(f"   Anzahl Datensätze: {len(df_prognose_lang):,}")

    # 5. HIERARCHIE-ZUSAMMENFASSUNG
    print(f"\n" + "-" * 80)
    print("HIERARCHIE-ZUSAMMENFASSUNG:")
    print(f"   {artikel_count:,} Artikel")
    print(f"   ↓ gruppiert in {teilegruppen_count:,} Teilegruppen")
    print(f"   ↓ verkauft an {kunden_count:,} Kunden")
    print(f"   ↓ ergeben {artikel_gesamt_menge:,.0f} Stück Gesamtmenge")

    # 6. KONZENTRATION ANALYSIEREN
    print(f"\n" + "-" * 80)
    print("KONZENTRATIONSANALYSE:")

    # Top 20% Artikel
    top_20_prozent_artikel = int(artikel_count * 0.2)
    top_artikel_menge = top_artikel.head(top_20_prozent_artikel).sum()
    artikel_konzentration = (top_artikel_menge / artikel_gesamt_menge) * 100

    print(
        f"   Top 20% der Artikel ({top_20_prozent_artikel} Artikel) machen {artikel_konzentration:.1f}% der Gesamtmenge aus"
    )

    # Top 20% Kunden
    top_20_prozent_kunden = int(kunden_count * 0.2)
    top_kunden_menge = kunden_agg.head(top_20_prozent_kunden)["Prognose_Menge"].sum()
    kunden_konzentration = (top_kunden_menge / artikel_gesamt_menge) * 100

    print(
        f"   Top 20% der Kunden ({top_20_prozent_kunden} Kunden) machen {kunden_konzentration:.1f}% der Gesamtmenge aus"
    )


def main():
    # 1. Teilaufgabe
    data = load_data()
    progonose_lang = structure_data(data)

    # Hierarchieebenen analysieren
    analysiere_hierarchieebenen(progonose_lang)

    df_agg_gesamt = bottom_up_sum(progonose_lang)
    # plot_trends(df_agg_gesamt)

    # 2. Teilaufgabe
    # umsatz_baumarkt = analyse_umsatz_pro_baumarkt(progonose_lang)


if __name__ == "__main__":
    main()
