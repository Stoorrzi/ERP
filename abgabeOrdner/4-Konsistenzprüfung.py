import pandas as pd
import numpy as np
import os

# --- KONFIGURATION ---
FILE_FORECAST_FINAL = "./output/final/Final_Forecast_2026_2027.xlsx"
FILE_PLAN = "output/agg_baumarktprogramm.xlsx"

def clean_keys(df, col_kunde='Kunde', col_monat='Monat'):
    """Stellt sicher, dass wir Text und Zahlen vergleichen können."""
    df[col_monat] = pd.to_numeric(df[col_monat], errors='coerce').fillna(0).astype(int)
    if col_kunde in df.columns:
        df[col_kunde] = df[col_kunde].astype(str).str.strip().str.upper()
    return df

def main():
    print("=== TEILAUFGABE 4: KONSISTENZPRÜFUNG ===")
    
    # 1. Daten laden
    if not os.path.exists(FILE_FORECAST_FINAL):
        print("❌ FEHLER: Finaler Forecast fehlt. Bitte erst Schritt 3 ausführen.")
        return

    print("1. Lade geglättete Artikeldaten...")
    df_final = pd.read_excel(FILE_FORECAST_FINAL)
    
    print("2. Lade ursprünglichen Vertriebsplan...")
    df_plan = pd.read_excel(FILE_PLAN)
    df_plan = df_plan.rename(columns={'Baumarkt': 'Kunde', 'Zahl': 'Ziel_Summe'})
    
    # WICHTIG: Gleiche Skalierung wie in Schritt 3 anwenden!
    df_plan['Ziel_Summe'] = df_plan['Ziel_Summe'] * 1000
    
    # Bereinigen
    df_final = clean_keys(df_final)
    df_plan = clean_keys(df_plan)

    # 2. Aggregation: Wir summieren die neuen Artikelwerte wieder hoch
    print("\n3. Prüfe Summen...")
    agg_check = df_final.groupby(['Kunde', 'Monat'])['Menge_Geglaettet'].sum().reset_index()
    agg_check = agg_check.rename(columns={'Menge_Geglaettet': 'Ist_Summe_Neu'})

    # 3. Vergleich mit dem Plan
    merged = pd.merge(agg_check, df_plan, on=['Kunde', 'Monat'], how='inner')
    
    # Differenz berechnen
    merged['Differenz'] = merged['Ist_Summe_Neu'] - merged['Ziel_Summe']
    merged['Differenz_Abs'] = merged['Differenz'].abs()
    
    # 4. Ergebnis-Analyse
    total_diff = merged['Differenz_Abs'].sum()
    max_diff = merged['Differenz_Abs'].max()
    
    print("-" * 60)
    print(f"Anzahl geprüfter Kunden/Monats-Kombinationen: {len(merged)}")
    print(f"Gesamte Abweichung (Summe über alle):         {total_diff:,.0f} Stück")
    print(f"Maximale Abweichung in einem Monat:           {max_diff:,.0f} Stück")
    print("-" * 60)
    
    # Bewertung
    # Wir tolerieren kleine Rundungsfehler (z.B. +/- 500 Stück bei Millionen-Umsatz ist okay)
    # Da wir auf Ganze Zahlen runden, kann pro Artikel 0.5 Abweichung entstehen.
    # Bei 1000 Artikeln wären das bis zu 500 Stück Abweichung pro Monat.
    
    if max_diff < 1000:
        print("✅ ERGEBNIS: KONSISTENT")
        print("   Die Abweichungen sind minimal und durch Rundung auf ganze Stückzahlen erklärbar.")
    else:
        print("⚠️ ERGEBNIS: ABWEICHUNGEN ERKANNT")
        print("   Schauen Sie sich diese Fälle genauer an:")
        print(merged[merged['Differenz_Abs'] > 1000].head())

    # Optional: Export der Prüfung
    merged.to_excel("./output/final/Konsistenz_Report.xlsx", index=False)
    print(f"\n   Detaillierter Report gespeichert: ./output/final/Konsistenz_Report.xlsx")

if __name__ == "__main__":
    main()