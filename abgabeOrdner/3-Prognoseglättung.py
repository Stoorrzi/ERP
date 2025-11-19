import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
import seaborn as sns

# --- KONFIGURATION ---
INPUT_FILE_ROHDATEN = "rohdaten.xlsx"
INPUT_FILE_PLAN = "output/agg_baumarktprogramm.xlsx"
OUTPUT_DIR = "./output/final"
OUTPUT_FILE_EXCEL = "Final_Forecast_2026_2027.xlsx"

# Erstelle Ausgabeordner
os.makedirs(OUTPUT_DIR, exist_ok=True)
sns.set_theme(style="whitegrid")

# ---------------------------------------------------------
# 1. HILFSFUNKTIONEN
# ---------------------------------------------------------

def clean_keys(df, col_kunde='Kunde', col_monat='Monat'):
    """Bereinigt Schlüssel für sauberen Merge."""
    # Monat zu Int
    df[col_monat] = pd.to_numeric(df[col_monat], errors='coerce').fillna(0).astype(int)
    # Kunde zu Upper-Case String
    if col_kunde in df.columns:
        df[col_kunde] = df[col_kunde].astype(str).str.strip().str.upper()
    return df

def calculate_factor(row):
    """Berechnet den Faktor pro Zeile."""
    ziel = row['Ziel_Summe']
    ist = row['Bottom_Up_Summe']
    
    if ist == 0: return 0.0    # Keine Basis -> 0
    if ziel == 0: return 1.0   # Kein Plan -> Prognose behalten (statt löschen)
    
    return ziel / ist

# ---------------------------------------------------------
# 2. DATEN LADEN
# ---------------------------------------------------------

def load_data():
    print("Step 1: Lade Daten...")
    
    # A) Prognose
    if not os.path.exists(INPUT_FILE_ROHDATEN):
        print(f"❌ Fehler: {INPUT_FILE_ROHDATEN} fehlt.")
        return pd.DataFrame(), pd.DataFrame()
        
    df_raw = pd.read_excel(INPUT_FILE_ROHDATEN)
    
    # Forecast zusammenbauen (Jahr 1 + 2)
    # Spaltennamen ggf. anpassen falls nötig
    try:
        p1 = df_raw[['matnr', 'Baumarkt', 'Baumarktartikel', 'progmo', 'prog_mg1']].copy()
        p1.columns = ['Artikel', 'Kunde', 'Gruppe', 'Monat', 'Menge']
        
        p2 = df_raw[['matnr', 'Baumarkt', 'Baumarktartikel', 'progmo2', 'prog_mg2']].copy()
        p2.columns = ['Artikel', 'Kunde', 'Gruppe', 'Monat', 'Menge']
        
        df_forecast = pd.concat([p1, p2], ignore_index=True)
        df_forecast = df_forecast.dropna(subset=['Monat', 'Menge'])
        df_forecast = df_forecast[df_forecast['Menge'] > 0]
        df_forecast = clean_keys(df_forecast)
        print(f"   ✅ Prognose geladen: {len(df_forecast)} Zeilen.")
        
    except KeyError as e:
        print(f"❌ Fehler: Spalte fehlt in Rohdaten: {e}")
        return pd.DataFrame(), pd.DataFrame()

    # B) Plan
    if not os.path.exists(INPUT_FILE_PLAN):
        print(f"❌ Fehler: {INPUT_FILE_PLAN} fehlt.")
        return pd.DataFrame(), pd.DataFrame()

    df_plan = pd.read_excel(INPUT_FILE_PLAN)
    df_plan = df_plan.rename(columns={'Baumarkt': 'Kunde', 'Zahl': 'Ziel_Summe'})
    
    # --- KORREKTUR: Einheiten anpassen ---
    print("   ⚠️  Info: Skaliere Vertriebsplan (Tausend -> Stück) mit Faktor 1000.")
    df_plan['Ziel_Summe'] = df_plan['Ziel_Summe'] * 1000
    # -------------------------------------
    
    df_plan = clean_keys(df_plan)
    print(f"   ✅ Plan geladen: {len(df_plan)} Zeilen.")

    return df_forecast, df_plan

# ---------------------------------------------------------
# 3. GLÄTTUNG (RECONCILIATION)
# ---------------------------------------------------------

def run_reconciliation(df_forecast, df_plan):
    print("\nStep 2: Führe Abgleich durch...")
    
    # 1. Aggregation Bottom-Up
    bu_agg = df_forecast.groupby(['Kunde', 'Monat'])['Menge'].sum().reset_index()
    bu_agg = bu_agg.rename(columns={'Menge': 'Bottom_Up_Summe'})
    
    # 2. Merge
    merged = pd.merge(bu_agg, df_plan, on=['Kunde', 'Monat'], how='inner')
    
    if merged.empty:
        print("❌ FEHLER: Keine Matches (Kunde/Monat) gefunden!")
        return pd.DataFrame()

    # 3. Faktor berechnen
    merged['Faktor'] = merged.apply(calculate_factor, axis=1)
    
    # --- STATISTIK CHECK (Das löst Ihre Verwirrung) ---
    avg_factor = merged['Faktor'].mean()
    
    # Gewichteter Faktor: (Summe aller Pläne) / (Summe aller Prognosen)
    total_ziel = merged['Ziel_Summe'].sum()
    total_ist = merged['Bottom_Up_Summe'].sum()
    weighted_factor = total_ziel / total_ist if total_ist > 0 else 0
    
    print(f"   📊 Statistik:")
    print(f"      - Arithmetischer Schnitt der Faktoren: {avg_factor:.2f} (Anfällig für Ausreißer)")
    print(f"      - Gewichteter Faktor (Real):           {weighted_factor:.2f} (Erwartet: ~1.76)")
    
    if weighted_factor < 0.1 or weighted_factor > 10:
        print("      ⚠️ WARNUNG: Auch der gewichtete Faktor ist extrem! Bitte Daten prüfen.")
    else:
        print("      ✅ Plausibilität OK.")

    # 4. Anwenden
    df_final = pd.merge(df_forecast, merged[['Kunde', 'Monat', 'Faktor']], on=['Kunde', 'Monat'], how='left')
    
    # Fallback für fehlende Pläne
    missing_count = df_final['Faktor'].isna().sum()
    if missing_count > 0:
        print(f"   ℹ️  Info: {missing_count} Zeilen ohne Plan behalten (Faktor 1.0).")
        
    df_final['Faktor'] = df_final['Faktor'].fillna(1.0)
    df_final['Menge_Geglaettet'] = (df_final['Menge'] * df_final['Faktor']).round(0).astype(int)
    
    return df_final

# ---------------------------------------------------------
# 4. MAIN
# ---------------------------------------------------------

def main():
    # Laden
    df_forecast, df_plan = load_data()
    if df_forecast.empty: return

    # Rechnen
    df_final = run_reconciliation(df_forecast, df_plan)
    if df_final.empty: return

    # Speichern
    out_path = os.path.join(OUTPUT_DIR, OUTPUT_FILE_EXCEL)
    cols = ['Artikel', 'Kunde', 'Gruppe', 'Monat', 'Menge', 'Faktor', 'Menge_Geglaettet']
    df_final[cols].to_excel(out_path, index=False)
    print(f"\n✅ FERTIG! Datei gespeichert: {out_path}")
    
    # Kleiner Plot zur Bestätigung
    try:
        plot_data = df_final.groupby('Monat')[['Menge', 'Menge_Geglaettet']].sum().reset_index()
        plot_data['Monat'] = plot_data['Monat'].astype(str)
        plt.figure(figsize=(10, 5))
        plt.plot(plot_data['Monat'], plot_data['Menge'], label='Original', linestyle='--')
        plt.plot(plot_data['Monat'], plot_data['Menge_Geglaettet'], label='Geglättet (Ziel)')
        plt.title("Gesamtvolumen Vorher vs. Nachher")
        plt.legend()
        plt.savefig(os.path.join(OUTPUT_DIR, "Final_Check.png"))
        print("   Plot gespeichert.")
    except:
        pass

if __name__ == "__main__":
    main()