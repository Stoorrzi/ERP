import pandas as pd
import os

INPUT_FILE_ROHDATEN = "rohdaten.xlsx"
INPUT_FILE_PLAN = "output/agg_baumarktprogramm.xlsx"

def debug_check():
    print("=== DEBUGGING DATEN-INTEGRITÄT ===")
    
    # 1. CHECK ROHDATEN (PROGNOSE)
    print(f"\n1. Prüfe Rohdaten: {INPUT_FILE_ROHDATEN}")
    try:
        df = pd.read_excel(INPUT_FILE_ROHDATEN)
        print(f"   Zeilen gesamt: {len(df)}")
        
        # Prüfe Spalte 'prog_mg1' (Menge)
        if 'prog_mg1' in df.columns:
            print(f"   Datentyp 'prog_mg1': {df['prog_mg1'].dtype}")
            print(f"   Summe 'prog_mg1' (Raw): {df['prog_mg1'].sum()}")
            print(f"   Beispielwerte (Head): {df['prog_mg1'].head(5).tolist()}")
            
            # Falls es Text ist, versuchen wir zu konvertieren und schauen was passiert
            if df['prog_mg1'].dtype == 'object':
                print("   ⚠️ ACHTUNG: Spalte ist OBJECT (Text)! Konvertierungstest:")
                try:
                    converted = pd.to_numeric(df['prog_mg1'], errors='coerce')
                    print(f"   Summe nach 'to_numeric': {converted.sum()}")
                    print(f"   Anzahl NaN nach Konvertierung: {converted.isna().sum()}")
                except:
                    print("   Konvertierung fehlgeschlagen.")
        else:
            print("   ❌ Spalte 'prog_mg1' fehlt!")

    except Exception as e:
        print(f"   ❌ Fehler beim Lesen: {e}")

    # 2. CHECK VERTRIEBSPLAN
    print(f"\n2. Prüfe Vertriebsplan: {INPUT_FILE_PLAN}")
    try:
        df_plan = pd.read_excel(INPUT_FILE_PLAN)
        print(f"   Zeilen gesamt: {len(df_plan)}")
        print(f"   Spalten: {df_plan.columns.tolist()}")
        
        # Check Summe VOR Multiplikation
        col_zahl = 'Zahl' if 'Zahl' in df_plan.columns else df_plan.columns[2] # Fallback
        sum_raw = df_plan[col_zahl].sum()
        print(f"   Summe Plan (Raw): {sum_raw}")
        
        # Check Summe NACH Multiplikation
        sum_times_1000 = sum_raw * 1000
        print(f"   Summe Plan (* 1000): {sum_times_1000}")
        
    except Exception as e:
        print(f"   ❌ Fehler beim Lesen: {e}")

    # 3. SCHÄTZUNG FAKTOR
    print("\n3. Erwarteter Faktor (Schätzung)")
    try:
        # Wir nehmen die berechneten Summen von oben
        forecast_sum = df['prog_mg1'].sum() # Achtung: Nimmt nur prog_mg1, aber reicht für Größenordnung
        plan_sum = sum_times_1000
        
        print(f"   Prognose-Basis (ca.): {forecast_sum:,.0f}")
        print(f"   Plan-Ziel (ca.):      {plan_sum:,.0f}")
        
        if forecast_sum > 0:
            factor = plan_sum / forecast_sum
            print(f"   -> Erwarteter Faktor wäre ca.: {factor:.6f}")
        else:
            print("   -> Faktor nicht berechenbar (Prognose ist 0 oder Fehler).")
            
    except:
        print("   Konnte Schätzung nicht durchführen.")

if __name__ == "__main__":
    debug_check()