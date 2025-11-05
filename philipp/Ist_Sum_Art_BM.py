import pandas as pd

# --- 1. DATEN LADEN ---
# Ändern Sie 'IHR_DATEINAME.xlsx' in den Pfad zu Ihrer Excel-Datei
# Wenn Ihre Daten nicht im ersten Tabellenblatt sind, fügen Sie sheet_name='NameDesBlatts' hinzu
try:
    # Pandas' read_excel verwenden, da Sie .xlsx-Dateien haben
    df = pd.read_excel('rohdaten.xlsx')
    
    print("Excel-Datei erfolgreich geladen.")
    
    # --- 2. ANNAHMEN DEFINIEREN ---
    # Diese Spaltennamen basieren auf der .csv-Datei, die Sie mir gesendet haben.
    # Passen Sie diese an, falls die Spalten in Ihrer Excel-Datei anders heißen.
    
    time_col = 'bedmo'       # Spalte für den Monat (z.B. 202511)
    group_col1 = 'Baumarkt'  # Spalte für den Baumarkt
    group_col2 = 'matnr'     # Spalte für die Artikelnummer
    value_col = 'bedmo_mg'   # Spalte für die verkaufte Menge
    
    # --- 3. AGGREGATION (ANALYSE) ---
    # Überprüfen, ob die Spalten vorhanden sind
    required_cols = [time_col, group_col1, group_col2, value_col]
    if not all(col in df.columns for col in required_cols):
        print(f"Fehler: Eine der Spalten {required_cols} wurde nicht gefunden.")
        print(f"Verfügbare Spalten sind: {df.columns.tolist()}")
    else:
        # Gruppieren nach Monat, Baumarkt und Artikelnummer
        # Anschließend die Mengen ('value_col') für jede Gruppe summieren
        aggregation = df.groupby([time_col, group_col1, group_col2])[value_col].sum().reset_index()
        
        # Die Wertespalte umbenennen für bessere Lesbarkeit
        aggregation = aggregation.rename(columns={value_col: 'Gesamtmenge'})
        
        # Sortieren nach Monat (absteigend, neueste zuerst)
        aggregation_sorted = aggregation.sort_values(by=time_col, ascending=False)
        
        # --- 4. ERGEBNISSE ANZEIGEN ---
        print("\n--- Aggregierte Verkaufsdaten ---")
        print(aggregation_sorted)
        
        # --- 5. ERGEBNISSE SPEICHERN ---
        output_filename = 'aggregierte_verkaeufe.csv'
        aggregation_sorted.to_csv(output_filename, index=False, sep=';', decimal=',')
        print(f"\nAnalyse erfolgreich. Ergebnisse in '{output_filename}' gespeichert.")

except FileNotFoundError:
    print(f"Fehler: Die Datei 'IHR_DATEINAME.xlsx' wurde nicht gefunden.")
except Exception as e:
    print(f"Ein Fehler ist aufgetreten: {e}")