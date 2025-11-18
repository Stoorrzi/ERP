import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import numpy as np

sns.set_theme(style="whitegrid")

def load_data(filepath="rohdaten.xlsx"):
    """
    Lädt die Excel-Rohdaten.
    """
    try:
        df_raw = pd.read_excel(filepath)
        print(f"Datei '{filepath}' erfolgreich geladen: {df_raw.shape[0]} Zeilen, {df_raw.shape[1]} Spalten.")
        
        df_raw['bedmo_date'] = pd.to_datetime(df_raw['bedmo'], format='%Y%m')
        
    except FileNotFoundError:
        print(f"FEHLER: Datei nicht gefunden: '{filepath}'")
        return None
    except Exception as e:
        print(f"Fehler beim Laden oder Umwandeln von 'bedmo': {e}")
        return None
        
    return df_raw

# --- Schritt 2: Effizient Aggregieren ---

def aggregate_data(data):
    """
    Aggregiert die Rohdaten auf die beiden geforderten Ebenen:
    1. Pro Baumarkt & Monat
    2. Pro Baumarktartikel & Monat
    """
    
    # Aggregations-Logik definieren
    agg_definition = {
        "wavor_bstlmg": "sum",
        "progmo": "first",
        "prog_mg1": "sum",
        "progmo2": "first",
        "prog_mg2": "sum",
    }
    
    # --- 1. Aggregation pro Baumarkt & Monat ---
    print("Aggregiere Daten pro Baumarkt und Monat...")
    df_baumarkt_agg = (
        data.groupby(["Baumarkt", "bedmo_date"])
        .agg(agg_definition)
        .reset_index()
    )
    df_baumarkt_agg = df_baumarkt_agg.sort_values(by=["Baumarkt", "bedmo_date"])
    
    # --- 2. Aggregation pro Baumarktartikel & Monat ---
    print("Aggregiere Daten pro Baumarktartikel und Monat...")
    df_artikelgruppe_agg = (
        data.groupby(["Baumarktartikel", "bedmo_date"])
        .agg(agg_definition)
        .reset_index()
    )
    df_artikelgruppe_agg = df_artikelgruppe_agg.sort_values(by=["Baumarktartikel", "bedmo_date"])

    print("Aggregation abgeschlossen.")
    return df_baumarkt_agg, df_artikelgruppe_agg


# --- Schritt 3: Störgrößen erkennen UND Glätten  ---

def detect_and_smooth(df_group, metric_col='wavor_bstlmg', window=3):
    """
    VERBESSERTE Version: Findet Ausreißer basierend auf prozentualer Abweichung.
    """
    df_group = df_group.copy()
    
    df_group['moving_avg'] = df_group[metric_col].rolling(window=window, center=True, min_periods=1).mean()
    df_group['pct_diff'] = (df_group[metric_col] - df_group['moving_avg']) / df_group['moving_avg']
    df_group['pct_diff'] = df_group['pct_diff'].replace([np.inf, -np.inf], 0).fillna(0)

    is_dropout = (df_group[metric_col] <= 0.1) & (df_group['moving_avg'] > 100) 
    is_stat_low = (df_group['pct_diff'] < -0.70) 

    df_group['is_outlier'] = is_dropout | is_stat_low
    
    df_group[f'{metric_col}_geglättet'] = df_group[metric_col]
    df_group.loc[df_group['is_outlier'], f'{metric_col}_geglättet'] = df_group['moving_avg']
    
    return df_group

# --- Schritt 4: PLOT-FUNKTIONEN FÜR DIE PRÄSENTATION ---

def plot_task_trends(df_baumarkt_agg, output_dir):
    """
    AUFGABE: Analyse von Trends (Gesamtmarkt)
    Erstellt einen Plot, der den Gesamt-Trend aller Verkäufe zeigt.
    """
    print("Erstelle Plot: 1_Gesamtmarkt_Trend.png")
    
    # Alle Baumärkte pro Monat summieren, um den Gesamtmarkt zu erhalten
    df_trend = df_baumarkt_agg.groupby('bedmo_date').agg(Gesamtvolumen=('wavor_bstlmg', 'sum')).reset_index()
    
    plt.figure(figsize=(12, 6))
    sns.lineplot(data=df_trend, x='bedmo_date', y='Gesamtvolumen', linewidth=2.5)
    
    # Einen geglätteten Trend (gleitender Durchschnitt) hinzufügen
    df_trend['Trend_geglaettet'] = df_trend['Gesamtvolumen'].rolling(window=6, center=True, min_periods=1).mean()
    sns.lineplot(data=df_trend, x='bedmo_date', y='Trend_geglaettet', color='red', linestyle='--', label='6-Monats-Trend')
    
    
    
    plt.title('Analyse: Gesamtmarkt-Trend (Alle Baumärkte)', fontsize=16)
    plt.ylabel('Summiertes Bestellvolumen (wavor_bstlmg)')
    plt.xlabel('Monat')
    plt.legend()
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "1_Gesamtmarkt_Trend.png"))
    plt.close()

def plot_task_seasonality(df_artikelgruppe_agg, output_dir):
    """
    AUFGABE: Analyse von Saisonalität (auf Teilegruppen-Ebene)
    
    NEU: Zeigt die 5 Gruppen mit der HÖCHSTEN SCHWANKUNG (Volatilität),
    nicht das höchste Gesamtvolumen.
    """
    print("Erstelle Plot: 2_Saisonalitaet_Staerste_Schwankung.png")
    
    # Berechne die Volatilität (Schwankung) für jede Gruppe
    # Wir nutzen den Variationskoeffizienten (Std / Mean)
    df_volatility = df_artikelgruppe_agg.groupby('Baumarktartikel')['wavor_bstlmg'].agg(
        std_dev='std',
        mean_val='mean'
    ).reset_index()
    
    # CV = Standardabweichung / Mittelwert. fillna(0) falls mean = 0
    df_volatility['cv'] = (df_volatility['std_dev'] / df_volatility['mean_val']).fillna(0)
    
    df_volatility = df_volatility[df_volatility['mean_val'] > 100] # Schwellenwert ggf. anpassen!

    # Finde die Top 5 Gruppen mit der höchsten Schwankung (CV)
    top_volatile_groups = df_volatility.nlargest(5, 'cv')['Baumarktartikel']

    df_top_groups = df_artikelgruppe_agg[df_artikelgruppe_agg['Baumarktartikel'].isin(top_volatile_groups)]

    plt.figure(figsize=(12, 7))
    sns.lineplot(
        data=df_top_groups,
        x='bedmo_date',
        y='wavor_bstlmg',
        hue='Baumarktartikel', 
        style='Baumarktartikel', 
        linewidth=2,
        markers=True
    )
    
    
    
    plt.title('Analyse: Saisonalität (Top 5 Gruppen mit höchster Schwankung)', fontsize=16)
    plt.ylabel('Bestellvolumen (wavor_bstlmg)')
    plt.xlabel('Monat')
    plt.legend(title='Artikelgruppe', bbox_to_anchor=(1.02, 1), loc='upper left')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "2_Saisonalitaet_Staerste_Schwankung.png"))
    plt.close()
    
def plot_task_outliers(df_baumarkt_smoothed, output_dir):
    """
    AUFGABE: Analyse von Ausreißern (auf Kunden-Ebene)
    Zeigt ein klares Beispiel für eine Störgröße und deren Glättung.
    """
    print("Erstelle Plot: 3_Ausreisser_Glaettung.png")
    
    # Finde den Baumarkt mit den meisten Ausreißern als gutes Beispiel
    outlier_counts = df_baumarkt_smoothed.groupby('Baumarkt')['is_outlier'].sum().nlargest(1)
    
    if outlier_counts.empty:
        print("Keine Ausreißer gefunden. Überspringe Plot.")
        return
        
    example_group_name = outlier_counts.index[0]
    df_group = df_baumarkt_smoothed[df_baumarkt_smoothed['Baumarkt'] == example_group_name]
    
    
    df_plot = df_group.copy()
    outliers = df_plot[df_plot['is_outlier']]

    plt.figure(figsize=(12, 6))
    
    sns.lineplot(data=df_plot, x='bedmo_date', y='wavor_bstlmg_geglättet', 
                 label='Geglättete Daten (Prognosebasis)', color='blue', linewidth=2.5, zorder=3)
                 
    sns.scatterplot(data=df_plot, x='bedmo_date', y='wavor_bstlmg', 
                    label='Original IST-Daten', color='gray', alpha=0.6, zorder=2)
                    
    if not outliers.empty:
        sns.scatterplot(data=outliers, x='bedmo_date', y='wavor_bstlmg', 
                        label='Erkannte Störgröße (z.B. Umbau)', color='red', s=150, zorder=5)

    
    
    plt.title(f'Analyse: Störgrößen & Glättung (Beispiel: Baumarkt {example_group_name})', fontsize=16)
    plt.ylabel('Bestellvolumen (wavor_bstlmg)')
    plt.xlabel('Monat')
    plt.legend()
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "3_Ausreisser_Glaettung.png"))
    plt.close()


# --- NEUE FUNKTION: Plot 4  ---

def plot_task_trends_per_baumarkt(df_baumarkt_agg, output_dir, top_n=10):
    """
    AUFGABE: Analyse von Trends pro Baumarkt (Top-Kunden)
    Erstellt einen Plot, der die Trends der Top N Baumärkte vergleicht.
    """
    print(f"Erstelle Plot: 4_Top_{top_n}_Baumarkt_Trends.png")
    
    top_baumaerkte = df_baumarkt_agg.groupby('Baumarkt')['wavor_bstlmg'].sum().nlargest(top_n).index
    df_top_baumaerkte = df_baumarkt_agg[df_baumarkt_agg['Baumarkt'].isin(top_baumaerkte)]
    
    plt.figure(figsize=(12, 7))
    sns.lineplot(
        data=df_top_baumaerkte,
        x='bedmo_date',
        y='wavor_bstlmg',
        hue='Baumarkt', 
        style='Baumarkt', 
        linewidth=2,
        markers=True
    )
    
    
    
    plt.title(f'Analyse: Kunden-Trends (Top {top_n} Baumärkte)', fontsize=16)
    plt.ylabel('Summiertes Bestellvolumen (wavor_bstlmg)')
    plt.xlabel('Monat')
    plt.legend(title='Baumarkt', bbox_to_anchor=(1.02, 1), loc='upper left')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, f"4_Top_{top_n}_Baumarkt_Trends.png"))
    plt.close()



def main():
    # Output-Verzeichnisse erstellen
    os.makedirs("./output", exist_ok=True)
    plot_dir = "./output/plots/1"
    os.makedirs(plot_dir, exist_ok=True)
    
    # 1. Laden
    data = load_data()
    if data is None:
        print("Daten konnten nicht geladen werden. Skript wird beendet.")
        return

    # 2. Aggregieren (Ebenen 3 und 4)
    df_baumarkt_agg, df_artikelgruppe_agg = aggregate_data(data)
    
    # 3. Glättungs-Daten berechnen (Notwendig für den Ausreißer-Plot)
    print("\nStarte Analyse & Glättung für 'Baumarkt'...")
    df_baumarkt_smoothed = (
        df_baumarkt_agg.groupby('Baumarkt')
        .apply(detect_and_smooth)
        .reset_index(drop=True)
    )
    
    # 4. PRÄSENTATIONS-PLOTS ERSTELLEN
    
    # Plot 1: Gesamt-Trend
    plot_task_trends(df_baumarkt_agg, plot_dir)
    
    # Plot 2: Saisonalität
    plot_task_seasonality(df_artikelgruppe_agg, plot_dir)
    
    # Plot 3: Ausreißer / Störgrößen
    plot_task_outliers(df_baumarkt_smoothed, plot_dir)
    
    # Plot 4: Trends pro Baumarkt
    plot_task_trends_per_baumarkt(df_baumarkt_agg, plot_dir, top_n=10)
    
    print(f"\nAlle Analyse-Plots wurden im Ordner '{plot_dir}' gespeichert.")

if __name__ == "__main__":
    main()


