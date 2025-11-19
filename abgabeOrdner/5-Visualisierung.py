import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import os

# --- KONFIGURATION ---
INPUT_FILE = "./output/final/Final_Forecast_2026_2027.xlsx"
OUTPUT_DIR_PLOTS = "./output/final/plots"

# Setup
os.makedirs(OUTPUT_DIR_PLOTS, exist_ok=True)
sns.set_theme(style="whitegrid") 

def load_data():
    if not os.path.exists(INPUT_FILE):
        print(f"❌ FEHLER: Datei '{INPUT_FILE}' fehlt.")
        return pd.DataFrame()
    
    print("1. Lade Daten für Visualisierung...")
    df = pd.read_excel(INPUT_FILE)
    # Monat als String für diskrete Achse
    df['Monat_Str'] = df['Monat'].astype(str)
    return df

# ---------------------------------------------------------
# PLOT 1: MANAGEMENT SUMMARY (Legende UNTEN)
# ---------------------------------------------------------
def plot_management_summary(df):
    print("2. Erstelle Management-Summary...")
    
    agg = df.groupby('Monat_Str')[['Menge', 'Menge_Geglaettet']].sum().reset_index()
    agg_melt = agg.melt(id_vars='Monat_Str', value_vars=['Menge', 'Menge_Geglaettet'], 
                        var_name='Typ', value_name='Stückzahl')
    
    agg_melt['Typ'] = agg_melt['Typ'].replace({
        'Menge': 'Ursprüngliche Prognose (Bottom-Up)', 
        'Menge_Geglaettet': 'Angepasster Plan (Final)'
    })

    plt.figure(figsize=(14, 8)) # Etwas höher für die Legende unten
    ax = sns.lineplot(data=agg_melt, x='Monat_Str', y='Stückzahl', hue='Typ', style='Typ', 
                      markers=True, dashes=False, linewidth=3)
    
    # Farben setzen
    palette = {'Ursprüngliche Prognose (Bottom-Up)': 'grey', 'Angepasster Plan (Final)': '#2ecc71'}
    for line in ax.lines:
        if line.get_label() in palette:
            line.set_color(palette[line.get_label()])

    # Formatierung Y-Achse
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: format(int(x), ',').replace(',', '.')))
    
    # --- FIX: LEGENDE NACH UNTEN VERSCHIEBEN ---
    # bbox_to_anchor=(x, y): (0.5, -0.2) bedeutet "Mittig, unterhalb der Achse"
    sns.move_legend(ax, "upper center", bbox_to_anchor=(0.5, -0.15), ncol=2, title=None, frameon=False)
    
    plt.title("Gesamtvolumen: Anpassung an den Vertriebsplan", pad=20, fontsize=16, fontweight='bold')
    plt.xlabel("")
    plt.ylabel("Absatzmenge (Stück)")
    plt.xticks(rotation=45)
    
    # Wichtig: Layout anpassen, damit Legende nicht abgeschnitten wird
    plt.tight_layout() 
    
    save_path = os.path.join(OUTPUT_DIR_PLOTS, "1_Management_Summary.png")
    plt.savefig(save_path, dpi=150, bbox_inches='tight') # bbox_inches='tight' schützt die Legende zusätzlich
    print(f"   ✅ Gespeichert: {save_path}")

# ---------------------------------------------------------
# PLOT 2: HEATMAP (Dynamische Größe gegen Quetschen)
# ---------------------------------------------------------
def plot_correction_heatmap(df):
    print("3. Erstelle Heatmap...")
    
    pivot_faktor = df.pivot_table(index='Kunde', columns='Monat_Str', values='Faktor', aggfunc='mean')
    
    # Dynamische Größe berechnen (verhindert Quetschen)
    n_customers = len(pivot_faktor.index)
    n_months = len(pivot_faktor.columns)
    
    fig_height = max(8, n_customers * 0.6)
    fig_width = max(12, n_months * 0.5)
    
    plt.figure(figsize=(fig_width, fig_height))
    
    sns.heatmap(pivot_faktor, cmap="vlag_r", center=1.0, annot=True, fmt=".2f", 
                linewidths=.5, square=True,
                cbar_kws={'label': 'Korrekturfaktor (1.0 = Neutral)', 'shrink': 0.8},
                annot_kws={"size": 9})
    
    plt.title("Intensität der Eingriffe pro Kunde (Rot = Kürzung, Blau = Erhöhung)", pad=20, fontsize=16, fontweight='bold')
    plt.xlabel("")
    plt.ylabel("")
    plt.xticks(rotation=45, ha='right')
    plt.yticks(rotation=0)
    plt.tight_layout()
    
    save_path = os.path.join(OUTPUT_DIR_PLOTS, "2_Korrektur_Heatmap.png")
    plt.savefig(save_path, dpi=150, bbox_inches='tight')
    print(f"   ✅ Gespeichert: {save_path}")

# ---------------------------------------------------------
# PLOT 3: DETAIL-STRUKTUR (Legende UNTEN)
# ---------------------------------------------------------
def plot_detail_structure(df):
    print("4. Erstelle Detail-Plot...")
    
    top_kunde = df.groupby('Kunde')['Menge'].sum().idxmax()
    col_gruppe = 'Gruppe' if 'Gruppe' in df.columns else df.columns[2]
    
    try:
        beispiel_gruppe = df[df['Kunde'] == top_kunde][col_gruppe].value_counts().index[0]
    except IndexError:
        print("   ⚠️ Keine Daten für Detail-Plot gefunden.")
        return

    data_subset = df[(df['Kunde'] == top_kunde) & (df[col_gruppe] == beispiel_gruppe)].copy()
    agg_subset = data_subset.groupby('Monat_Str')[['Menge', 'Menge_Geglaettet']].sum().reset_index()
    
    plt.figure(figsize=(14, 8)) # Etwas höher
    ax1 = plt.gca()
    ax2 = ax1.twinx()
    
    l1 = ax1.plot(agg_subset['Monat_Str'], agg_subset['Menge'], color='grey', linestyle='--', label='Original (Links)', linewidth=2)
    l2 = ax2.plot(agg_subset['Monat_Str'], agg_subset['Menge_Geglaettet'], color='blue', label='Geglättet (Rechts)', linewidth=3)
    
    ax1.set_ylabel('Original Menge', color='grey', fontsize=12)
    ax2.set_ylabel('Geglättete Menge', color='blue', fontsize=12)
    
    # Legende Kombinieren und nach UNTEN schieben
    lns = l1 + l2
    labs = [l.get_label() for l in lns]
    # bbox_to_anchor=(0.5, -0.15) -> Unter das Diagramm
    ax1.legend(lns, labs, loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2, frameon=False)
    
    plt.title(f"Struktur-Check: {top_kunde} / {beispiel_gruppe}", pad=20, fontsize=16)
    ax1.set_xticklabels(agg_subset['Monat_Str'], rotation=45)
    plt.tight_layout()
    
    save_path = os.path.join(OUTPUT_DIR_PLOTS, "3_Detail_Struktur.png")
    plt.savefig(save_path, dpi=150, bbox_inches='tight')
    print(f"   ✅ Gespeichert: {save_path}")

def main():
    print("=== TEILAUFGABE 5: VISUALISIERUNG (FIXED LAYOUT) ===")
    df = load_data()
    if df.empty: return
    
    plot_management_summary(df)
    plot_correction_heatmap(df)
    plot_detail_structure(df)
    
    print("\n✅ Fertig! Plots befinden sich in ./output/final/plots/")

if __name__ == "__main__":
    main()