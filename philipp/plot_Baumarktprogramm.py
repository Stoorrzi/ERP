import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os


def load_baumarktprogramm():
    """
    Lädt die Baumarktprogramm.xlsx Datei
    """
    try:
        df = pd.read_excel("BAUMARKTPROGRAMM.xlsx")
        print(f"Baumarktprogramm geladen: {df.shape[0]} Zeilen, {df.shape[1]} Spalten")
        print(f"Spalten: {list(df.columns)}")
        return df
    except FileNotFoundError:
        print("❌ Datei 'BAUMARKTPROGRAMM.xlsx' nicht gefunden!")
        return None
    except Exception as e:
        print(f"❌ Fehler beim Laden: {e}")
        return None


def extract_data_for_plotting(df):
    """
    Extrahiert und strukturiert die Daten für das Plotting
    Spaltenaufteilung:
    - 2025: E-P (Spalten 4-15)
    - 2026: R-AC (Spalten 17-28)
    - 2027: AE-AP (Spalten 30-41)
    - 2028: AR-BC (Spalten 43-54)
    """
    # Annahme: Erste Spalte enthält Baumarkt-Namen
    baumärkte = df.iloc[:, 0].dropna().unique()

    # Monatsnamen für Labels
    monate = [
        "Jan",
        "Feb",
        "Mär",
        "Apr",
        "Mai",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Okt",
        "Nov",
        "Dez",
    ]

    # Spalten-Mapping (Excel-Spalten zu Index)
    # A=0, B=1, C=2, D=3, E=4, F=5, ..., P=15, Q=16, R=17, ..., AC=28, AD=29, AE=30, ..., AP=41, AQ=42, AR=43, ..., BC=54
    spalten_mapping = {
        "2025": list(range(4, 16)),  # E bis P (Spalten 4-15)
        "2026": list(range(17, 29)),  # R bis AC (Spalten 17-28)
        "2027": list(range(30, 42)),  # AE bis AP (Spalten 30-41)
        "2028": list(range(43, 55)),  # AR bis BC (Spalten 43-54)
    }

    print("Spalten-Mapping:")
    for jahr, spalten in spalten_mapping.items():
        print(
            f"  {jahr}: Spalten {spalten[0]}-{spalten[-1]} ({chr(65+spalten[0])}-{chr(65+spalten[-1]) if spalten[-1] < 26 else 'A' + chr(65+spalten[-1]-26)})"
        )

    # Dictionary für strukturierte Daten
    plot_data = {}

    for baumarkt in baumärkte:
        if pd.isna(baumarkt) or baumarkt == "":
            continue

        baumarkt_str = str(baumarkt)
        print(f"Verarbeite Baumarkt: {baumarkt_str}")

        # Filter für aktuellen Baumarkt
        baumarkt_rows = df[df.iloc[:, 0] == baumarkt]

        if len(baumarkt_rows) == 0:
            continue

        # Daten für 2025, 2026, 2027, 2028 extrahieren
        jahre_data = {}

        for jahr, spalten_indices in spalten_mapping.items():
            jahr_daten = []

            for spalte_idx in spalten_indices:
                try:
                    if spalte_idx < len(df.columns):
                        wert = (
                            baumarkt_rows.iloc[0, spalte_idx]
                            if len(baumarkt_rows) > 0
                            else 0
                        )
                        jahr_daten.append(float(wert) if pd.notna(wert) else 0)
                    else:
                        jahr_daten.append(0)
                except:
                    jahr_daten.append(0)

            # Sicherstellen, dass genau 12 Monate vorhanden sind
            while len(jahr_daten) < 12:
                jahr_daten.append(0)
            jahre_data[jahr] = jahr_daten[:12]

        plot_data[baumarkt_str] = jahre_data

        # Debug: Zeige erste paar Werte
        print(f"  Beispieldaten für {baumarkt_str}:")
        for jahr in ["2025", "2026", "2027", "2028"]:
            summe = sum(jahre_data[jahr])
            print(
                f"    {jahr}: Summe = {summe:.2f}, erste 3 Monate = {jahre_data[jahr][:3]}"
            )

    return plot_data, monate


def plot_baumarkt_vergleich(plot_data, monate):
    """
    Erstellt Liniendiagramme für jeden Baumarkt mit durchgehender Zeitlinie
    """
    if not plot_data:
        print("❌ Keine Daten zum Plotten verfügbar")
        return

    # Anzahl Baumärkte
    n_baumärkte = len(plot_data)

    # Layout berechnen
    cols = 3  # 3 Spalten
    rows = (n_baumärkte + cols - 1) // cols

    # Figure erstellen
    fig, axes = plt.subplots(rows, cols, figsize=(18, 6 * rows))
    fig.suptitle(
        "Baumarktprogramm - Zeitverlauf 2025-2028\n(Liniendiagramme)",
        fontsize=16,
        fontweight="bold",
    )

    # Wenn nur eine Zeile oder Spalte, axes zu Liste machen
    if rows == 1 and cols == 1:
        axes = [axes]
    elif rows == 1:
        axes = [axes]
    elif cols == 1:
        axes = [[ax] for ax in axes]

    # Farben für die Jahre
    farben = {
        "2025": "#1f77b4",  # Blau
        "2026": "#ff7f0e",  # Orange
        "2027": "#2ca02c",  # Grün
        "2028": "#d62728",  # Rot
    }

    # Durchgehende X-Achse: 48 Monate (4 Jahre × 12 Monate)
    x_gesamt = list(range(48))  # 0 bis 47

    # Labels für X-Achse (alle 6 Monate)
    x_labels = []
    x_ticks = []
    for jahr_idx, jahr in enumerate(["2025", "2026", "2027", "2028"]):
        for monat_idx, monat in enumerate(monate):
            x_pos = jahr_idx * 12 + monat_idx
            if monat_idx % 6 == 0:  # Alle 6 Monate ein Label
                x_labels.append(f"{monat} {jahr}")
                x_ticks.append(x_pos)

    baumarkt_names = list(plot_data.keys())

    for i, baumarkt in enumerate(baumarkt_names):
        row = i // cols
        col = i % cols

        # Aktueller Subplot
        if rows == 1 and cols == 1:
            ax = axes
        elif rows == 1:
            ax = axes[col]
        elif cols == 1:
            ax = axes[row]
        else:
            ax = axes[row][col]

        # Daten für aktuellen Baumarkt
        daten = plot_data[baumarkt]

        # Durchgehende Datenliste erstellen (alle 4 Jahre hintereinander)
        y_werte = []
        x_werte = []

        for jahr_idx, jahr in enumerate(["2025", "2026", "2027", "2028"]):
            for monat_idx in range(12):
                x_pos = jahr_idx * 12 + monat_idx
                x_werte.append(x_pos)
                y_werte.append(daten[jahr][monat_idx])

        # Hauptlinie plotten
        ax.plot(
            x_werte,
            y_werte,
            linewidth=2,
            marker="o",
            markersize=4,
            color="#1f77b4",
            label="Zeitverlauf",
        )

        # Optionale Jahres-Markierungen (verschiedene Farben für Segmente)
        for jahr_idx, jahr in enumerate(["2025", "2026", "2027", "2028"]):
            start_idx = jahr_idx * 12
            end_idx = (jahr_idx + 1) * 12
            ax.plot(
                x_werte[start_idx:end_idx],
                y_werte[start_idx:end_idx],
                linewidth=3,
                alpha=0.7,
                color=farben[jahr],
                label=jahr,
            )

        # Plot formatieren
        ax.set_title(f"{baumarkt}", fontsize=14, fontweight="bold")
        ax.set_xlabel("Zeit (2025-2028)")
        ax.set_ylabel("Werte")
        ax.legend()
        ax.grid(True, alpha=0.3)

        # X-Achse formatieren
        ax.set_xticks(x_ticks)
        ax.set_xticklabels(x_labels, rotation=45, ha="right")

        # Y-Achse bei 0 beginnen
        ax.set_ylim(bottom=0)

    # Leere Subplots ausblenden
    for i in range(n_baumärkte, rows * cols):
        row = i // cols
        col = i % cols
        if rows == 1:
            axes[col].set_visible(False)
        elif cols == 1:
            axes[row].set_visible(False)
        else:
            axes[row][col].set_visible(False)

    plt.tight_layout()

    # Plot speichern
    os.makedirs("./output/images", exist_ok=True)
    plt.savefig(
        "./output/images/baumarktprogramm_jahresvergleich.png",
        dpi=300,
        bbox_inches="tight",
    )
    print("✅ Plot gespeichert: ./output/images/baumarktprogramm_jahresvergleich.png")

    plt.show()


def plot_gesamt_übersicht(plot_data, monate):
    """
    Erstellt eine Gesamtübersicht aller Baumärkte in einem Plot
    """
    if not plot_data:
        return

    plt.figure(figsize=(15, 10))

    # Farben für die Jahre
    farben = {
        "2025": "#1f77b4",
        "2026": "#ff7f0e",
        "2027": "#2ca02c",
        "2028": "#d62728",
    }

    # X-Positionen
    x = np.arange(len(monate))
    n_baumärkte = len(plot_data)
    width = 0.8 / (n_baumärkte * 4)  # Breite angepasst an Anzahl Baumärkte (4 Jahre)

    # Für jeden Baumarkt und jedes Jahr
    for i, (baumarkt, daten) in enumerate(plot_data.items()):
        for j, jahr in enumerate(["2025", "2026", "2027", "2028"]):
            offset = (i * 4 + j - (n_baumärkte * 4 - 1) / 2) * width
            plt.bar(
                x + offset,
                daten[jahr],
                width,
                label=f"{baumarkt} {jahr}" if i == 0 or j == 0 else "",
                color=farben[jahr],
                alpha=0.7,
            )

    plt.title(
        "Baumarktprogramm - Gesamtübersicht\n(Alle Baumärkte und Jahre 2025-2028)",
        fontsize=16,
        fontweight="bold",
    )
    plt.xlabel("Monate")
    plt.ylabel("Werte")
    plt.xticks(x, monate)
    plt.legend(bbox_to_anchor=(1.05, 1), loc="upper left")
    plt.grid(True, alpha=0.3)

    plt.tight_layout()
    plt.savefig(
        "./output/images/baumarktprogramm_gesamtuebersicht.png",
        dpi=300,
        bbox_inches="tight",
    )
    print(
        "✅ Gesamtübersicht gespeichert: ./output/images/baumarktprogramm_gesamtuebersicht.png"
    )

    plt.show()


def main():
    print("🎨 Baumarktprogramm Plotting gestartet...")

    # 1. Daten laden
    df = load_baumarktprogramm()
    if df is None:
        return

    # 2. Daten für Plotting strukturieren
    plot_data, monate = extract_data_for_plotting(df)

    if not plot_data:
        print("❌ Keine verwendbaren Daten gefunden")
        return

    print(f"✅ Daten für {len(plot_data)} Baumärkte gefunden")

    # 3. Einzelne Plots pro Baumarkt
    plot_baumarkt_vergleich(plot_data, monate)

    print("🎉 Plotting abgeschlossen!")


if __name__ == "__main__":
    main()
