import pandas as pd
import numpy as np
import os
from datetime import datetime
import warnings
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

warnings.filterwarnings("ignore")


def load_rohdaten():
    """
    Lädt die Rohdaten aus rohdaten.xlsx

    Returns:
        pd.DataFrame: Rohdaten mit Bestellinformationen
    """
    try:
        df_raw = pd.read_excel("rohdaten.xlsx")
        print(
            f"✅ Rohdaten geladen: {df_raw.shape[0]} Zeilen, {df_raw.shape[1]} Spalten"
        )

        return df_raw

    except FileNotFoundError:
        print("❌ Datei 'rohdaten.xlsx' nicht gefunden!")
        print("🔍 Verfügbare Excel-Dateien:")
        for file in os.listdir("."):
            if file.endswith(".xlsx"):
                print(f"   - {file}")
        return None

    except Exception as e:
        print(f"❌ Fehler beim Laden der Rohdaten: {e}")
        return None


def load_baumarktprogramm():
    """
    Lädt das Baumarktprogramm aus BAUMARKTPROGRAMM.xlsx

    Returns:
        pd.DataFrame: Baumarktprogramm-Daten mit Prognosen
    """
    try:
        df = pd.read_excel("BAUMARKTPROGRAMM.xlsx")
        print(
            f"✅ Baumarktprogramm geladen: {df.shape[0]} Zeilen, {df.shape[1]} Spalten"
        )

        return df

    except FileNotFoundError:
        print("❌ Datei 'BAUMARKTPROGRAMM.xlsx' nicht gefunden!")
        print("🔍 Verfügbare Excel-Dateien:")
        for file in os.listdir("."):
            if file.endswith(".xlsx"):
                print(f"   - {file}")
        return None

    except Exception as e:
        print(f"❌ Fehler beim Laden des Baumarktprogramms: {e}")
        return None


def agg_Rohdaten(data):
    """
    Aggregiert Rohdaten mit Integration der Prognosedaten.
    Prognosemonate werden unter wavor_bstlmg angezeigt.
    Bei doppelten Daten wird der größere Wert genommen.
    """

    # Schritt 1: Normale Bestelldaten aggregieren
    bestelldaten_agg = (
        data.groupby(["Baumarkt", "bedmo"]).agg({"wavor_bstlmg": "sum"}).reset_index()
    )

    # Schritt 2: Prognosedaten als zusätzliche 'bedmo' behandeln
    prognose1 = (
        data.groupby(["Baumarkt", "progmo"])
        .agg({"prog_mg1": "sum"})
        .reset_index()
        .copy()
    )
    prognose1 = prognose1.rename(
        columns={"progmo": "bedmo", "prog_mg1": "wavor_bstlmg"}
    )

    prognose2 = (
        data.groupby(["Baumarkt", "progmo2"])
        .agg({"prog_mg2": "sum"})
        .reset_index()
        .copy()
    )
    prognose2 = prognose2.rename(
        columns={"progmo2": "bedmo", "prog_mg2": "wavor_bstlmg"}
    )

    # Schritt 3: Zusammenfügen (Bestellungen + Prognosen)
    combined = pd.concat(
        [bestelldaten_agg, prognose1, prognose2], ignore_index=True, sort=False
    )

    # Schritt 4: Fehlende oder ungültige Monate entfernen und Monat normalisieren
    combined = combined.dropna(subset=["bedmo", "wavor_bstlmg"])

    def _normalize_month(val):
        try:
            if isinstance(val, float) and np.isnan(val):
                return np.nan
            if isinstance(val, float):
                # z.B. 202610.0 -> 202610
                return int(val)
            if isinstance(val, (int, np.integer)):
                return int(val)
            if isinstance(val, (pd.Timestamp, datetime)):
                return int(val.strftime("%Y%m"))
            s = str(val).strip()
            if s.endswith(".0"):
                s = s[:-2]
            return int(s)
        except Exception:
            return np.nan

    combined["bedmo"] = combined["bedmo"].apply(_normalize_month)
    combined = combined.dropna(subset=["bedmo"])
    combined["bedmo"] = combined["bedmo"].astype(int)

    # Schritt 5: Bei doppelten Baumarkt/bedmo den größeren Wert nehmen
    finale_daten = combined.groupby(["Baumarkt", "bedmo"], as_index=False).agg(
        {"wavor_bstlmg": "max"}
    )

    # Schritt 6: Spalten umbenennen und sortieren
    finale_daten = finale_daten.rename(
        columns={"bedmo": "Monat", "wavor_bstlmg": "Zahl", "Baumarkt": "Baumarkt"}
    )
    finale_daten = finale_daten.sort_values(["Baumarkt", "Monat"]).reset_index(
        drop=True
    )

    return finale_daten


def agg_Baumarktprogramm(data):
    """
    Wandelt das BAUMARKTPROGRAMM-DataFrame in langes Format um:
    Spalten: ['Baumarkt', 'Monat', 'Zahl']
    Monat ist im Format JJJJMM (int). Fehlende Werte werden als 0 behandelt.
    Mapping der Spaltenbereiche:
      2025: E-P  (Index 4-15)
      2026: R-AC (Index 17-28)
      2027: AE-AP (Index 30-41)
      2028: AR-BC (Index 43-54)
    Entfernt Zeilen, bei denen die Baumarkt-Spalte den Text "Baumarkt" enthält.
    """
    if data is None or data.empty:
        return pd.DataFrame(columns=["Baumarkt", "Monat", "Zahl"])

    # Spaltenbereiche (Index)
    spalten_mapping = {
        "2025": list(range(4, 16)),  # E-P -> 4..15
        "2026": list(range(17, 29)),  # R-AC -> 17..28
        "2027": list(range(30, 42)),  # AE-AP -> 30..41
        "2028": list(range(43, 55)),  # AR-BC -> 43..54
    }

    rows = []
    for _, row in data.iterrows():
        baumarkt = row.iloc[0]
        if pd.isna(baumarkt):
            continue
        bname = str(baumarkt).strip()
        if bname == "" or bname.lower() == "baumarkt":
            continue
        for jahr, indices in spalten_mapping.items():
            for m_idx in range(12):
                col_idx = indices[m_idx] if m_idx < len(indices) else None
                wert = 0.0
                if col_idx is not None and col_idx < len(row):
                    val = row.iloc[col_idx]
                    if pd.notna(val):
                        try:
                            wert = float(val)
                        except Exception:
                            try:
                                wert = float(
                                    str(val).replace(",", ".").replace(" ", "")
                                )
                            except Exception:
                                wert = 0.0
                monat_code = int(f"{jahr}{m_idx+1:02d}")
                rows.append({"Baumarkt": bname, "Monat": monat_code, "Zahl": wert})

    result = pd.DataFrame(rows, columns=["Baumarkt", "Monat", "Zahl"])
    # Falls mehrere Zeilen für gleichen Baumarkt/Monat existieren, zusammenfassen (Summe)
    result = result.groupby(["Baumarkt", "Monat"], as_index=False).agg({"Zahl": "sum"})
    # Entferne eventuelle verbleibende Zeilen mit dem Wort "baumarkt"
    result = result[
        ~result["Baumarkt"].astype(str).str.strip().str.lower().eq("baumarkt")
    ]
    result = result.sort_values(["Baumarkt", "Monat"]).reset_index(drop=True)
    return result


def plot_vergleich_baumarkt(rohdaten_agg, baumarkt_prog, out_dir="./output/images"):
    """
    Vergleichsplots pro Baumarkt:
    - Maßstab der Achsen ist angepasst
    """
    os.makedirs(out_dir, exist_ok=True)

    # Prüfung der benötigten Spalten
    required = {"Baumarkt", "Monat", "Zahl"}
    if not required.issubset(set(rohdaten_agg.columns)) or not required.issubset(
        set(baumarkt_prog.columns)
    ):
        raise ValueError(
            "Beide DataFrames müssen die Spalten 'Baumarkt', 'Monat' und 'Zahl' enthalten."
        )

    # Alle Baumärkte aus beiden DataFrames
    baumaerkte = sorted(
        set(rohdaten_agg["Baumarkt"].dropna().unique()).union(
            set(baumarkt_prog["Baumarkt"].dropna().unique())
        )
    )

    def _to_datetime_month(series):
        # Konvertieren von Monat (JJJJMM) in datetime (erster Tag im Monat)
        s = series.dropna().astype(str).str.strip()
        s = s.str.replace(r"\.0$", "", regex=True)
        s = s[s.str.match(r"^\d{6}$")]
        result = pd.to_datetime(
            series.astype(str).str.replace(r"\.0$", "", regex=True),
            format="%Y%m",
            errors="coerce",
        )
        return result

    for bm in baumaerkte:
        df_r = rohdaten_agg[rohdaten_agg["Baumarkt"] == bm][["Monat", "Zahl"]].copy()
        df_p = baumarkt_prog[baumarkt_prog["Baumarkt"] == bm][["Monat", "Zahl"]].copy()

        if df_r.empty and df_p.empty:
            continue

        # Monat -> datetime
        df_r["Month_dt"] = (
            _to_datetime_month(df_r["Monat"])
            if not df_r.empty
            else pd.Series(dtype="datetime64[ns]")
        )
        df_p["Month_dt"] = (
            _to_datetime_month(df_p["Monat"])
            if not df_p.empty
            else pd.Series(dtype="datetime64[ns]")
        )

        # Drop Zeilen ohne gültiges Datum
        if not df_r.empty:
            df_r = df_r.dropna(subset=["Month_dt"]).copy()
        if not df_p.empty:
            df_p = df_p.dropna(subset=["Month_dt"]).copy()

        # gemeinsamer Zeitraum bestimmen
        min_dt = None
        max_dt = None
        if not df_r.empty:
            min_dt = df_r["Month_dt"].min()
            max_dt = df_r["Month_dt"].max()
        if not df_p.empty:
            if min_dt is None or (df_p["Month_dt"].min() < min_dt):
                min_dt = df_p["Month_dt"].min()
            if max_dt is None or (df_p["Month_dt"].max() > max_dt):
                max_dt = df_p["Month_dt"].max()

        if min_dt is None or max_dt is None:
            continue

        # vollständige Monatsreihe
        dates = pd.date_range(start=min_dt, end=max_dt, freq="MS")
        df_full = pd.DataFrame({"Month_dt": dates})

        # Merge und Serien erstellen (fehlende Monate -> 0)
        series_r = (
            pd.merge(df_full, df_r[["Month_dt", "Zahl"]], on="Month_dt", how="left")
            .set_index("Month_dt")["Zahl"]
            .fillna(0)
            .astype(float)
        )
        series_p = (
            pd.merge(df_full, df_p[["Month_dt", "Zahl"]], on="Month_dt", how="left")
            .set_index("Month_dt")["Zahl"]
            .fillna(0)
            .astype(float)
        )

        # Skalierungsfaktor berechnen (auf Basis des Maximums)
        max_r = series_r.max()
        max_p = series_p.max()
        factor = 1.0
        if max_p > 0 and max_r > 0:
            factor = max_r / max_p

        # Plot erstellen
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.plot(
            dates,
            series_r.values,
            label="Rohdaten",
            color="C0",
            marker="o",
            linewidth=1,
        )
        if series_p.sum() > 0:
            ax.plot(
                dates,
                (series_p * factor).values,
                label=f"Baumarktprogramm (skaliert)",
                color="C1",
                linestyle="--",
                marker="s",
                linewidth=1,
            )

        # Formatierung
        ax.set_title(f"{bm} — Rohdaten vs. Baumarktprogramm")
        ax.set_xlabel("Monat")
        ax.set_ylabel("Zahl (Programm skaliert)")
        ax.xaxis.set_major_locator(mdates.AutoDateLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
        plt.setp(ax.get_xticklabels(), rotation=45, ha="right")
        ax.legend()
        ax.grid(alpha=0.3)

        # Speichern
        safe_name = (
            "".join(c for c in bm if c.isalnum() or c in (" ", "_", "-"))
            .strip()
            .replace(" ", "_")
        )
        out_path = os.path.join(out_dir, f"{safe_name}_vergleich.png")
        fig.tight_layout()
        fig.savefig(out_path, dpi=150)
        plt.close(fig)

    return


def main():
    print("Abweichungsanalyse - Datenimport")
    print("=" * 50)

    rohdaten = load_rohdaten()
    baumarktprogramm = load_baumarktprogramm()

    print("🔬 Abweichungsanalyse - Aufbereitung")
    print("=" * 50)
    rohdaten_agg = agg_Rohdaten(rohdaten)
    os.makedirs("./output", exist_ok=True)
    rohdaten_agg.to_excel("./output/agg_rohdaten.xlsx", index=False)

    # Export des Baumarktprogramms im langen Format
    baumarktProgamm_agg = agg_Baumarktprogramm(baumarktprogramm)
    baumarktProgamm_agg.to_excel("./output/agg_baumarktprogramm.xlsx", index=False)

    plot_vergleich_baumarkt(rohdaten_agg, baumarktProgamm_agg, out_dir="./output/plots/2")


if __name__ == "__main__":
    main()
