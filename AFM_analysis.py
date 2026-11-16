# -*- coding: utf-8 -*-
"""
Analyse de données AFM (fichiers Excel multi-onglets, export instrument).

Principe : les données étant copiées/collées depuis un appareil, AUCUNE
position fixe n'est supposée. Tout est repéré par libellé :
  - 1 bloc = 1 image, délimité par la chaîne "Parameters AFM"
    (recherchée n'importe où dans la feuille) ;
  - un paramètre = une cellule dont le texte correspond exactement à un
    libellé connu ; sa valeur = première cellule numérique à droite sur
    la même ligne, son unité = première cellule texte suivant cette
    valeur ; la recherche s'arrête au libellé connu suivant de la ligne
    (les profils sont donc trouvés quel que soit leur décalage en
    colonnes, et s'il y a 5 ou 10 profils par image) ;
  - les paramètres surfaciques (niveau image : Roughness Ra, Sq, Sa...)
    et les paramètres de profil (Ra, Rq, Rz, Δa...) partagent le même
    mécanisme ; les paramètres surfaciques sont rattachés au profil 0.
"""

from itertools import combinations
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns
from scipy.stats import kruskal, mannwhitneyu
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

# ======================================================================
# USER SETTINGS
# ======================================================================

INPUT_FILE = "AFM-data-Ra-Sa.xlsx"
OUTPUT_EXCEL = "AFM_Report.xlsx"
OUTPUT_PLOT_DIR = Path("AFM_Graphs")
OUTPUT_PLOT_DIR.mkdir(exist_ok=True)

BLOCK_HEADER = "Parameters AFM"
BLOCK_FALLBACK_HEADER = "Sample Name"

# Libellés de paramètres de profil (1D) à extraire
PROFILE_PARAMETERS = [
    "Roughness average (Ra)",
    "Root mean square roughness (Rq)",
    "Maximum height of the roughness (Rt)",
    "Maximum roughness valley depth (Rv)",
    "Maximum roughness peak height (Rp)",
    "Average maximum height of the roughness (Rtm)",
    "Average maximum roughness valley depth (Rvm)",
    "Average maximum roughness peak height (Rpm)",
    "Average third highest peak to third lowest valley height (R3z)",
    "Average third highest peak to third lowest valley height (R3z ISO)",
    "Average maximum height of the profile (Rz)",
    "Average maximum height of the roughness (Rz ISO)",
    "Maximum peak to valley roughness (Ry = Rmax)",
    "Skewness (Rsk)",
    "Kurtosis (Rku)",
    "Waviness average (Wa)",
    "Root mean square waviness (Wq)",
    "Waviness maximum height (Wy = Wmax)",
    "Maximum height of the profile (Pt)",
    "Mean spacing of profile irregularities (Sm)",
    "Average wavelength of the profile (λa)",
    "Root mean square (RMS) wavelength of the profile (λq)",
    "Average absolute slope (Δa)",
    "Root mean square (RMS) slope (Δq)",
    "Length (L)",
    "Developed profile length (L₀)",
    "Profile length ratio (lr)",
    "Swedish height (H)",
]

# Libellés de paramètres surfaciques (2D, niveau image)
SURFACE_PARAMETERS = [
    "Roughness Ra",
    "Calibration",
    "RMS roughness (Sq)",
    "Mean roughness (Sa)",
    "Skew (Ssk)",
    "Excess kurtosis",
    "Maximum peak height (Sp)",
    "Maximum pit depth (Sv)",
    "Maximum height (Sz)",
]

# Libellés structurants (délimitent la recherche de valeur sur une ligne)
SECTION_LABELS = [
    "Parameters AFM",
    "Sample Name",
    "Statistical Quantities",
    "Roughness Parameters",
    "Amplitude",
    "Spatial",
    "Hybrid",
    "Functional",
    "File",
    "Image",
    "Selected line",
    "Cut-off",
    # Petites étiquettes de synthèse de l'en-tête de bloc (Ra, Rq, ...)
    # : servent de délimiteurs pour ne pas absorber les colonnes
    # voisines non libellées.
    "Ra",
    "Rq",
    "Rz",
    "Rsk",
    "Rku",
]

# Paramètres affichés dans le dashboard
KEY_PARAMETERS = [
    "Roughness average (Ra)",
    "Root mean square roughness (Rq)",
    "Skewness (Rsk)",
    "Kurtosis (Rku)",
    "Average absolute slope (Δa)",
    "Root mean square (RMS) slope (Δq)",
    "Mean roughness (Sa)",
    "RMS roughness (Sq)",
    "Skew (Ssk)",
    "Excess kurtosis",
]

ALL_PARAMETER_LABELS = set(PROFILE_PARAMETERS) | set(SURFACE_PARAMETERS)
ALL_LABELS = ALL_PARAMETER_LABELS | set(SECTION_LABELS)


# ======================================================================
# HELPERS
# ======================================================================

def norm_txt(x):
    """Normalise un libellé pour comparaison (espaces, unicode)."""
    return " ".join(str(x).replace("\xa0", " ").split())


def to_num(x):
    return pd.to_numeric(x, errors="coerce")


def cohens_d(x, y):
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    nx, ny = len(x), len(y)
    if nx < 2 or ny < 2:
        return np.nan
    pooled_sd = np.sqrt(
        ((nx - 1) * np.var(x, ddof=1) + (ny - 1) * np.var(y, ddof=1))
        / (nx + ny - 2)
    )
    if pooled_sd == 0 or np.isnan(pooled_sd):
        return np.nan
    return (np.mean(x) - np.mean(y)) / pooled_sd


def value_and_unit(cells_after):
    """
    cells_after : liste [(col, valeur), ...] des cellules non nulles
    situées à droite d'un libellé, jusqu'au prochain libellé connu.

    Renvoie (valeur, unité) : valeur = premier numérique rencontré,
    unité = premier texte suivant la valeur. Si une cellule texte
    apparaît AVANT toute valeur numérique, il n'y a pas de valeur
    (cas des libellés sans donnée) -> (NaN, None).
    """
    value, unit = np.nan, None
    for _, v in cells_after:
        if pd.isna(value):
            n = to_num(v)
            if pd.notna(n) and str(v).strip() != "":
                value = n
            continue
        if unit is None and not isinstance(v, (int, float, np.integer, np.floating)):
            candidate = str(v).strip()
            if 0 < len(candidate) <= 8:
                unit = candidate
                break
    return value, unit


def row_cells(df, r):
    """Cellules non nulles d'une ligne : [(col, valeur), ...]."""
    return [
        (c, df.iloc[r, c])
        for c in range(df.shape[1])
        if pd.notna(df.iloc[r, c])
    ]


def find_block_starts(df):
    """
    Debut de chaque bloc image. Un bloc commence normalement par
    'Parameters AFM' ; si cette cellule a ete supprimee lors d'un
    nettoyage du fichier, la ligne 'Sample Name' qui suit immediatement
    sert de delimiteur de secours (candidats a moins de 3 lignes d'ecart
    fusionnes pour ne pas couper un bloc en deux).
    """
    candidates = []
    for r in range(len(df)):
        for c in range(df.shape[1]):
            t = norm_txt(df.iloc[r, c])
            if t == BLOCK_HEADER or t == BLOCK_FALLBACK_HEADER:
                candidates.append(r)
                break
    merged = []
    for r in candidates:
        if merged and r - merged[-1] <= 2:
            continue
        merged.append(r)
    return merged


def extract_block(df, startrow, endrow, sheet):
    """Extrait tous les enregistrements d'un bloc (une image AFM)."""
    block_records = []
    profile_label_cols = set()
    label_occurrences = []

    for r in range(startrow, endrow):
        cells = row_cells(df, r)
        for pos, (c, v) in enumerate(cells):
            txt = norm_txt(v)
            if txt not in ALL_LABELS:
                continue
            span_end = df.shape[1]
            for c2, v2 in cells[pos + 1:]:
                if norm_txt(v2) in ALL_LABELS:
                    span_end = c2
                    break
            cells_after = [(c2, v2) for c2, v2 in cells if c < c2 < span_end]
            label_occurrences.append((r, c, txt, cells_after))
            if txt in PROFILE_PARAMETERS:
                profile_label_cols.add(c)

    profile_cols = sorted(profile_label_cols)
    profile_index = {c: i + 1 for i, c in enumerate(profile_cols)}

    image_name = None
    sample_name = None
    length_values = []

    for r, c, txt, cells_after in label_occurrences:
        value, unit = value_and_unit(cells_after)

        if txt == "Image" and image_name is None:
            for _, v in cells_after:
                if not isinstance(v, (int, float, np.integer, np.floating)):
                    image_name = str(v).strip()
                    break
            continue

        if txt == "Sample Name" and sample_name is None:
            for _, v in cells_after:
                if not isinstance(v, (int, float, np.integer, np.floating)):
                    sample_name = str(v).strip()
                    break
            continue

        if txt == "Length (L)" and pd.notna(value) and 0.5 < value < 20:
            length_values.append(value)

        if txt not in ALL_PARAMETER_LABELS or pd.isna(value):
            continue

        # Le format export place un 0 dans la colonne d'unité quand la
        # donnée surfacique est absente : valeur 0 sans unité = donnée
        # manquante, on ne l'enregistre pas.
        if txt in SURFACE_PARAMETERS and value == 0 and unit is None:
            continue

        block_records.append({
            "Sample": sheet,
            "Image": image_name,
            "SampleName": sample_name,
            "Profile": profile_index.get(c, 0) if txt in PROFILE_PARAMETERS else 0,
            "Parameter": txt,
            "Value": value,
            "Unit": unit or "",
        })

    n_profiles = len(profile_cols)
    if n_profiles == 0:
        return [], None, "Unknown"

    l_mean = np.mean(length_values) if length_values else np.nan
    if 1.5 <= l_mean <= 3.0:
        scan_size = "2x2"
    elif 8.0 <= l_mean <= 12.0:
        scan_size = "10x10"
    else:
        scan_size = "Unknown"

    for rec in block_records:
        rec["ScanSize"] = scan_size

    return block_records, image_name, scan_size


# ======================================================================
# EXTRACT DATA
# ======================================================================

records = []

xls = pd.ExcelFile(INPUT_FILE)

for sheet in xls.sheet_names:
    print("\nTraitement de l'onglet :", sheet)

    df = pd.read_excel(INPUT_FILE, sheet_name=sheet, header=None)

    block_starts = find_block_starts(df)
    print("Nombre d'images trouvées :", len(block_starts))

    image_counter = 0

    for idx, startrow in enumerate(block_starts):
        endrow = block_starts[idx + 1] if idx + 1 < len(block_starts) else len(df)

        block_records, image_name, scan_size = extract_block(df, startrow, endrow, sheet)

        if not block_records:
            print(f"  Bloc {idx + 1} (ligne {startrow}) : aucune donnée, ignoré")
            continue

        image_counter += 1
        for rec in block_records:
            rec["Image"] = image_name or f"Image_{image_counter}"
        records.extend(block_records)

        print(f"  Bloc {idx + 1} (ligne {startrow}) : {image_name} "
              f"| {scan_size} | {len(block_records)} valeurs")

# =====================================================================
# DATAFRAME
# =====================================================================

print("\nNombre total d'enregistrements :", len(records))

data = pd.DataFrame(records)

if data.empty:
    raise SystemExit("Aucune donnée extraite : vérifier INPUT_FILE / format.")

print("\nDimensions du dataframe :")
print(data.shape)

print("\nAperçu :")
print(data.head())

print("\nRépartition 2x2 / 10x10 :")
print(data["ScanSize"].value_counts())

print("\nRépartition par échantillon :")
print(data.groupby(["Sample", "ScanSize"]).size())

# =====================================================================
# DESCRIPTIVE STATS
# =====================================================================

stats_table = (
    data
    .groupby(["ScanSize", "Sample", "Image", "Parameter"], dropna=False)
    .agg(Mean=("Value", "mean"), Std=("Value", "std"), Count=("Value", "count"))
    .reset_index()
)

# =====================================================================
# STATISTICAL TESTS
# =====================================================================

test_results = []

for scan in data["ScanSize"].unique():
    for parameter in data["Parameter"].unique():
        sub = data[(data["ScanSize"] == scan) & (data["Parameter"] == parameter)]
        samples = list(sub["Sample"].unique())
        group_data = [
            sub[sub["Sample"] == s]["Value"].dropna().values
            for s in samples
        ]
        group_data = [g for g in group_data if len(g) > 0]

        kw_stat, kw_p = np.nan, np.nan
        if len(group_data) >= 2:
            try:
                unique_values = np.unique(np.concatenate(group_data))
                if len(unique_values) > 1:
                    kw_stat, kw_p = kruskal(*group_data)
            except Exception:
                pass

        for s1, s2 in combinations(samples, 2):
            d1 = sub[sub["Sample"] == s1]["Value"].dropna()
            d2 = sub[sub["Sample"] == s2]["Value"].dropna()
            if len(d1) < 2 or len(d2) < 2:
                continue
            try:
                mw_stat, mw_p = mannwhitneyu(d1, d2, alternative="two-sided")
            except Exception:
                mw_stat, mw_p = np.nan, np.nan
            test_results.append({
                "ScanSize": scan,
                "Parameter": parameter,
                "Sample1": s1,
                "Sample2": s2,
                "Kruskal_p": kw_p,
                "MW_p": mw_p,
                "Cohens_d": abs(cohens_d(d1, d2)),
            })

tests_df = pd.DataFrame(test_results)

# =====================================================================
# PARAMETER RANKING
# =====================================================================

ranking = tests_df.copy()
ranking["Significance"] = -np.log10(ranking["MW_p"].fillna(1))
ranking = ranking.sort_values("Significance", ascending=False)

# =====================================================================
# BOXPLOTS - BY SAMPLE, KEY PARAMETERS
# =====================================================================

print("\n================ KEY PARAMETERS ================\n")

for parameter in KEY_PARAMETERS:
    n = len(data[data["Parameter"] == parameter])
    print(parameter)
    print("Nombre de lignes :", n)
    print("-" * 40)

generated_figures = []

for scan in data["ScanSize"].unique():
    scan_df = data[data["ScanSize"] == scan]

    for parameter in KEY_PARAMETERS:
        sub = scan_df[scan_df["Parameter"] == parameter]
        if len(sub) == 0:
            continue

        plt.figure(figsize=(12, 8))
        ax = sns.boxplot(data=sub, x="Sample", y="Value", showmeans=True)
        sns.stripplot(data=sub, x="Sample", y="Value", color="black",
                      alpha=0.5, size=4)

        ax.set_xlabel("")
        units = sub["Unit"].dropna().astype(str)
        units = units[units != ""].unique()
        unit = units[0] if len(units) == 1 else "a.u."
        ax.set_ylabel(f"({unit})")

        stats_local = (sub.groupby("Sample")
                       .agg(Mean=("Value", "mean"), Std=("Value", "std"),
                            N=("Value", "count"))
                       .reset_index())

        # Position entiere sur l'axe (l'index textuel du groupby ne se
        # projette pas sur la bonne categorie)
        tick_labels = [t.get_text() for t in ax.get_xticklabels()]
        for i, label in enumerate(tick_labels):
            row = stats_local[stats_local["Sample"] == label]
            if row.empty:
                continue
            row = row.iloc[0]
            ax.text(i, -0.10, f"{row['Mean']:.2f}±{row['Std']:.2f}",
                    ha="center", va="top", fontsize=9,
                    transform=ax.get_xaxis_transform())
            ax.text(i, -0.16, f"n={int(row['N'])}",
                    ha="center", va="top", fontsize=9,
                    transform=ax.get_xaxis_transform())

        plt.title(f"{parameter} ({scan}) - By Sample")
        plt.tight_layout()

        fig_name = OUTPUT_PLOT_DIR / f"{parameter}_{scan}_BySample.png"
        plt.savefig(fig_name, dpi=300, bbox_inches="tight")
        plt.close()
        generated_figures.append(str(fig_name))

print(f"\nNombre total de graphiques générés : {len(generated_figures)}")

# =====================================================================
# SAVE EXCEL REPORT
# =====================================================================

with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
    data.to_excel(writer, sheet_name="Raw_Data", index=False)
    counts_table = (data.groupby(["Sample", "ScanSize", "Parameter"])["Value"]
                    .count().reset_index(name="Count"))
    counts_table.to_excel(writer, sheet_name="Counts", index=False)
    stats_table.to_excel(writer, sheet_name="Statistics", index=False)
    tests_df.to_excel(writer, sheet_name="Statistical_Tests", index=False)
    ranking.to_excel(writer, sheet_name="Parameter_Ranking", index=False)
    pd.DataFrame().to_excel(writer, sheet_name="Dashboard")

# =====================================================================
# INSERT KEY PARAMETERS FIGURES INTO DASHBOARD
# =====================================================================

wb = load_workbook(OUTPUT_EXCEL)
ws = wb["Dashboard"]
row_pos = 1

for scan in ["2x2", "10x10"]:
    for parameter in KEY_PARAMETERS:
        fig_file = OUTPUT_PLOT_DIR / f"{parameter}_{scan}_BySample.png"
        if fig_file.exists():
            img = XLImage(str(fig_file))
            img.width, img.height = 550, 300
            ws.add_image(img, f"A{row_pos}")
            row_pos += 18

wb.save(OUTPUT_EXCEL)

print("\nRapport sauvegardé :", OUTPUT_EXCEL)
print("Graphiques dans :", OUTPUT_PLOT_DIR)
