# AFM Analysis

Analyse de données AFM exportées vers Excel (multi-onglets, 1 onglet = 1 échantillon).

## Usage

1. Placer le fichier de données Excel à côté du script et ajuster `INPUT_FILE` en tête de script si besoin, ou le passer en argument.
2. Lancer :

```bash
pip install pandas openpyxl matplotlib seaborn scipy numpy
python AFM_analysis.py            # utilise INPUT_FILE
python AFM_analysis.py data.xlsx  # ou fichier passé en argument
```

## Sorties

- `AFM_Report.xlsx` : données brutes, statistiques descriptives, tests statistiques (Kruskal-Wallis, Mann-Whitney, Cohen's d), classement des paramètres, dashboard avec figures.
- `AFM_Graphs/` : boxplots par échantillon pour les paramètres clés.
- `AFM_Graphs/Bubble_*.png` : bubble plots skewness vs kurtosis — `Ssk` vs `Sku` (bulles ~ Sa, 1 point = 1 image) et `Rsk` vs `Sku` (bulles ~ Ra, 1 point = 1 profil) ; variante `Rsk` vs `Rku` également produite.

L'extraction est basée sur les libellés des paramètres (aucune position fixe n'est supposée) : les blocs d'images sont délimités par la cellule `Parameters AFM`, la valeur de chaque paramètre est la première cellule numérique à droite de son libellé, et le nombre de profils par image est détecté dynamiquement.
