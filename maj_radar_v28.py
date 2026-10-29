#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
maj_radar_v28.py — Harmonisation des libellés (préparation démo, 29/10/2026)
============================================================================

Entrée  : VC_Database_Standardisee_v27.xlsx
Sortie  : VC_Database_Standardisee_v28.xlsx

Remplace uniquement des variantes strictement équivalentes d'un même libellé dans l'onglet
Participations (anglais / français, accord, abréviations de pays), pour que les filtres et
fiches du Radar n'affichent pas « Series B » à côté de « Série B ». Les valeurs détaillées
(« Actif (Série C 2022, 50M$) », « Incubation (2017) ; Série A… ») ne sont pas touchées.
Chaque champ modifié est résumé dans une ligne de l'onglet Anomalies (avant → après, nombre).

Usage : python3 maj_radar_v28.py
"""
from __future__ import annotations

from collections import Counter
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v27.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v28.xlsx"

EQUIVALENCES = {
    "Stage financé": {
        "Series A": "Série A", "Series B": "Série B", "Series C": "Série C", "Series F": "Série F",
        "Pre-seed": "Pré-seed", "Amorçage": "Seed", "Amorçage (seed)": "Seed",
    },
    "Statut start-up (détail)": {"Actif": "Active", "Exit / cédée": "Exit", "n/d": "non trouvé"},
    "Pays d’origine": {"US": "États-Unis", "UK": "Royaume-Uni"},
}


def main() -> int:
    feuilles = pd.read_excel(ENTREE, sheet_name=None)
    part = feuilles["Participations"]
    ano = feuilles["Anomalies"]
    suivant = int(ano["Anomalie_ID"].str.extract(r"(\d+)")[0].astype(int).max()) + 1
    lignes = []
    for champ, table in EQUIVALENCES.items():
        compte = Counter(v for v in part[champ] if v in table)
        if not compte:
            continue
        part[champ] = part[champ].map(lambda v: table.get(v, v) if isinstance(v, str) else v)
        lignes.append({
            "Anomalie_ID": f"ANO-{suivant:04d}", "Type d'anomalie": "Harmonisation de libellés",
            "Onglet source": "Participations", "Table ou bloc source": "Préparation démo 29/10/2026",
            "Ligne source": None, "Entité concernée": f"{sum(compte.values())} lignes", "Champ concerné": champ,
            "Valeur source": " ; ".join(f"{k} ×{n}" for k, n in compte.items()),
            "Valeur retenue": " ; ".join(f"{k} → {table[k]}" for k in compte),
            "Candidats éventuels": None, "Score de similarité": None, "Niveau de confiance": "Élevé",
            "Traitement appliqué": "Variantes strictement équivalentes fusionnées",
            "Commentaire": "Valeurs détaillées (avec précisions entre parenthèses) laissées telles quelles.",
        })
        suivant += 1
    feuilles["Participations"] = part
    feuilles["Anomalies"] = pd.concat([ano, pd.DataFrame(lignes)], ignore_index=True)
    with pd.ExcelWriter(SORTIE, engine="openpyxl") as writer:
        for nom, df in feuilles.items():
            df.to_excel(writer, sheet_name=nom, index=False)
    print(f"Classeur produit : {SORTIE.name} · Anomalies : +{len(lignes)} (jusqu'à ANO-{suivant - 1:04d})")
    for l in lignes:
        print(f"  {l['Champ concerné']} : {l['Valeur retenue']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
