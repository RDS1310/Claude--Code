#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
maj_radar_v30.py — Cohérence « tour d'entrée » / « investi par le fonds » et doublons de deals
==============================================================================================

Entrée  : VC_Database_Standardisee_v29.xlsx
Sortie  : VC_Database_Standardisee_v30.xlsx

Remarque de l'utilisateur (29/10/2026, fiche BlackFin) : Akur8 affichait 18,2 M€ investis pour un tour
de 8,2 M€, et Descartes Underwriting apparaissait deux fois.

- Akur8 (BlackfinTech 1) : seul cas de la base où l'investi dépasse le tour. 8,2 M€ = Série A d'entrée
  (8,9 M$) ; les sources citées indiquent que BlackFin a réinvesti en B et C → Nb tours (fonds) = 3 et
  montant explicitement présenté comme cumul (montant lui-même non retrouvé dans les sources : à vérifier).
- Descartes Underwriting : le champ Deal du fonds portait une coquille (« UnderWritting ») et figurait
  sur les deux véhicules BlackFin. Coquille corrigée ; doublon retiré de BlackfinTech 1 ; les 12 M€ sont
  reportés sur la participation BlackfinTech II (véhicule nommé dans les sources), avec une date de 2018
  dans le champ Deal incohérente avec la Série B de 2022 → signalé à vérifier.
- Concentric : « Superscript (Digital Risks) » aligné sur le libellé de la participation.
- Astorya.vc : « Embrea » (Deal) vs « Embea » (participation) — non tranché, journalisé.

Usage : python3 maj_radar_v30.py
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from maj_radar_v27 import Journal

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v29.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v30.xlsx"
ETIQUETTE = "[Cohérence 29/10/2026]"


def main() -> int:
    feuilles = pd.read_excel(ENTREE, sheet_name=None)
    fonds = feuilles["Fonds"].astype(object)
    part = feuilles["Participations"].astype(object)
    j = Journal(feuilles["Anomalies"])

    def ligne(df, **crit):
        m = pd.Series(True, index=df.index)
        for k, v in crit.items():
            m &= df[k] == v
        assert m.sum() == 1, (crit, int(m.sum()))
        return df.index[m][0]

    # Akur8 : montant cumulé sur 3 tours
    i = ligne(part, **{"Véhicule": "BlackfinTech 1", "Start-up": "Akur8"})
    part.at[i, "Nb tours (fonds)"] = 3
    part.at[i, "Stage financé"] = "Série A"
    part.at[i, "Commentaires"] = (f"{part.at[i, 'Commentaires']} {ETIQUETTE} Tour d'entrée = Série A (8,2 M€) ; "
                                  "18,2 M€ = mise cumulée de BlackFin sur les tours A, B et C (montant à confirmer).")
    j.add("Valeur approximative conservée", "BlackFin Capital Partners / Akur8", "Montant investi par le fonds (M€)",
          "18,2 affiché face à un tour de 8,2", "18,2 = cumul A + B + C ; Nb tours (fonds) 1 → 3", "Moyen",
          "Présenté comme cumul sur la page", "Réinvestissements en B et C cités par PitchBook ; le montant cumulé "
          "lui-même n'est pas retrouvé dans les sources citées.")

    # Descartes Underwriting : coquille, doublon entre véhicules, montant rattaché au bon véhicule
    i1 = ligne(fonds, **{"Véhicule": "BlackfinTech 1"})
    i2 = ligne(fonds, **{"Véhicule": "BlackfinTech II"})
    assert fonds.at[i1, "Deal_02_Nom"] == "Descartes UnderWritting" and fonds.at[i2, "Deal_01_Nom"] == "Descartes UnderWritting"
    for k in ("Nom", "Secteur", "Stage", "Montant_investi_M€", "Date_invest"):
        fonds.at[i1, f"Deal_02_{k}"] = None
    fonds.at[i2, "Deal_01_Nom"] = "Descartes Underwriting"
    p2 = ligne(part, **{"Véhicule": "BlackfinTech II", "Start-up": "Descartes Underwriting"})
    part.at[p2, "Montant investi par le fonds (M€)"] = 12.0
    part.at[p2, "Confiance"] = "Moyen"
    part.at[p2, "Commentaires"] = (f"{part.at[p2, 'Commentaires']} {ETIQUETTE} 12 M€ repris du champ « deal » du fonds ; "
                                   "ce champ datait l'opération de 2018, incohérent avec la Série B de 2022 : à vérifier.")
    j.add("Risque de double comptage", "BlackFin Capital Partners / Descartes Underwriting", "Deal_0x (Fonds)",
          "« Descartes UnderWritting » 12 M€ (2018) sur BlackfinTech 1 et BlackfinTech II",
          "Coquille corrigée ; deal retiré de BlackfinTech 1 ; 12 M€ portés sur la participation BlackfinTech II", "Moyen",
          "Doublon supprimé de la fiche fonds", "Date 2018 du champ deal incohérente avec la Série B 2022 (110 M€) : à vérifier.",
          onglet="Fonds")

    # Concentric : libellé aligné
    i = ligne(fonds, **{"Nom du fonds": "Concentric"})
    if fonds.at[i, "Deal_03_Nom"] == "Superscript (Digital Risks)":
        fonds.at[i, "Deal_03_Nom"] = "Superscript (ex-Digital Risks)"
        j.add("Libellé ambigu", "Concentric / Superscript", "Deal_03_Nom", "Superscript (Digital Risks)",
              "Superscript (ex-Digital Risks)", "Élevé", "Aligné sur la participation (évite un doublon)", "", onglet="Fonds")

    # Astorya.vc : orthographe non tranchée
    j.add("Libellé ambigu", "Astorya.vc / Embea", "Deal_01_Nom", "Embrea (champ deal)", "Inchangé (participation : Embea)",
          "Faible", "Non tranché — orthographe à vérifier", "Le site affiche une seule ligne grâce au rapprochement souple des noms.",
          onglet="Fonds")

    feuilles.update({"Fonds": fonds, "Participations": part, "Anomalies": j.table()})
    with pd.ExcelWriter(SORTIE, engine="openpyxl") as writer:
        for nom, df in feuilles.items():
            df.to_excel(writer, sheet_name=nom, index=False)
    print(f"Classeur produit : {SORTIE.name} · Anomalies +{len(j.lignes)} (jusqu'à ANO-{j.suivant - 1:04d})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
