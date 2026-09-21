#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
purger_cvc_hors_europe.py — Retire les CVC hors périmètre géographique (Europe)
==================================================================================

L'utilisateur cible l'Europe pour le benchmark concurrentiel de GVI : les CVC à
société mère non européenne mélangent des stratégies/échelles non comparables
(tickets, marchés, thèses) et sont retirés plutôt que de rester en base à des fins
de complétude. Retire les 13 véhicules suivants (et leurs participations) ajoutés
lors de la campagne de recherche CVC externe : MS&AD Ventures, Sompo, Tokio Marine
Future Fund, Nationwide Ventures, MassMutual Ventures, Guardian Strategic Ventures,
New York Life Ventures, Optum Ventures, QBE Ventures, IAG Firemark Ventures, Intact
Ventures, Liberty Mutual Strategic Ventures, Transamerica Ventures.

Conservés (société mère européenne) : Allianz X (DE), MAIF Avenir (FR), Munich Re
Ventures (DE — société mère européenne malgré une équipe basée à San Francisco),
Generali Ventures (IT), Aviva Ventures (UK), Zurich Insurance Group (CH), Achmea
Innovation Fund (NL), plus les 11 CVC déjà en base (tous européens) et GVI/Groupama.

Entrée  : VC_Database_Standardisee_v18.xlsx
Sortie  : VC_Database_Standardisee_v19.xlsx

Usage : python3 purger_cvc_hors_europe.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v18.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v19.xlsx"

PREFIXES_HORS_EUROPE = [
    "msad-ventures_", "sompo_", "tokio-marine-future-fund_", "nationwide-ventures_",
    "massmutual-ventures_", "guardian-strategic-ventures_", "new-york-life-ventures_",
    "optum-ventures_", "qbe-ventures_", "iag-firemark-ventures_", "intact-ventures_",
    "liberty-mutual-strategic-ventures_", "transamerica-ventures_",
]


def hors_europe(fonds_id: str) -> bool:
    return any(fonds_id.startswith(p) for p in PREFIXES_HORS_EUROPE)


def main() -> int:
    fonds = pd.read_excel(ENTREE, sheet_name="Fonds")
    part = pd.read_excel(ENTREE, sheet_name="Participations")
    tours = pd.read_excel(ENTREE, sheet_name="Tours_de_table")
    dico = pd.read_excel(ENTREE, sheet_name="Dictionnaire")
    ano = pd.read_excel(ENTREE, sheet_name="Anomalies")

    mask_f = fonds["Fonds_ID"].apply(hors_europe)
    mask_p = part["Fonds_ID"].apply(hors_europe)

    fonds_retires = fonds.loc[mask_f, "Nom du fonds"].drop_duplicates().tolist()
    fonds_out = fonds.loc[~mask_f].reset_index(drop=True)
    part_out = part.loc[~mask_p].reset_index(drop=True)

    prochain_id = int(ano["Anomalie_ID"].str.extract(r"(\d+)")[0].astype(int).max()) + 1
    nouvelle_anomalie = {
        "Anomalie_ID": f"ANO-{prochain_id:04d}",
        "Type d'anomalie": "Recadrage de périmètre géographique", "Onglet source": "Fonds",
        "Table ou bloc source": "Purge post CVC-Lots (21/09/2026)", "Ligne source": None,
        "Entité concernée": "13 CVC à société mère non européenne",
        "Champ concerné": "Ensemble du véhicule (Fonds + Participations + Contacts)",
        "Valeur source": ", ".join(fonds_retires),
        "Valeur retenue": "Retirés de la base",
        "Candidats éventuels": None, "Score de similarité": None, "Niveau de confiance": "Élevé",
        "Traitement appliqué": (
            "L'utilisateur a précisé que le benchmark CVC cible l'Europe : les véhicules à société mère "
            "japonaise (MS&AD Ventures, Sompo, Tokio Marine Future Fund), nord-américaine (Nationwide, "
            "MassMutual, Guardian, New York Life, Optum, Liberty Mutual, Transamerica) et "
            "australienne/canadienne (QBE, IAG Firemark, Intact) mélangeaient des stratégies/échelles non "
            "comparables et ont été retirés plutôt que conservés pour la complétude. Conservés : Allianz X "
            "(DE), MAIF Avenir (FR), Munich Re Ventures (DE), Generali Ventures (IT), Aviva Ventures (UK), "
            "Zurich Insurance Group (CH), Achmea Innovation Fund (NL) — sociétés mères européennes malgré, "
            "pour certains, des équipes basées hors Europe.",
        ),
        "Commentaire": f"{mask_f.sum()} fonds et {mask_p.sum()} participations retirés.",
    }
    ano_out = pd.concat([ano, pd.DataFrame([nouvelle_anomalie])], ignore_index=True)

    with pd.ExcelWriter(SORTIE, engine="openpyxl") as writer:
        fonds_out.to_excel(writer, sheet_name="Fonds", index=False)
        part_out.to_excel(writer, sheet_name="Participations", index=False)
        tours.to_excel(writer, sheet_name="Tours_de_table", index=False)
        dico.to_excel(writer, sheet_name="Dictionnaire", index=False)
        ano_out.to_excel(writer, sheet_name="Anomalies", index=False)

    print(f"Classeur produit : {SORTIE}")
    print(f"Fonds retirés : {mask_f.sum()}  ({len(fonds)} → {len(fonds_out)})")
    print(f"Participations retirées : {mask_p.sum()}  ({len(part)} → {len(part_out)})")
    nb_cvc = (fonds_out["Type d'investisseur"] == "CVC").sum()
    print(f"CVC restants : {nb_cvc}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
