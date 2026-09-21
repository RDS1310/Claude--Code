#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fusionner_kima.py — Complète le portefeuille Kima Ventures (Gap-Lot 8)
====================================================================================

Kima Ventures déclarait 10 participations insurtech pour 4 documentées
(Alan, Luko, Panora, Welfaire). Recherche dédiée : 3 nouvelles confirmées
(Stoïk, Assurup, Cautioneo). Une 4e piste (Acheel) a été volontairement
écartée — les sources citent Xavier Niel à titre personnel, pas le fonds
Kima Ventures explicitement, distinction importante que l'agent a repérée
et respectée.

Entrée  : VC_Database_Standardisee_v13.xlsx
Sortie  : VC_Database_Standardisee_v14.xlsx

Usage : python3 fusionner_kima.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v13.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v14.xlsx"

FONDS_ID = "kima-venture-capital_kima-venture"
NOM_FONDS = "Kima Venture Capital"
VEHICULE = "Kima Venture"

ROWS = [
    # startup, secteur, pays, tour, date, taille, devise, role, statut, confiance, commentaires, sources
    ("Stoïk", "InsurTech (cœur)", "France", "Seed", 2022.0, 3.8, "€",
     "Co-investisseur (avec Alven Capital, Anthemis Group, business angels dont R. Vullierme/Luko, E. Schalit/Dashlane, H. Kravis)",
     "Actif (Série B 2024, Série C 2026)", "Élevé",
     "Assurance cyber pour PME (courtier/porteur de risque + logiciel de monitoring de sécurité). Société fondée mars 2021.",
     "FrenchWeb ; Global Security Mag"),
    ("Assurup", "InsurTech (cœur)", "France", "Seed (1er tour)", 2017.0, 1.0, "€",
     "Co-investisseur (avec Keyrus Innovation Factory, business angels dont F. Mazzella/BlaBlaCar, F. Nappez, C. Vermeulen)",
     "Actif", "Élevé",
     "Courtier d'assurance 100% digital pour start-ups. Fondée 2015 par David Carasso et Jérémy Dahan.",
     "FrenchWeb ; CFNews ; Fusacq"),
    ("Cautioneo", "FinTech liée assurance", "France", "Pré-seed/Seed", 2018.0, 0.35, "€",
     "Co-investisseur (avec Founders Future, Marc Ménasé)",
     "Statut à confirmer", "Élevé",
     "Intermédiaire en assurance/garantie locative pour indépendants et profils précaires. Basée à Lille, fondée par Julien Chenet et David Edery.",
     "Maddyness ; Tribune de l'Assurance ; Univers Freebox"),
]


def conv(montant, devise):
    return round(montant * {"€": 1.0, "$": 0.92, "£": 1.17}.get(devise, 1.0), 2)


def main() -> int:
    fonds = pd.read_excel(ENTREE, sheet_name="Fonds")
    part = pd.read_excel(ENTREE, sheet_name="Participations").copy()
    tours = pd.read_excel(ENTREE, sheet_name="Tours_de_table")
    dico = pd.read_excel(ENTREE, sheet_name="Dictionnaire")
    ano = pd.read_excel(ENTREE, sheet_name="Anomalies")

    existants = {(r["Fonds_ID"], r["Start-up"]) for _, r in part.iterrows()}

    nouvelles_lignes = []
    for startup, secteur, pays, tour, date, taille, devise, role, statut, confiance, commentaires, sources in ROWS:
        if (FONDS_ID, startup) in existants:
            continue
        nouvelles_lignes.append({
            "Fonds_ID": FONDS_ID, "Nom du fonds": NOM_FONDS, "Véhicule": VEHICULE,
            "Start-up": startup, "Secteur": secteur, "Stage financé": tour,
            "Montant investi par le fonds (M€)": None, "Total levé par la start-up (M€)": None,
            "Date de création": None, "Date d’investissement": date, "Pays d’origine": pays,
            "Nb pays d’implantation": None,
            "Positionnement (Leader/Minoritaire)": "Minoritaire",
            "Capital social (k€)": None, "CA (M€)": None, "Valorisation (M€)": None,
            "Nb fonds investisseurs": None, "Nb tours (fonds)": None, "Nb tours (total)": None,
            "Repositionnement": None, "Exit (O/N)": "N", "MoC exit": None, "MoEP exit": None,
            "Taille du tour (M€)": conv(taille, devise), "% de détention": None,
            "Rôle du véhicule (détail)": role, "Statut start-up (détail)": statut,
            "Confiance": confiance, "Commentaires": commentaires, "Source(s)": sources,
        })

    part_out = pd.concat([part, pd.DataFrame(nouvelles_lignes)], ignore_index=True)

    prochain_id = int(ano["Anomalie_ID"].str.extract(r"(\d+)")[0].astype(int).max()) + 1
    nouvelle_anomalie = {
        "Anomalie_ID": f"ANO-{prochain_id:04d}",
        "Type d'anomalie": "Participation non retenue (investisseur personnel vs véhicule)",
        "Onglet source": "Participations",
        "Table ou bloc source": "Gap-Lot 8 (recherche Kima Ventures, 21/09/2026)",
        "Ligne source": None, "Entité concernée": "Kima Venture Capital",
        "Champ concerné": "Start-up",
        "Valeur source": "Acheel (Seed 2023, 29M€) — sources citent « Xavier Niel » comme investisseur personnel, pas le fonds Kima Ventures",
        "Valeur retenue": "Non intégrée",
        "Candidats éventuels": None, "Score de similarité": None, "Niveau de confiance": "Moyen",
        "Traitement appliqué": "Écartée par prudence : aucune source primaire ne nomme explicitement le fonds Kima Ventures (par opposition à Xavier Niel à titre personnel, qui investit aussi en direct hors du véhicule).",
        "Commentaire": "Jean de La Rochebrochard (associé gérant de Kima) aurait facilité la rencontre des fondateurs, mais ceci n'établit pas un investissement du fonds. Fusacq, Next-Finance.",
    }
    ano_out = pd.concat([ano, pd.DataFrame([nouvelle_anomalie])], ignore_index=True)

    with pd.ExcelWriter(SORTIE, engine="openpyxl") as writer:
        fonds.to_excel(writer, sheet_name="Fonds", index=False)
        part_out.to_excel(writer, sheet_name="Participations", index=False)
        tours.to_excel(writer, sheet_name="Tours_de_table", index=False)
        dico.to_excel(writer, sheet_name="Dictionnaire", index=False)
        ano_out.to_excel(writer, sheet_name="Anomalies", index=False)

    print(f"Classeur produit : {SORTIE}")
    print(f"Participations avant : {len(part)}  →  après : {len(part_out)}  (+{len(nouvelles_lignes)})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
