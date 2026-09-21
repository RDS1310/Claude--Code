#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fusionner_cvc_gvi.py — Distinction CVC vs VC indépendant + ajout de GVI (Groupama)
====================================================================================

Deux opérations en une seule passe :
1. Classification rétroactive : ajoute deux colonnes au sheet Fonds,
   "Type d'investisseur" (CVC / VC indépendant) et "Société mère (si CVC)",
   sur les 111 fonds déjà documentés. 11 véhicules déjà en base sont en
   réalité des CVC assurantiels (Macif Innovation, SCOR Ventures, Ethias
   Ventures, Open CNP, UNIQA Ventures, Helsana HealthInvest, Insurtech
   Capital/Apicil, Howden Ventures, CommerzVentures, NCA, Kickstart
   Innovation) et sont reclassés en conséquence ; tous les autres restent
   "VC indépendant".
2. Ajout de Groupama Vol'terre Investissement (GVI) comme nouveau véhicule
   CVC, à partir de la donnée interne gvi_data/GVI_Status_Report_Startups.xlsx
   (Volt'terre GVI = "Oui" → 2 participations en portefeuille : Certificall,
   AlloBrain). Aucune recherche web nécessaire : donnée interne déjà fiable.

Entrée  : VC_Database_Standardisee_v14.xlsx
Sortie  : VC_Database_Standardisee_v15.xlsx

Usage : python3 fusionner_cvc_gvi.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v14.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v15.xlsx"
GVI_SOURCE = BASE_DIR / "gvi_data" / "GVI_Status_Report_Startups.xlsx"

CVC_MAP = {
    "Macif Innovation": "MACIF",
    "SCOR Ventures": "SCOR",
    "Ethias Ventures": "Ethias",
    "Open CNP": "CNP Assurances",
    "UNIQA Ventures": "UNIQA",
    "Helsana HealthInvest": "Helsana",
    "Insurtech Capital": "Apicil",
    "Howden Ventures": "Howden (courtier)",
    "CommerzVentures": "Commerzbank (banque)",
    "NCA (Next Commerce Accelerator)": "AXA / BNP Paribas Cardif / La Poste (consortium)",
    "Kickstart Innovation": "Consortium corporate suisse (Mobiliar, Helvetia, Zurich Insurance, SBB, Six, Post, etc.)",
}

FONDS_ID_GVI = "groupama_gvi"


def main() -> int:
    fonds = pd.read_excel(ENTREE, sheet_name="Fonds")
    part = pd.read_excel(ENTREE, sheet_name="Participations").copy()
    tours = pd.read_excel(ENTREE, sheet_name="Tours_de_table")
    dico = pd.read_excel(ENTREE, sheet_name="Dictionnaire")
    ano = pd.read_excel(ENTREE, sheet_name="Anomalies")

    # --- 1. Classification rétroactive ---
    fonds["Type d'investisseur"] = fonds["Nom du fonds"].map(
        lambda n: "CVC" if n in CVC_MAP else "VC indépendant"
    )
    fonds["Société mère (si CVC)"] = fonds["Nom du fonds"].map(CVC_MAP)

    nouvelle_ligne_dico = {
        "Colonne": "Type d'investisseur",
        "Description": "CVC (Corporate Venture Capital, véhicule d'investissement d'un groupe industriel/assurantiel) ou VC indépendant (fonds tiers, LP institutionnels/privés).",
        "Valeurs possibles": "CVC ; VC indépendant",
    }
    dico_cols = set(dico.columns)
    if {"Colonne", "Description", "Valeurs possibles"} <= dico_cols:
        dico = pd.concat([dico, pd.DataFrame([nouvelle_ligne_dico])], ignore_index=True)

    # --- 2. Ajout de GVI (Groupama) ---
    su = pd.read_excel(GVI_SOURCE, "Start-up")
    su = su.astype(object).where(pd.notna(su), None)
    portefeuille = su[su["Volt'terre \nGVI"].astype(str).str.strip().str.lower() == "oui"]

    SECTEUR_GVI = "InsurTech (adjacent)"  # outils/IA servant les opérations d'un assureur (Groupama)

    deals_cols = {}
    for i, (_, r) in enumerate(portefeuille.iterrows(), start=1):
        deals_cols[f"Deal_{i:02d}_Nom"] = r.get("Start-up")
        deals_cols[f"Deal_{i:02d}_Secteur"] = SECTEUR_GVI
        deals_cols[f"Deal_{i:02d}_Stage"] = r.get("Maturité")
        deals_cols[f"Deal_{i:02d}_Montant_investi_M€"] = None
        deals_cols[f"Deal_{i:02d}_Date_invest"] = None
    for i in range(len(portefeuille) + 1, 4):
        deals_cols[f"Deal_{i:02d}_Nom"] = None
        deals_cols[f"Deal_{i:02d}_Secteur"] = None
        deals_cols[f"Deal_{i:02d}_Stage"] = None
        deals_cols[f"Deal_{i:02d}_Montant_investi_M€"] = None
        deals_cols[f"Deal_{i:02d}_Date_invest"] = None

    ligne_gvi = {
        "Nom du fonds": "Groupama", "Véhicule": "Groupama Vol'terre Investissement (GVI)",
        "AuM (M€)": None, "DPI": None, "RVPI": None, "TVPI": None, "MoC (MoM)": None,
        "IRR (TRI)": None, "Quartile": None, "Montant levé (M€)": None, "Montant alloué (M€)": None,
        "Ticket min (M€)": None, "Ticket max (M€)": None, "Statut": "Actif", "Phase": "Gestion",
        "Millésime": None, "Géographie": "France",
        "Stratégie": "InsurTech (adjacent) ; Innovation assurance ; Outils internes Groupama",
        "Stages pratiqués": "Pré-seed ; Seed",
        "Pré-Seed": 1.0, "Seed": 1.0, "Pré-Série A": 0.0, "Série A": 0.0, "Série B": 0.0,
        "Série C": 0.0, "Série D": 0.0,
        "Nb participations (déclaré)": len(portefeuille), "Nb participations documentées": len(portefeuille),
        "Nb exits": 0.0, "Nb InsurTech": len(portefeuille), "Fonds_ID": FONDS_ID_GVI,
        "Source_AuM": (
            "Donnée interne — gvi_data/GVI_Status_Report_Startups.xlsx (suivi interne GVI, "
            f"{len(su)} start-ups analysées, {len(portefeuille)} en portefeuille actif "
            "[Volt'terre GVI = Oui] au 21/09/2026). Voir aussi la page « Stratégie & Pipe GVI » "
            "pour la fiche complète de chaque start-up (pipe nominatif consultable)."
        ),
        "Date_MAJ": pd.Timestamp.now(),
        **deals_cols,
        "Site web": None,
        "Type d'investisseur": "CVC",
        "Société mère (si CVC)": "Groupama",
    }
    fonds_out = pd.concat([fonds, pd.DataFrame([ligne_gvi])], ignore_index=True)

    nouvelles_participations = []
    for _, r in portefeuille.iterrows():
        nom = r.get("Start-up")
        total_leve_txt = str(r.get("Total  levé") or "")
        nouvelles_participations.append({
            "Fonds_ID": FONDS_ID_GVI, "Nom du fonds": "Groupama", "Véhicule": "Groupama Vol'terre Investissement (GVI)",
            "Start-up": nom, "Secteur": SECTEUR_GVI, "Stage financé": r.get("Maturité"),
            "Montant investi par le fonds (M€)": None, "Total levé par la start-up (M€)": None,
            "Date de création": r.get("Création"), "Date d’investissement": None,
            "Pays d’origine": r.get("Pays\nd'origine"), "Nb pays d’implantation": None,
            "Positionnement (Leader/Minoritaire)": "Minoritaire",
            "Capital social (k€)": None, "CA (M€)": None, "Valorisation (M€)": None,
            "Nb fonds investisseurs": None, "Nb tours (fonds)": None, "Nb tours (total)": None,
            "Repositionnement": None, "Exit (O/N)": "N", "MoC exit": None, "MoEP exit": None,
            "Taille du tour (M€)": None, "% de détention": None,
            "Rôle du véhicule (détail)": "Investisseur (CVC Groupama, programme Vol'terre)",
            "Statut start-up (détail)": "Actif (portefeuille GVI)",
            "Confiance": "Élevé",
            "Commentaires": r.get("Pitch"),
            "Source(s)": (
                "Donnée interne — gvi_data/GVI_Status_Report_Startups.xlsx (Volt'terre GVI = Oui) ; "
                "fiche complète disponible sur la page Stratégie & Pipe GVI (pipe nominatif). "
                f"Total levé déclaré : {total_leve_txt or 'non précisé'}."
            ),
        })

    part_out = pd.concat([part, pd.DataFrame(nouvelles_participations)], ignore_index=True)

    with pd.ExcelWriter(SORTIE, engine="openpyxl") as writer:
        fonds_out.to_excel(writer, sheet_name="Fonds", index=False)
        part_out.to_excel(writer, sheet_name="Participations", index=False)
        tours.to_excel(writer, sheet_name="Tours_de_table", index=False)
        dico.to_excel(writer, sheet_name="Dictionnaire", index=False)
        ano.to_excel(writer, sheet_name="Anomalies", index=False)

    print(f"Classeur produit : {SORTIE}")
    print(f"Fonds avant : {len(fonds)}  →  après : {len(fonds_out)}  (+1 : GVI)")
    print(f"Participations avant : {len(part)}  →  après : {len(part_out)}  (+{len(nouvelles_participations)})")
    print(f"CVC reclassés : {len(CVC_MAP)}  |  VC indépendants : {len(fonds) - len(CVC_MAP)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
