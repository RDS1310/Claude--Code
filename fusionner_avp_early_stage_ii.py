#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fusionner_avp_early_stage_ii.py — Intègre la recherche dédiée AVP early stage II
====================================================================================

La presse (avril 2021) mentionne "16 investissements insurtech" dans le
portefeuille d'Atlantic Vantage Point (ex-AXA Venture Partners) sans jamais
les nommer. Une recherche dédiée (Gap-Lot 1, relancée après une interruption
de rate-limit) en a identifié 6, avec un niveau de preuve acceptable, sur la
page portefeuille officielle axavp.com et des communiqués concordants :
Thimble, Gravie, Idelic, Dayforward, Limelight Health, ARTA (cette dernière,
plus logtech qu'assurance pure, marquée Confiance Moyenne et catégorie
adjacente). Le reste (10-12 sociétés) demeure non identifié malgré une
recherche approfondie et l'exclusion explicite d'une trentaine de pistes
vérifiées et infirmées (Vesttoo, Openly, Bestow, Wefox, Getsafe, Zego, etc.).

Entrée  : VC_Database_Standardisee_v10.xlsx
Sortie  : VC_Database_Standardisee_v11.xlsx

Usage : python3 fusionner_avp_early_stage_ii.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v10.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v11.xlsx"

FONDS_ID = "atlantic-vantage-point_avp-early-stage-ii"
VEHICULE = "AVP early stage II"

ROWS = [
    # startup, secteur, pays, tour, date, taille($M), role, statut, confiance, commentaires, sources
    ("Thimble (ex-Verifly)", "InsurTech (cœur)", "US", "Série A", 2019.0, 22,
     "Participant (lead : IAC, avec Slow Ventures, Open Ocean)", "Actif", "Élevé",
     "Assurance à la demande pour TPE/indépendants. Page portefeuille officielle axavp.com/avp/thimble confirmée.",
     "axavp.com/avp/thimble ; Insurance Journal ; Global Venturing ; Coverager"),
    ("Gravie", "InsurTech (cœur)", "US", "Série D", 2021.0, 28,
     "Lead (chef de file)", "Actif", "Élevé",
     "Marketplace d'assurance santé collective (employee benefits). Communiqué officiel Gravie « led by AXA "
     "Venture Partners ». Ticket ($28M) dépasse le plafond de $6M initialement annoncé pour Early Stage II — "
     "pourrait relever d'un véhicule Growth plutôt que du early-stage strict, non tranché.",
     "gravie.com (communiqué) ; insurtechinsights.com ; axavp.com/avp/gravie"),
    ("Idelic", "InsurTech (cœur)", "US", "Série B", 2021.0, 20,
     "Participant (lead : Highland Capital Partners)", "Actif", "Élevé",
     "Plateforme sécurité conducteurs/flottes réduisant le risque assurantiel (trucking).",
     "axavp.com/idelic-raises-a-20m-series-b ; BusinessWire ; Carrier Management"),
    ("Dayforward", "InsurTech (cœur)", "US", "Levée", 2023.0, 25,
     "Lead (chef de file, avec HSCM Ventures, Juxtapose, Munich Re Ventures)", "Actif", "Élevé",
     "Assurance-vie temporaire digitale. Ticket ($25M) dépasse le plafond initial $6M — même remarque que Gravie "
     "sur un possible véhicule Growth.",
     "axavp.com/avp/dayforward ; PR Newswire ; Coverager ; FinSMEs ; Insurance Business"),
    ("Limelight Health", "InsurTech (adjacent)", "US", "Série C", 2019.0, None,
     "Participant (lead : Principal Life)", "Racheté par Sun Life (à reconfirmer)", "Faible",
     "Plateforme de devis/souscription pour avantages sociaux employeurs. Round du 17/01/2019 à la charnière "
     "entre la clôture d'Early Stage I et l'annonce d'Early Stage II (22/01/2019) — véhicule exact non tranchable. "
     "Montant AVP non isolé du total ($33,5M).",
     "axavp.com/avp/limelight-health ; PR Newswire ; TechStartups"),
    ("ARTA", "InsurTech (adjacent)", "US", "Série A", 2022.0, 11,
     "Lead (AVP a mené le tour)", "Actif (a priori)", "Moyen",
     "Plateforme de fulfillment/expédition pour biens de valeur (art, bijoux, objets de collection) intégrant des "
     "produits d'assurance (partenariat Chubb). Classement « insurtech » discutable : cœur de métier = logistique, "
     "assurance = produit accessoire intégré. Ticket et calendrier cohérents avec Early Stage II.",
     "axavp.com/arta-raises-11m ; FinTech Global ; FinSMEs ; Coverager"),
]


def main() -> int:
    fonds = pd.read_excel(ENTREE, sheet_name="Fonds")
    part = pd.read_excel(ENTREE, sheet_name="Participations").copy()
    tours = pd.read_excel(ENTREE, sheet_name="Tours_de_table")
    dico = pd.read_excel(ENTREE, sheet_name="Dictionnaire")
    ano = pd.read_excel(ENTREE, sheet_name="Anomalies")

    existants = {(r["Fonds_ID"], r["Start-up"]) for _, r in part.iterrows()}

    nouvelles_lignes = []
    for startup, secteur, pays, tour, date, taille_usd, role, statut, confiance, commentaires, sources in ROWS:
        if (FONDS_ID, startup) in existants:
            continue
        nouvelles_lignes.append({
            "Fonds_ID": FONDS_ID,
            "Nom du fonds": "Atlantic Vantage Point",
            "Véhicule": VEHICULE,
            "Start-up": startup,
            "Secteur": secteur,
            "Stage financé": tour,
            "Montant investi par le fonds (M€)": None,
            "Total levé par la start-up (M€)": None,
            "Date de création": None,
            "Date d’investissement": date,
            "Pays d’origine": pays,
            "Nb pays d’implantation": None,
            "Positionnement (Leader/Minoritaire)": "Leader" if "lead" in role.lower() else "Minoritaire",
            "Capital social (k€)": None,
            "CA (M€)": None,
            "Valorisation (M€)": None,
            "Nb fonds investisseurs": None,
            "Nb tours (fonds)": None,
            "Nb tours (total)": None,
            "Repositionnement": None,
            "Exit (O/N)": "N",
            "MoC exit": None,
            "MoEP exit": None,
            "Taille du tour (M€)": round(taille_usd * 0.92, 2) if taille_usd is not None else None,
            "% de détention": None,
            "Rôle du véhicule (détail)": role,
            "Statut start-up (détail)": statut,
            "Confiance": confiance,
            "Commentaires": commentaires,
            "Source(s)": sources,
        })

    part_out = pd.concat([part, pd.DataFrame(nouvelles_lignes)], ignore_index=True)

    with pd.ExcelWriter(SORTIE, engine="openpyxl") as writer:
        fonds.to_excel(writer, sheet_name="Fonds", index=False)
        part_out.to_excel(writer, sheet_name="Participations", index=False)
        tours.to_excel(writer, sheet_name="Tours_de_table", index=False)
        dico.to_excel(writer, sheet_name="Dictionnaire", index=False)
        ano.to_excel(writer, sheet_name="Anomalies", index=False)

    print(f"Classeur produit : {SORTIE}")
    print(f"Participations avant : {len(part)}  →  après : {len(part_out)}  (+{len(nouvelles_lignes)})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
