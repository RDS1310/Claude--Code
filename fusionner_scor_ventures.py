#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fusionner_scor_ventures.py — Complète le portefeuille SCOR Ventures
====================================================================================

SCOR Ventures (CVC du réassureur SCOR, 130M€, 100% dédié assurance/réassurance)
déclarait 25 participations pour seulement 4 documentées dans notre base. Une
recherche dédiée (Gap-Lot 4) en identifie 12 de plus, dont 8 avec un lien
assurance suffisamment clair pour intégration (les 4 autres — FH Ortho SAS,
Finanzguru, Kontempo, Novisto — sont hors périmètre thématique du radar et
volontairement écartées, journalisées en anomalie).

Entrée  : VC_Database_Standardisee_v11.xlsx
Sortie  : VC_Database_Standardisee_v12.xlsx

Usage : python3 fusionner_scor_ventures.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v11.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v12.xlsx"

FONDS_ID = "scor-ventures_scor-ventures"
NOM_FONDS = "SCOR Ventures"
VEHICULE = "SCOR Ventures"

ROWS = [
    # startup, secteur, pays, tour, date, taille($M), devise, role, statut, confiance, commentaires, sources
    ("Human API", "InsurTech (adjacent)", "US", "Série C (+ tour 2019)", 2019.0, 10, "$",
     "Investisseur (SCOR Life & Health Ventures), avec Guardian Life, Allianz Life Ventures, CNO Financial",
     "Exit — racheté par LexisNexis Risk Solutions (24/04/2023)", "Élevé",
     "Plateforme de partage de données de santé pour souscription vie simplifiée. Investissement initial "
     "fév. 2019, suivi Série C oct. 2020 (~20M$).",
     "scor.com/en/news/scor-life-health-ventures-invests-human-api ; risk.lexisnexis.com (communiqué 2023-04-25)"),
    ("Doma (ex-States Title)", "InsurTech (cœur)", "US", "Série C", 2020.0, 123, "$",
     "Investisseur (SCOR Global P&C Ventures), avec Greenspring, Foundation Capital, Assurant, FifthWall, Lennar Ventures",
     "Exit — IPO via SPAC (juillet 2021, Doma Holdings)", "Élevé",
     "Assurance titre immobilier pilotée par la donnée (predictive title insurance).",
     "coverager.com ; alta.org (2020-05-21) ; scor.com Ventures 2.0 2021 Year in Review"),
    ("Snapsheet", "InsurTech (cœur)", "US", "Partenariat/investissement", 2021.0, None, None,
     "Investisseur + partenaire commercial", "Actif", "Moyen-Élevé",
     "Plateforme de gestion de sinistres digitale end-to-end (claims, paiements, expertise virtuelle). Collaboration "
     "ultérieure avec Branch Insurance (autre participation SCOR).",
     "scor.com/en/news/scor-partners-snapsheet ; reinsurancene.ws"),
    ("Hokodo", "InsurTech (cœur)", "UK/France", "Série A", 2021.0, 12.5, "$",
     "Investisseur + capacité de souscription (Lloyd's Channel Syndicate)",
     "Cessation d'activité fin 2025", "Élevé",
     "« Trade Credit as a Service » pour marchands B2B européens, souscrit via SCOR at Lloyd's. Réinvesti en "
     "Série B (2022, ~40M$). Partenariats également avec AIG, BNP Paribas, Citi, Munich Re.",
     "insuranceinsider.com ; businesswire.com (2021-06-10) ; hokodo.co (rétrospective)"),
    ("Energetic Insurance", "InsurTech (cœur)", "US", "Série A", 2019.0, 2.5, "$",
     "Investisseur (SCOR P&C Ventures) + capacité de réassurance (7 renouvellements consécutifs)",
     "Actif (relation renouvelée en 2026)", "Élevé",
     "Assurance crédit paramétrique couvrant le risque de défaut de paiement sur projets solaires commerciaux (PPA).",
     "reinsurancene.ws ; businesswire (renouvellement 2026-05-20)"),
    ("ifeel", "InsurTech (adjacent)", "Espagne", "Série A", 2021.0, None, None,
     "Investisseur (SCOR Life & Health Ventures)", "Actif", "Moyen",
     "Plateforme digitale d'accès à des solutions de santé mentale pour assurés. Date exacte du tour à confirmer "
     "(sources partiellement contradictoires).",
     "scor.com/en/news/scor-life-health-ventures-invests-ifeel ; coverager.com ; insuranceassetrisk.com"),
    ("iBeat", "InsurTech (adjacent)", "US", "Investissement stratégique initial", 2018.0, None, None,
     "Investisseur — 1er investissement de SCOR Life & Health Ventures (lancement du véhicule)",
     "Statut incertain (pas d'actualité récente)", "Moyen",
     "Montre connectée à capteurs médicaux (détection d'arrêt cardiaque), données pour souscription vie/santé. "
     "Co-investisseur : Transamerica Ventures.",
     "scor.com/en/news/scor-global-life-launches-scor-life-health-ventures ; reinsurancene.ws"),
    ("Measured (Analytics and Insurance)", "InsurTech (cœur)", "US", "Non nommés", 2021.0, None, None,
     "Investisseur + fournisseur de capacité de longue date (souscription cyber)", "Actif", "Moyen-Élevé",
     "Cyber-assurance data-driven pour PME (produit CyberGuard). SCOR cité comme « long-standing capacity "
     "provider » avant l'arrivée de Canopius.",
     "theinsurer.com ; scor.com Ventures 2.0 2021 Year in Review"),
]


def convertir(montant, devise):
    if montant is None:
        return None
    taux = {"€": 1.0, "$": 0.92, "£": 1.17}.get(devise, 1.0)
    return round(float(montant) * taux, 2)


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
            "Fonds_ID": FONDS_ID,
            "Nom du fonds": NOM_FONDS,
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
            "Positionnement (Leader/Minoritaire)": "Minoritaire",
            "Capital social (k€)": None,
            "CA (M€)": None,
            "Valorisation (M€)": None,
            "Nb fonds investisseurs": None,
            "Nb tours (fonds)": None,
            "Nb tours (total)": None,
            "Repositionnement": None,
            "Exit (O/N)": "O" if "exit" in statut.lower() or "ipo" in statut.lower() else "N",
            "MoC exit": None,
            "MoEP exit": None,
            "Taille du tour (M€)": convertir(taille, devise),
            "% de détention": None,
            "Rôle du véhicule (détail)": role,
            "Statut start-up (détail)": statut,
            "Confiance": confiance,
            "Commentaires": commentaires,
            "Source(s)": sources,
        })

    part_out = pd.concat([part, pd.DataFrame(nouvelles_lignes)], ignore_index=True)

    prochain_id = int(ano["Anomalie_ID"].str.extract(r"(\d+)")[0].astype(int).max()) + 1
    nouvelle_anomalie = {
        "Anomalie_ID": f"ANO-{prochain_id:04d}",
        "Type d'anomalie": "Participations écartées (hors périmètre thématique)",
        "Onglet source": "Participations",
        "Table ou bloc source": "Gap-Lot 4 (recherche SCOR Ventures, 17/09/2026)",
        "Ligne source": None,
        "Entité concernée": "SCOR Ventures",
        "Champ concerné": "Start-up",
        "Valeur source": "FH Ortho SAS (dispositifs médicaux orthopédiques), Finanzguru (assistant financier "
                          "personnel généraliste), Kontempo (BNPL/trade finance B2B), Novisto (SaaS reporting ESG)",
        "Valeur retenue": "Non intégrées (aucun lien assurance direct)",
        "Candidats éventuels": None,
        "Score de similarité": None,
        "Niveau de confiance": "Moyen",
        "Traitement appliqué": "Écartées du radar malgré leur présence confirmée dans le portefeuille SCOR "
                                "Ventures — hors des catégories InsurTech cœur/adjacent/FinTech liée assurance.",
        "Commentaire": "FH Ortho notamment à faire valider par Groupama si sa pertinence est jugée différente "
                       "(présence confirmée via Crunchbase).",
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
    print(f"Anomalies avant : {len(ano)}  →  après : {len(ano_out)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
