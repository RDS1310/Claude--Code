#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fusionner_gap_lots.py — Intègre les Gap-Lots 2 et 3 (véhicules "vides" réexaminés)
====================================================================================

Une partie des véhicules "Fermé" avaient été explorés de façon non exhaustive
(quota épuisé) et rapportés comme "aucune insurtech trouvée" sans garantie
d'exhaustivité. Une nouvelle passe avec quota frais (Gap-Lot 2 : XAnge/NewFund ;
Gap-Lot 3 : 360 Capital/Samaipata/TomCat/Alven/ISAI) a trouvé 5 nouvelles
participations et confirmé à zéro plusieurs véhicules supplémentaires
(XAnge Capital 2, XAnge Digital 3, Newfund NAEH, Samaipata I/II, 360 Life II,
TomCat Ventures I).

Entrée  : VC_Database_Standardisee_v9.xlsx
Sortie  : VC_Database_Standardisee_v10.xlsx

Usage : python3 fusionner_gap_lots.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v9.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v10.xlsx"

TAUX_CHANGE = {"€": 1.0, "$": 0.92, "£": 1.17, "CHF": 1.07}

ROWS = [
    # nom_fonds, fonds_id, vehicule, startup, secteur, pays, tour, date, taille, devise,
    # role, statut, confiance, commentaires, sources
    ("XAnge Capital", "xange-capital_xange-4", "XAnge 4", "Dattak", "InsurTech (cœur)", "France",
     "Seed", 2022.0, 7, "€", "Investisseur principal (avec Matthieu Bébéar, ex-AXA, business angel)",
     "Active (3500 clients assurés, 20M€ de primes fin 2025)", "Moyen",
     "MGA cyber-assurance PME/ETI (couverture + prévention, distribution via courtiers). Réinvesti en Série A "
     "(2023-07-20, 11M€, co-lead avec Breega et Bpifrance, avec Raise Ventures/Partech ; total levé 18M€). "
     "Véhicule XAnge exact non nommé, XAnge 4 déduit par calendrier.",
     "frenchweb.fr ; dattak.io (communiqué) ; siparex.com ; presse.bpifrance.fr ; planetecsca.fr"),
    ("NewFund Capital", "newfund-capital_newfund-2", "NewFund 2", "Coverd", "InsurTech (cœur)", "France",
     "1er tour", 2020.0, 1.2, "€", "Investisseur (avec business angels)",
     "Incertain — pas d'actualité récente (2023-2026), statut actuel non vérifié", "Faible",
     "Assurtech affinitaire : assurance smartphone/high-tech 100% en ligne. Véhicule exact non nommé ; "
     "ni Coverd ni Olino basés en Nouvelle-Aquitaine/Pays basque, donc probablement pas NAEH — NewFund 2 déduit par calendrier.",
     "cfnews.net ; frenchweb.fr ; newfundcap.com (portfolio, titre indexé)"),
    ("NewFund Capital", "newfund-capital_newfund-2", "NewFund 2", "Olino (ex-Riskee)", "InsurTech (cœur)", "France",
     "Seed", 2022.0, 2.2, "€", "Investisseur principal/lead (avec Astorya.VC, M Capital)",
     "Active (site en ligne 2026)", "Moyen",
     "Assurance embarquée B2B (indépendants, PME, plateformes SaaS/fintech/legaltech). Co-investissement notable "
     "avec Astorya.VC (spécialiste insurtech). Véhicule exact non nommé, NewFund 2 déduit par calendrier.",
     "maddyness.com ; olino.fr/blog ; frenchweb.fr"),
    ("360 Capital Partners", "360-capital-partners_360-square-ii-seed-fund", "360 Square II Seed Fund", "Korint",
     "InsurTech (cœur)", "France", "Amorçage", 2022.0, 1.3, "€",
     "Investisseur participant (2e tour mené par Ventech)", "Actif", "Moyen",
     "Plateforme permettant aux assureurs/insurtechs/courtiers grossistes de déployer rapidement des produits "
     "digitaux. Extension 2024-09-06 (5M€, total cumulé 7M€ depuis fin 2022, mené par Ventech). Véhicule exact "
     "non nommé (possible aussi 360 fund V) — Square II déduit par calibre du ticket.",
     "CFNews (« Korint assure son amorçage », « Un deuxième VC sécurise Korint ») ; lemondedudroit.fr ; korint.io/ressources"),
    ("Alven Capital", "alven-capital_alven-v", "Alven V", "TheGuarantors", "FinTech liée assurance", "US",
     "Série A", 2017.0, 11.7, "$", "Co-lead (avec White Star Capital), puis participant en Série B (2019) et C (2022)",
     "Actif (Série C 2022, 50M$)", "Moyen",
     "Assurance-caution locative / remplacement dépôt de garantie (lease guarantee, security deposit insurance), "
     "New York. Confirmé par la page portefeuille officielle Alven (qualifie explicitement « InsurTech »). "
     "Rattachement Alven V (vs IV) déduit par calendrier (closing ~2016-2017), non cité nommément.",
     "alven.co (portfolio TheGuarantors) ; alleywatch.com ; White Star Capital ; Crunchbase"),
]


def convertir(montant, devise):
    if montant is None:
        return None
    taux = TAUX_CHANGE.get(devise, 1.0)
    return round(float(montant) * taux, 2)


def main() -> int:
    fonds = pd.read_excel(ENTREE, sheet_name="Fonds")
    part = pd.read_excel(ENTREE, sheet_name="Participations").copy()
    tours = pd.read_excel(ENTREE, sheet_name="Tours_de_table")
    dico = pd.read_excel(ENTREE, sheet_name="Dictionnaire")
    ano = pd.read_excel(ENTREE, sheet_name="Anomalies")

    fonds_idx = fonds.set_index("Fonds_ID")
    existants = {(r["Fonds_ID"], r["Start-up"]) for _, r in part.iterrows()}

    nouvelles_lignes = []
    for (nom_fonds, fonds_id, vehicule, startup, secteur, pays, tour, date, taille, devise,
         role, statut, confiance, commentaires, sources) in ROWS:
        if fonds_id not in fonds_idx.index:
            continue
        if (fonds_id, startup) in existants:
            continue
        nouvelles_lignes.append({
            "Fonds_ID": fonds_id,
            "Nom du fonds": nom_fonds,
            "Véhicule": vehicule,
            "Start-up": startup,
            "Secteur": secteur,
            "Stage financé": tour,
            "Montant investi par le fonds (M€)": None,
            "Total levé par la start-up (M€)": None,
            "Date de création": None,
            "Date d’investissement": date,
            "Pays d’origine": pays,
            "Nb pays d’implantation": None,
            "Positionnement (Leader/Minoritaire)": "Leader" if "lead" in role.lower() or "principal" in role.lower() else "Minoritaire",
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
            "Taille du tour (M€)": convertir(taille, devise),
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
