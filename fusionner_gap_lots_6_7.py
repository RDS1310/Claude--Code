#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fusionner_gap_lots_6_7.py — Intègre les Gap-Lots 6 et 7
====================================================================================

Complète 4 véhicules dont le "Nb InsurTech" déclaré dépassait largement le
nombre de participations documentées : Truffle Capital Fund II (+3),
Serena III (+1), Start Venture I (+7), Ethias Ventures (+2). Plusieurs
sociétés trouvées dans le portefeuille d'Ethias Ventures ont été écartées
car hors périmètre thématique (mobilité, santé, énergie sans lien assurance
direct) — journalisé en anomalie.

Entrée  : VC_Database_Standardisee_v12.xlsx
Sortie  : VC_Database_Standardisee_v13.xlsx

Usage : python3 fusionner_gap_lots_6_7.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v12.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v13.xlsx"

TAUX = {"€": 1.0, "$": 0.92, "£": 1.17}


def conv(montant, devise):
    if montant is None:
        return None
    return round(float(montant) * TAUX.get(devise, 1.0), 2)


ROWS = [
    # nom_fonds, fonds_id, vehicule, startup, secteur, pays, tour, date, taille, devise,
    # role, statut, confiance, commentaires, sources
    ("Truffle Capital", "truffle-capital_truffle-fintech-et-insurtech-fund-ii", "Truffle FinTech & InsurTech Fund II",
     "Cachet", "InsurTech (cœur)", "Estonie", "Série A", 2022.0, 5.5, "€",
     "Chef de file (lead)", "Actif", "Moyen",
     "Marketplace d'assurance pour travailleurs de plateformes (gig economy), couverture à l'heure/à l'usage. "
     "Avec Uniqa Ventures et Icebreaker.vc (investisseur historique). Tour antérieur 1,1M€ (seed) sans Truffle.",
     "truffle.com/portfolio-truffle-capital/cachet ; arcticstartup.com ; en.ain.ua ; fintech.global"),
    ("Truffle Capital", "truffle-capital_truffle-fintech-et-insurtech-fund-ii", "Truffle FinTech & InsurTech Fund II",
     "WeGroup", "InsurTech (adjacent)", "Belgique", "Série A", 2020.0, 3.0, "€",
     "Co-investisseur (avec Seeder Fund)", "Actif", "Moyen-Élevé",
     "SaaS d'aide à la vente et d'analyse de besoins/risques connectant assureurs et clients digitaux. Autres "
     "investisseurs : InsurTech Hub Munich, Start It X.",
     "truffle.com/portfolio-truffle-capital/wegroup ; medium.com/wegroup ; crunchbase.com"),
    ("Truffle Capital", "truffle-capital_truffle-fintech-et-insurtech-fund-ii", "Truffle FinTech & InsurTech Fund II",
     "Particeep", "InsurTech (adjacent)", "France", "Prise de participation minoritaire", 2021.0, 2.0, "€",
     "Co-investisseur (avec Sopra Steria/Sopra Banking Software)", "Sorti — cédé à Kereis (courtier), déc. 2022", "Élevé",
     "Plateforme SaaS de commercialisation digitale de produits bancaires/assurance/investissement en marque "
     "blanche. Sert explicitement Groupama comme client, en plus de BNP Paribas, Crédit Agricole, Arkéa, Nexity.",
     "truffle.com (annonce + fiche portefeuille) ; planet-fintech.com (cession Kereis) ; blog.particeep.com"),
    ("Serena Capital", "serena-capital_serena-iii", "Serena III",
     "Goodvest", "FinTech liée assurance", "France", "Série B", 2025.0, 12.0, "€",
     "Co-investisseur (avec Ring Capital, Polytechnique Ventures, AG2R La Mondiale, business angels)", "Actif", "Faible-Moyen",
     "Assurance-vie/PER responsable (contrat GOODVIE). Serena Capital a mené ce tour, investisseur distinct de "
     "Ring Capital/Mission I (déjà documenté pour ce même tour). Vintage incertain : tour daté juste avant le 1er "
     "closing de Serena IV (oct. 2025, 200M€) — pourrait relever de Serena III en fin de déploiement ou d'un pont "
     "vers Serena IV, à vérifier avant confirmation définitive.",
     "maddyness.com (24/09/2025) ; comparateurbanque.com ; jaimelesstartups.fr"),
    ("Start Venture", "start-venture_start-venture-i", "Start Venture I",
     "WeatherPromise", "InsurTech (cœur)", "US", "Série A", 2026.0, 12.8, "$",
     "Co-investisseur (lead Maveron, avec 1Sharpe, Lerer Hippeau, Clocktower, Commerce Ventures, MS Transverse, 1Flourish)",
     "Actif", "Élevé",
     "Garanties météo paramétriques pour voyages/événements. Confirmé par un billet de blog du fonds lui-même.",
     "startventures.vc ; theinsurer.com ; pulse2.com ; prnewswire.com"),
    ("Start Venture", "start-venture_start-venture-i", "Start Venture I",
     "ClaimSorted", "InsurTech (cœur)", "UK", "Seed", 2025.0, 13.3, "$",
     "Co-investisseur (lead Atomico, avec Eurazeo, Y Combinator, firstminute capital)", "Actif", "Élevé",
     "Plateforme IA de gestion de sinistres (TPA nouvelle génération). Nommé explicitement « Start Ventures "
     "Capital » dans plusieurs communiqués.",
     "fintech.global ; eu-startups.com ; ibsintelligence.com ; forbes.com"),
    ("Start Venture", "start-venture_start-venture-i", "Start Venture I",
     "Anansi", "InsurTech (cœur)", "UK", "Seed", 2021.0, 1.5, "£",
     "Co-investisseur (lead Octopus Ventures, avec UNIQA, Sie Ventures)", "Actif au moment de l'annonce", "Élevé",
     "Assurance embarquée « goods-in-transit » pour e-commerçants/3PL.",
     "startventures.vc ; maddyness.com (UK) ; business-money.com ; sieventures.medium.com"),
    ("Start Venture", "start-venture_start-venture-i", "Start Venture I",
     "Claims Carbon (Claims Carbon Institute)", "InsurTech (adjacent)", "Suède", "Seed", 2022.0, 1.04, "$",
     "Co-investisseur (lead Vaens, avec astorya.vc, Claesson & Anderzén)", "Actif au moment de l'annonce", "Moyen-Élevé",
     "Données carbone/climat pour la chaîne de valeur des assureurs (SaaS décarbonation). « BiG Start Ventures » "
     "cité nommément.",
     "fintech.global ; insurance-edge.net ; claimscarbon.com"),
    ("Start Venture", "start-venture_start-venture-i", "Start Venture I",
     "Coverflex", "InsurTech (adjacent)", "Portugal", "Série A", 2023.0, 15.0, "€",
     "Co-investisseur (lead SCOR Ventures, avec Breega, Armilar, Stableton, MS&AD, Shilling)", "Actif", "Élevé",
     "Avantages salariés incluant assurance santé/garanties. Investisseur distinct de Breega Capital/Breega "
     "Venture III (déjà documenté pour cette même société) — co-investissement légitime.",
     "techcrunch.com ; eu-startups.com ; fintech.global ; coverflex.com"),
    ("Start Venture", "start-venture_start-venture-i", "Start Venture I",
     "Drivit", "InsurTech (adjacent)", "Portugal", "Exit", 2020.0, None, None,
     "Actionnaire cédant", "Sorti — racheté par Zego (UK)", "Élevé",
     "Télématique/UBI (usage-based insurance) — données de conduite vendues aux assureurs auto. Confirmé par "
     "communiqué juridique du cabinet ayant conseillé la vente.",
     "abreuadvogados.com ; drivit.com"),
    ("Start Venture", "start-venture_start-venture-i", "Start Venture I",
     "Quantee", "InsurTech (cœur)", "Pologne", "Seed", 2022.0, 0.7, "$",
     "Investisseur (rôle exact/tour précis non confirmé)", "Sorti — racheté par Guidewire (avril 2025)", "Faible",
     "Moteur de tarification IA (pricing dynamique) pour assureurs. Listé comme investisseur par un agrégateur "
     "(CBInsights), aucune source primaire ne confirme le tour exact — à vérifier.",
     "techstartups.com ; cbinsights.com (agrégateur non revérifié)"),
    ("Ethias Ventures", "ethias-ventures_ethias-ventures", "Ethias Ventures",
     "Ai5", "InsurTech (adjacent)", "Belgique", "Investissement direct (corporate)", 2026.0, 1.5, "€",
     "Investisseur, en synergie avec NRB (filiale IT du groupe)", "Actif", "Élevé",
     "IA agentique B2B ; premier cas d'usage = rédaction automatisée des rapports d'inspection des préventeurs Ethias.",
     "newsroom.ethias.be ; datanews.levif.be ; lalibre.be"),
    ("Ethias Ventures", "ethias-ventures_ethias-ventures", "Ethias Ventures",
     "Sourse (ex-Stratos Solution)", "InsurTech (adjacent)", "Belgique", "Levée", 2025.0, 3.65, "€",
     "Co-investisseur (avec Namur Invest et investisseurs historiques)", "Actif", "Élevé",
     "IA + observation terrestre (satellite) pour détecter/prévenir risques catastrophes naturelles et "
     "infrastructures critiques.",
     "dhnet.be"),
]


def main() -> int:
    fonds = pd.read_excel(ENTREE, sheet_name="Fonds")
    part = pd.read_excel(ENTREE, sheet_name="Participations").copy()
    tours = pd.read_excel(ENTREE, sheet_name="Tours_de_table")
    dico = pd.read_excel(ENTREE, sheet_name="Dictionnaire")
    ano = pd.read_excel(ENTREE, sheet_name="Anomalies")

    fonds_idx = set(fonds["Fonds_ID"])
    existants = {(r["Fonds_ID"], r["Start-up"]) for _, r in part.iterrows()}

    nouvelles_lignes = []
    for (nom_fonds, fonds_id, vehicule, startup, secteur, pays, tour, date, taille, devise,
         role, statut, confiance, commentaires, sources) in ROWS:
        if fonds_id not in fonds_idx or (fonds_id, startup) in existants:
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
            "Positionnement (Leader/Minoritaire)": "Leader" if "lead" in role.lower() or "chef de file" in role.lower() else "Minoritaire",
            "Capital social (k€)": None,
            "CA (M€)": None,
            "Valorisation (M€)": None,
            "Nb fonds investisseurs": None,
            "Nb tours (fonds)": None,
            "Nb tours (total)": None,
            "Repositionnement": None,
            "Exit (O/N)": "O" if any(k in statut.lower() for k in ("sorti", "cédé", "racheté", "exit")) else "N",
            "MoC exit": None,
            "MoEP exit": None,
            "Taille du tour (M€)": conv(taille, devise),
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
        "Table ou bloc source": "Gap-Lot 6 (recherche Ethias Ventures, 17/09/2026)",
        "Ligne source": None,
        "Entité concernée": "Ethias Ventures",
        "Champ concerné": "Start-up",
        "Valeur source": "moveUP, Cascador Health (santé), Aidoptation (mobilité/défense), BattMobility, Ethias "
                          "Lease, Sparki (mobilité/énergie VE), June Energy (énergie) — 7 sociétés confirmées dans "
                          "le portefeuille du fonds",
        "Valeur retenue": "Non intégrées (aucun lien assurance direct/principal)",
        "Candidats éventuels": None,
        "Score de similarité": None,
        "Niveau de confiance": "Élevé",
        "Traitement appliqué": "Écartées du radar : la thèse réelle d'Ethias Ventures couvre « InsurTech + "
                                "écosystèmes Mobilité/Santé/Habitation liés à l'assurance », pas l'InsurTech au "
                                "sens strict des catégories du radar. Seules Ai5 et Sourse (déjà intégrées) "
                                "correspondent aux catégories cœur/adjacent/FinTech liée assurance.",
        "Commentaire": "Ces 7 sociétés + Linkbycar + Ai5 + Sourse totalisent exactement les 10 participations "
                       "déclarées par le fonds — recherche exhaustive, pas un gap restant.",
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
