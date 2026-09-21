#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fusionner_cvc_externes.py — Ajoute 6 nouveaux véhicules CVC assurantiels (CVC-Lots 1-3)
========================================================================================

Intègre les résultats de recherche (WebSearch, méthodologie anti-fabrication identique
aux lots précédents) sur : Allianz X, MAIF Avenir, Munich Re Ventures (+HSB Ventures),
MS&AD Ventures, Sompo (Digital Lab/Light Vortex), Generali Ventures (fonds de fonds,
0 participation directe sourcée). Reclasse aussi Atlantic Vantage Point (ex-AXA Venture
Partners) : reste "VC indépendant" (MBO août 2024, AXA désormais LP ancre minoritaire,
plus un CVC captif), avec la filiation AXA documentée en commentaire — aucune nouvelle
participation nécessaire (Policygenius et Idelic déjà en base sous ce véhicule).

Entrée  : VC_Database_Standardisee_v15.xlsx
Sortie  : VC_Database_Standardisee_v16.xlsx

Usage : python3 fusionner_cvc_externes.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v15.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v16.xlsx"

TAUX = {"€": 1.0, "$": 0.92, "£": 1.17, "AUD": 0.60}


def conv(montant, devise):
    if montant is None:
        return None
    return round(montant * TAUX.get(devise, 1.0), 2)


FONDS_NOUVEAUX = [
    {
        "Nom du fonds": "Allianz X", "Véhicule": "Allianz X", "Fonds_ID": "allianz-x_allianz-x",
        "AuM (M€)": None, "Statut": "Actif", "Phase": "Gestion", "Millésime": 2013.0,
        "Géographie": "International", "Stratégie": "InsurTech ; FinTech ; Croissance",
        "Stages pratiqués": "Série B ; Série C ; Série D ; Croissance",
        "Pré-Seed": 0.0, "Seed": 0.0, "Pré-Série A": 0.0, "Série A": 0.0,
        "Série B": 1.0, "Série C": 1.0, "Série D": 1.0,
        "Nb participations (déclaré)": 25.0, "Nb participations documentées": 11, "Nb exits": 2.0,
        "Nb InsurTech": 11.0, "Site web": "https://www.allianzx.com",
        "Source_AuM": (
            "[CVC-Lot 1, 21/09/2026] Société mère : Allianz SE. AuM déclaré évolutif selon les sources "
            "(1 Md€ en 2019 à 1,5-2 Md€ en 2025-2026, aucun chiffre unique officiel stable) — non retenu "
            "par prudence anti-fabrication. Sources : munich-startup.de, allianz.com (PDF officiel). "
            "Portefeuille total >25 sociétés (jusqu'à 12 licornes) tous secteurs, non ventilé insurtech "
            "par le fonds lui-même (Tracxn/PitchBook)."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "Allianz SE",
    },
    {
        "Nom du fonds": "MAIF Avenir", "Véhicule": "MAIF Avenir", "Fonds_ID": "maif-avenir_maif-avenir",
        "AuM (M€)": 250.0, "Statut": "Renommé (Ternel depuis 2024)", "Phase": "Gestion", "Millésime": 2015.0,
        "Géographie": "France",
        "Stratégie": "InsurTech ; FinTech ; Avenir du travail ; Santé ; Data ; Consommation responsable",
        "Stages pratiqués": "Seed ; Série A ; Série B",
        "Pré-Seed": 0.0, "Seed": 1.0, "Pré-Série A": 0.0, "Série A": 1.0,
        "Série B": 1.0, "Série C": 0.0, "Série D": 0.0,
        "Nb participations (déclaré)": 28.0, "Nb participations documentées": 1, "Nb exits": None,
        "Nb InsurTech": 1.0, "Site web": "https://maif-avenir.fr",
        "Source_AuM": (
            "[CVC-Lot 1, 21/09/2026] Société mère : MAIF. 250 M€ levés (125 M€ 2015 + 125 M€ 2017). "
            "28 participations actives tous secteurs déclarées (mind Fintech), une seule participation "
            "insurtech/fintech-assurance documentée en source ouverte (Lovys). En 2024, rapprochement avec "
            "Capital Croissance et renommage du véhicule en « Ternel » (MAIF reste partenaire/actionnaire) — "
            "peut expliquer la faible couverture presse récente sous le nom « MAIF Avenir ». "
            "Sources : entreprise.maif.fr, Journal du Net, Argus de l'Assurance, Carenews."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "MAIF",
    },
    {
        "Nom du fonds": "Munich Re Ventures", "Véhicule": "Munich Re Ventures (incl. HSB Ventures)",
        "Fonds_ID": "munich-re-ventures_munich-re-ventures",
        "AuM (M€)": 1104.0, "Statut": "Actif", "Phase": "Gestion", "Millésime": 2014.0,
        "Géographie": "International",
        "Stratégie": "InsurTech ; Climate ; Cyber ; Risque industriel (HSB)",
        "Stages pratiqués": "Seed ; Série A ; Série B ; Série C ; Série D",
        "Pré-Seed": 0.0, "Seed": 1.0, "Pré-Série A": 0.0, "Série A": 1.0,
        "Série B": 1.0, "Série C": 1.0, "Série D": 1.0,
        "Nb participations (déclaré)": 36.0, "Nb participations documentées": 5, "Nb exits": 3.0,
        "Nb InsurTech": 5.0, "Site web": "https://www.munichre.com/mrv",
        "Source_AuM": (
            "[CVC-Lot 2, 21/09/2026] Société mère : Munich Re. AuM ~1,2 Md$ confirmé mai 2025 "
            "(PRNewswire, clôture « HSB Fund II » 125 M$ pour le 10e anniversaire du véhicule) → conv. 1104 M€. "
            "Nombre de participations total très divergent selon sources (36 à 123 selon Artemis.bm/vcsheet/"
            "CBInsights, non réconcilié) — 36 retenu comme la valeur la plus prudente. Combine les fonds "
            "génériques Munich Re Ventures (insurtech/climate/cyber) et les fonds dédiés HSB (financés par la "
            "filiale Hartford Steam Boiler, orientés risque industriel « built world »)."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "Munich Re",
    },
    {
        "Nom du fonds": "MS&AD Ventures", "Véhicule": "MS&AD Ventures", "Fonds_ID": "msad-ventures_msad-ventures",
        "AuM (M€)": 368.0, "Statut": "Actif", "Phase": "Gestion", "Millésime": 2018.0,
        "Géographie": "International",
        "Stratégie": "InsurTech ; FinTech ; IA/Analytics ; Santé digitale ; Durabilité",
        "Stages pratiqués": "Seed ; Série A ; Série B ; Croissance",
        "Pré-Seed": 0.0, "Seed": 1.0, "Pré-Série A": 0.0, "Série A": 1.0,
        "Série B": 1.0, "Série C": 0.0, "Série D": 0.0,
        "Nb participations (déclaré)": 190.0, "Nb participations documentées": 9, "Nb exits": 2.0,
        "Nb InsurTech": 9.0, "Site web": "https://www.msad.vc",
        "Source_AuM": (
            "[CVC-Lot 2, 21/09/2026] Société mère : MS&AD Insurance Group Holdings (Mitsui Sumitomo + Aioi "
            "Nissay Dowa). Capacité d'investissement portée à 400 M$ en 2024 (contre 200 M$ en 2021) → conv. "
            "368 M€ (ms-ad-hd.com, communiqué juillet 2021 + màj 2024). « 190 investissements » cité par une "
            "source secondaire (Altss), non confirmé par une source primaire MS&AD — confiance faible sur ce "
            "chiffre, retenu comme déclaratif large (tous secteurs, y compris tours de suivi)."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "MS&AD Insurance Group Holdings",
    },
    {
        "Nom du fonds": "Sompo", "Véhicule": "Sompo Digital Lab / Sompo Light Vortex",
        "Fonds_ID": "sompo_digital-lab-light-vortex",
        "AuM (M€)": None, "Statut": "Actif", "Phase": "Investissement", "Millésime": None,
        "Géographie": "International",
        "Stratégie": "InsurTech ; Innovation stratégique",
        "Stages pratiqués": "Seed ; Série A ; Série B ; Série C",
        "Pré-Seed": 0.0, "Seed": 1.0, "Pré-Série A": 0.0, "Série A": 1.0,
        "Série B": 1.0, "Série C": 1.0, "Série D": 0.0,
        "Nb participations (déclaré)": None, "Nb participations documentées": 5, "Nb exits": 1.0,
        "Nb InsurTech": 5.0, "Site web": "https://www.sompo.io/en",
        "Source_AuM": (
            "[CVC-Lot 2, 21/09/2026] Société mère : Sompo Holdings. Pas d'AuM de fonds VC classique publié — "
            "seul chiffre trouvé : >20 M$ investis cumulativement dans des startups israéliennes (jetro.go.jp). "
            "Digital Lab actif depuis ~2016 (hub Tel Aviv ouvert 2018) ; Light Vortex est une filiale annoncée "
            "séparément (année précise non confirmée). Attention méthodologique : plusieurs investissements "
            "(Cover Genius, Zego, Trov) sont attribués dans la presse à « Sompo Holdings » sans que les sources "
            "précisent s'il s'agit de Digital Lab, Light Vortex ou d'un investissement corporate direct — "
            "confiance Moyen sur l'entité exacte, confiance Élevé sur le fait de l'investissement. Sompo a par "
            "ailleurs créé en 2024 un véhicule distinct « SOMPO Growth Partners » (Tokyo, portefeuille FLUX/"
            "Honest/Kin), hors périmètre de cette recherche (mandat Digital Lab/Light Vortex uniquement)."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "Sompo Holdings",
    },
    {
        "Nom du fonds": "Generali Ventures", "Véhicule": "Generali Ventures",
        "Fonds_ID": "generali-ventures_generali-ventures",
        "AuM (M€)": 250.0, "Statut": "Actif", "Phase": "Investissement", "Millésime": 2022.0,
        "Géographie": "Europe ; Amérique",
        "Stratégie": "Fonds de fonds VC (InsurTech/FinTech/SaaS)",
        "Stages pratiqués": None,
        "Pré-Seed": 0.0, "Seed": 0.0, "Pré-Série A": 0.0, "Série A": 0.0,
        "Série B": 0.0, "Série C": 0.0, "Série D": 0.0,
        "Nb participations (déclaré)": 5.0, "Nb participations documentées": 0, "Nb exits": None,
        "Nb InsurTech": 0.0, "Site web": "https://www.generali.com",
        "Source_AuM": (
            "[CVC-Lot 3, 21/09/2026] Société mère : Assicurazioni Generali. Engagement dédié de 250 M€ "
            "annoncé (plan stratégique « Lifetime Partner 24 »), à déployer sur des fonds VC jusqu'à fin 2024. "
            "⚠️ Constat structurel : Generali Ventures n'est PAS un CVC à investissements directs en startups, "
            "mais un véhicule de type fonds de fonds (LP) — il a investi comme commanditaire dans 5 fonds VC "
            "tiers : Mundi Ventures (spécialiste insurtech), Speedinvest (déjà suivi séparément dans cette "
            "base), Dawn Capital, Headline, StepStone Group. Aucune participation directe en startup insurtech/"
            "fintech-assurance sourcée au niveau du véhicule lui-même (les startups du portefeuille Mundi "
            "Ventures, par ex., sont des participations de Mundi, pas de Generali Ventures — non incluses ici "
            "pour éviter toute extrapolation non sourcée). À distinguer de « The Generali Innovation Fund », "
            "programme interne d'intrapreneuriat, pas un CVC externe. Sources : Sifted, Generali.com, "
            "Coverager, Reinsurance News, Global Venturing."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "Assicurazioni Generali",
    },
    {
        "Nom du fonds": "Aviva Ventures", "Véhicule": "Aviva Ventures", "Fonds_ID": "aviva-ventures_aviva-ventures",
        "AuM (M€)": 292.5, "Statut": "Actif", "Phase": "Investissement", "Millésime": 2015.0,
        "Géographie": "Royaume-Uni ; International",
        "Stratégie": "InsurTech ; IoT prévention ; Croissance",
        "Stages pratiqués": "Seed ; Série A ; Série B",
        "Pré-Seed": 0.0, "Seed": 1.0, "Pré-Série A": 0.0, "Série A": 1.0,
        "Série B": 1.0, "Série C": 0.0, "Série D": 0.0,
        "Nb participations (déclaré)": 40.0, "Nb participations documentées": 8, "Nb exits": 3.0,
        "Nb InsurTech": 8.0, "Site web": "https://www.aviva.com",
        "Source_AuM": (
            "[CVC-Lot 3, 21/09/2026] Société mère : Aviva plc. Lancé déc. 2015, opérationnel 2016 avec "
            "~£100m (déploiement visé £20m/an sur 5 ans), complété par £150m en 2023 — soit ~£250m cumulés "
            "→ conv. 292,5 M€ (aviva.com, fintech.global avril 2023). Co-opéré avec Founders Factory. "
            "30 à 47 investissements selon les bases tierces (PitchBook/Crunchbase/CBInsights, chiffres non "
            "réconciliés) tous secteurs. À noter : un véhicule distinct « Aviva Investors Venture & Growth "
            "Capital LTAF » (2025, ~£150m, modèle fonds de fonds) existe séparément — non confondu avec ce "
            "véhicule, lien entre les deux structures non clarifié par les sources."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "Aviva plc",
    },
]

# (fonds_id, nom_fonds, vehicule, startup, secteur, pays, stage, date, taille_tour_devise, devise,
#  role, statut, confiance, commentaires, sources)
PARTICIPATIONS_NOUVELLES = [
    # --- Allianz X ---
    ("allianz-x_allianz-x", "Allianz X", "Allianz X", "Pie Insurance", "InsurTech (cœur)", "États-Unis",
     "Série C", 2021.0, 118, "$", "Co-lead (avec Acrew Capital ; Greycroft, SVB Capital, SiriusPoint, Elefund)",
     "Actif", "Élevé",
     "Assurance workers' compensation pour PME, distribution digitale ; 1er investissement d'Allianz X dans une "
     "insurtech B2B nord-américaine. Tour ultérieur de $315M co-lead Allianz X (date précise non confirmée).",
     "allianz.com (25/03/2021), Crowdfund Insider, Medium Allianz X ($315M)"),
    ("allianz-x_allianz-x", "Allianz X", "Allianz X", "Coalition, Inc.", "InsurTech (cœur)", "États-Unis",
     "Série F", 2022.0, 250, "$", "Co-lead (avec Valor Equity Partners et Kinetic Partners)", "Actif", "Élevé",
     "Cyber-assurance combinée à des outils de cybersécurité proactive, valorisation $5Md, >160 000 clients.",
     "allianz.com (08/07/2022), GlobeNewswire"),
    ("allianz-x_allianz-x", "Allianz X", "Allianz X", "NEXT Insurance", "InsurTech (cœur)", "États-Unis",
     "Stratégique", 2023.0, 265, "$", "Co-investisseur avec Allstate", "Actif", "Élevé",
     "Assurance PME/travailleurs indépendants 100% digitale. Ouvre un partenariat stratégique de réassurance "
     "élargi avec Allianz Re.", "allianz.com (02/11/2023), CNBC"),
    ("allianz-x_allianz-x", "Allianz X", "Allianz X", "Coterie Insurance", "InsurTech (cœur)", "États-Unis",
     "Série C", 2025.0, None, "$", "Investissement stratégique en actions (lead)", "Actif", "Moyen",
     "MGA insurtech fournissant des produits P&C aux PME américaines. Montant du tour non divulgué "
     "publiquement.", "Medium Allianz X, Coterie newsroom, Freshfields (09/2025)"),
    ("allianz-x_allianz-x", "Allianz X", "Allianz X", "simplesurance", "InsurTech (cœur)", "Allemagne",
     "Acquisition 100%", 2022.0, None, "$", "Acquéreur (rachat total)", "Filiale intégrée (Allianz X)", "Élevé",
     "Pionnier allemand de l'assurance embarquée (e-commerce/mobilité/travel/fintech), partenaire N26/heycar/"
     "OnePlus. Synergies approfondies avec Allianz Partners.",
     "Medium Allianz X, Insurance Journal, Insurance Insider"),
    ("allianz-x_allianz-x", "Allianz X", "Allianz X", "Cambridge Mobile Telematics", "InsurTech (adjacent)",
     "États-Unis", "Stratégique", 2026.0, 350, "$", "Co-lead avec TPG et State Farm", "Actif", "Élevé",
     "Plateforme mondiale de télématique/sécurité routière alimentée par IA, utilisée par de nombreux assureurs "
     "auto pour la tarification et prévention du risque.", "TPG.com, BusinessWire (24/03/2026), Coverager"),
    ("allianz-x_allianz-x", "Allianz X", "Allianz X", "Ualá", "FinTech liée assurance", "Argentine",
     "Série E", 2026.0, 195, "$", "Lead (avec Stone Ridge, Tencent, TABLE Holdings, Soros Fund Management, "
     "D1 Capital)", "Actif", "Moyen",
     "Néobanque latino-américaine, lancement 2026 de produits d'assurance-vie/accident 100% digitaux embarqués "
     "(>300 000 devis en quelques semaines). ⚠️ Montant divergent entre sources Allianz X : $195M (fintech."
     "global, LatamList) vs $300M (Medium Allianz X) — probable confusion entre 2 tours, à vérifier.",
     "fintech.global (05/03/2026), LatamList, Medium Allianz X"),
    ("allianz-x_allianz-x", "Allianz X", "Allianz X", "Lemonade", "InsurTech (cœur)", "États-Unis",
     "Série C", 2017.0, 120, "$", "Co-investisseur", "Exit", "Moyen",
     "Assureur habitation/locataire propulsé par IA, licorne cotée NYSE. Entrée Series B (2017), suivi Series C "
     "(2017, $120M total) et D (2019). Allianz X et Grays Peak Capital sont depuis sortis du capital "
     "(date de sortie non trouvée).", "Insurance Journal (11/04/2019), Mergr, allianzx.com"),
    ("allianz-x_allianz-x", "Allianz X", "Allianz X", "Openly", "InsurTech (cœur)", "États-Unis",
     "Croissance", 2024.0, 193, "$", "Co-lead avec Eden Global Partners", "Actif", "Élevé",
     "Assureur habitation premium, distribution 100% via courtiers indépendants. $70M de dette antérieure "
     "d'Allianz X ; nouveau tour de croissance annoncé le 22/04/2026 (montant non précisé).",
     "Medium Allianz X (30/01/2024), Yahoo Finance, citybiz"),
    ("allianz-x_allianz-x", "Allianz X", "Allianz X", "Nauto", "InsurTech (adjacent)", "États-Unis",
     "Série A", 2016.0, None, "$", "Co-investisseur (avec Greylock Partners, Playground Global)", "Actif",
     "Moyen",
     "IA embarquée pour flottes (sécurité conducteur), données utilisées pour la tarification/prévention du "
     "risque assurantiel. Total levé par Nauto : $215M (dont Series C juin 2023, montant non divulgué).",
     "Tracxn Nauto, allianzx.com/ourcompanies/nauto"),
    ("allianz-x_allianz-x", "Allianz X", "Allianz X", "Innovation Group", "InsurTech (cœur)", "Royaume-Uni",
     "Acquisition 100%", 2022.0, None, "$", "Acquéreur (rachat total)", "Filiale intégrée (Allianz X)", "Élevé",
     "Plateforme SaaS « Gateway » de gestion digitale du parcours sinistre (FNOL → réparation → règlement) "
     "pour >1200 clients assurance/automobile dans le monde. Management conservé.",
     "InsurTech Digital (10/10/2022), Insurance Business, Medium Allianz X"),
    # --- MAIF Avenir ---
    ("maif-avenir_maif-avenir", "MAIF Avenir", "MAIF Avenir", "Lovys", "InsurTech (cœur)", "France",
     "Seed", 2019.0, 3.3, "€", "Co-investisseur historique (Portugal Ventures, Plug and Play en 2019 ; "
     "Heartcore, NewAlpha, Raise Ventures, Bpifrance en Série A)", "Actif", "Élevé",
     "Assurtech française fondée en 2017 par João Cardoso, assurance habitation/auto 100% digitale et "
     "résiliable à tout moment (« assurance par abonnement »). Suivi en Série A (19/01/2021, 17 M€).",
     "entreprise.maif.fr, Maddyness (03/07/2019), Argus de l'Assurance, Alliancy (Série A 2021)"),
    # --- Munich Re Ventures ---
    ("munich-re-ventures_munich-re-ventures", "Munich Re Ventures", "Munich Re Ventures (incl. HSB Ventures)",
     "NEXT Insurance", "InsurTech (cœur)", "États-Unis", "Série A", 2017.0, 29, "$",
     "Lead (avec Markel, Nationwide)", "Exit", "Élevé",
     "Assurance PME digitale, courtage/MGA. Participation à la Series C en 2018. Cas d'école : rachat à 100% "
     "par ERGO (filiale Munich Re) pour $2,6Md, annoncé le 19/03/2025 — plus grosse acquisition insurtech P&C "
     "à ce jour. Montant du tour divergent : $29M (Carrier Management/Insurance Journal) vs $35M (Crunchbase).",
     "munichre.com/mrv/en/portfolio/next.html, fintech.global (25/03/2025), Insurance Journal (03/05/2017), "
     "Carrier Management (04/05/2017)"),
    ("munich-re-ventures_munich-re-ventures", "Munich Re Ventures", "Munich Re Ventures (incl. HSB Ventures)",
     "At-Bay", "InsurTech (cœur)", "États-Unis", "Série B", 2020.0, 34, "$",
     "Co-lead via fonds HSB (avec Acrew Capital)", "Exit", "Élevé",
     "Cyber-assurance (MGA) et cybersécurité proactive pour PME. Rachat annoncé par Munich Re Group pour $575M "
     "le 19/08/2026 (clôture prévue T1 2027) — ~2x les primes brutes émises ($278M GWP), en retrait par "
     "rapport à la valorisation post-money $1,35Md de la Série D 2021. ⚠️ Événement M&A potentiellement à "
     "ajouter à l'Étude de Marché M&A.",
     "theinsurer.com (24/02/2020), Insurance Business Mag (19/08/2026), munichre.com (19/08/2026)"),
    ("munich-re-ventures_munich-re-ventures", "Munich Re Ventures", "Munich Re Ventures (incl. HSB Ventures)",
     "Trov", "InsurTech (cœur)", "États-Unis", "Série D", 2017.0, 45, "$",
     "Lead (Sompo Holdings co-investisseur du même tour — voir aussi Sompo)", "Exit partiel", "Élevé",
     "Assurance à la demande (« on-demand insurance »), opérations UK/Australie. Équipe et technologie "
     "rachetées par Travelers en 2019.", "businessinsurance.com, reinsurancene.ws, intelligentinsurer.com"),
    ("munich-re-ventures_munich-re-ventures", "Munich Re Ventures", "Munich Re Ventures (incl. HSB Ventures)",
     "Cape Analytics", "InsurTech (adjacent)", "États-Unis", "Série C", 2021.0, 44, "$",
     "Co-investisseur (lead : Pivot Investment Partners)", "Exit", "Moyen",
     "Intelligence géospatiale/IA sur le risque immobilier pour assureurs (toiture, végétation, exposition "
     "catastrophes). Présent depuis le Seed/Série A. Rachetée par Moody's en 2025.",
     "Confluence.vc, Carrier Management (14/07/2021), ir.moodys.com (2025)"),
    ("munich-re-ventures_munich-re-ventures", "Munich Re Ventures", "Munich Re Ventures (incl. HSB Ventures)",
     "ShipIn Systems", "InsurTech (adjacent)", "Israël", None, 2022.0, None, "$",
     "Investisseur (« premier investissement maritime » de Munich Re Ventures)", "Actif", "Moyen",
     "Plateforme de gestion de flotte visuelle par IA, réduction du risque maritime/assurance marine.",
     "Coverager (04/2022), smartmaritimenetwork.com (08/04/2022), shipin.ai"),
    # --- MS&AD Ventures ---
    ("msad-ventures_msad-ventures", "MS&AD Ventures", "MS&AD Ventures", "Vouch", "InsurTech (cœur)",
     "États-Unis", None, None, None, "$", "Co-investisseur", "Actif", "Moyen",
     "Plateforme d'assurance commerciale digitale pour startups (cyber, RC, biens). Total levé par Vouch : "
     "$70M toutes sources confondues (Ribbit Capital, Y Combinator, Index Ventures).",
     "msad.vc/portfolio/vouch, vouch.us/venture/msad, reinsurancene.ws"),
    ("msad-ventures_msad-ventures", "MS&AD Ventures", "MS&AD Ventures", "Socotra", "InsurTech (cœur)",
     "États-Unis", "Série C", 2021.0, 50, "$", "Co-investisseur non-lead (lead : Insight Partners)", "Actif",
     "Élevé",
     "Plateforme cœur de gestion de polices (policy administration) pour assureurs. Tour initial de $5,2M "
     "(avec Nationwide, 8VC, date précise non trouvée).",
     "intelligentinsurer.com, ms-ad-hd.com/en/group/innovation/venture/socotra.html"),
    ("msad-ventures_msad-ventures", "MS&AD Ventures", "MS&AD Ventures", "INSHUR", "InsurTech (cœur)",
     "Royaume-Uni", None, 2024.0, 19, "$", "Co-investisseur (lead : Viola Growth)", "Actif", "Élevé",
     "Assurance auto commerciale digitale pour chauffeurs VTC/livreurs.",
     "reinsurancene.ws (16/10/2024), Coverager, inshurgroup.com"),
    ("msad-ventures_msad-ventures", "MS&AD Ventures", "MS&AD Ventures", "Fairmatic", "InsurTech (cœur)",
     "États-Unis", "Série B", None, 46, "$", "Co-investisseur (lead : Battery Ventures)", "Actif", "Moyen",
     "Assurance auto commerciale basée sur télématique/IA. Part précise de MS&AD dans le tour non trouvée.",
     "Insurance Business Mag, iireporter.com, vc-mapping.gilion.com"),
    ("msad-ventures_msad-ventures", "MS&AD Ventures", "MS&AD Ventures", "Hippo Insurance", "InsurTech (cœur)",
     "États-Unis", "Série E", 2020.0, 150, "$", "Co-investisseur", "Exit", "Élevé",
     "Assurance habitation US, devenue publique par SPAC en 2021. ⚠️ À distinguer d'un investissement séparé "
     "de $350M par Mitsui Sumitomo Insurance (filiale opérationnelle du groupe MS&AD, pas le fonds MS&AD "
     "Ventures), annoncé en novembre 2020 avec un traité de réassurance.",
     "iireporter.com, Insurance Journal (25/11/2020), ms-ad-hd.com (25/11/2020)"),
    ("msad-ventures_msad-ventures", "MS&AD Ventures", "MS&AD Ventures", "NEXT Insurance", "InsurTech (cœur)",
     "États-Unis", None, None, None, "$", "Non précisé", "Exit", "Moyen",
     "Listée comme participation du portefeuille officiel MS&AD Ventures (msad.vc) ; exit confirmé par rachat "
     "ERGO/Munich Re en 2025 (voir aussi Munich Re Ventures). Round MS&AD spécifique non identifié.",
     "msad.vc/portfolio/next-insurance"),
    ("msad-ventures_msad-ventures", "MS&AD Ventures", "MS&AD Ventures", "Agio Ratings", "InsurTech (adjacent)",
     "États-Unis", None, 2025.0, 6, "$", "Co-investisseur", "Actif", "Élevé",
     "Notation de risque de contrepartie en temps réel pour institutions financières et assureurs (actifs "
     "numériques). Total levé >$11M. Partenariat ultérieur (début 2025) avec l'assureur Relm Insurance pour un "
     "produit « crypto exchange default ».", "BusinessWire (06/10/2025), msad.vc/portfolio/agio-ratings, "
     "pulse2.com, Coverager"),
    ("msad-ventures_msad-ventures", "MS&AD Ventures", "MS&AD Ventures", "Assured Allies", "InsurTech (adjacent)",
     None, None, None, None, "$", "Non précisé", "Actif", "Faible",
     "Prévention/technologie de vieillissement à domicile liée à l'assurance dépendance/soins de longue durée. "
     "Présence confirmée sur le portefeuille officiel uniquement, aucun détail financier trouvé.",
     "msad.vc/portfolio/assured-allies"),
    ("msad-ventures_msad-ventures", "MS&AD Ventures", "MS&AD Ventures", "Jupiter Intelligence",
     "InsurTech (adjacent)", "États-Unis", None, None, None, "$", "Non précisé", "Actif", "Faible",
     "Intelligence de risque climatique (modélisation catastrophes, reporting TCFD) pour assureurs et "
     "entreprises. Présence confirmée sur le portefeuille officiel uniquement.",
     "Portefeuille officiel MS&AD (vertical sustainability/climate)"),
    # --- Sompo ---
    ("sompo_digital-lab-light-vortex", "Sompo", "Sompo Digital Lab / Sompo Light Vortex", "Chainproof",
     "InsurTech (cœur)", "Bermudes", "Seed/fondateur", 2022.0, None, "$",
     "Investisseur fondateur (« foundational partner »), avec Munich Re en soutien de capacité de "
     "réassurance", "Actif", "Élevé",
     "Premier assureur régulé de smart contracts (DeFi), licence IGB régulée par la Bermuda Monetary "
     "Authority.", "theinsurer.com, PRNewswire, hubbis.com"),
    ("sompo_digital-lab-light-vortex", "Sompo", "Sompo Digital Lab / Sompo Light Vortex", "Cover Genius",
     "FinTech liée assurance", "Australie", "Série C", 2021.0, None, "AUD",
     "Lead (Sompo Holdings Asia Pte Ltd)", "Actif", "Moyen",
     "Assurance embarquée (embedded insurance), opérations globales. ~68M AUD apportés par Sompo Holdings Asia "
     "sur un tour d'~100M AUD (≈42 M€ au taux approximatif). ⚠️ Entité exacte incertaine : sources nomment "
     "« Sompo Holdings (Asia) », pas explicitement Digital Lab ou Light Vortex.",
     "Bloomberg (28/09/2021), theinsurer.com, PYMNTS.com"),
    ("sompo_digital-lab-light-vortex", "Sompo", "Sompo Digital Lab / Sompo Light Vortex", "Zego",
     "InsurTech (cœur)", "Royaume-Uni", "Stratégique", 2026.0, 28, "$",
     "Co-investisseur stratégique (Sompo Holdings)", "Actif", "Moyen",
     "Assurance auto/flotte digitale basée télématique (UK). Partenariat stratégique pour développer "
     "l'assurance télématique au Japon. ⚠️ Entité exacte incertaine (Sompo Holdings, pas explicitement Digital "
     "Lab/Light Vortex).", "fintech.global (16/04/2026), zego.com, dealroom.co"),
    ("sompo_digital-lab-light-vortex", "Sompo", "Sompo Digital Lab / Sompo Light Vortex", "Trov",
     "InsurTech (cœur)", "États-Unis", "Série D", 2017.0, 45, "$",
     "Co-investisseur (tour mené par Munich Re/HSB Ventures — voir aussi Munich Re Ventures)", "Exit partiel",
     "Moyen",
     "Assurance à la demande. ⚠️ Entité exacte incertaine (Sompo Holdings, pas explicitement Digital Lab).",
     "reinsurancene.ws, intelligentinsurer.com"),
    ("sompo_digital-lab-light-vortex", "Sompo", "Sompo Digital Lab / Sompo Light Vortex", "Nexar",
     "InsurTech (adjacent)", "Israël", None, None, None, "$",
     "Investisseur (Sompo Digital Lab confirmé nommément)", "Actif", "Moyen",
     "Plateforme IA de dashcams générant des données de risque routier, avec produit dédié « insurance "
     "platform » pour assureurs auto. Inclus dans l'enveloppe globale >$20M investis dans des startups "
     "israéliennes du Digital Lab, sans détail du tour spécifique.",
     "jetro.go.jp, wfmz.com/prnewswire"),
    # --- Aviva Ventures ---
    ("aviva-ventures_aviva-ventures", "Aviva Ventures", "Aviva Ventures", "Cocoon", "InsurTech (adjacent)",
     "Royaume-Uni", None, 2016.0, None, "£", "Investisseur (1er deal du fonds)", "Actif (à la date source)",
     "Moyen",
     "Dispositif IoT de sécurité domestique (détection sonore de mouvement), lié à des offres d'assurance "
     "habitation (partenariat réduction prime avec PolicyCastle).",
     "aviva.com (déc. 2015/2016), cocoon.life/blog, Insurance Times"),
    ("aviva-ventures_aviva-ventures", "Aviva Ventures", "Aviva Ventures", "Neos", "InsurTech (cœur)",
     "Royaume-Uni", "Investissement initial", 2017.0, 5, "£",
     "Investisseur puis actionnaire majoritaire (nov. 2018)", "Exit", "Élevé",
     "Assurance habitation connectée (smart home insurance), capteurs + police d'assurance. Participation "
     "cédée à Sky UK (date de sortie divergente entre sources : 2021 vs 2022).",
     "aviva.com (nov. 2018), Insurance Times, FStech, Insurance Business Mag"),
    ("aviva-ventures_aviva-ventures", "Aviva Ventures", "Aviva Ventures", "Roost", "InsurTech (adjacent)",
     "États-Unis", "Série B", 2017.0, 10.4, "$",
     "Co-investisseur (avec Desjardins Insurance, Fosun RZ Capital)", "Actif (à la date source)", "Élevé",
     "Télématique domestique (capteurs de fuite d'eau etc.) utilisée en prévention par les assureurs "
     "partenaires. 1er investissement US d'Aviva Ventures.",
     "aviva.com (08/2017), Insurance Journal, Canadian Underwriter, Insurance Post"),
    ("aviva-ventures_aviva-ventures", "Aviva Ventures", "Aviva Ventures", "Savari", "InsurTech (adjacent)",
     "États-Unis", "Série B", 2018.0, 12, "$", "Lead investisseur (avec SAIC Capital, Flex)", "Exit", "Élevé",
     "Capteurs V2X pour véhicules autonomes, destinés à alimenter le développement de produits d'assurance "
     "auto autonome par Aviva. Acquise par Harman International.",
     "Global Venturing (04/2018), GlobeNewswire, insurtechanalyst.com"),
    ("aviva-ventures_aviva-ventures", "Aviva Ventures", "Aviva Ventures", "Acre",
     "InsurTech (cœur) / FinTech liée assurance", "Royaume-Uni", "Investissement initial", 2019.0, 5, "£",
     "Co-investisseur (avec Sesame Bankhall) puis investisseur historique (follow-on £6,5M avril 2023)",
     "Exit", "Élevé",
     "Plateforme « next-gen mortgage, protection and general insurance » pour courtiers (blockchain), "
     "distribution de produits de protection/assurance générale. Sortie signalée janvier 2026, nature exacte "
     "(acquisition ou autre) non confirmée par communiqué primaire.",
     "mortgagesolutions.co.uk, moneymarketing.co.uk, mortgagefinancegazette.com, acresoftware.com"),
    ("aviva-ventures_aviva-ventures", "Aviva Ventures", "Aviva Ventures", "Shepper", "InsurTech (adjacent)",
     "Royaume-Uni", "Série A", 2018.0, 5.4, "$", "Lead investisseur (avec Idekapital)", "Actif (à la date "
     "source)", "Élevé",
     "Plateforme d'inspections à la demande (réseau gig-economy), utilisée notamment pour inspections de "
     "biens/sinistres. Aviva a placé un dirigeant au board puis comme CEO.",
     "Insurance Post, Founders Factory (foundersfactory.com/aviva), Lions Financial"),
    ("aviva-ventures_aviva-ventures", "Aviva Ventures", "Aviva Ventures", "Meshed", "InsurTech (cœur)",
     "Royaume-Uni", "Pré-seed", 2025.0, 0.95, "£", "Co-investisseur (via Founders Factory et Exponential "
     "Science Foundation, lead : Haatch)", "Actif", "Élevé",
     "« Premier courtier d'assurance IA du Royaume-Uni » pour PME (agents IA vocaux/navigateur pour devis, "
     "relance assureurs, gestion de polices). Réduction de coûts d'assurance revendiquée de 52%.",
     "Insurance Times, Insurance Age, businesscloud.co.uk, tech.eu"),
    ("aviva-ventures_aviva-ventures", "Aviva Ventures", "Aviva Ventures", "Indico Data", "InsurTech (adjacent)",
     "États-Unis", "Stratégique", 2025.0, None, "$", "Investisseur stratégique (Chief Innovation Officer "
     "d'Aviva rejoint le board comme observateur/conseiller)", "Actif", "Élevé",
     "Automatisation IA des opérations d'assurance (traitement documentaire pour assureurs IARD mondiaux). "
     "Dernier investissement confirmé du fonds à la date de la recherche (28/10/2025).",
     "PRNewswire, Indico Data (blog officiel), Insurance Journal, BeBeez, FinSMEs"),
]

ANOMALIES_NOUVELLES = [
    ("Changement de statut de véhicule", "Fonds", "Atlantic Vantage Point (ex-AXA Venture Partners)",
     "Type d'investisseur / Société mère",
     "AXA Venture Partners — CVC captif d'AXA depuis 2015 (ex-AXA Strategic Ventures)",
     "Reste classé « VC indépendant » (pas de changement)",
     "Depuis le MBO d'août 2024, l'équipe de management d'AVP a racheté la participation de 70% détenue par "
     "AXA ; le véhicule s'est rebaptisé Atlantic Vantage Point (avpcap.com) et AXA n'est plus qu'un LP ancre "
     "minoritaire (nouveau fonds Growth I de 1,5 Md€, aux côtés du FEI). Conservé en « VC indépendant » "
     "(statut actuel) plutôt que reclassé CVC, avec la filiation AXA historique documentée ici pour mémoire. "
     "Aucune nouvelle participation ajoutée : Policygenius et Idelic, retrouvées par la recherche CVC-Lot 1, "
     "étaient déjà documentées en base sous ce véhicule.",
     "Élevé",
     "Private Equity Wire, Global Venturing, Coverager (2024-2025)"),
    ("Structure de véhicule non conforme au périmètre attendu", "Fonds", "Generali Ventures",
     "Nb participations documentées",
     "Attendu : investissements directs en startups insurtech",
     "0 participation directe — fonds de fonds (LP dans 5 fonds VC)",
     "Generali Ventures s'avère être un véhicule de type fonds de fonds (LP), pas un CVC à investissements "
     "directs. Il est commanditaire de Mundi Ventures, Speedinvest (déjà suivi séparément), Dawn Capital, "
     "Headline et StepStone Group. Aucune startup insurtech/fintech-assurance financée en direct par Generali "
     "Ventures n'a pu être sourcée — les portefeuilles des fonds sous-jacents ne sont pas comptés ici pour "
     "éviter toute extrapolation non sourcée.",
     "Élevé",
     "Sifted, Generali.com, Coverager, Reinsurance News"),
    ("Participation non retenue (hors périmètre sectoriel)", "Participations", "Allianz X",
     "Start-up", "SafeBoda (Ouganda) — co-lead Series B 2019, intention déclarée de produits d'assurance",
     "Non intégrée", "Activité principale = ride-hailing/mobilité, pas assurance au moment de l'investissement.",
     "Moyen", "Ventureburn, Global Legal Chronicle"),
    ("Participation non retenue (hors périmètre sectoriel)", "Participations", "Allianz X",
     "Start-up", "EthiFinance — investissement le plus récent d'Allianz X (01/08/2026)",
     "Non intégrée", "Agence de notation ESG, aucun lien clair avec l'assurance trouvé dans les sources.",
     "Moyen", "Recherche CVC-Lot 1"),
    ("Incohérence de montant entre sources primaires", "Participations", "Allianz X / Ualá",
     "Taille du tour (M€)", "$195M (fintech.global, LatamList) vs $300M (Medium Allianz X, même émetteur)",
     "$195M retenu par prudence (2 sources indépendantes concordantes)",
     "Deux communications d'Allianz X elles-mêmes donnent des montants différents pour le même tour Ualá — "
     "probable confusion entre deux tours distincts, à vérifier auprès d'une source primaire supplémentaire.",
     "Faible", "fintech.global, LatamList, Medium Allianz X"),
    ("Ambiguïté d'entité CVC", "Participations", "Sompo (Cover Genius, Zego, Trov)",
     "Rôle du véhicule (détail)", "Investisseur nommé dans la presse : « Sompo Holdings » / « Sompo Holdings "
     "Asia »", "Rattaché à Sompo Digital Lab / Light Vortex par défaut, confiance Moyen sur l'entité exacte",
     "Aucune source ne précise si ces investissements relèvent de Sompo Digital Lab, Sompo Light Vortex ou "
     "d'un investissement corporate direct du groupe hors de ces deux véhicules. Le lien capitalistique avec "
     "le groupe Sompo est lui bien confirmé ; seule l'entité CVC précise est incertaine.",
     "Moyen", "Bloomberg, theinsurer.com, fintech.global, PYMNTS.com"),
    ("Participation non retenue (lien assurance non établi)", "Participations", "Aviva Ventures",
     "Start-up", "Opun (proptech travaux, £3m Série A 2016) et Tembo Money (courtage hypothécaire, £2,5-5M "
     "2021-2023)", "Non intégrées",
     "Aucun produit d'assurance embarqué confirmé dans les sources consultées pour ces deux sociétés — "
     "activité principale hors périmètre (proptech / courtage hypothécaire pur).",
     "Moyen", "uktech.news, cityam.com, finextra.com, uktechnews.info"),
    ("Anomalie de source tierce (probable erreur d'agrégateur)", "Participations", "Aviva Ventures",
     "Start-up", "« Haast » (compliance IA, Série A $12M avril 2026, investisseurs confirmés : Peak XV "
     "Partners, DST Global Partners, Airtree, Aura Ventures, Black Sheep Capital)", "Non intégrée",
     "Une base tierce agrégée mentionne Aviva Ventures comme investisseur, mais Aviva n'apparaît dans aucune "
     "source primaire ou communiqué de presse relatif à ce tour. Probable confusion de nom/classification "
     "erronée de l'agrégateur plutôt qu'un investissement réel.",
     "Faible", "Recherche CVC-Lot 3 (Crunchbase agrégé, non confirmé par source primaire)"),
    ("Anomalie de source tierce (données incohérentes)", "Participations", "MS&AD Ventures",
     "Start-up", "« Volocopter » et « JFrog » cités comme exits confirmés par une source secondaire "
     "(Fundz.net)", "Non intégrées",
     "Aucun lien avec l'assurance identifiable pour ces deux sociétés (mobilité aérienne / DevOps) ; la "
     "cohérence de cette donnée est douteuse, possible erreur d'attribution de la source secondaire.",
     "Faible", "Fundz.net (source secondaire non corroborée)"),
]


def main() -> int:
    fonds = pd.read_excel(ENTREE, sheet_name="Fonds")
    part = pd.read_excel(ENTREE, sheet_name="Participations").copy()
    tours = pd.read_excel(ENTREE, sheet_name="Tours_de_table")
    dico = pd.read_excel(ENTREE, sheet_name="Dictionnaire")
    ano = pd.read_excel(ENTREE, sheet_name="Anomalies")

    fonds_out = pd.concat([fonds, pd.DataFrame(FONDS_NOUVEAUX)], ignore_index=True)

    existants = {(r["Fonds_ID"], r["Start-up"]) for _, r in part.iterrows()}
    nouvelles_lignes = []
    for (fid, nom_fonds, vehicule, startup, secteur, pays, stage, date, taille, devise,
         role, statut, confiance, commentaires, sources) in PARTICIPATIONS_NOUVELLES:
        if (fid, startup) in existants:
            continue
        nouvelles_lignes.append({
            "Fonds_ID": fid, "Nom du fonds": nom_fonds, "Véhicule": vehicule,
            "Start-up": startup, "Secteur": secteur, "Stage financé": stage,
            "Montant investi par le fonds (M€)": None, "Total levé par la start-up (M€)": None,
            "Date de création": None, "Date d’investissement": date, "Pays d’origine": pays,
            "Nb pays d’implantation": None,
            "Positionnement (Leader/Minoritaire)": "Minoritaire",
            "Capital social (k€)": None, "CA (M€)": None, "Valorisation (M€)": None,
            "Nb fonds investisseurs": None, "Nb tours (fonds)": None, "Nb tours (total)": None,
            "Repositionnement": None, "Exit (O/N)": "O" if statut and "Exit" in statut else "N",
            "MoC exit": None, "MoEP exit": None,
            "Taille du tour (M€)": conv(taille, devise), "% de détention": None,
            "Rôle du véhicule (détail)": role, "Statut start-up (détail)": statut,
            "Confiance": confiance, "Commentaires": commentaires, "Source(s)": sources,
        })
    part_out = pd.concat([part, pd.DataFrame(nouvelles_lignes)], ignore_index=True)

    prochain_id = int(ano["Anomalie_ID"].str.extract(r"(\d+)")[0].astype(int).max()) + 1
    nouvelles_anomalies = []
    for i, (type_ano, onglet, entite, champ, val_source, val_retenue, commentaire, confiance, sources) in \
            enumerate(ANOMALIES_NOUVELLES):
        nouvelles_anomalies.append({
            "Anomalie_ID": f"ANO-{prochain_id + i:04d}",
            "Type d'anomalie": type_ano, "Onglet source": onglet,
            "Table ou bloc source": "CVC-Lots 1-3 (recherche CVC assurantiels, 21/09/2026)",
            "Ligne source": None, "Entité concernée": entite, "Champ concerné": champ,
            "Valeur source": val_source, "Valeur retenue": val_retenue,
            "Candidats éventuels": None, "Score de similarité": None, "Niveau de confiance": confiance,
            "Traitement appliqué": commentaire, "Commentaire": sources,
        })
    ano_out = pd.concat([ano, pd.DataFrame(nouvelles_anomalies)], ignore_index=True)

    with pd.ExcelWriter(SORTIE, engine="openpyxl") as writer:
        fonds_out.to_excel(writer, sheet_name="Fonds", index=False)
        part_out.to_excel(writer, sheet_name="Participations", index=False)
        tours.to_excel(writer, sheet_name="Tours_de_table", index=False)
        dico.to_excel(writer, sheet_name="Dictionnaire", index=False)
        ano_out.to_excel(writer, sheet_name="Anomalies", index=False)

    print(f"Classeur produit : {SORTIE}")
    print(f"Fonds avant : {len(fonds)}  →  après : {len(fonds_out)}  (+{len(FONDS_NOUVEAUX)})")
    print(f"Participations avant : {len(part)}  →  après : {len(part_out)}  (+{len(nouvelles_lignes)})")
    print(f"Anomalies avant : {len(ano)}  →  après : {len(ano_out)}  (+{len(nouvelles_anomalies)})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
