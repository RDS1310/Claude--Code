#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
verifier_participations_confiance.py — Passe de vérification/upgrade de confiance
====================================================================================

Suite de la campagne "véhicules Fermé" : une seconde vague de recherche ciblée
(6 lots V1-V6) a tenté de confirmer nommément, ou de corriger, les ~68 lignes de
Participations à Confiance "Moyen"/"Faible"/"Faible-Moyen" (dataset complet,
véhicules Ouvert ET Fermé confondus). Ce script applique les résultats :
- upgrades de Confiance quand une source primaire nomme le véhicule/confirme le fait
- corrections factuelles (date, tour, montant, ou changement de véhicule) quand la
  recherche a trouvé une contradiction avec les données existantes
- suppression de 1 ligne dont le véhicule attribué s'est révélé incorrect
  (Alan / "Index Ventures XII" — le tour relève en réalité du fonds Growth d'Index,
  non suivi dans notre base), avec anomalie journalisée
- fusion de 2 lignes qui décrivaient en réalité le même véhicule (Clark chez White
  Star Capital : la Série C 2021 a été à tort attribuée à WSC III, qui n'a clôturé
  qu'en octobre 2021 — elle relève de WSC II comme la Série B)

Entrée  : VC_Database_Standardisee_v7.xlsx
Sortie  : VC_Database_Standardisee_v8.xlsx

Usage : python3 verifier_participations_confiance.py
"""

from __future__ import annotations

import math
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v7.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v8.xlsx"


def vide(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    if isinstance(v, str) and v.strip().lower() in ("", "n.d.", "n.d", "n/d", "nan"):
        return True
    return False


# ---------------------------------------------------------------------------
# Mises à jour simples : (Nom du fonds, Véhicule, Start-up) -> champs à écraser.
# Ne couvre que les champs qui changent ; le reste de la ligne est inchangé.
# ---------------------------------------------------------------------------
UPDATES = {
    ("BlackFin Capital Partners", "BlackfinTech 1", "Akur8"): {
        "Confiance": "Élevé",
        "Commentaires": "PitchBook identifie explicitement « BlackFin Tech Fund 1 » comme véhicule investisseur (Série A 8,9 M$, mars 2020, avec MTech Capital ; réinvesti en B et C).",
        "Source(s)": "TechCrunch ; PitchBook (BlackFin Tech Fund 1) ; BlackFin — page Akur8",
    },
    ("BlackFin Capital Partners", "BlackfinTech 1", "Descartes Underwriting"): {
        "Confiance": "Élevé",
        "Commentaires": "PitchBook liste Descartes Underwriting dans « BlackFin Tech Fund 1 ». BlackFin = tout premier VC au capital (seed ~2M€, closing 06/02/2019), réinvestissements ultérieurs.",
        "Source(s)": "Crunchbase (seed round) ; Reinsurance News ; BlackFin — page Descartes",
    },
    ("Portage Venture", "Portag3 Venture I", "Alan"): {
        "Confiance": "Élevé",
        "Commentaires": "Tour seed 12M€ (oct. 2016) réunissait OpenCNP, Partech Ventures, Financière Power et Portag3 Ventures LP. Véhicule « Portag3 Ventures Fund » (1er véhicule) millésime 2016, cohérent.",
        "Source(s)": "Maddyness (2016) ; Portage (profil Hélène Falchier) ; PitchBook",
    },
    ("Portage Venture", "Portag3 Venture II", "Clark"): {
        "Confiance": "Élevé",
        "Stage financé": "Série B",
        "Date d’investissement": 2018.0,
        "Taille du tour (M€)": 26.68,
        "Commentaires": "Série B 29M$, annoncée 24/04/2018, menée par Portag3 Ventures (Fund II) et White Star Capital, avec Coparion, Kulczyk Investments, Yabeo Capital.",
        "Source(s)": "TechCrunch ; talent4boards ; PSP Investments (Fund II close)",
    },
    ("Partech", "Partech Seed (I to IV)", "Alan"): {
        "Confiance": "Élevé",
        "Commentaires": "Communiqué Partech liste explicitement Alan parmi les participations de Partech Entrepreneur II (fonds seed lancé 2015, closing ~100M€ déc. 2016).",
        "Source(s)": "Partech « Alan, a legendary company in the making » ; Tech.eu (closing Partech Entrepreneur II)",
    },
    ("Ring Capital", "Mission I", "Goodvest"): {
        "Confiance": "Élevé",
        "Commentaires": "Ring Capital confirme Goodvest dans son portefeuille via « Ring Mission ». Tour nov. 2023 (10M€) confirmé ; Série B datée du 24/09/2025 pour 12M€ (menée par Serena), avec Ring Capital, Polytechnique Ventures, ALM Innovation, Globivest.",
        "Source(s)": "Ring Capital (fiche Goodvest) ; Ring Capital (fonds Ring Mission) ; Maddyness (nov. 2023) ; Tech.eu (sept. 2025)",
    },
    ("Founders Future VC", "Founders Future Good", "Neat"): {
        "Confiance": "Moyen-Élevé",
        "Commentaires": "Participation confirmée par plusieurs sources datées (seed 10M€ oct. 2022, Série A 50M€ sept. 2024, Founders Future « investisseur historique »). Nom exact du véhicule « Founders Future Good » non cité nommément — PitchBook confirme l'existence du fonds (millésime 2020, 35 participations), compatible mais déductif.",
        "Source(s)": "Tech.eu (oct. 2022) ; Coverager ; mind.eu (Série A sept. 2024) ; PitchBook (Founders Future Good)",
    },
    ("Founders Future VC", "Founders Future Good", "Napo"): {
        "Confiance": "Faible",
        "Commentaires": "Signal contradictoire : aucune source datée ne confirme Founders Future au capital de Napo. Listes d'investisseurs détaillées et concordantes pour chaque tour (Série A nov. 2022 17,3$M menée DN Capital ; Série B fév. 2025 12-15£M menée Mercia Ventures) NE mentionnent JAMAIS Founders Future. Un extrait indexé du site foundersfuture.com suggère Napo dans leur portefeuille en ligne, mais site inaccessible pour vérifier. Recommandation : ne pas confirmer tant qu'aucun communiqué daté n'est trouvé.",
        "Source(s)": "Crunchbase (Série A Napo) ; ffnews (Série B Napo) ; Mercia (annonce Série B) ; PitchBook (Napo)",
    },
    ("EOS Venture", "EVP I", "Westhill"): {
        "Confiance": "Élevé",
        "Commentaires": "EOS a mené la Série A (relation dès nov. 2018, closing déc. 2018 selon CB Insights, annonce publique 04/06/2019). J. Kalman (associé fondateur EOS) a rejoint le board. EVP I était l'unique fonds d'EOS à l'époque.",
        "Source(s)": "Coverager (Westhill Series A) ; GlobeNewswire (04/06/2019) ; Coverager (EOS closes fund) ; Insurance Journal (avril 2018)",
    },
    ("EOS Venture", "EVP I", "Roadzen"): {
        "Confiance": "Élevé",
        "Date d’investissement": 2020.0,
        "Commentaires": "Date corrigée : Série A en janvier 2020 (pas mai 2019). EOS confirmé comme « Portfolio Company » sur eosvc.com, et investisseur nommé sur Crunchbase/Tracxn avec WI Harper Group/Yukti Securities.",
        "Source(s)": "eosvc.com (portfolio Roadzen) ; Crunchbase ; Tracxn",
    },
    ("EOS Venture", "EVP I", "Ticker"): {
        "Confiance": "Élevé",
        "Stage financé": "Série A",
        "Commentaires": "Précision : c'est une Série A (pas Série B) menée par EOS en 2021. Munich Re Ventures a soutenu le lancement initial 2019 (tour distinct et antérieur) — rien ne confirme que Munich Re co-investissait dans le tour EOS 2021 lui-même.",
        "Source(s)": "Insurance Times ; Businesswire (Ticker/Abacai sept. 2021) ; Reinsurance News",
    },
    ("EOS Venture", "EVP I", "Neos Insurance"): {
        "Confiance": "Moyen",
        "Commentaires": "EOS Venture Partners fondée en 2016 (pas 2018). L'investissement seed oct. 2016 (~1,22£M) est confirmé nommément comme « Eos Venture Partners ». Mais la levée formelle du fonds EVP I n'a démarré qu'en avril 2018 — le deal Neos 2016 précède la levée officielle, pourrait relever d'un véhicule antérieur/informel de la même équipe. Investisseur confirmé, véhicule exact incertain.",
        "Source(s)": "UKTN ; Tracxn (Neos) ; venturecapitalarchive.com (EOS Fund I)",
    },
    ("Atlantic Vantage Point", "AVP early stage I", "ClimateSecure"): {
        "Confiance": "Élevé",
        "Commentaires": "AXA a lancé/investi via AXA Strategic Ventures (fév. 2015) ; portefeuille AXA Seed Factory transféré à AXA Strategic Ventures. Lignée : AXA Seed Factory → AXA Strategic Ventures (2015) → AXA Venture Partners (rebrand 2018) → Atlantic Vantage Point (rebrand 2024). Nom d'époque = « AXA Strategic Ventures ».",
        "Source(s)": "Artemis.bm ; Carrier Management (2015) ; FinTech Global (rebrand AVP)",
    },
    ("Atlantic Vantage Point", "AVP early stage I", "Particeep"): {
        "Confiance": "Élevé",
        "Commentaires": "Source primaire AXA : « Particeep… financée par AXA Strategic Ventures », opération complémentaire 500K€. Partenariat AXA Creditor/Particeep annoncé oct. 2015.",
        "Source(s)": "axa.com (magazine) ; newsassurancespro.com ; communiqué PDF axa.com",
    },
    ("Atlantic Vantage Point", "AVP early stage I", "FundShop"): {
        "Confiance": "Élevé",
        "Date d’investissement": 2014.0,
        "Stage financé": "Seed",
        "Taille du tour (M€)": 0.3,
        "Commentaires": "Levée 300K€ en septembre 2014 (tour seed) avec AXA Seed Factory/AXA Strategic Ventures (avec Ardian) — 1er investissement d'AXA Seed Factory dans FundShop.",
        "Source(s)": "Maddyness (sept. 2014) ; CB Insights (profil AXA Seed Factory)",
    },
    ("Atlantic Vantage Point", "AVP early stage I", "Widmee"): {
        "Confiance": "Élevé",
        "Commentaires": "Confirmation précise : « AXA SEED FACTORY investit 210K€ dans WIDMEE », tour clos le 13/11/2014.",
        "Source(s)": "Fusacq ; Dynamique Mag",
    },
    ("Atlantic Vantage Point", "AVP early stage II", "PolicyGenius"): {
        "Commentaires": "Chronologie précise : AXA Strategic Ventures a investi en A (2015, $5,9M), B (2015), C (mai 2017, $30M) — plausiblement AVP ES I. Puis AXA Venture Partners (déjà rebrandé) a participé à la D ($100M, 30/01/2020) — période où ES II (vintage 2019) était potentiellement actif, mais rien ne confirme s'il s'agit d'un nouvel investissement via ES II ou d'un suivi depuis ES I. Série E (2022, $125M) liste « Axa » directement — pourrait être un investissement corporate direct plutôt que via un véhicule AVP. Impossible de trancher entre AVP ES I et AVP ES II.",
        "Source(s)": "FinSMEs (Série C 2017 ; Série D janv. 2020) ; avpcap.com (Série D) ; Global Venturing (Série E 2022) ; PitchBook (AVP ES II)",
    },
    ("Serena Capital", "Serena III", "Acheel"): {
        "Confiance": "Faible-Moyen",
        "Commentaires": "Tour confirmé et daté précisément : 29M€ annoncé en mai 2021, avec Xavier Niel (NJJ), Serena Capital et Portag3 Ventures. Cohérent avec la période de déploiement de Serena III (closing final fév. 2021), mais aucune source ne cite nommément « Serena III ».",
        "Source(s)": "FrenchWeb ; FUSACQ Buzz ; Nordic9",
    },
    ("Breega Capital", "Breega Venture III", "Coverflex"): {
        "Confiance": "Élevé",
        "Commentaires": "Le fonds officiel « Breega Capital Venture Three » a clôturé le 24/03/2021 à 130M$ (~110M€). Le tour Coverflex (5M€, mi-avril 2021) tombe juste après ce closing : forte cohérence chronologique.",
        "Source(s)": "TechCrunch (24/03/2021, closing du fonds) ; Fintech.global (16/04/2021, tour Coverflex) ; data.gouv.fr/annuaire-entreprises",
    },
    ("Breega Capital", "Breega Seed II", "Cuvva"): {
        "Commentaires": "Point de vigilance : le tour Cuvva (15£M, Série A, 03/12/2019) précède de 15 mois le closing du fonds « Breega Capital Venture Three » (24/03/2021), qui pourrait être le même véhicule légal que notre « Breega Seed II ». Plusieurs sources indiquent que c'est plutôt le 1er fonds « Venture » de Breega (levé 2017, ~105-106M€, non suivi dans notre base Fonds) qui a financé Cuvva. Rattachement à Breega Seed II conservé par prudence en l'absence de certitude, mais à vérifier directement auprès de Breega.",
        "Source(s)": "TechCrunch (03/12/2019) ; TechCrunch (23/06/2022, rétrospective fonds 2017) ; TechCrunch (24/03/2021, closing Breega Seed II/Venture Three)",
    },
    ("Cathay", "Cathay Innovation II", "Coherent"): {
        "Confiance": "Élevé",
        "Commentaires": "Série A 14M$ confirmée (annoncée 09-10/11/2020), menée par Cathay Innovation avec Franklin Templeton. Aucune source ne nomme le véhicule précis « Cathay Innovation II » — juste « Cathay Innovation » (la société de gestion).",
        "Source(s)": "cathaycapital.com ; TechCrunch (09/11/2020) ; cathayinnovation.com (page Coherent)",
    },
    ("Cathay", "Cathay Innovation II", "Igloo"): {
        "Confiance": "Élevé",
        "Commentaires": "Confirmé Série A+ 2020 (alors « Axinan », 16M$ total) ET Série B 2022 (Cathay a mené à 19M$ mars 2022, extension 46M$ nov. 2022 comme investisseur revenant). Véhicule « Cathay Innovation II » non nommé nommément.",
        "Source(s)": "fintechnews.sg ; technode.global ; TechCrunch (28/11/2022)",
    },
    ("Cathay", "Cathay Innovation II", "Qover"): {
        "Confiance": "Élevé",
        "Date d’investissement": 2021.0,
        "Commentaires": "Date corrigée : le tour 25M$ mené par Prime Ventures avec Cathay Innovation date du 27/04/2021 (pas janvier 2023). Un Série C distinct de 30M$ a eu lieu juillet 2023 mais SANS Cathay Innovation (Alven, Anthemis, Kreos Capital, Zurich Global Ventures). Véhicule précis toujours non nommé.",
        "Source(s)": "tech.eu (27/04/2021) ; qover.com/press ; cathaycapital.com",
    },
    ("Cathay", "Cathay Innovation II", "Coverfy"): {
        "Confiance": "Moyen",
        "Commentaires": "Lien confirmé par source primaire (page portefeuille officielle Cathay Innovation) — plus une mention isolée. Date/montant du ticket Cathay introuvables.",
        "Source(s)": "cathayinnovation.com/company/coverify ; Tracxn",
    },
    ("Truffle Capital", "Truffle FinTech & InsurTech Fund II", "MoneyTrack"): {
        "Confiance": "Élevé",
        "Date d’investissement": 2021.0,
        "Commentaires": "Nom officiel exact du véhicule : « Truffle Financial Innovation Fund » (alias « Fonds Institutionnel FinTech-InsurTech »), clos à 140M€ (1er closing nov. 2017, final ~2019) — pas de « Fund II » distinct sur ce segment chez Truffle. Levée MoneyTrack de 2,3M€ datée précisément de 2021, concomitante au rachat de Progexia.",
        "Source(s)": "groupebpce.com ; frenchweb.fr ; truffle.com/portfolio-truffle-capital/moneytrack",
    },
    ("speedInvest", "Speedinvest IV", "Grace"): {
        "Confiance": "Élevé",
        "Commentaires": "Seed 5,9M€ (03/04/2025) co-mené FinTech Collective ET Speedinvest, confirmé (avec Firstminute Capital, Purple, Kima Ventures, Bpifrance). Véhicule « Speedinvest IV » non nommé spécifiquement.",
        "Source(s)": "TechCrunch (03/04/2025) ; traded.co",
    },
    ("speedInvest", "Speedinvest IV", "Bliss (Saúde Bliss)"): {
        "Confiance": "Élevé",
        "Commentaires": "Confirmé : Speedinvest a soutenu le seed 2023 et est revenu dans la Série A R$57M (~11M$US, 24/03/2026), avec Kfund, Grupo Bradesco, Actyus, Clocktower Ventures, Canary.",
        "Source(s)": "speedinvest.com/portfolio ; businessweek.com.br ; startupi.com.br",
    },
    ("13books Capital (ex-Element Ventures)", "Fonds I & II", "hepster"): {
        "Confiance": "Élevé",
        "Commentaires": "Série A 10M$ menée par Element Ventures annoncée 04/03/2021 ; fonds fondateur 130M$ (Fonds I) lancé/annoncé 02/08/2021 avec hepster déjà cité en portefeuille.",
        "Source(s)": "TechCrunch (04/03/2021 ; 02/08/2021) ; 13bookscapital.com/portfolio/hepster",
    },
    ("13books Capital (ex-Element Ventures)", "Fonds I & II", "Roadzen"): {
        "Confiance": "Élevé",
        "Commentaires": "2e source indépendante trouvée : page portefeuille officielle 13books liste Roadzen explicitement, aux côtés de Hepster, Coincover, Runa, Billhop, Thirdfort, Duco, nCino, Fenergo, ErisX. Source primaire, plus une mention isolée.",
        "Source(s)": "13bookscapital.com/portfolio/roadzen",
    },
    ("Astorya.vc", "Astorya.vc", "Embea"): {
        "Confiance": "Élevé",
        "Commentaires": "Tour seed 4M€ mené par Atlantic Labs, avec astorya.vc.",
        "Source(s)": "L'Assurance en Mouvement",
    },
    ("Astorya.vc", "Astorya.vc", "Riskwolf"): {
        "Confiance": "Élevé",
        "Commentaires": "1er tour externe (pre-seed) CHF 750K, mené par SICTIC, astorya.vc parmi +10 investisseurs.",
        "Source(s)": "startupticker.ch ; EU-Startups",
    },
    ("Astorya.vc", "Astorya.vc", "SesameIT"): {
        "Confiance": "Élevé",
        "Commentaires": "Levée 10M€ (30/03/2023) auprès de 115K, Banque des Territoires, et investisseurs historiques BNP Paribas Développement et astorya.vc (entré lors d'un tour antérieur).",
        "Source(s)": "Communiqué SesameIT ; Solutions Numériques ; GlobalSecurityMag",
    },
    ("Astorya.vc", "Astorya.vc", "Weecover"): {
        "Confiance": "Élevé",
        "Commentaires": "Tour seed 2,3M€ mené par Nauta (2021), astorya.vc participant.",
        "Source(s)": "Tracxn ; Tech.eu",
    },
    ("Astorya.vc", "Astorya.vc", "Wenalyze"): {
        "Confiance": "Élevé",
        "Commentaires": "Levée 1,7M€, menée par Athos Capital, avec Bankinter, astoryaVC et GoHub. Florian Graillot (astorya.vc) cité dans l'annonce.",
        "Source(s)": "Medium/GoHub Ventures",
    },
    ("Astorya.vc", "Astorya.vc", "Wetterheld"): {
        "Confiance": "Élevé",
        "Commentaires": "Seed 300K€ (30/09/2019), astorya.vc cité parmi 4-5 investisseurs.",
        "Source(s)": "Crunchbase (recoupé avec portefeuille astorya.vc)",
    },
    ("115K", "115K", "Cartan Trade"): {
        "Confiance": "Faible",
        "Commentaires": "Aucune source indépendante ne confirme 115K investisseur de Cartan Trade. Tours 2022 (8M€, SV One/SCOR, Bpifrance, Quattro Holding) et 2025 (9M€, SV One/SCOR + Intact) ne citent jamais 115K. Seul lien : Damien Launoy (Managing Partner 115K) a siégé au board — ne prouve pas un investissement du véhicule. Candidat à suppression si non corroboré.",
        "Source(s)": "Bpifrance presse ; Cartan Trade (réorg. actionnariat 2025) ; Option Finance ; CFNEWS",
    },
    ("115K", "115K", "Covalt (ex-Pono)"): {
        "Confiance": "Élevé",
        "Commentaires": "Pono (devenu Covalt) a levé 3M€ (10/12/2022) auprès de Newfund, 115K, Kima Ventures. 115K liste Pono dans son propre portefeuille.",
        "Source(s)": "Maddyness ; 115K portfolio ; Societe.tech",
    },
    ("115K", "115K", "Zaion"): {
        "Confiance": "Élevé",
        "Commentaires": "Levée 11M€ (24/01/2025), tour mené par 115K. 3e levée de la société.",
        "Source(s)": "Communiqué La Banque Postale/115K ; Truffle Capital news ; En-Contact",
    },
    ("Venpace", "Venpace", "Complero"): {
        "Confiance": "Élevé",
        "Commentaires": "Investissement « high six-digit » (16/12/2022) de VENPACE et ISB (investisseur existant).",
        "Source(s)": "Venpace (Medium) ; WSS Redpoint",
    },
    ("Venpace", "Venpace", "DC Connected Car"): {
        "Confiance": "Élevé",
        "Commentaires": "Seed 2,1M€ (août 2024), mené par VENPACE avec Borusan Ventures, Atlas Ventures, Network.VC, APX, Bloomhaus Ventures.",
        "Source(s)": "BusinessWire ; Vestbee ; Silicon Canals",
    },
    ("Venpace", "Venpace", "HealthCaters"): {
        "Confiance": "Élevé",
        "Commentaires": "Seed 1,2M€, mené par Barmenia Next Strategies avec Venpace (co-investisseur), DvH Ventures.",
        "Source(s)": "Silicon Canals ; IBB Ventures portfolio",
    },
    ("Venpace", "Venpace", "hypt."): {
        "Confiance": "Élevé",
        "Commentaires": "Seed CHF 1,65M (~1,77M€, sept. 2025), mené par VENPACE avec SixThirty Ventures, COREangels, SICTIC, Gateway Ventures, NCA.",
        "Source(s)": "startupticker.ch ; FinSMEs ; Vestbee",
    },
    ("BPI France", "Large Ventures", "Shift Technology"): {
        "Confiance": "Élevé",
        "Commentaires": "Bpifrance entre au capital via Large Venture lors de la Série D 220M$ (mai 2021), menée par Advent International/Advent Tech avec Avenir Growth. Shift devient licorne française.",
        "Source(s)": "Bpifrance communiqué officiel ; Dechert ; FrenchWeb ; Alliancy",
    },
    ("Concentric", "Concentric", "Black Insurance"): {
        "Confiance": "Élevé",
        "Commentaires": "Black Insurance (blockchain, Estonie) a levé ~800K$ (sept. 2018) avec Concentric, BlackPearls, Fineqia, Siena. Concentric liste Black Insurance dans sa page Investments officielle.",
        "Source(s)": "Concentric (page Investments) ; PR.com",
    },
    ("Insurtech Capital", "Insurtech Capital (Groupe Apicil)", "Nalo"): {
        "Confiance": "Élevé",
        "Commentaires": "Investissement le 21/04/2023 ; Apicil devenu actionnaire de référence puis détient >90% du capital — prise de participation majoritaire progressive, pas un simple ticket VC minoritaire.",
        "Source(s)": "PitchBook ; newsassurancespro.com ; planet-fintech.com",
    },
    ("Kima Venture Capital", "Kima Venture", "Alan"): {
        "Confiance": "Élevé",
        "Stage financé": "Série A",
        "Date d’investissement": 2018.0,
        "Taille du tour (M€)": 23.0,
        "Commentaires": "Correction : Kima Ventures N'A PAS participé au seed initial (oct. 2016, 12M€). Entré en Série A (avril 2018, 23M€, menée par Index Ventures, avec Partech, CNP Assurances, Portag3, Xavier Niel/Kima Ventures).",
        "Source(s)": "Index Ventures communiqué ; FrenchWeb ; mind Health",
    },
    ("Macif Innovation", "Macif Innovation", "Liberty Rider"): {
        "Confiance": "Élevé",
        "Commentaires": "Seed 1,6M€ avec Matmut, Inter Mutuelles Assistance, Macif, Mutuelle des Motards, Racer. Succès emblématique du portefeuille Macif Innovation.",
        "Source(s)": "Espace presse Matmut ; Maddyness (portrait Macif Innovation)",
    },
    ("NCA (Next Commerce Accelerator)", "NCA (Next Commerce Accelerator)", "Etvas"): {
        "Confiance": "Élevé",
        "Commentaires": "Etvas a intégré la 5e promotion NCA (oct. 2019) ; NCA décrit comme tout premier investisseur (ticket ~25-150K€ contre equity).",
        "Source(s)": "startupcity.hamburg ; HTGF (portfolio Etvas) ; hamburg-startups.net",
    },
    ("SCOR Ventures", "SCOR Ventures", "Branch Insurance"): {
        "Confiance": "Élevé",
        "Commentaires": "Série A 24M$ (2020, co-menée HSCM/Greycroft) puis Série B 50M$ (2021, menée Anthemis Group) — SCOR P&C Ventures participant aux deux.",
        "Source(s)": "PRNewswire ; IntelligentInsurer ; ReinsuranceNews",
    },
    ("SCOR Ventures", "SCOR Ventures", "Protex AI"): {
        "Confiance": "Élevé",
        "Commentaires": "Tour combiné seed+Série A 18M$ (28/08/2022), mené par Notion Capital, avec Firstminute Capital, Flexport, et investissement stratégique SCOR Ventures.",
        "Source(s)": "Protex AI communiqué ; Irish Times ; Silicon Canals",
    },
    ("Seraphim Space", "Seraphim Space Ventures (SSVII + SSIT)", "Adaptive Insurance"): {
        "Confiance": "Élevé",
        "Commentaires": "Seed 5M$, mené par Congruent Ventures, avec Generation Space (bras US de Seraphim Space) et Montauk Climate.",
        "Source(s)": "Seraphim Space communiqué ; PRWeb",
    },
    ("Sharpstone Capitale", "Sharpstone Capitale", "Nalo"): {
        "Confiance": "Élevé",
        "Commentaires": "Investisseur historique (tour 4M€ avec business angels), antérieur à l'entrée Apicil/Insurtech Capital (2023). Deux investisseurs distincts à périodes différentes, pas d'erreur d'attribution.",
        "Source(s)": "mind Fintech ; Maddyness (portrait Sharpstone Capital)",
    },
    ("Tenity", "Tenity Incubation Fund I & II", "CyberTide"): {
        "Confiance": "Élevé",
        "Commentaires": "Sélectionnée parmi 9 startups du programme incubation Tenity (promo Suisse), investissement initial CHF 50K du Tenity Incubation Fund II. Antler aussi investisseur.",
        "Source(s)": "Tenity communiqué officiel ; Tracxn",
    },
    ("The Net Street Capital", "The Net Street Capital", "Coverfy"): {
        "Confiance": "Faible",
        "Commentaires": "Point de vigilance : recherche exhaustive des 6 tours documentés de Coverfy (seed 2016, Série A, convertible 2020) — AUCUNE source indépendante ne mentionne « The Net Street Capital ». Seules des bases agrégées peu détaillées (Capboard/OpenVC) associent les deux sans détail. Coverfy signalée fermée (Crunchbase). Risque d'erreur d'attribution réel — vérifier la source d'origine.",
        "Source(s)": "Webcapitalriesgo (Série A ; Seed) ; Ecommerce News",
    },
    ("UNIQA Ventures", "UNIQA Ventures", "Finabro"): {
        "Confiance": "Élevé",
        "Commentaires": "Premier tour 600K€, UNIQA lead investor (400K€) + 200K€ business angels.",
        "Source(s)": "Finabro blog officiel ; wien.wirtschaftszeit.at",
    },
    ("UNIQA Ventures", "UNIQA Ventures", "Luko"): {
        "Confiance": "Élevé",
        "Commentaires": "Participation mentionnée de façon récurrente et indépendante (CB Insights étude de cas, Vestbee, PitchBook, Tracxn/Golden). Round précis non confirmé (communiqués Luko 2019-2020 ne citent pas UNIQA nommément).",
        "Source(s)": "CB Insights (étude de cas) ; Vestbee ; PitchBook",
    },
}

# ---------------------------------------------------------------------------
# Changement de véhicule : Continuity relève d'Alpha II, pas de DV4 (source
# primaire CFNews + Bpifrance Presse, confirmé sans ambiguïté).
# ---------------------------------------------------------------------------
CHANGEMENT_VEHICULE = {
    ("Elaia Partners", "DV4", "Continuity"): {
        "Fonds_ID": "elaia-partners_alpha-ii",
        "Véhicule": "Alpha II",
        "Confiance": "Élevé",
        "Commentaires": "Correction : CFNews confirme explicitement que le capital a été injecté par Elaia Partners via le fonds Alpha II et Bpifrance Digital Venture — pas DV4.",
        "Source(s)": "CFNews (« Continuity assure son amorçage ») ; Bpifrance Presse (1er closing Alpha II à 30M€)",
    },
}

# ---------------------------------------------------------------------------
# Fusion : la Série C 2021 de Clark, attribuée à tort à WSC III (qui n'a
# clôturé qu'en octobre 2021, après ce tour), relève en réalité de WSC II.
# On enrichit la ligne WSC II/Clark et on supprime la ligne WSC III/Clark.
# ---------------------------------------------------------------------------
FUSION_WSC_CLARK_CIBLE = ("White Star Capital", "WSC II", "Clark")
FUSION_WSC_CLARK_SOURCE = ("White Star Capital", "WSC III", "Clark")
FUSION_WSC_CLARK_UPDATE = {
    "Confiance": "Élevé",
    "Commentaires": (
        "Communiqué officiel du 2e fonds White Star Capital (juin 2018, $180M) cite "
        "explicitement Clark parmi ses investissements (Série B avril 2018, 29$M). "
        "La Série C de Clark (69M€, janvier 2021) lui est également rattachée : elle a "
        "eu lieu AVANT le closing final de WSC III ($360M, octobre 2021 — Forbes), qui "
        "ne pouvait donc pas la financer — corrigé depuis « WSC III »."
    ),
    "Source(s)": "TechCrunch (03/06/2018) ; White Star Capital/Medium ; Nordic9 ; The Logic ; Forbes (25/10/2021, closing WSC III)",
}

# ---------------------------------------------------------------------------
# Suppression : le tour de croissance d'Alan (mars 2026, 100M€) relève du
# fonds Growth d'Index Ventures, pas d'« Index Ventures XII » (fonds
# early-stage distinct) — véhicule non suivi dans notre base Fonds.
# ---------------------------------------------------------------------------
SUPPRESSION = [("Index Venture", "Index Ventures XII", "Alan")]


def main() -> int:
    fonds = pd.read_excel(ENTREE, sheet_name="Fonds")
    part = pd.read_excel(ENTREE, sheet_name="Participations")
    tours = pd.read_excel(ENTREE, sheet_name="Tours_de_table")
    dico = pd.read_excel(ENTREE, sheet_name="Dictionnaire")
    ano = pd.read_excel(ENTREE, sheet_name="Anomalies")

    part = part.copy()
    cle_index = {
        (row["Nom du fonds"], row["Véhicule"], row["Start-up"]): idx
        for idx, row in part.iterrows()
    }

    n_maj = 0
    non_trouvees = []
    for cle, champs in UPDATES.items():
        idx = cle_index.get(cle)
        if idx is None:
            non_trouvees.append(cle)
            continue
        for champ, valeur in champs.items():
            part.at[idx, champ] = valeur
        n_maj += 1

    for cle, champs in CHANGEMENT_VEHICULE.items():
        idx = cle_index.get(cle)
        if idx is None:
            non_trouvees.append(cle)
            continue
        for champ, valeur in champs.items():
            part.at[idx, champ] = valeur
        n_maj += 1

    idx_cible = cle_index.get(FUSION_WSC_CLARK_CIBLE)
    idx_source = cle_index.get(FUSION_WSC_CLARK_SOURCE)
    if idx_cible is not None and idx_source is not None:
        for champ, valeur in FUSION_WSC_CLARK_UPDATE.items():
            part.at[idx_cible, champ] = valeur
        part = part.drop(index=idx_source)

    n_suppr = 0
    for cle in SUPPRESSION:
        idx = cle_index.get(cle)
        if idx is not None and idx in part.index:
            part = part.drop(index=idx)
            n_suppr += 1

    part_out = part.reset_index(drop=True)

    # Anomalie : rattachement Index Ventures / Alan invalidé.
    prochain_id = int(ano["Anomalie_ID"].str.extract(r"(\d+)")[0].astype(int).max()) + 1
    nouvelle_anomalie = {
        "Anomalie_ID": f"ANO-{prochain_id:04d}",
        "Type d'anomalie": "Participation retirée — véhicule attribué invalidé",
        "Onglet source": "Participations",
        "Table ou bloc source": "Vérif-Lot V4 (passe de vérification confiance, 16/09/2026)",
        "Ligne source": None,
        "Entité concernée": "Index Venture",
        "Champ concerné": "Véhicule",
        "Valeur source": "Alan, levée de croissance 100M€ (2026-03), attribuée à « Index Ventures XII »",
        "Valeur retenue": "Ligne retirée de Participations",
        "Candidats éventuels": "Fonds Growth d'Index Ventures (Growth VII, ~2,2Md$ 2026) — non suivi dans la base Fonds",
        "Score de similarité": None,
        "Niveau de confiance": "Moyen",
        "Traitement appliqué": "Ligne supprimée (véhicule non tracké), signalée pour ajout éventuel d'un nouveau véhicule",
        "Commentaire": (
            "Recherche confirmant que le tour relève du fonds Growth d'Index Ventures "
            "(distinct d'Index Ventures XII, early-stage vintage 2024) — aucune source "
            "ne nomme le fonds Growth précisément pour Alan. maddyness.com (11/03/2026), "
            "thesaasnews.com, TechCrunch (09/07/2024, existence du fonds Growth)."
        ),
    }
    ano_out = pd.concat([ano, pd.DataFrame([nouvelle_anomalie])], ignore_index=True)

    with pd.ExcelWriter(SORTIE, engine="openpyxl") as writer:
        fonds.to_excel(writer, sheet_name="Fonds", index=False)
        part_out.to_excel(writer, sheet_name="Participations", index=False)
        tours.to_excel(writer, sheet_name="Tours_de_table", index=False)
        dico.to_excel(writer, sheet_name="Dictionnaire", index=False)
        ano_out.to_excel(writer, sheet_name="Anomalies", index=False)

    print(f"Classeur produit : {SORTIE}")
    print(f"Participations avant : {len(pd.read_excel(ENTREE, sheet_name='Participations'))}  →  après : {len(part_out)}")
    print(f"Lignes mises à jour : {n_maj}  |  fusionnées : {1 if idx_source is not None else 0}  |  supprimées : {n_suppr}")
    print(f"Anomalies avant : {len(ano)}  →  après : {len(ano_out)}")
    if non_trouvees:
        print(f"Clés non trouvées dans la base (à vérifier) : {len(non_trouvees)}")
        for c in non_trouvees:
            print("  -", c)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
