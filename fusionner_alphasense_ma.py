#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fusionner_alphasense_ma.py — Intègre le rapport AlphaSense (M&A assurance/insurtech)
====================================================================================

L'utilisateur a fourni deux documents issus d'une recherche AlphaSense (17/09/2026) :
un rapport PDF (17 opérations/tendances) et un fichier Excel de suivi personnel
(23 lignes, plus détaillé, avec notes méthodologiques sur les sources à recouper).

Traitement :
- Les opérations déjà présentes dans ma_data/deals.json (At-Bay, Prima Assicurazioni,
  Magnolia, Ociane Matmut/Mgéfi) ne sont PAS dupliquées. Une seule (At-Bay) est
  enrichie d'un multiple manquant, purement additif, sans rien écraser. Les autres
  gardent leurs données existantes (déjà plus riches ou légèrement différentes de la
  version AlphaSense) plutôt que d'introduire une incohérence numérique.
- Les lignes de type "Tendance" du rapport (Fitch, FTI Consulting, Oliver Wyman,
  S&P Global) sont des commentaires d'analystes agrégés, pas des opérations
  discrètes — elles ne correspondent à aucun schéma existant (confirmees/rumeurs
  attendent un acquéreur+une cible) et ne sont donc pas intégrées telles quelles.
  Le chiffre FTI Consulting (789 opérations Europe 2025, +14%) était déjà présent
  dans contexteMarche, à l'identique — aucune mise à jour nécessaire.
- 24 nouvelles opérations confirmées + 2 rumeurs sont ajoutées, avec le same schéma
  que l'existant. Deux lignes portent une réserve explicite sur leur source
  (BNP Paribas Cardif/BCC Vita : lien source erroné signalé par l'utilisateur lui-même ;
  Malakoff Humanis/Unofi : source secondaire à recouper).

Entrée  : ma_data/deals.json
Sortie  : ma_data/deals.json (modifié en place, une sauvegarde .bak est conservée)

Usage : python3 fusionner_alphasense_ma.py
"""

from __future__ import annotations

import json
import shutil
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
FICHIER = BASE_DIR / "ma_data" / "deals.json"

TAUX = {"€": 1.0, "$": 0.92, "£": 1.17, "CHF": 1.07}


def conv(montant, devise):
    if montant is None:
        return None
    return round(montant * TAUX.get(devise, 1.0))


NOUVELLES_CONFIRMEES = [
    dict(date="02/09/2026", acquereur="Alan", cible="Tanel (Sénégal)",
         secteur="Assurtech santé — Afrique de l'Ouest", type="Acquisition",
         statut="Finalisée ; montant non communiqué",
         valorisation=None, valorisationM=None, multiple=None,
         theme="Assurtech — expansion internationale",
         resume="Alan a finalisé le rachat de Tanel, acteur de santé digitale basé à Dakar — première implantation du groupe en Afrique. Les fondateurs de Tanel conservent la direction des opérations locales.",
         source="Agence Ecofin ; Maddyness",
         url="https://www.maddyness.com/2026/09/02/alan-met-le-cap-sur-lafrique-avec-une-acquisition-au-senegal/"),
    dict(date="12/06/2026", acquereur="Belfius Insurance", cible="Insurlytech (maison mère de Leocare)",
         secteur="Assurtech IARD — distribution digitale (France)", type="Acquisition (100 %)",
         statut="Accord signé ; intégration à partir du 01/07/2026",
         valorisation=None, valorisationM=None, multiple=None,
         theme="Assurtech — distribution digitale",
         resume="Belfius Insurance acquiert 100 % d'Insurlytech — première opération du bancassureur belge hors de Belgique. Leocare revendique 1,3 M d'utilisateurs et plus de 750 courtiers/partenaires B2B2C.",
         source="La Libre ; Bretagne Économique",
         url="https://www.lalibre.be/economie/entreprises-startup/2026/06/12/belfius-sort-de-ses-frontieres-avec-une-acquisition-qui-lui-ouvre-les-portes-du-marche-de-lassurance-digitale-en-france-IQZCUAMHH5CRJGAWHPMO7CNJD4/"),
    dict(date="05/06/2026", acquereur="Macif Santé Prévoyance (Aéma Groupe)", cible="Math (absorption) + partenariat MNFCT",
         secteur="Mutuelle santé-prévoyance (France)", type="Absorption + partenariat",
         statut="Projet annoncé, sous réserve des autorisations réglementaires",
         valorisation=None, valorisationM=None, multiple=None,
         theme="Prévoyance & mutualité",
         resume="Macif Santé Prévoyance prévoit d'absorber la mutuelle Math et renforce son partenariat avec MNFCT sur la protection sociale complémentaire des agents territoriaux/hospitaliers — ensemble protégeant près de 2,3 M de personnes.",
         source="Asquare Partners (reprend le communiqué Macif)",
         url="https://presse.macif.fr/actualites/la-macif-finalise-l-unification-de-ses-activites-sante-et-prevoyance-au-service-de-son-developpement-strategique-8a26e-821df.html"),
    dict(date="22/06/2026", acquereur="BNP Paribas Cardif", cible="BCC Vita (Italie) — 51 % → 70 %",
         secteur="Bancassurance vie (Italie)", type="Renforcement de participation majoritaire",
         statut="Annoncée",
         valorisation=None, valorisationM=None, multiple=None,
         theme="Bancassurance",
         resume="Cardif acquiert 19 % supplémentaires de BCC Vita et prolonge le partenariat avec BCC Iccrea jusqu'en 2039.",
         source="Asquare Partners — ⚠️ réserve : le lien source original pointe vers un article sans rapport avec l'opération ; à vérifier directement sur le site BNP Paribas Cardif avant citation externe",
         url=None),
    dict(date="30/04/2026", acquereur="LCL / Crédit Agricole Assurances", cible="Groupe Milleis (Milleis Vie → Spirica)",
         secteur="Banque privée / assurance vie haut de gamme (France)", type="Acquisition",
         statut="Finalisée",
         valorisation=None, valorisationM=None, multiple=None,
         theme="Épargne patrimoniale",
         resume="LCL et CAA finalisent le rachat de Milleis, troisième banque privée indépendante française (64 000 familles clientes, 13 Md€ d'encours). Milleis Vie est intégrée à Spirica.",
         source="Asquare Partners (reprend le communiqué Crédit Agricole)",
         url="https://presse.credit-agricole.com/lcl-et-credit-agricole-assurances-finalisent-lacquisition-du-groupe-milleis-et-consolident-leur-strategie-patrimoniale/?lang=fra"),
    dict(date="11/2025", acquereur="Malakoff Humanis", cible="Unofi (85 %)",
         secteur="Épargne notariale / assurance vie (France)", type="Acquisition majoritaire",
         statut="Finalisée (novembre 2025 selon la source)",
         valorisation=None, valorisationM=None, multiple=None,
         theme="Prévoyance & mutualité",
         resume="Malakoff Humanis prend 85 % d'Unofi dans le cadre de sa stratégie de diversification vers l'épargne. ⚠️ Source secondaire uniquement (France Épargne Research) — date et pourcentage à recouper avec un communiqué officiel.",
         source="France Épargne Research — source secondaire, à vérifier",
         url="https://www.france-epargne.fr/research/fr/french-insurance-2026-positioning-paper-mutuelle-prevoyance-borrower-insurance"),
    dict(date="02/03/2026", acquereur="Zurich Insurance Group", cible="Beazley plc (UK)",
         secteur="Assurance de spécialités / cyber / Lloyd's", type="Acquisition (100 %, numéraire)",
         statut="Toutes autorisations obtenues ; audience du tribunal le 22/09/2026 ; prise d'effet attendue le 01/10/2026",
         valorisation="8 100 M£", valorisationM=conv(8100, "£"), multiple=None,
         theme="Consolidation spécialités",
         resume="Offre 100 % numéraire à 1 335 pence/action (prime de 59,8 % sur le cours du 16/01/2026), financée par cash, dette nouvelle et une augmentation de capital de 5 Md$. L'ensemble revendiquerait ≈15 Md$ de primes de spécialités pro forma.",
         source="Reinsurance News ; Insurance Times ; communiqué Zurich",
         url="https://www.reinsurancene.ws/zurichs-8-1bn-beazley-acquisition-to-complete-october-1-following-regulatory-clearance/"),
    dict(date="08/2026", acquereur="AXA XL", cible="S-RM (UK) — solde du capital",
         secteur="Services de prévention cyber", type="Acquisition (contrôle total)",
         statut="Accord annoncé",
         valorisation=None, valorisationM=None, multiple=None,
         theme="Cyber & prévention",
         resume="AXA XL prend le contrôle total de S-RM (conseil en renseignement et cybersécurité) pour développer une offre de prévention au-delà de la couverture d'assurance.",
         source="InsurTech.me Investment Intelligence Report",
         url="https://insurtech.me/reports/2026-08-08/"),
    dict(date="08/2026", acquereur="Maybank", cible="30,95 % de Maybank Ageas Holdings (cédant : Ageas)",
         secteur="Bancassurance / takaful (Malaisie)", type="Cession minoritaire → 100 % pour l'acquéreur",
         statut="Accord conclu",
         valorisation="1 100 M€ (pour 30,95 % ; 100 % valorisé ≈3 500 M€)", valorisationM=1100,
         multiple=2,
         theme="Bancassurance",
         resume="Ageas sort d'une coentreprise de 25 ans avec Maybank ; plus-value nette estimée ≈450 M€ pour Ageas. Multiple ≈P/B 2x des fonds propres IFRS 2025.",
         source="InsurTech.me Investment Intelligence Report — source secondaire, à recouper avec le communiqué Ageas",
         url="https://insurtech.me/reports/2026-08-08/"),
    dict(date="07/12/2025", acquereur="Ageas", cible="25 % d'AG Insurance (cédant : BNP Paribas Fortis) → 100 %",
         secteur="Assurance multibranche (Belgique)", type="Renforcement de participation majoritaire",
         statut="Annoncée",
         valorisation="1 900 M€ (pour 25 %)", valorisationM=1900, multiple=None,
         theme="Bancassurance",
         resume="Ageas rachète la participation minoritaire de BNP Paribas Fortis dans AG Insurance (n°1 belge vie et non-vie) et formalise un partenariat de long terme avec BNP Paribas.",
         source="Communiqué Ageas",
         url="https://www.ageas.com/en/newsroom/ageas-to-take-full-ownership-of-ag-insurance-and-formalise-long-term-partnership-with-bnp-paribas-3201192"),
    dict(date="05/12/2025", acquereur="Helvetia", cible="Baloise — fusion → Helvetia Baloise Holding",
         secteur="Assurance multibranche (Suisse/Europe)", type="Fusion entre égaux",
         statut="Finalisée",
         valorisation="11 600 M$ selon Milliman", valorisationM=conv(11600, "$"),
         multiple=None,
         theme="Consolidation domestique",
         resume="Plus grande opération vie/santé mondiale de 2025 selon Milliman (parité : 1 action Baloise = 1,0119 action HBAN). L'ensemble devient le premier assureur multibranche suisse (≈20 % de part de marché domestique).",
         source="Communiqué Helvetia Baloise ; Milliman",
         url="https://www.helvetia-baloise.com/corporate/hb/en/home/news-stories/publications/media-releases/2025/20251205.html"),
    dict(date="15/04/2026", acquereur="Standard Life PLC", cible="Aegon UK",
         secteur="Assurance vie & épargne (UK)", type="Acquisition (100 %)",
         statut="En cours (attente autorisations)",
         valorisation="2 000 M£", valorisationM=conv(2000, "£"), multiple=0.83,
         theme="Épargne & Retraite",
         resume="Standard Life a conclu un accord pour acquérir 100 % d'Aegon UK pour 2 Md£ (0,83x P/2025A Unrestricted Tier 1), financé par 750 M£ de cash et l'émission d'actions (15,3 % du groupe élargi). Finalisation attendue fin 2026.",
         source="Standard Life Half Year 2026 Results (document investisseur)",
         url=None),
    dict(date="29/06/2026", acquereur="Sixth Street", cible="Monument Re",
         secteur="Consolidation / réassurance vie (Europe)", type="Acquisition majoritaire",
         statut="Signée, clôture attendue fin 2026",
         valorisation=None, valorisationM=None, multiple=None,
         theme="Consolidation réassurance vie",
         resume="Sixth Street signe pour acquérir une participation majoritaire dans le consolidateur européen Monument Re ; Hannover Re reste actionnaire de référence. Vise à accélérer le transfert de risques de portefeuilles vie en Europe.",
         source="BusinessWire (communiqué Sixth Street)",
         url=None),
    dict(date="12/02/2026", acquereur="Admiral Group PLC (via Admiral Pioneer)", cible="Flock Ltd.",
         secteur="Insurtech / télématique flottes (UK)", type="Acquisition",
         statut="Complétée (clôture le 29/05/2026)",
         valorisation="80 M£", valorisationM=conv(80, "£"), multiple=None,
         theme="Assurtech — télématique flottes",
         resume="Admiral Group acquiert l'insurtech britannique Flock, télématique connectée pour flottes automobiles, pour intégrer sa tarification en temps réel au projet pilote Admiral Pioneer.",
         source="RBC Capital Market Research",
         url=None),
    dict(date="28/07/2026", acquereur="Cover Genius", cible="Friendsurance",
         secteur="Insurtech / bancassurance (Allemagne)", type="Acquisition",
         statut="Complétée",
         valorisation=None, valorisationM=None, multiple=None,
         theme="Embedded Insurance",
         resume="Cover Genius acquiert l'insurtech berlinoise Friendsurance, dont la technologie de distribution intégrée équipe des applications bancaires DACH — permet le déploiement de produits de protection conformes GDPR/PSD2.",
         source="Cover Genius (communiqué officiel)",
         url="https://www.covergenius.com/press/cover-genius-acquires-friendsurance-to-accelerate-embedded-protection-across-european-banking"),
    dict(date="31/07/2026", acquereur="Mapfre", cible="Tuio (38,9 %)",
         secteur="Insurtech — IA & tarification (Espagne)", type="Prise de participation minoritaire (augmentation de capital)",
         statut="Signée, attente feu vert réglementaire",
         valorisation=None, valorisationM=None, multiple=None,
         theme="Assurtech — IA & tarification",
         resume="Mapfre rachète 38,9 % de l'insurtech espagnole Tuio via une augmentation de capital, pour s'appuyer sur son expertise IA de tarification et financer son expansion Europe/Amérique latine.",
         source="Banco Sabadell Broker Research ; Beinsure",
         url=None),
    dict(date="07/07/2026", acquereur="KAPIA-RGI (filiale de RGI Group)", cible="Cegid Assurex Solutions",
         secteur="Logiciels d'assurance (France)", type="Acquisition d'activité",
         statut="Complétée",
         valorisation=None, valorisationM=None, multiple=None,
         theme="Logiciels d'assurance",
         resume="KAPIA-RGI rachète l'activité de Cegid Assurex dédiée aux éditeurs/distributeurs de produits d'assurance de personnes — consolide sa position sur les marchés vie/prévoyance/retraite en France (+20 clients institutionnels).",
         source="RGI Group (communiqué de presse)",
         url=None),
    dict(date="01/03/2026", acquereur="Athora", cible="Pension Insurance Corporation Group (PICG)",
         secteur="Épargne & retraite (UK)", type="Acquisition",
         statut="Complétée",
         valorisation=None, valorisationM=None, multiple=None,
         theme="Transfert de risques de pension",
         resume="Athora finalise l'acquisition de PICG, spécialiste de la gestion des régimes de retraite à prestations définies au Royaume-Uni, financée en partie par une levée de fonds propres de 3,5 Md€.",
         source="Athora HY 2026 Results (présentation investisseurs)",
         url=None),
    dict(date="18/09/2025", acquereur="Radian US Holdings, Inc.", cible="Inigo Ltd.",
         secteur="Assurance de spécialité / réassurance", type="Acquisition (100 %, numéraire)",
         statut="Complétée (clôture le 02/02/2026)",
         valorisation="1 670 M$", valorisationM=conv(1670, "$"), multiple=None,
         theme="Réassurance spécialités",
         resume="Radian rachète 100 % d'Inigo Ltd auprès de ses actionnaires (dont la Caisse de dépôt et placement du Québec et Qatar Investment Authority), en numéraire.",
         source="Financial Data — M&A Deals Database",
         url=None),
    dict(date="28/08/2026", acquereur="Tokio Marine HCC", cible="Direct Commercial Ltd.",
         secteur="Assurance dommages commerciale", type="Acquisition (100 %, numéraire)",
         statut="Complétée (clôture le 09/09/2026)",
         valorisation="359,94 M$", valorisationM=conv(359.94, "$"), multiple=7.7,
         theme="Assurance dommages commerciale",
         resume="Tokio Marine HCC finalise l'acquisition de Direct Commercial Ltd pour 359,94 M$, soit un multiple EV/Revenue de 7,70x — comparable de valorisation utile pour le segment commercial.",
         source="Financial Data — M&A Deals Database",
         url=None),
    dict(date="27/01/2026", acquereur="AUB Group Ltd.", cible="Prestige Insurance Holdings Ltd.",
         secteur="Courtage d'assurance (UK)", type="Acquisition (95,90 %, numéraire)",
         statut="Complétée (clôture le 10/03/2026)",
         valorisation="299,74 M$", valorisationM=conv(299.74, "$"), multiple=None,
         theme="Courtage — consolidation",
         resume="AUB Group (courtier australien) acquiert 95,90 % de Prestige Insurance Holdings, basé au Royaume-Uni, en numéraire.",
         source="Financial Data — M&A Deals Database",
         url=None),
    dict(date="16/02/2026", acquereur="Reale Mutua di Assicurazioni", cible="Lifenet Srl",
         secteur="Services de santé (Italie)", type="Acquisition majoritaire (78 %, numéraire)",
         statut="Complétée (clôture le 21/07/2026)",
         valorisation="710,97 M$", valorisationM=conv(710.97, "$"), multiple=None,
         theme="Intégration verticale santé",
         resume="Reale Mutua acquiert 78 % de l'entreprise italienne de cliniques et services médicaux Lifenet — intégration verticale d'un assureur santé vers les soins.",
         source="Financial Data — M&A Deals Database",
         url=None),
]

NOUVELLES_RUMEURS = [
    dict(date="15/05/2026", acteurs_pressentis="Intact Financial (Canada)", cible="Hiscox (UK)",
         secteur="Assurance de spécialités / Lloyd's", valorisation_indicative=None, valorisationM=None, multiple=None,
         theme="Consolidation spécialités",
         statut="Rumeur de marché — aucune offre formelle, discussions exploratoires rapportées",
         resume="Selon Insurance Post, Intact étudierait une offre sur Hiscox ; le titre a bondi d'environ 15 %. Aucune confirmation des deux groupes à ce stade.",
         source="Insurance Business ; Global Banking & Finance (citant Insurance Post)",
         url="https://www.insurancebusinessmag.com/uk/news/mergers-acquisitions/intacthiscox-deal-talk-adds-to-wave-of-foreign-bids-for-uk-names-report-575592.aspx"),
    dict(date="01/07/2026", acteurs_pressentis="KKR", cible="Partenariats d'assurance vie britanniques (non nommés)",
         secteur="Épargne & retraite / Pension Risk Transfer", valorisation_indicative=None, valorisationM=None, multiple=None,
         theme="Transfert de risques de pension",
         statut="En cours de discussions",
         resume="KKR négocierait des partenariats stratégiques avec de grands assureurs vie britanniques pour se positionner sur le marché des transferts de risques de pension (PRT), en déployant du capital via des véhicules gérés par KKR plutôt qu'en acquérant un acteur — stratégie inspirée du projet étudié avec Standard Life.",
         source="Financial Times via MT Newswires",
         url=None),
]


def main() -> int:
    shutil.copy(FICHIER, FICHIER.with_suffix(".json.bak"))
    data = json.loads(FICHIER.read_text(encoding="utf-8"))

    existants_cibles = {(d["acquereur"], d["cible"]) for d in data["confirmees"]}

    # Enrichissement additif : multiple manquant pour At-Bay (purement additif, ne remplace rien).
    for d in data["confirmees"]:
        if d["cible"] == "At-Bay" and d.get("multiple") is None:
            d["multiple"] = 2
            d["resume"] += " Multiple ≈2x les primes brutes émises (278 M$), selon Insurance Business."

    ajoutees = 0
    for d in NOUVELLES_CONFIRMEES:
        if (d["acquereur"], d["cible"]) in existants_cibles:
            continue
        data["confirmees"].append(d)
        ajoutees += 1

    ajoutees_rumeurs = 0
    existants_rumeurs = {(r.get("acteurs_pressentis"), r["cible"]) for r in data["rumeurs"]}
    for r in NOUVELLES_RUMEURS:
        if (r["acteurs_pressentis"], r["cible"]) in existants_rumeurs:
            continue
        data["rumeurs"].append(r)
        ajoutees_rumeurs += 1

    data["derniereMaj"] = "21/09/2026"

    FICHIER.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Confirmées : {len(data['confirmees']) - ajoutees} → {len(data['confirmees'])} (+{ajoutees})")
    print(f"Rumeurs    : {len(data['rumeurs']) - ajoutees_rumeurs} → {len(data['rumeurs'])} (+{ajoutees_rumeurs})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
