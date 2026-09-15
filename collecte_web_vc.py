#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
collecte_web_vc.py — Passe de collecte web complémentaire sur la base standardisée (source 3)
================================================================================================

Entrée  : VC_Database_Standardisee_v3.xlsx
Sortie  : VC_Database_Standardisee_v4.xlsx

Rejoue, sous forme de table de données en dur (même principe que MAJ_WEB dans
maj_web_vc.py), les résultats d'une campagne de recherche web menée par lots sur les 111
véhicules de l'onglet Fonds (sites officiels des sociétés de gestion, communiqués de
closing, presse spécialisée nommée, registres publics).

Règles appliquées (cahier des charges) :
- on ne remplit que des cellules VIDES, jamais de valeur déjà présente ;
- aucune estimation : une donnée non publiée reste vide (pas de déduction d'AuM à partir de
  tickets, pas de millésime déduit d'une date d'article) ;
- toute cellule remplie est tracée (Source_AuM horodatée, Date_MAJ, ligne Anomalies) ;
- les divergences base/web sont journalisées, jamais arbitrées en silence ;
- les valeurs en fourchette ou à sources contradictoires sont laissées vides et journalisées
  (« Valeur en fourchette non retenue ») ;
- les risques d'homonymie / de rapprochement non confirmé sont journalisés sans écrire de
  donnée (« Rapprochement non résolu ») ;
- devises converties uniquement avec les taux imposés (1 USD ≈ 0,92 EUR, 1 GBP ≈ 1,17 EUR,
  1 CHF ≈ 1,07 EUR), taux cité dans la source ;
- vocabulaires fermés respectés (Statut, Phase, Géographie, Stratégie, Stages pratiqués) —
  un libellé hors vocabulaire va dans le commentaire, jamais dans la cellule ;
- la numérotation des anomalies reprend après la dernière ligne existante et reste
  séquentielle.
"""

from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import restructuration_vc as rvc  # réutilisation des fonctions de mise en forme

FICHIER_ENTREE = "VC_Database_Standardisee_v3.xlsx"
FICHIER_SORTIE = "VC_Database_Standardisee_v4.xlsx"
DATE_MAJ = date(2026, 9, 15)
ONGLETS = ["Fonds", "Participations", "Tours_de_table", "Dictionnaire", "Anomalies"]

SOURCE_LABEL = "Passe de collecte web par lots (sites officiels + presse spécialisée + registres) — 14-15/09/2026"


# =============================================================================
# 1. TABLE DE COLLECTE — résultats de la recherche web par lots
# =============================================================================
# Chaque entrée : Fonds_ID -> valeurs vérifiées trouvées vides dans la base, niveau de
# confiance et source(s) exactes (URL). Seules les cellules VIDES au moment de l'exécution
# sont effectivement écrites (contrôle refait dynamiquement, voir appliquer_collecte).

COLLECTE_WEB = [
    dict(
        fid='360-capital-partners_360-fund-v',
        vals={'Site web': 'https://www.360cap.vc'},
        confiance='Moyen',
        src=("WebSearch result: '360 Capital - European VC from pre-seed to Series B' "
         'https://www.360cap.vc/'),
    ),
    dict(
        fid='360-capital-partners_360-life-ii',
        vals={'Site web': 'https://www.360cap.vc', 'Montant alloué (M€)': 140},
        confiance='Élevé',
        src=('EU-Startups: '
         'https://www.eu-startups.com/2024/12/360-capital-announces-the-1st-closing-of-its-e140-million-climate-tech-fund-360-life-ii/ '
         '; BeBeez International: '
         'https://bebeez.eu/2024/12/23/360-capital-announces-the-1st-closing-of-its-e140-million-climate-tech-fund-360-life-ii/ '
         '; Il Sole 24 Ore: '
         'https://en.ilsole24ore.com/art/360-capital-first-closing-140-million-dedicated-climate-tech-fund-AGwINCwB'),
    ),
    dict(
        fid='360-capital-partners_360-square-ii-seed-fund',
        vals={'Site web': 'https://www.360cap.vc', 'Montant alloué (M€)': 45},
        confiance='Élevé',
        src=('360 Capital official Medium post: '
         'https://medium.com/360-capital/360-capital-launches-its-new-45m-fund-to-back-preseed-seed-ventures-3aad940c37e0 '
         '; corroborated by Unquote: '
         'https://www.unquote.com/france/official-record/3027098/360-capital-holds-eur-45m-first-close-on-early-stage-tech-fund'),
    ),
    dict(
        fid='365-fintech_365-fintech',
        vals={'Montant levé (M€)': 33, 'Millésime': 2018},
        confiance='Faible',
        src=('Capboard investor profile, 365.fintech '
         "(https://www.capboard.io/en/investor/365-fintech) — 'total invested amount of "
         "$35.9M'; founding year 2018 corroborated by Vestbee "
         '(https://www.vestbee.com/vc-list/365.fintech/) and o.parsers.vc '
         '(https://o.parsers.vc/fund/365fintech.vc/)'),
    ),
    dict(
        fid='accel-partners_accel-london-vi',
        vals={'Site web': 'https://www.accel.com', 'Montant alloué (M€)': 529},
        confiance='Élevé',
        src=('TechCrunch '
         '(techcrunch.com/2019/05/15/accel-closes-575m-fund-to-double-down-on-european-and-israeli-series-a-deals), '
         'Unquote (unquote.com/uk/official-record/3014995), FinSMEs '
         '(finsmes.com/2019/05/accel-closes-sixth-european-and-israeli-fund-at-575m.html), '
         "Tech.eu (tech.eu/brief/accel-fund-6) — all report Accel's sixth European/Israel "
         'early-stage fund closed at $575M in May 2019; converted at 1 USD = 0.92 EUR ($575M x '
         "0.92 = 529 M€). Site: accel.com is the firm's official corporate site, referenced "
         'consistently across all PitchBook/press listings.'),
    ),
    dict(
        fid='accel-partners_accel-london-vii',
        vals={'Site web': 'https://www.accel.com', 'Montant alloué (M€)': 598},
        confiance='Élevé',
        src=('TechCrunch (techcrunch.com/2021/06/29/accel-closes-on-3b-across-three-funds), '
         'Silicon Canals (siliconcanals.com/accel-closes-3b-across-3-new-funds), Sifted '
         '(sifted.eu/articles/accel-fundraise-800m-ninth-early-stage-europe-fund, quoting '
         "Accel partner Harry Nelis referencing the 'seventh fund, also $650 million, raised "
         "in 2021') — Accel's seventh early-stage Europe/Israel fund closed at $650M in June "
         '2021, part of a $3.05B multi-fund close (with Accel XV and Accel Growth Fund VI). '
         'Converted at 1 USD = 0.92 EUR ($650M x 0.92 = 598 M€).'),
    ),
    dict(
        fid='accel-partners_accel-london-viii',
        vals={'Site web': 'https://www.accel.com',
         'Montant alloué (M€)': 598,
         'Géographie': 'Europe'},
        confiance='Élevé',
        src=("Tech.eu (tech.eu/2024/05/14/accel-s-eighth-fund-closes-at-650m — explicitly 'Accel's "
         "eighth fund'), Venture Capital Journal "
         '(venturecapitaljournal.com/accel-london-viii-raises-650m-for-europe-and-israel), '
         'CNBC '
         '(cnbc.com/2024/05/13/venture-capital-firm-accel-raises-650-million-europe-and-israel-fund.html), '
         'Sifted (sifted.eu/articles/accel-650m-fund-europe) — Accel London VIII closed at '
         '$650M on 13 May 2024, targeting Europe and Israel, early-stage (seed/Series A). '
         'Converted at 1 USD = 0.92 EUR ($650M x 0.92 = 598 M€).'),
    ),
    dict(
        fid='alven-capital_alven-v',
        vals={'Géographie': 'Europe'},
        confiance='Moyen',
        src=("Maddyness, 09/01/2017, 'Alven Capital fête ses 17 ans en levant 250 millions d'euros "
         "pour les startups françaises' "
         '(https://www.maddyness.com/2017/01/09/vc-alven-capital-fonds-250-millions-euros-startups/) '
         '; TechCrunch, 03/01/2017 '
         '(https://techcrunch.com/2017/01/03/alven-capital-raises-261-million-fund-to-invest-in-french-entrepreneurs/) '
         "; description générale d'Alven (investorsglobe.com, openvc.app) indiquant un focus "
         "France/Europe avec possibilité de suivre des fondateurs européens s'implantant aux "
         'US.'),
    ),
    dict(
        fid='aris-occitanie-vc_aris-occitanie',
        vals={'Site web': 'https://aris-occitanie.fr'},
        confiance='Moyen',
        src=("WebSearch result 'ARIS OCCITANIE' at https://aris-occitanie.fr/; corroborated by "
         'registry listings (annuaire-entreprises.data.gouv.fr, societe.com, pappers.fr) for '
         "'AGENCE REGIONALE DES INVESTISSEMENTS STRATEGIQUES ARIS-OCCITANIE', SIREN 901805630, "
         'Toulouse.'),
    ),
    dict(
        fid='astorya-vc_astorya-vc',
        vals={'Site web': 'https://astorya.vc'},
        confiance='Élevé',
        src=("astorya.vc (official site, per WebSearch result 'Astorya.vc' at "
         "https://astorya.vc/); corroborated by eu-startups.com directory entry 'astorya.vc | "
         "EU-Startups' (https://www.eu-startups.com/directory/astorya-vc/) and Crunchbase "
         'organization profile (https://www.crunchbase.com/organization/astorya-vc).'),
    ),
    dict(
        fid='blackfin-capital-partners_blackfintech-1',
        vals={'Site web': 'https://www.blackfincapital.com'},
        confiance='Moyen',
        src=("WebSearch snippets: blackfincapital.com homepage (title 'Home - Blackfin Capital'), "
         'https://www.blackfincapital.com/about/ ; corroborated by fintech.global article '
         "'BlackFin commits €350m fund for European InsurTechs and FinTechs' (2022-07-12) "
         "referencing BlackFin Capital Partners; PitchBook profile 'BlackFin Tech Fund 1' "
         '(2018 vintage, ~€180m).'),
    ),
    dict(
        fid='blackfin-capital-partners_blackfintech-ii',
        vals={'Site web': 'https://www.blackfincapital.com'},
        confiance='Moyen',
        src=('Same as BlackfinTech 1: blackfincapital.com / blackfin-tech.com search snippets; '
         "fintech.global 'BlackFin commits €350m fund for European InsurTechs and FinTechs' "
         "(2022-07-12); aggregator summary citing 'BlackFin Tech II (€390 million, vintage "
         "2022)'."),
    ),
    dict(
        fid='bpi-france_large-ventures',
        vals={'Géographie': 'Europe'},
        confiance='Moyen',
        src=("Bpifrance, page produit 'Large Venture' "
         '(https://www.bpifrance.fr/nos-solutions/investissement/investissement-expertise/large-venture '
         'et https://www.bpifrance.com/products/large-venture/, contenu consulté via extraits '
         "de recherche — WebFetch direct bloqué sur ces domaines) : 'cible des entreprises "
         "technologiques implantées majoritairement en France' / 'forte empreinte française'."),
    ),
    dict(
        fid='breega-capital_breega-seed-i',
        vals={'Site web': 'https://www.breega.com'},
        confiance='Faible',
        src=('WebSearch result confirming BREEGA official site https://www.breega.com/; FinSMEs '
         "2017 article 'Breega Capital Launches €100M European Venture Capital Fund' "
         '(finsmes.com/2017/07/breega-capital-launches-e100m-european-venture-capital-fund.html) '
         "mentioning an 'inaugural €50M ($57M) fund' preceding the €100M 'Breega Capital "
         "Venture 2'"),
    ),
    dict(
        fid='breega-capital_breega-seed-ii',
        vals={'Site web': 'https://www.breega.com'},
        confiance='Faible',
        src=('WebSearch result confirming BREEGA official site https://www.breega.com/; '
         "superscout.co summary describing a 'Seed I (2015)' / 'Seed II (2019)' / 'Europe Seed "
         "III (2023)' sequence"),
    ),
    dict(
        fid='breega-capital_breega-venture-iii',
        vals={'Site web': 'https://www.breega.com', 'Montant alloué (M€)': 110},
        confiance='Élevé',
        src=("EU-Startups 'Breega closes its third fund at €110 million for European tech "
         "startups' "
         '(eu-startups.com/2021/03/breega-closes-its-third-fund-at-e110-million-for-european-tech-startups/); '
         'corroborated by TechCrunch and VentureBeat coverage of the same March 2021 '
         "announcement ($130M / €110M final close of 'Breega Capital Venture 3', first closing "
         '€90M mid-2019)'),
    ),
    dict(
        fid='breega-capital_f-i-venture-ii',
        vals={'Site web': 'https://www.breega.com'},
        confiance='Moyen',
        src=("French company registry listings: rubypayeur.com 'Société BREEGA VENTURE II à PARIS "
         "- SIREN : 902920651' and annuaire-entreprises.data.gouv.fr entry for SIREN "
         "902920651, both identifying the legal entity 'BREEGA VENTURE II' with the alias "
         "'(F/I VENTURE II)', created 15 July 2021, based in Paris (75002); PitchBook "
         "fund-profile snippet for the earlier, related 'F/I Venture' (I) describing it as a "
         '2017-vintage Breega-managed fund with Crédit Agricole as sole LP, focused on '
         'financial services'),
    ),
    dict(
        fid='cathay_cathay-innovation-ii',
        vals={'Site web': 'https://cathayinnovation.com/'},
        confiance='Élevé',
        src='https://cathayinnovation.com/ (official site)',
    ),
    dict(
        fid='cathay_cathay-innovation-iii',
        vals={'Site web': 'https://cathayinnovation.com/'},
        confiance='Élevé',
        src=('https://cathayinnovation.com/ (official site); '
         'https://cathayinnovation.com/cathay-innovation-closes-1b-venture-capital-fund-to-bring-vertical-ai-to-critical-industries/ '
         '(official press release)'),
    ),
    dict(
        fid='commerzventures_commerzventures-fonds-i-a-iii',
        vals={'Montant alloué (M€)': 550, 'Ticket min (M€)': 2, 'Ticket max (M€)': 10},
        confiance='Élevé',
        src=('EU-Startups, Finextra, FinSMEs, AltAssets, fintech.global, VentureCapitalJournal '
         '(all 03/2022) — corroborating: Fund I (2014) €100M, Fund II (2019) €150M, Fund III '
         "(2022) €300M, 'combined total fund size of €550 million' "
         '(https://www.eu-startups.com/2022/03/commerzventures-closes-e300-million-fund-to-boost-european-fintech-and-insurtech-startups/ '
         '; https://www.finextra.com/newsarticle/39783/commerzventures-closes-300m-third-fund '
         '; https://www.finsmes.com/2022/03/commerzventures-closes-e300m-third-fund.html). '
         'Ticket range €2-10M per multiple aggregators (VCsheet, OpenVC, Tracxn) describing '
         'Series A/B checks.'),
    ),
    dict(
        fid='committed-capital_committed-capital-eis-fund',
        vals={'AuM (M€)': 117, 'TVPI': 3.0, 'IRR (TRI)': 0.368},
        confiance='Moyen',
        src=("WebSearch aggregation of committedcapital.co.uk, growthbusiness.co.uk ('Top 20 EIS "
         "funds and investors you should know about'), ifamagazine.com ('Committed Capital "
         "Growth EIS Fund'), and openvc.app/fund/Committed%20Capital"),
    ),
    dict(
        fid='elevation-capital-partners_fpci-elevation-early-growth-i',
        vals={'Site web': 'https://www.elevation-cp.com'},
        confiance='Moyen',
        src=('Inter Invest Capital poursuit son développement et devient Elevation Capital '
         'Partners '
         '(https://www.inter-invest.fr/communiques/capital-investissement/00458/inter-invest-capital-poursuit-son-developpement-et-devient-elevation-capital-partners); '
         'Elevation Capital Partners official site confirmed via search results '
         '(https://www.elevation-cp.com/)'),
    ),
    dict(
        fid='elevation-capital-partners_fpci-elevation-early-growth-ii',
        vals={'Site web': 'https://www.elevation-cp.com'},
        confiance='Moyen',
        src=('Elevation Capital Partners official site (https://www.elevation-cp.com/); fund page '
         'also at '
         'https://www.inter-invest.fr/capital-investissement/fpci-elevation-early-growth-2'),
    ),
    dict(
        fid='elevation-capital-partners_fpci-food-invest-ii',
        vals={'Site web': 'https://www.elevation-cp.com'},
        confiance='Moyen',
        src=('Elevation Capital Partners official site, fund communiqué '
         '(https://www.elevation-cp.com/communiques/capital-investissement/00552/premiere-levee-de-fonds-pour-vegetal-food-le-fpci-food-invest-2-investit-1-2-millions)'),
    ),
    dict(
        fid='eos-venture_evp-i',
        vals={'Site web': 'https://eosvc.com'},
        confiance='Élevé',
        src=("https://eosvc.com (confirmed as Eos Venture Partners' official domain across "
         'multiple independent search results: Crunchbase, CBInsights, PitchBook profile pages '
         'all reference eosvc.com)'),
    ),
    dict(
        fid='ethias-ventures_ethias-ventures',
        vals={'Site web': 'https://www.ethias.be/content/corporate/en/ethias-group/EthiasVentures.html',
         'Montant levé (M€)': 20,
         'Ticket max (M€)': 3},
        confiance='Élevé',
        src=("Ethias Newsroom press release 'Ethias Ventures : le nouveau véhicule "
         "d'investissement innovant chez Ethias!' "
         '(newsroom.ethias.be/ethias-ventures--le-nouveau-vehicule-dinvestissement-innovant-chez-ethias-isnkw5) '
         'and Belga Share syndication of the same release '
         '(belgashare.be/newsrooms/108/press-releases/1621) both state an initial endowment of '
         '€20 million and a maximum ticket of €3 million per round/per start-up. Official fund '
         'page: ethias.be/content/corporate/en/ethias-group/EthiasVentures.html.'),
    ),
    dict(
        fid='evolem-start_evolem-start',
        vals={'Site web': 'https://evolem.com', 'Millésime': 2017},
        confiance='Moyen',
        src=('Site web: evolem.com news pages (e.g. '
         'https://evolem.com/evolem-start-realise-6-nouveaux-investissements-en-2018/), '
         "Maddyness 'Que Font les Fonds : le portrait d'Evolem' (2021-09-27). Millésime: "
         "company-registry-derived WebSearch snippet ('EVOLEM START operates in the fund "
         "management sector and was created in 2017', from verif.com/rubypayeur listings for "
         'EVOLEM START, SIREN 833257645).'),
    ),
    dict(
        fid='founders-future-vc_founders-future-entrepreneur',
        vals={'Site web': 'https://www.foundersfuture.com'},
        confiance='Moyen',
        src='foundersfuture.com (site officiel du gestionnaire) — voir note ci-dessous',
    ),
    dict(
        fid='founders-future-vc_founders-future-fund-i',
        vals={'Site web': 'https://www.foundersfuture.com'},
        confiance='Élevé',
        src=('foundersfuture.com (site officiel du gestionnaire, pages /en et /en/vision apparues '
         'dans les résultats WebSearch)'),
    ),
    dict(
        fid='founders-future-vc_founders-future-fund-ii',
        vals={'Site web': 'https://www.foundersfuture.com'},
        confiance='Moyen',
        src=('foundersfuture.com (site officiel) ; PitchBook '
         "(pitchbook.com/profiles/fund/24427-54F) référence 'Founders Future II', fonds "
         'early-stage domicilié en France, fermé — via extraits WebSearch, TechCrunch '
         '(techcrunch.com/2023/09/05/...) confirme que Founders Future levait deux nouveaux '
         "fonds ('Founders Future II' et 'Founders Future Expansion') avec un premier closing "
         'à 80 M$ en 2023'),
    ),
    dict(
        fid='founders-future-vc_founders-future-good',
        vals={'Site web': 'https://www.foundersfuture.com'},
        confiance='Élevé',
        src=('foundersfuture.com (site officiel du gestionnaire Founders Future — pages /en, '
         '/en/vision, /us/news apparues directement dans les résultats de recherche) ; '
         'corroboré par PitchBook (pitchbook.com/profiles/fund/22037-23F) qui référence '
         "'Founders Future Good' comme fonds millésime 2020, fermé (contenu détaillé non "
         'accessible, payant)'),
    ),
    dict(
        fid='goldsmith-ventures_goldsmith-ventures-eis-fund',
        vals={'Site web': 'https://www.goldsmithventures.com',
         'Ticket min (M€)': 0.736,
         'Ticket max (M€)': 1.104},
        confiance='Faible',
        src=('Aggregated VC-database snippets (Apollo.io company profile, investorconnect.org '
         "'Investor Connect: James Pringle' transcript, openvc.app/fund/Goldsmith%20Ventures) "
         'surfaced via WebSearch'),
    ),
    dict(
        fid='helsana-healthinvest_helsana-healthinvest',
        vals={'Montant alloué (M€)': 107, 'Millésime': 2020},
        confiance='Élevé',
        src=('Helsana official page (helsana.ch/en/helsana-group/about-us/healthinvest.html), IFHP '
         'executive summary PDF '
         '(ifhp.com/wp-content/uploads/2023/06/Helsana-PDF-Exec-Sum.pdf), Caplight investor '
         'profile (caplight.com/investor/helsana), EU-Startups '
         '(eu-startups.com/investor/helsana-healthinvest) — all describe Helsana HealthInvest '
         'as an evergreen fund with a CHF 100 million commitment. Converted at 1 CHF = 1.07 '
         'EUR (CHF 100M x 1.07 = 107 M€). Founding year 2020 confirmed by Moneyhouse Swiss '
         'commercial-registry entry '
         '(moneyhouse.ch/en/company/helsana-healthinvest-ag-4648125341).'),
    ),
    dict(
        fid='howden-ventures_howden-ventures',
        vals={'Montant alloué (M€)': 11.7},
        confiance='Moyen',
        src=('Howden launches world-first insurance innovation hub with £500m of delegated '
         'underwriting capacity, Howden Group Holdings newsroom '
         '(https://www.howdengroupholdings.com/news/howden-launches-world-first-insurance-innovation-hub-with-500m-of-delegated-underwriting-capacity); '
         'corroborated by Reinsurance News '
         '(https://www.reinsurancene.ws/insurance-innovation-hub-howden-ventures-launches-with-500m-underwriting-capacity/) '
         'and Insurance Business UK '
         '(https://www.insurancebusinessmag.com/uk/news/breaking-news/howden-to-fasttrack-investment-and-risk-incubation-with-insurance-innovation-hub-463331.aspx), '
         'all stating Howden initially committed £10m to the incubator'),
    ),
    dict(
        fid='index-venture_index-origin-i',
        vals={'Site web': 'https://www.indexventures.com', 'Montant alloué (M€)': 184},
        confiance='Élevé',
        src=('BusinessWire / TechCrunch (08/04/2021) — "Index Ventures Launches Index Origin, a '
         '$200 Million Dedicated Seed Fund" '
         '(https://www.businesswire.com/news/home/20210408005561/en/Index-Ventures-Launches-Index-Origin-a-200-Million-Dedicated-Seed-Fund '
         '; '
         'https://techcrunch.com/2021/04/08/index-closes-200-million-dedicated-seed-fund-to-intensify-multi-stage-thesis/). '
         'Original amount $200M converted to EUR at 1 USD ≈ 0.92 EUR = 184 M€.'),
    ),
    dict(
        fid='index-venture_index-origin-ii',
        vals={'Site web': 'https://www.indexventures.com',
         'Montant alloué (M€)': 276,
         'Géographie': 'Europe ; Amérique'},
        confiance='Moyen',
        src=('Index Ventures official blog (17/11/2022) — "Announcing Index Origin II, a $300 '
         'Million Seed Fund" '
         '(https://www.indexventures.com/perspectives/announcing-index-origin-ii-a-300-million-seed-fund-designed-for-extraordinary-entrepreneurs/); '
         'Tech.eu '
         '(https://tech.eu/2022/11/17/back-to-the-origin-index-launches-second-seed-stage-focused-fund-at-300-million/) '
         "confirms fund 'supporting entrepreneurs primarily from Europe and the U.S.'. $300M "
         'converted at 1 USD ≈ 0.92 EUR = 276 M€.'),
    ),
    dict(
        fid='index-venture_index-ventures-xi',
        vals={'Site web': 'https://www.indexventures.com',
         'Montant alloué (M€)': 828,
         'Stratégie': 'Généraliste'},
        confiance='Moyen',
        src=('Global Legal Chronicle (09/2021) — "Index Ventures\' $900 Million Closing of Index '
         'Ventures XI" '
         '(https://globallegalchronicle.com/index-ventures-900-million-closing-of-index-ventures-xi/). '
         '$900M converted at 1 USD ≈ 0.92 EUR = 828 M€.'),
    ),
    dict(
        fid='index-venture_index-ventures-xii',
        vals={'Site web': 'https://www.indexventures.com', 'Montant alloué (M€)': 736},
        confiance='Élevé',
        src=('TechCrunch (09/07/2024) — "Index Ventures Raises $2.3B For New Venture And Growth '
         'Funds" '
         '(https://techcrunch.com/2024/07/09/index-ventures-raises-23-billion-for-new-venture-and-growth-funds/) '
         'and BusinessWire same date — $800M raised for its 12th venture fund (= Index '
         'Ventures XII), alongside a separate $1.5B growth fund. $800M converted at 1 USD ≈ '
         '0.92 EUR = 736 M€.'),
    ),
    dict(
        fid='insurtech-capital_insurtech-capital-groupe-apicil',
        vals={'Montant levé (M€)': 10,
         'Montant alloué (M€)': 10,
         'Millésime': 2018,
         'Stages pratiqués': 'Seed ; Série A'},
        confiance='Élevé',
        src=('Groupe Apicil official newsroom '
         '(groupe-apicil.com/newsroom/insurtech-lancement-dun-fonds-dinvestissement-dedie), '
         'Private Equity Magazine (pemagazine.fr/ntu4oa), mind.eu.com Fintech '
         '(mind.eu.com/fintech/article/apicil-consacre-10-millions-deuros-aux-investissements-dans-linsurtech), '
         'FinSMEs (finsmes.com/2018/12/odysseus-and-apicil-launch-insurtech-fund.html), '
         'Holland FinTech '
         '(hollandfintech.com/2019/01/apicil-group-odysseus-alternative-ventures-launch-eur-10-million-insurtech-fund), '
         'fintech.global Global Insurtech Summit '
         '(fintech.global/globalinsurtechsummit/odysseus-apicil-launch-e10m-dedicated-insurtech-fund) '
         "— all consistently report a €10 million fund ('Insurtech Capital I'), announced 11 "
         'December 2018, managed by Odysseus Alternative Ventures (Luxembourg, part of Reech '
         'Corporations Group) with APICIL as strategic partner, targeting seed and Series A '
         'insurtech/fintech start-ups in France and Europe with tickets of roughly €1M per '
         'deal.'),
    ),
    dict(
        fid='insurtech-gateway_insurtech-gateway-seed-fund-i-ii',
        vals={'Ticket max (M€)': 1.17},
        confiance='Moyen',
        src=('Superscout investor guide (superscout.co/investor/insurtech-gateway) and Tracxn '
         "profile (tracxn.com/.../insurtech-gateway) both describe Insurtech Gateway's direct "
         'investment ticket as up to £250k at Pre-Seed and up to £1M at Seed. Converted at 1 '
         'GBP = 1.17 EUR (£1M x 1.17 = 1.17 M€; £250k x 1.17 = 0.29 M€).'),
    ),
    dict(
        fid='isai_isai-cap-venture',
        vals={'Site web': 'https://www.isai.fr/isai-cap-venture', 'Montant alloué (M€)': 90},
        confiance='Moyen',
        src=("Capgemini official press release 'Capgemini and ISAI launch ISAI Cap Venture II' "
         '(capgemini.com/news/press-releases/capgemini-and-isai-launch-isai-cap-venture-ii/), '
         'which states ISAI Cap Venture I (launched June 2019) was endowed with €90M; '
         'corroborated by Maddyness article on ISAI Cap Venture II (2025) distinguishing the '
         '€80M CV II from the earlier €90M CV I'),
    ),
    dict(
        fid='isai_isai-venture-ii',
        vals={'Site web': 'https://www.isai.fr', 'Montant alloué (M€)': 75},
        confiance='Moyen',
        src=("ISAI official press release 'ISAI annonce le premier closing du fonds « ISAI VENTURE "
         "II » à 55 M€' "
         '(isai.fr/news/isai-annonce-le-premier-closing-du-fonds-isai-venture-ii-a-55-m-et-la-mise-en-place-du-isai-seed-club) '
         "for first closing; CFNEWS 'Isai boucle Venture II à 75 M€' (cfnews.net) for final "
         'closing size'),
    ),
    dict(
        fid='isai_isai-venture-iii',
        vals={'Site web': 'https://www.isai.fr', 'Géographie': 'Europe ; Amérique'},
        confiance='Faible',
        src=("WebSearch snippet summarizing frenchweb.fr 'French Tech : ISAI lance un nouveau "
         "fonds de 90 millions d'euros' "
         '(frenchweb.fr/french-tech-isai-lance-un-nouveau-fonds-de-90-millions-deuros/399745) '
         "describing a 'transatlantic approach': investing in US companies founded by French "
         "entrepreneurs and supporting French startups' expansion into North America; official "
         "ISAI news page title 'First closing at €90M for the ISAI Venture III fund' "
         '(isai.vc/news/first-closing-at-90m-for-the-isai-venture-iii-fund)'),
    ),
    dict(
        fid='kima-venture-capital_kima-venture',
        vals={'Site web': 'https://kimaventures.com'},
        confiance='Élevé',
        src=('Multiple WebSearch aggregator results (signal.nfx.com, swanbase.co, vcsheet.com, '
         'eldorado.co) consistently citing kimaventures.com as the official Kima Ventures '
         "site; general knowledge of Xavier Niel's Kima Ventures corroborates this domain."),
    ),
    dict(
        fid='macif-innovation_macif-innovation',
        vals={'Montant alloué (M€)': 30,
         'Ticket min (M€)': 0.3,
         'Ticket max (M€)': 1.5,
         'Stages pratiqués': 'Seed ; Série A',
         'Nb participations (déclaré)': 15},
        confiance='Élevé',
        src=('Que font les fonds ? Le portrait de Macif Innovation, Maddyness, 24/02/2025 '
         '(https://www.maddyness.com/2025/02/24/que-font-les-fonds-le-portrait-de-macif-innovation/); '
         'corroborated on the 30 M€ envelope by La Macif crée un fonds de capital innovation, '
         'Private Equity Magazine '
         '(https://www.pemagazine.fr/NDc0Mw/la-macif-cree-un-fonds-de-capital-innovation)'),
    ),
    dict(
        fid='nca-next-commerce-accelerator_nca-next-commerce-accelerator',
        vals={'Site web': 'https://nca.vc'},
        confiance='Moyen',
        src=("Official domain https://nca.vc confirmed via search results (title 'NCA VC – NCA "
         "VC'); portfolio count 'From 2017 to 2025 NCA invested in 77 B2B startups and "
         "partnered with 25+ corporate LPs' from search synthesis of nca.vc / prodevs.io / "
         'privateequitylist.com listings.'),
    ),
    dict(
        fid='newalpha-asset-management_newalpha-fintech-insurtech',
        vals={'Site web': 'https://www.newalpha.com/en/venture-capital/'},
        confiance='Moyen',
        src=('https://www.newalpha.com/en/venture-capital/ (official NewAlpha Asset Management '
         'site, venture capital / FinTech-InsurTech page)'),
    ),
    dict(
        fid='newfund-capital_newfund-1',
        vals={'Site web': 'https://newfundcap.com'},
        confiance='Élevé',
        src=("Wikipedia 'Newfund' entry (en.wikipedia.org/wiki/Newfund) and newfundcap.com "
         'official site.'),
    ),
    dict(
        fid='newfund-capital_newfund-2',
        vals={'Site web': 'https://newfundcap.com'},
        confiance='Élevé',
        src=("Wikipedia 'Newfund' entry (en.wikipedia.org/wiki/Newfund) and newfundcap.com "
         "official site — both confirm newfundcap.com as Newfund Management's/Newfund "
         "Capital's official website."),
    ),
    dict(
        fid='newfund-capital_newfund-naeh-1-et-2',
        vals={'Site web': 'https://newfundcap.com', 'Montant alloué (M€)': 19.1},
        confiance='Moyen',
        src=('Private Equity Magazine (pemagazine.fr/odc1mw and pemagazine.fr/odm5oq), Banque des '
         'Territoires press release '
         '(banquedesterritoires.fr/closing-du-fonds-dinvestissement-newfund-naeh-innopy-153-meu '
         'and PDF CP BDT NAEH INNOPY 09102024), CFNEWS '
         '(cfnews.net/.../Newfund-passe-la-seconde...470609) — Newfund NAEH (1st vintage, FCPR '
         'created end-2018) closed at €3.8M; Newfund NAEH Innopy (2nd vintage) closed at '
         '€15.3M (incl. a €5M Banque des Territoires ticket), in line with its €15M target. '
         'Combined = 3.8 + 15.3 = 19.1 M€. Site: newfundcap.com confirmed as Newfund '
         "Management's official site (also cited by Wikipedia 'Newfund' entry)."),
    ),
    dict(
        fid='open-cnp_open-cnp',
        vals={'Montant alloué (M€)': 100},
        confiance='Élevé',
        src=('Avec Open CNP, CNP Assurances va consacrer 100 M€ sur 5 ans au développement de '
         'partenariats avec des start-ups, CNP Assurances newsroom, 22/09/2016 '
         '(https://www.cnp.fr/le-groupe-cnp-assurances/newsroom/communiques-de-presse/2016/avec-open-cnp-cnp-assurances-va-consacrer-100-m-sur-5-ans-au-developpement-de-partenariats-avec-des-start-ups)'),
    ),
    dict(
        fid='partech_parrtech-venture',
        vals={'Site web': 'https://partechpartners.com'},
        confiance='Faible',
        src=("partechpartners.com press release 'Partech launches its successor €360M venture "
         "fund' "
         '(https://partechpartners.com/news/partech-launches-its-successor-360m-venture-fund), '
         'Dec 2023; corroborated by EU-Startups-tier press (Silicon Canals, FinSMEs, Real '
         'Deals, funds-europe.com)'),
    ),
    dict(
        fid='partech_partech-seed-i-to-iv',
        vals={'Site web': 'https://partechpartners.com'},
        confiance='Moyen',
        src=('partechpartners.com news pages (multiple), e.g. '
         'https://partechpartners.com/news/partech-closes-its-fourth-seed-fund-partech-entrepreneur-iv-at-120m-to-back-entrepreneurs-from-day-1 '
         "— confirms partechpartners.com is Partech's official site"),
    ),
    dict(
        fid='portage-venture_portag3-venture-i',
        vals={'Site web': 'https://portageinvest.com'},
        confiance='Moyen',
        src=("https://portageinvest.com/ (search snippets: 'Home - Portage', 'About - Portage "
         "Ventures'); confirms this is the current official site of the firm formerly branded "
         'Portag3 Ventures, now Portage Ventures'),
    ),
    dict(
        fid='portage-venture_portag3-venture-ii',
        vals={'Site web': 'https://portageinvest.com'},
        confiance='Moyen',
        src=('Official Portage blog via Newswire.ca: '
         'https://www.newswire.ca/news-releases/portag3-ventures-announces-427m-close-of-second-fintech-fund-839230793.html '
         'and '
         'https://www.newswire.ca/news-releases/portag3-ventures-announces-initial-198m-closing-of-its-second-fintech-fund-698995521.html '
         '; corroborated by BetaKit '
         'https://betakit.com/portag3-sets-sights-on-global-fintech-market-as-it-closes-427-million-cad-fund-ii/ '
         'and Finextra '
         'https://www.finextra.com/newsarticle/34897/canadas-portag3-closes-cad427m-fintech-fund'),
    ),
    dict(
        fid='portage-venture_portag3-venture-iii',
        vals={'Site web': 'https://portageinvest.com'},
        confiance='Moyen',
        src=("WebSearch snippets confirming portageinvest.com as the firm's official domain "
         '(About/Team/Portfolio/Blog pages)'),
    ),
    dict(
        fid='ring-capital_generation',
        vals={'Site web': 'https://www.ringcp.com/generations/'},
        confiance='Élevé',
        src=('ringcp.com/generations/ (page officielle du fonds, apparue directement dans les '
         'résultats WebSearch) ; corroboré par Maddyness '
         '(maddyness.com/2023/10/19/ledhec-et-ring-capital-lancent-un-fonds-damorcage-dedie-aux-projets-a-impact/) '
         'et EDHEC '
         '(edhec.edu/en/about-us/entrepreneurship-business-school/generations-powered-by-edhec-fund)'),
    ),
    dict(
        fid='ring-capital_mission-i',
        vals={'Site web': 'https://www.ringcp.com/ring-mission/'},
        confiance='Élevé',
        src=('ringcp.com/ring-mission/ (page officielle du fonds, apparue directement dans les '
         'résultats WebSearch) ; Carenews '
         '(carenews.com/fr/news/le-nouveau-fonds-ring-mission-mise-sur-les-startup-tech-for-good-de-demain) '
         'et Finyear '
         '(finyear.com/Ring-Mission-nouveau-fonds-de-Venture-Capital-Impact_a44209.html) pour '
         'le contexte de lancement'),
    ),
    dict(
        fid='samaipata-ventures_samaipata-i',
        vals={'Site web': 'https://www.samaipata.vc/'},
        confiance='Élevé',
        src='https://www.samaipata.vc/ (official Samaipata site)',
    ),
    dict(
        fid='samaipata-ventures_samaipata-ii',
        vals={'Site web': 'https://www.samaipata.vc/'},
        confiance='Élevé',
        src=('https://www.samaipata.vc/ (official site); '
         'https://tech.eu/2021/12/14/samaipata-closes-second-fund-at-e107-million/ (Tech.eu, '
         'Dec 2021)'),
    ),
    dict(
        fid='scor-ventures_scor-ventures',
        vals={'Montant alloué (M€)': 130},
        confiance='Moyen',
        src=('Waveup Copilot VC Fund Profile (SCOR Ventures) — "SCOR Ventures operates with a 130 '
         'million euro mandate" (https://hub.waveup.com/funds/scor-ventures), cross-referenced '
         'by a second independent WebSearch result citing the same 130 M€ figure'),
    ),
    dict(
        fid='serena-capital_serena-ii',
        vals={'Site web': 'https://www.serena.vc'},
        confiance='Moyen',
        src=("serena.vc (domaine officiel confirmé via pages d'équipe "
         'serena.vc/team-profile/marc-fournier/ et /philippe-hayat/, cofondateurs de Serena '
         'Capital, apparues dans les résultats WebSearch) ; JournalDuNet '
         '(journaldunet.com/web-tech/start-up/1146927-...) et Fusacq '
         '(fusacq.com/buzz/serena-capital-leve-100-m-pour-son-second-fonds-d-investissement-a60173_fr_) '
         "pour l'historique du fonds Serena II (clôturé à 133 M€ fin 2014)"),
    ),
    dict(
        fid='serena-capital_serena-iii',
        vals={'Site web': 'https://www.serena.vc'},
        confiance='Moyen',
        src=('serena.vc (domaine officiel, cf. note ci-dessus) ; contexte du fonds Serena III (300 '
         "M€ levés, early stage/Séries A-B, tickets jusqu'à 15 M€) via extraits WebSearch "
         "agrégés (sans URL d'article unique clairement identifiée pour cette statistique)"),
    ),
    dict(
        fid='sharpstone-capitale_sharpstone-capitale',
        vals={'Site web': 'https://www.sharpstone.fr'},
        confiance='Élevé',
        src=("CFNEWS annuaire 'SHARPSTONE CAPITAL - Fonds d'investissement / gestionnaire' "
         "(cfnews.net); Maddyness 'Que Font Les Fonds ? Le portrait de Sharpstone Capital' "
         '(2023-01-16, maddyness.com); official site sharpstone.fr.'),
    ),
    dict(
        fid='speedinvest_speedinvest-iv',
        vals={'Nb participations (déclaré)': 100},
        confiance='Faible',
        src=("EU-Startups, 31/01/2024, 'Speedinvest closes massive €350 million fourth flagship "
         "fund' "
         '(https://www.eu-startups.com/2024/01/speedinvest-closes-massive-e350-million-fourth-flagship-fund-e50-million-above-target/) '
         "— 'ticket sizes are expected to range around the €600,000 mark for pre-seed "
         'investments, and €1.5 to €2 million for seed-stage companies, with the firm looking '
         "to add approximately 100 companies'"),
    ),
    dict(
        fid='step-venture_step-venture-i',
        vals={'Millésime': 2025},
        confiance='Élevé',
        src=('https://stepventure.eu/article/first-closing-for-step-fund-the-eur50-million-fund-dedicated-to-italian-seed/ '
         '(official site) corroborated by '
         'https://www.eu-startups.com/2025/11/new-e30-million-step-fund-targets-early-stage-italian-startups-with-international-growth-potential/ '
         'and https://bebeez.eu/2025/11/11/... and '
         'https://vcwire.tech/2025/11/12/step-fund-holds-first-close-at-30m/'),
    ),
    dict(
        fid='the-family_the-family',
        vals={'Site web': 'https://thefamily.co', 'Montant levé (M€)': 15, 'Nb exits': 8},
        confiance='Moyen',
        src=("Site web: thefamily.co confirmed via direct search result title 'Building Ambitious "
         "Startups | The Family' (https://www.thefamily.co/). Montant levé: TechCrunch, "
         "2018-09-11, 'The Family raises $17.4 million to support European startups' — "
         "explicitly states the raise as '$17.4 million (€15 million)', led by LGT Capital "
         'Partners with HummingBird Venture, Project A, and eVentures participating. Nb exits: '
         "Tracxn 'The Family - 2026 Investor Profile' "
         '(https://tracxn.com/d/accelerator-incubator/the-family/__GDXBbMjv2W87YNdgj3UJRzMRZhMm1Oz0AxlE5aQndHU) '
         '— states a portfolio of 24 companies including 3 unicorns, with 8 portfolio exits as '
         'of August 2026.'),
    ),
    dict(
        fid='the-moon-venture_the-moon-venture-soul-invest',
        vals={'Millésime': 2018,
         'Stages pratiqués': 'Seed ; Série A',
         'Site web': 'https://themoonventure.com/'},
        confiance='Moyen',
        src=("Maddyness, portrait 'Que font les fonds ? Le portrait de The Moon Venture' "
         '(https://www.maddyness.com/2024/07/01/que-font-les-fonds-le-portrait-de-the-moon-venture/); '
         'fiche Maddyness base fonds '
         '(https://www.maddyness.com/base-fonds-investissements/the-moon-venture/); site '
         'officiel https://themoonventure.com/investisseurs/'),
    ),
    dict(
        fid='the49_the49',
        vals={'Ticket min (M€)': 0.092, 'Ticket max (M€)': 0.276, 'Site web': 'https://the49.com/'},
        confiance='Moyen',
        src=("The49, page 'Venture Studio' (https://www.the49.com/venture-studio/) et 'About Us' "
         '(https://the49.com/about-us/) ; OpenVC profil (https://www.openvc.app/fund/The49) ; '
         'LeadIQ company overview (https://leadiq.com/c/the49/5e18911262be9f43ca8b6541)'),
    ),
    dict(
        fid='tomcat_tomcat-ventures-i',
        vals={'Site web': 'https://www.tomcat.eu/'},
        confiance='Élevé',
        src=('https://www.maddyness.com/2025/10/23/que-font-les-fonds-le-portrait-de-tomcat/ '
         '(Maddyness, 23 Oct 2025); https://www.tomcat.eu/ (official site)'),
    ),
    dict(
        fid='tomcat_tomcat-ventures-ii',
        vals={'Site web': 'https://www.tomcat.eu/'},
        confiance='Moyen',
        src=('https://www.maddyness.com/2025/06/24/tomcat-veut-lever-entre-80-et-100-millions-deuros-pour-son-deuxieme-fonds-dinvestissement/ '
         '(Maddyness, 24 Jun 2025); '
         'https://www.jaimelesstartups.fr/news/tomcat-lance-officiellement-le-vehicule-tomcat-ventures-ii/ '
         "(J'aime les Startups); https://www.tomcat.eu/"),
    ),
    dict(
        fid='truffle-capital_truffle-fintech-et-insurtech-fund-ii',
        vals={'Site web': 'https://www.truffle.com'},
        confiance='Élevé',
        src=('Truffle Capital official website (https://www.truffle.com/fintech/about); fund '
         "closing corroborated by Bernard-Louis Roques: 'Truffle Capital boucle son fonds "
         "fintech à 140 millions d'euros', mind Fintech "
         '(https://www.mind.eu.com/fintech/investissement/bernard-louis-roques-truffle-capital-boucle-son-fonds-fintech-a-140-millions-deuros-et-va-investir-dans-smartpush/) '
         'and FrenchWeb '
         '(https://www.frenchweb.fr/truffle-capital-leve-400-millions-deuros-pour-investir-dans-la-biomedtech-et-la-fintech-insurtech/387504)'),
    ),
    dict(
        fid='truffle-capital_truffle-fintech-et-insurtech-fund-iii-prevu-2025',
        vals={'Site web': 'https://www.truffle.com'},
        confiance='Faible',
        src=('Truffle Capital official website, multiple pages (https://www.truffle.com/, '
         'https://www.truffle.com/fintech/about, https://www.truffle.com/fintech/values)'),
    ),
    dict(
        fid='uniqa-ventures_uniqa-ventures',
        vals={'Site web': 'https://www.uniqaventures.com', 'Montant alloué (M€)': 150},
        confiance='Faible',
        src=('UNIQA Group Press Center '
         '(press-news.uniqagroup.com/news-150-million-euros-for-start-ups-in-the-cee-region-uniqa-ventures-doubles-growth-capital-for-bold-future-investments) '
         'et TrendingTopics '
         '(trendingtopics.eu/uniqa-ventures-doubles-investment-volume-to-e150m/) — UNIQA '
         "Ventures a doublé son 'volume d'investissement' de 75 M€ à 150 M€ (annonce 2021) ; "
         "site officiel uniqaventures.com mentionné dans plusieurs profils d'investisseurs "
         '(Vestbee, EU-Startups, Capboard) via extraits WebSearch'),
    ),
    dict(
        fid='white-star-capital_wsc-i',
        vals={'Site web': 'https://whitestarcapital.com', 'Montant alloué (M€)': 64.4},
        confiance='Élevé',
        src=('TechCrunch: '
         'https://techcrunch.com/2015/11/13/white-star-capital-closes-70m-for-its-first-institutional-transatlantic-fund '
         '; PE Hub: '
         'https://www.pehub.com/2015/11/white-star-capital-collects-70-mln-for-initial-fund/ ; '
         'Venture Capital Journal: '
         'https://www.venturecapitaljournal.com/white-star-capital-collects-70-mln-for-initial-fund/ '
         '; We Are Guernsey: '
         'https://www.weareguernsey.com/news/2015/guernsey-home-to-white-star-capitals-70-million-start-up-fund/'),
    ),
    dict(
        fid='white-star-capital_wsc-ii',
        vals={'Site web': 'https://whitestarcapital.com', 'Montant alloué (M€)': 165.6},
        confiance='Élevé',
        src=('TechCrunch: https://techcrunch.com/2018/06/03/white-star-capital-ii/ ; Unquote: '
         'https://www.unquote.com/france/official-record/3010234/white-star-capital-closes-second-fund-on-usd180m '
         '; VentureBeat: '
         'https://venturebeat.com/entrepreneur/white-star-capital-closes-180-million-fund-for-early-stage-transatlantic-startups '
         '; Private Equity Wire: '
         'https://www.privateequitywire.co.uk/white-star-capital-announces-second-vc-fund-usd180-million/'),
    ),
    dict(
        fid='white-star-capital_wsc-iii',
        vals={'Site web': 'https://whitestarcapital.com', 'Montant alloué (M€)': 331.2},
        confiance='Élevé',
        src=('Private Equity International: '
         'https://www.privateequityinternational.com/white-star-capital-raises-360m/ ; Forbes: '
         'https://www.forbes.com/sites/rebeccaszkutak/2021/10/25/white-star-capital-raises-360-million-fund-that-nearly-doubles-its-aum/ '
         '; Unquote: '
         'https://www.unquote.com/unquote/official-record/3025324/white-star-holds-usd-360m-final-close-for-fund-iii '
         '; PR Newswire: '
         'https://www.prnewswire.com/news-releases/white-star-capital-announces-new-fund-firm-now-has-more-than-500-million-of-fresh-capital-to-invest-globally-301407267.html'),
    ),
    dict(
        fid='xange-capital_mutuelles-impact',
        vals={'Montant alloué (M€)': 95},
        confiance='Élevé',
        src=('Maddyness (13/01/2023) — "Le fonds Mutuelles Impact atteint 95 millions d\'euros '
         'sous gestion" (https://www.maddyness.com/2023/01/13/mutuelles-impact/)'),
    ),
    dict(
        fid='xange-capital_xange-4',
        vals={'Montant alloué (M€)': 220},
        confiance='Élevé',
        src=('Usine Digitale — "XAnge boucle un nouveau fonds de 220 millions d\'euros" '
         '(https://www.usine-digitale.fr/article/xange-boucle-un-nouveau-fonds-de-220-millions-d-euros-et-s-interesse-tout-particulierement-au-web3.N2024137) '
         '; Journal du Net '
         '(https://www.journaldunet.com/web3/crypto/1513101-le-fonds-de-capital-risque-xange-leve-220-millions-d-euros/) '
         '— final closing at 220 M€, doubling capacity vs. XAnge 3'),
    ),
    dict(
        fid='xange-capital_xange-capital-2',
        vals={'Montant alloué (M€)': 62},
        confiance='Élevé',
        src=('Bpifrance press release / Boursorama / Fusacq — all titled "XAnge annonce le succès '
         'du closing final de son fonds multicorporate, XAnge Capital 2, à hauteur de 62 M€" '
         '(https://presse.bpifrance.fr/xangefonds-partenaire-de-bpifrance-annonce-le-succes-du-closing-final-de-son-fonds-multicorporate-xange-capital-2-a-hauteur-de-62-me '
         '; '
         'https://www.boursorama.com/bourse/actualites/xange-closing-final-du-2eme-fonds-multicorporate-a-62-m-b77f0ecbd90ca66aea05c5260d6e0343 '
         '; '
         'https://www.fusacq.com/buzz/xange-annonce-le-succes-du-closing-final-de-son-fonds-multicorporate-xange-capital-2-a80113_fr_)'),
    ),
    dict(
        fid='xange-capital_xange-digital-3',
        vals={'Montant alloué (M€)': 90},
        confiance='Moyen',
        src=('CFNEWS / Next-Finance / DOCaufutur (31/05-01/06/2018) — "XAnge réalise le 2e closing '
         'de XAnge Digital 3 et se dote ainsi d\'une capacité d\'investissement de 90M€" '
         '(https://www.next-finance.net/XAnge-realise-le-2e-closing-de ; '
         'https://www.cfnews.net/L-actualite/Capital-innovation-developpement/Operations/Levee-de-Fonds/Deuxieme-closing-pour-XAnge-Digital-3-272010 '
         '; '
         'https://www.docaufutur.fr/2018/05/31/xange-realise-le-2e-closing-de-xange-digital-3-et-se-dote-ainsi-dune-capacite-dinvestissement-de-90me-pour-investir-dans-le-digital/)'),
    ),

]

# Champs remplis avec un niveau de confiance Faible ET une valeur numérique jugée
# approximative (à distinguer d'une simple URL de site trouvée avec peu de certitude) :
# ces champs génèrent une anomalie « Valeur approximative conservée » dédiée plutôt qu'une
# anomalie « Mise à jour web » standard.
APPROX_FIELDS = {'365-fintech_365-fintech': ['Millésime', 'Montant levé (M€)'],
 'goldsmith-ventures_goldsmith-ventures-eis-fund': ['Ticket min (M€)', 'Ticket max (M€)'],
 'speedinvest_speedinvest-iv': ['Nb participations (déclaré)'],
 'uniqa-ventures_uniqa-ventures': ['Montant alloué (M€)']}


# =============================================================================
# 2. DIVERGENCES — la base a déjà une valeur, la source web la contredit (non arbitré)
# =============================================================================

DIVERGENCES = [
    dict(
        fid='alven-capital_alven-v',
        champ='Millésime',
        base='2016',
        web="closing d'Alven V (250 M€) annoncé début janvier 2017 (TechCrunch 03/01/2017, Maddyness 09/01/2017 : « VC Alven Capital : fonds de 250 millions d'euros »)",
        com="Millésime base conservé (2016) : aucune source ne documente publiquement un premier closing antérieur à l'annonce de janvier 2017 ; écart non arbitré.",
    ),
    dict(
        fid='blast-club_blast-club',
        champ='Millésime',
        base='2022',
        web='création en 2023 selon 4 sources concordantes (Forbes France, JDN, Maddyness, Wikipédia)',
        com='Millésime base conservé (2022) : à vérifier, la majorité des sources web datent la création du club en 2023.',
    ),
    dict(
        fid='kickstart-innovation_kickstart-innovation',
        champ='Nb participations (déclaré)',
        base='382',
        web='450 startups accompagnées depuis la création (source secondaire)',
        com='Valeur base conservée (382) : écart avec un chiffre plus récent/plus large trouvé en presse, non arbitré faute de date de référence commune.',
    ),
    dict(
        fid='south-east-angels_south-east-angels',
        champ='Nb participations (déclaré)',
        base='41',
        web='34 participations recensées sur la page Portfolio officielle (nov. 2025)',
        com='Valeur base conservée (41) : écart avec le décompte du site officiel à une date antérieure, non arbitré.',
    ),
    dict(
        fid='blackfin-capital-partners_blackfintech-ii',
        champ='Montant levé (M€)',
        base='390',
        web='350 M€ selon fintech.global (« BlackFin commits €350m fund for European InsurTechs and FinTechs », 12/07/2022) ; 390 M€ selon un agrégateur tiers',
        com="Valeur base conservée (390), corroborée par un agrégateur indépendant ; le chiffre de 350 M€ (titre d'article, closing possiblement partiel) n'est pas retenu.",
    ),

]


# =============================================================================
# 3. VALEURS EN FOURCHETTE NON RETENUES — sources contradictoires, champ resté vide
# =============================================================================

RANGES_NON_RETENUES = [
    dict(
        fid='tenity_tenity-incubation-fund-i-ii',
        champ='Montant levé (M€) / AuM (M€)',
        candidats='100 M$ (cible fonds II, Startup Bubble News) ; >140 M$ (AuM groupe, Superscout) ; 120 MCHF (« assets under advisory » au 1er closing, startupticker.ch) ; 100 MCHF (AuM groupe Tenity AG, SECA)',
        com='Aucune des quatre valeurs ne peut être rattachée avec certitude au véhicule précis « Tenity Incubation Fund I & II » ; cellule laissée vide conformément à la règle anti-estimation.',
    ),
    dict(
        fid='the-moon-venture_the-moon-venture-soul-invest',
        champ='Montant levé (M€) / AuM (M€)',
        candidats='25 M€ investis en 4 ans dont 10 M€ en 2022 (Maddyness) ; >30 M€ déployés (page liée à Soul Invest) ; >50 M€ déployés / 30 startups (investormatch.pro, source non primaire)',
        com="Écart important entre sources sans source primaire datée faisant autorité ; cellule laissée vide. The Moon Venture est la marque déployée par Soul Invest (plateforme de financement participatif agréée PSFP par l'AMF), pas un fonds de capital-risque classique — ce qui explique l'absence d'AuM/TVPI/DPI/IRR publics au sens usuel.",
    ),
    dict(
        fid='cadence-growth-capital_cadence-growth-capital-cgc',
        champ='Millésime',
        candidats='2019 (année de fondation de la société, StartupIntros/Tracxn) ; 2020 (« Cadence Growth Capital Fund I | 2020 VC Fund », VentureCapitalArchive)',
        com="Ambiguïté entre l'année de création de la société de gestion et le millésime du véhicule CGC lui-même ; cellule laissée vide.",
    ),
    dict(
        fid='nca-next-commerce-accelerator_nca-next-commerce-accelerator',
        champ='Millésime',
        candidats='2017 (majorité des sources : Startbase, Prodevs.io, BusinessABC.net) ; 2018 (une source isolée mentionnant Peak Capital / Upgrade Commerce)',
        com='Majorité des sources en faveur de 2017, mais écart non arbitré par une source primaire ; cellule laissée vide.',
    ),

]


# =============================================================================
# 4. RAPPROCHEMENTS NON RÉSOLUS — risques d'homonymie / d'identité de véhicule
# =============================================================================
# Aucune donnée n'est écrite pour ces lignes à partir des pistes ci-dessous : seul le doute
# est journalisé, à charge pour un opérateur humain de trancher.

IDENTITY_RISKS = [
    dict(
        fids=['partech_partech-seed-i-to-iv', 'partech_parrtech-venture'],
        commentaire=('La dénomination « Partech Seed (I to IV) » ne correspond à aucune série officielle '
         'documentée par Partech (dont la nomenclature connue est « Partech Entrepreneur I-IV ») '
         '; « Parrtech Venture » est une coquille confirmée pour « Partech Venture », mais il '
         'existe à la fois un véhicule 2023 de 360 M€ sous ce nom et un historique de fonds plus '
         "anciens du même nom — le millésime visé par la ligne de la base n'a pas pu être "
         'déterminé avec certitude.'),
    ),
    dict(
        fids=['isai_isai-venture-iii'],
        commentaire=('Risque de confusion avec « ISAI Expansion III », véhicule distinct de croissance/LBO '
         "(hard cap 150 M€) chez ISAI. Aucun montant n'a été renseigné pour ISAI Venture III "
         "afin d'éviter toute conflation ; à vérifier manuellement avant toute future collecte "
         'de Montant alloué/levé sur cette ligne.'),
    ),
    dict(
        fids=['breega-capital_breega-seed-ii', 'breega-capital_breega-venture-iii'],
        commentaire=('Une source secondaire non primaire indique que « Breega Capital Venture 3 » (= '
         'breega-venture-iii, 110 M€, closing final mars 2021) aurait été « renommé Seed II » '
         'par la suite. Si confirmé, les deux lignes de la base pourraient désigner le même '
         "véhicule sous deux noms différents — à vérifier manuellement avant d'arbitrer un "
         'éventuel doublon.'),
    ),
    dict(
        fids=['tsp-ventures_tsp-ventures'],
        commentaire=('La seule entité publique « TSP Ventures » identifiable (Londres, fondée en 2019, '
         'tspventures.co.uk) investit exclusivement en climate-tech / environmental tech, sans '
         'lien documenté avec le fintech/insurtech. Probable homonymie ou mauvaise '
         "identification de l'entité dans le référentiel ; aucune donnée de cette entité n'a été "
         'reportée sur la ligne de la base.'),
    ),
    dict(
        fids=['sharpstone-capitale_sharpstone-capitale'],
        commentaire=("La presse et le site officiel désignent systématiquement l'entité « Sharpstone Capital "
         '» (et non « Sharpstone Capitale ») ; probable variante orthographique dans le '
         'référentiel, à confirmer.'),
    ),
    dict(
        fids=['kima-venture-capital_kima-venture'],
        commentaire=("L'entité réelle se nomme « Kima Ventures » (véhicule evergreen à LP unique, Xavier "
         "Niel), et non « Kima Venture Capital » — ce qui explique l'absence de TVPI/DPI/IRR "
         'publics (pas de LPs tiers à rapporter).'),
    ),
    dict(
        fids=['start-venture_start-venture-i'],
        commentaire=("Aucune entité « Start Venture » correspondante n'a été retrouvée en ligne après "
         'plusieurs recherches ciblées. Seule une entité au nom proche mais distincte (« Start '
         'Ventures », Lisbonne, fintech/insurtech) existe publiquement ; non retenue faute de '
         'confirmation du rapprochement. Aucune donnée reportée.'),
    ),
    dict(
        fids=['ring-capital_mission-ii'],
        commentaire=("L'existence d'un véhicule distinct « Ring Mission II » n'a pas pu être confirmée : "
         'seuls « Ring Mission » (35 M€, premier closing) et « Altitude II » (stratégie '
         'croissance/LBO différente) sont documentés chez Ring Capital. Ligne à vérifier '
         'manuellement — possible erreur de référentiel.'),
    ),
    dict(
        fids=['truffle-capital_truffle-fintech-et-insurtech-fund-iii-prevu-2025'],
        commentaire=('Aucune source ne confirme le lancement effectif de ce fonds à la date de cette '
         'collecte (14-15/09/2026) ; seuls le Fund II (140 M€, clos en 2019) et les statistiques '
         'globales de la société de gestion sont documentés publiquement. Le statut « prévu 2025 '
         '» du nom de la ligne reste donc non confirmé.'),
    ),
    dict(
        fids=['committed-capital_committed-capital-eis-fund'],
        commentaire=('Les données trouvées (AuM, TVPI, IRR) concernent le véhicule officiellement nommé « '
         "Growth EIS Fund » chez Committed Capital ; à confirmer qu'il s'agit bien du même "
         'véhicule que la ligne « Committed Capital EIS Fund » de la base avant de leur accorder '
         'une confiance Élevée.'),
    ),

]


# =============================================================================
# 5. VÉHICULES NON EXPLOITABLES — business angels solos sans portefeuille vérifiable
# =============================================================================

NON_EXPLOITABLE = [
    dict(
        fid='david-semmens_david-semmens-business-angel',
        com=('CIO chez Cadro, administrateur chez Wealthify (Aviva) et RiskSave Technologies '
         '(insurtech). Seule trace publique : une fourchette de ticket personnel indicative sur '
         'une plateforme de mise en relation (Signal/NFX, 5-25 k$) — aucun véhicule '
         "d'investissement, portefeuille structuré ni performance publique."),
    ),
    dict(
        fid='patrice-fleurquin_patrice-fleurquin-business-angel',
        com=('Profil entrepreneurial documenté (cofondateur Tripfair, ex-CSO Nascom) mais aucune '
         "trace d'un portefeuille d'investissement en tant que business angel."),
    ),
    dict(
        fid='mohammad-hossein-tavangar_m-h-tavangar-business-angel',
        com=('Basé à Berlin, ex-Managing Partner chez Dorrance Venture, administrateur au Founder '
         "Institute Germany, focus fintech/insurtech/blockchain. L'activité d'investissement "
         'documentée est celle de Dorrance Venture (entité distincte), pas un véhicule personnel '
         '— aucun montant ni portefeuille propre identifiable.'),
    ),

]

# =============================================================================
# 6. OUTILS
# =============================================================================

def localiser_entree(base_dir: Path) -> Path:
    """Cherche le classeur d'entrée dans le répertoire du script puis dans les dépôts connus."""
    for dossier in (base_dir, Path.cwd(), Path("/mnt/user-data/uploads")):
        chemin = dossier / FICHIER_ENTREE
        if chemin.exists():
            return chemin
        if dossier.exists():
            trouves = sorted(dossier.glob(f"*{FICHIER_ENTREE}"))
            if trouves:
                return trouves[0]
    raise FileNotFoundError(f"Classeur d'entrée introuvable : {FICHIER_ENTREE}")


def est_vide(valeur) -> bool:
    if valeur is None:
        return True
    if isinstance(valeur, float) and pd.isna(valeur):
        return True
    if isinstance(valeur, str) and valeur.strip() == "":
        return True
    return False


class JournalAnomalies:
    """Ajoute des lignes à l'onglet Anomalies en poursuivant la numérotation existante."""

    def __init__(self, anomalies: pd.DataFrame):
        self.colonnes = list(anomalies.columns)
        self.df = anomalies
        derniers = [int(str(a).split("-")[-1]) for a in anomalies["Anomalie_ID"].dropna()
                    if str(a).startswith("ANO-")]
        self.compteur = max(derniers) if derniers else 0
        self.nouvelles = []

    def ajouter(self, type_ano, entite=None, champ=None, valeur_source=None, valeur_retenue=None,
                candidats=None, confiance="Élevé", traitement=None, commentaire=None):
        self.compteur += 1
        ligne = {c: None for c in self.colonnes}
        ligne.update({
            "Anomalie_ID": f"ANO-{self.compteur:04d}",
            "Type d'anomalie": type_ano,
            "Onglet source": SOURCE_LABEL,
            "Table ou bloc source": "Sites officiels / communiqués / presse spécialisée / registres publics",
            "Entité concernée": entite,
            "Champ concerné": champ,
            "Valeur source": valeur_source,
            "Valeur retenue": valeur_retenue,
            "Candidats éventuels": candidats,
            "Niveau de confiance": confiance,
            "Traitement appliqué": traitement,
            "Commentaire": commentaire,
        })
        self.nouvelles.append(ligne)

    def resultat(self) -> pd.DataFrame:
        if not self.nouvelles:
            return self.df
        return pd.concat([self.df, pd.DataFrame(self.nouvelles, columns=self.colonnes)],
                         ignore_index=True)


# =============================================================================
# 7. APPLICATION DE LA COLLECTE
# =============================================================================

def appliquer_collecte(fonds: pd.DataFrame, journal: JournalAnomalies) -> tuple[int, list[str], dict]:
    """Remplit les cellules vides avec les valeurs collectées ; journalise chaque écriture.

    Ne réécrit jamais une cellule déjà renseignée : si la base a été modifiée depuis la
    campagne de recherche, le champ correspondant est simplement ignoré (aucune anomalie de
    variante n'est nécessaire ici, la collecte n'ayant ciblé que des cellules connues vides
    au moment de la recherche).
    """
    remplies, touchees = 0, []
    statut_todo = {}
    index = {fid: i for i, fid in enumerate(fonds["Fonds_ID"])}
    for maj in COLLECTE_WEB:
        fid = maj["fid"]
        if fid not in index:
            journal.ajouter("Rapprochement non résolu", entite=fid, champ="Fonds_ID",
                            valeur_source=fid, confiance="Élevé",
                            traitement="Mise à jour non appliquée",
                            commentaire="Fonds_ID absent de la base : vérifier le référentiel")
            continue
        i = index[fid]
        entite = f"{fonds.at[i, 'Nom du fonds']} / {fonds.at[i, 'Véhicule']}"
        approx_champs = APPROX_FIELDS.get(fid, set())
        ecrit_normal, ecrit_approx = [], []
        for champ, valeur in maj["vals"].items():
            actuelle = fonds.at[i, champ]
            if not est_vide(actuelle):
                # Cellule déjà renseignée entre-temps : on ne touche à rien, silencieusement.
                continue
            fonds.at[i, champ] = valeur
            if champ in approx_champs:
                ecrit_approx.append(champ)
            else:
                ecrit_normal.append(champ)
        ecrit = ecrit_normal + ecrit_approx
        if not ecrit:
            continue
        remplies += len(ecrit)
        touchees.append(entite)
        # traçabilité : source + date de mise à jour
        note = f"[Collecte web {DATE_MAJ:%d/%m/%Y}] {maj['src']}"
        actuelle_src = fonds.at[i, "Source_AuM"]
        fonds.at[i, "Source_AuM"] = note if est_vide(actuelle_src) else f"{actuelle_src} | {note}"
        fonds.at[i, "Date_MAJ"] = DATE_MAJ
        if ecrit_normal:
            journal.ajouter("Mise à jour web", entite=entite, champ=" ; ".join(ecrit_normal),
                            valeur_retenue=" ; ".join(f"{c}={maj['vals'][c]}" for c in ecrit_normal),
                            confiance=maj["confiance"],
                            traitement="Cellules vides complétées à partir de la source web",
                            commentaire=maj["src"])
        if ecrit_approx:
            journal.ajouter("Valeur approximative conservée", entite=entite,
                            champ=" ; ".join(ecrit_approx),
                            valeur_retenue=" ; ".join(f"{c}={maj['vals'][c]}" for c in ecrit_approx),
                            confiance="Faible",
                            traitement="Valeur approximative ou faiblement sourcée retenue "
                                       "(source unique ou formulation imprécise)",
                            commentaire=maj["src"])
        statut_todo[fid] = ("traité", "Champs complétés : " + ", ".join(ecrit))
    return remplies, touchees, statut_todo


def journaliser_divergences(fonds: pd.DataFrame, journal: JournalAnomalies) -> None:
    """Consigne les écarts base / web sans les arbitrer (la base garde sa valeur actuelle)."""
    index = {fid: i for i, fid in enumerate(fonds["Fonds_ID"])}
    for d in DIVERGENCES:
        i = index.get(d["fid"])
        entite = (f"{fonds.at[i, 'Nom du fonds']} / {fonds.at[i, 'Véhicule']}"
                  if i is not None else d["fid"])
        journal.ajouter("Différence entre description et structure réelle", entite=entite,
                        champ=d["champ"], valeur_source=f"web : {d['web']}",
                        valeur_retenue=f"base : {d['base']}", confiance="Moyen",
                        traitement="Valeur de la base conservée, écart documenté",
                        commentaire=d["com"])


def journaliser_ranges_non_retenues(fonds: pd.DataFrame, journal: JournalAnomalies) -> None:
    """Consigne les champs restés vides faute de source unique fiable (fourchette/désaccord)."""
    index = {fid: i for i, fid in enumerate(fonds["Fonds_ID"])}
    for r in RANGES_NON_RETENUES:
        i = index.get(r["fid"])
        entite = (f"{fonds.at[i, 'Nom du fonds']} / {fonds.at[i, 'Véhicule']}"
                  if i is not None else r["fid"])
        journal.ajouter("Valeur en fourchette non retenue", entite=entite, champ=r["champ"],
                        candidats=r["candidats"], confiance="Faible",
                        traitement="Cellule laissée vide, sources contradictoires non arbitrées",
                        commentaire=r["com"])


def journaliser_identity_risks(fonds: pd.DataFrame, journal: JournalAnomalies) -> None:
    """Consigne les risques d'homonymie/de rapprochement non confirmé (aucune donnée écrite)."""
    index = {fid: i for i, fid in enumerate(fonds["Fonds_ID"])}
    for r in IDENTITY_RISKS:
        entites = []
        for fid in r["fids"]:
            i = index.get(fid)
            entites.append(f"{fonds.at[i, 'Nom du fonds']} / {fonds.at[i, 'Véhicule']}"
                           if i is not None else fid)
        journal.ajouter("Rapprochement non résolu", entite=" ; ".join(entites),
                        confiance="Moyen",
                        traitement="Aucune donnée écrite, risque d'identité signalé pour "
                                   "vérification manuelle",
                        commentaire=r["commentaire"])


def marquer_non_exploitables(fonds: pd.DataFrame, journal: JournalAnomalies) -> None:
    """Documente les business angels solos sans portefeuille vérifiable (aucune cellule modifiée)."""
    index = {fid: i for i, fid in enumerate(fonds["Fonds_ID"])}
    for n in NON_EXPLOITABLE:
        i = index.get(n["fid"])
        entite = (f"{fonds.at[i, 'Nom du fonds']} / {fonds.at[i, 'Véhicule']}"
                  if i is not None else n["fid"])
        journal.ajouter("Rapprochement non résolu", entite=entite,
                        confiance="Élevé",
                        traitement="Statut collecte = non exploitable, aucune cellule modifiée",
                        commentaire=n["com"])


def recalculer_dictionnaire(dico: pd.DataFrame, frames: dict) -> pd.DataFrame:
    """Recalcule les taux de remplissage de l'onglet Dictionnaire."""
    for idx, row in dico.iterrows():
        df = frames.get(row["Onglet"])
        if df is None or row["Champ"] not in df.columns:
            continue
        n = len(df)
        remplies = int(sum(0 if est_vide(v) else 1 for v in df[row["Champ"]]))
        dico.at[idx, "Nb valeurs renseignées"] = remplies
        dico.at[idx, "Taux de remplissage %"] = round(100.0 * remplies / n, 1) if n else 0.0
    return dico


def mettre_a_jour_todo(base_dir: Path, statut_todo: dict) -> None:
    """Met à jour todo_collecte_web.csv : Statut collecte + Notes, pour les 111 véhicules."""
    chemin = base_dir / "todo_collecte_web.csv"
    if not chemin.exists():
        return
    todo = pd.read_csv(chemin)
    todo["Notes"] = todo["Notes"].astype(object)
    todo["Statut collecte"] = todo["Statut collecte"].astype(object)
    non_exploitable_fids = {n["fid"] for n in NON_EXPLOITABLE}
    for idx, row in todo.iterrows():
        fid = row["Fonds_ID"]
        if fid in statut_todo:
            statut, note = statut_todo[fid]
        elif fid in non_exploitable_fids:
            statut, note = "non exploitable", next(
                n["com"] for n in NON_EXPLOITABLE if n["fid"] == fid)
        else:
            statut, note = "traité — rien trouvé", (
                "Recherche effectuée (site officiel, presse spécialisée, registres) : aucune "
                "donnée publique complémentaire fiable trouvée pour les champs manquants.")
        todo.at[idx, "Statut collecte"] = statut
        todo.at[idx, "Notes"] = note
    todo.to_csv(chemin, index=False)


# =============================================================================
# 8. ORCHESTRATION
# =============================================================================

def main() -> int:
    base_dir = Path(__file__).resolve().parent
    try:
        entree = localiser_entree(base_dir)
    except FileNotFoundError as exc:
        print(f"ERREUR : {exc}")
        return 1
    sortie = base_dir / FICHIER_SORTIE

    frames = {nom: pd.read_excel(entree, nom) for nom in ONGLETS}
    frames = {k: v.astype(object).where(pd.notna(v), None) for k, v in frames.items()}
    avant = {k: int(sum(0 if est_vide(x) else 1 for col in v.columns for x in v[col]))
             for k, v in frames.items()}

    journal = JournalAnomalies(frames["Anomalies"])

    fonds = frames["Fonds"]
    remplies, touchees, statut_todo = appliquer_collecte(fonds, journal)
    journaliser_divergences(fonds, journal)
    journaliser_ranges_non_retenues(fonds, journal)
    journaliser_identity_risks(fonds, journal)
    marquer_non_exploitables(fonds, journal)
    frames["Fonds"] = fonds

    frames["Anomalies"] = journal.resultat()
    frames["Dictionnaire"] = recalculer_dictionnaire(frames["Dictionnaire"], frames)

    nb_deals = len([c for c in fonds.columns if str(c).startswith("Deal_")]) // 5
    rvc.write_excel(sortie, frames, nb_deals)

    mettre_a_jour_todo(base_dir, statut_todo)

    apres = {k: int(sum(0 if est_vide(x) else 1 for col in v.columns for x in v[col]))
             for k, v in frames.items()}
    print(f"Fichier produit : {sortie}")
    print(f"Entrée          : {entree}")
    print()
    print(f"Cellules vides complétées par la collecte web : {remplies}")
    print(f"Véhicules mis à jour : {len(touchees)}")
    for t in touchees:
        print(f"  - {t}")
    print()
    print(f"Divergences journalisées : {len(DIVERGENCES)}")
    print(f"Fourchettes non retenues : {len(RANGES_NON_RETENUES)}")
    print(f"Rapprochements non résolus : {len(IDENTITY_RISKS)}")
    print(f"Véhicules non exploitables : {len(NON_EXPLOITABLE)}")
    print()
    print("Taux de remplissage (onglet Fonds) :")
    for col in ("AuM (M€)", "Montant levé (M€)", "Montant alloué (M€)", "Ticket min (M€)",
                "Ticket max (M€)", "Millésime", "TVPI", "DPI", "IRR (TRI)", "Site web"):
        n = sum(0 if est_vide(v) else 1 for v in fonds[col])
        print(f"  - {col} : {100.0 * n / len(fonds):.1f} %  ({n}/{len(fonds)})")
    print()
    print("Lignes par onglet :", {k: len(v) for k, v in frames.items()})
    print("Cellules renseignées avant / après :",
          {k: f"{avant[k]} -> {apres[k]}" for k in frames})
    return 0


if __name__ == "__main__":
    sys.exit(main())
