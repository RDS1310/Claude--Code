#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fusionner_participations_fermes.py — Intègre la recherche web « véhicules Fermé »
====================================================================================

Suite de fusionner_participations_ouverts.py : cette fois la recherche a été
menée pour les véhicules au statut "Fermé" de 27 sociétés de gestion (campagne
Closed-Lots A à G + une passe de rattrapage dédiée EOS Venture / Atlantic
Vantage Point / Index Ventures), avec attribution au véhicule précis quand la
presse le permet, ou par déduction de millésime sinon (Confiance abaissée en
conséquence).

Entrée  : VC_Database_Standardisee_v6.xlsx
Sortie  : VC_Database_Standardisee_v7.xlsx

Usage : python3 fusionner_participations_fermes.py
"""

from __future__ import annotations

import math
import re
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v6.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v7.xlsx"

TAUX_CHANGE = {"€": 1.0, "$": 0.92, "£": 1.17, "CHF": 1.07, "AUD": 0.60}


def vide(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    if isinstance(v, str) and v.strip().lower() in ("", "n.d.", "n.d", "n/d", "nan"):
        return True
    return False


def convertir(montant, devise) -> float | None:
    if vide(montant):
        return None
    taux = TAUX_CHANGE.get(str(devise).strip()) if not vide(devise) else 1.0
    if taux is None:
        return None
    return round(float(montant) * taux, 2)


def positionnement(role: str) -> str | None:
    if vide(role):
        return None
    r = role.lower()
    if any(k in r for k in ("actionnaire de référence", "actionnaire principal", "investisseur unique")):
        return "Leader"
    if r.startswith("lead") or "lead" in r.split(" (")[0] or r.startswith("chef de file"):
        return "Leader"
    if any(k in r for k in ("co-lead", "co-chef de file", "co-investisseur", "investisseur existant",
                             "investisseur historique", "investisseur (", "participant", "co-fondateur")):
        return "Minoritaire"
    return None


def exit_on(statut: str) -> str | None:
    if vide(statut):
        return None
    s = statut.lower()
    if any(k in s for k in ("exit", "cédée", "cédé", "acquis", "racheté", "rachetée")):
        return "O"
    if "arrêt" in s or "défaillance" in s:
        return "N"
    if "actif" in s or "active" in s:
        return "N"
    return None


# ---------------------------------------------------------------------------
# Table de correspondance (Nom du fonds, Véhicule exact) -> Fonds_ID, tirée
# de l'onglet Fonds de v6. Les libellés de véhicule ci-dessous dans PARTICIPATIONS
# ont été nettoyés (qualificatifs "déduit"/"probable"/"non confirmé" retirés)
# pour matcher exactement ces clés.
# ---------------------------------------------------------------------------
VEHICULE_ID = {
    ("BlackFin Capital Partners", "BlackfinTech 1"): "blackfin-capital-partners_blackfintech-1",
    ("BlackFin Capital Partners", "BlackfinTech II"): "blackfin-capital-partners_blackfintech-ii",
    ("Portage Venture", "Portag3 Venture I"): "portage-venture_portag3-venture-i",
    ("Portage Venture", "Portag3 Venture II"): "portage-venture_portag3-venture-ii",
    ("Partech", "Partech Seed (I to IV)"): "partech_partech-seed-i-to-iv",
    ("Ring Capital", "Mission I"): "ring-capital_mission-i",
    ("Founders Future VC", "Founders Future Good"): "founders-future-vc_founders-future-good",
    ("EOS Venture", "EVP I"): "eos-venture_evp-i",
    ("Atlantic Vantage Point", "AVP early stage I"): "atlantic-vantage-point_avp-early-stage-i",
    ("Atlantic Vantage Point", "AVP early stage II"): "atlantic-vantage-point_avp-early-stage-ii",
    ("Elaia Partners", "Elaia Delta"): "elaia-partners_elaia-delta",
    ("Elaia Partners", "DV4"): "elaia-partners_dv4",
    ("Alven Capital", "Alven VI"): "alven-capital_alven-vi",
    ("Serena Capital", "Serena III"): "serena-capital_serena-iii",
    ("ISAI", "ISAI Cap Venture"): "isai_isai-cap-venture",
    ("White Star Capital", "WSC II"): "white-star-capital_wsc-ii",
    ("White Star Capital", "WSC III"): "white-star-capital_wsc-iii",
    ("Accel Partners", "Accel London VI"): "accel-partners_accel-london-vi",
    ("Accel Partners", "Accel London VII"): "accel-partners_accel-london-vii",
    ("Breega Capital", "Breega Venture III"): "breega-capital_breega-venture-iii",
    ("Breega Capital", "Breega Seed II"): "breega-capital_breega-seed-ii",
    ("Cathay", "Cathay Innovation II"): "cathay_cathay-innovation-ii",
    ("Truffle Capital", "Truffle FinTech & InsurTech Fund II"): "truffle-capital_truffle-fintech-et-insurtech-fund-ii",
    ("Index Venture", "Index Ventures XI"): "index-venture_index-ventures-xi",
    ("Index Venture", "Index Ventures XII"): "index-venture_index-ventures-xii",
    ("speedInvest", "Speedinvest IV"): "speedinvest_speedinvest-iv",
    ("CommerzVentures", "CommerzVentures Fonds I à III"): "commerzventures_commerzventures-fonds-i-a-iii",
    ("13books Capital (ex-Element Ventures)", "Fonds I & II"): "13books-capital-ex-element-ventures_fonds-i-ii",
}

# ---------------------------------------------------------------------------
# Nouvelles participations identifiées (véhicules Fermé). Un seul enregistrement
# par (Véhicule, Start-up) : quand plusieurs tours existent pour la même paire,
# le premier tour connu sert de Date/Taille de référence, les tours suivants
# sont résumés dans Commentaires.
# ---------------------------------------------------------------------------
ROWS = [
    # nom_fonds, vehicule, startup, secteur, pays, tour, date(YYYY-MM), taille, devise,
    # role, statut, confiance, commentaires, sources
    ("BlackFin Capital Partners", "BlackfinTech 1", "FRISS", "InsurTech (cœur)", "Pays-Bas",
     "Série A", "2017-12", 15, "€", "Co-investisseur (lead Aquiline Technology Growth)", "À vérifier", "Élevé",
     '"Our first investment" selon BlackFin Tech.',
     "FRISS press release ; Medium BlackFin Tech ; Unquote"),
    ("BlackFin Capital Partners", "BlackfinTech 1", "Akur8", "InsurTech (cœur)", "France",
     "Série A", "2020-03", 8.9, "$", "Co-lead (avec MTech Capital)", "Active", "Moyen",
     "Réinvesti en Série B (2021, 30 M$) et Série C (sept. 2023, 25 M$, avec FinTLV/Guidewire).",
     "PE Wire ; Unquote ; Kamet Ventures ; Crunchbase"),
    ("BlackFin Capital Partners", "BlackfinTech 1", "Descartes Underwriting", "InsurTech (cœur)", "France",
     "Seed", "2019-02", 2.5, "$", "Investisseur", "Active", "Moyen", None,
     "Crunchbase ; Reinsurance News ; Cathay Capital"),
    ("BlackFin Capital Partners", "BlackfinTech II", "Descartes Underwriting", "InsurTech (cœur)", "France",
     "Série B", "2022-01", 120, "$", "Investisseur existant", "Active", "Élevé",
     "Véhicule nommé explicitement (« BlackFin Tech 2 Fund »).",
     "Medium BlackFin Tech ; The Insurer ; Orrick ; Artemis.bm"),
    ("BlackFin Capital Partners", "BlackfinTech 1", "Bdeo", "InsurTech (adjacent)", "Espagne",
     "Série A", "2020-11", 5, "€", "Chef de file", "Active", "Élevé",
     "Réinvesti en Série B (juin 2023, 7,5 M€, avec Hollard, Wayra-Telefónica Seguros).",
     "Crunchbase ; Medium BlackFin Tech ; TechFundingNews"),
    ("BlackFin Capital Partners", "BlackfinTech 1", "Epsor", "FinTech liée assurance", "France",
     "Série A", "2019-07", 6, "€", "Chef de file (co-invest Partech)", "Active", "Élevé",
     "Réinvesti en Série B (mai 2021, 20 M€) et Série C (mars 2025, 16 M€, mené par FST).",
     "Fusacq ; Usine Digitale ; Partech news"),

    ("Portage Venture", "Portag3 Venture I", "Alan", "InsurTech (cœur)", "France",
     "Seed", "2016-10", 12, "€", "Co-investisseur (avec Index Ventures, CNP, Partech)", "Active", "Moyen",
     "Millésime cohérent (2016), véhicule non nommé explicitement.",
     "Partech news ; Alan.com blog"),
    ("Portage Venture", "Portag3 Venture II", "Clark", "InsurTech (cœur)", "Allemagne",
     "n/d", None, None, None, "Investisseur", "À vérifier", "Moyen",
     '"Fund II" cité explicitement, tour/date/montant non retrouvés.',
     "PSP Investments ; Seedtable"),

    ("Partech", "Partech Seed (I to IV)", "Alan", "InsurTech (cœur)", "France",
     "Seed", "2016-10", 12, "€", "Co-investisseur", "Active", "Moyen", None,
     "Partech news ; Alan.com blog"),
    ("Partech", "Partech Seed (I to IV)", "Epsor", "FinTech liée assurance", "France",
     "Série A", "2019-07", 6, "€", "Co-investisseur (BlackFin lead)", "Active", "Élevé",
     "Communiqué co-signé nommant Partech explicitement.",
     "Partech news ; Fusacq ; Usine Digitale"),
    ("Partech", "Partech Seed (I to IV)", "Orus", "InsurTech (cœur)", "France",
     "Seed", "2022-07", 5, "€", "Co-investisseur", "Active", "Moyen",
     "Réinvesti en Série A (2023-08/10, ~10-11 M€). Série B (juin 2025, 25 M€) sans Partech cité.",
     "EU-Startups ; Partech news ; TechCrunch"),

    ("Ring Capital", "Mission I", "Goodvest", "FinTech liée assurance", "France",
     "n/d", "2023-11", 10, "€", "Investisseur (nouvel entrant)", "Active", "Moyen",
     "Réinvesti en Série B (2025-09, 12 M€, mené par Serena Capital InnovAllianz III) comme « investisseur historique ». Rattachement Mission I déduit du calibre du ticket, non nommé.",
     "Maddyness ; Les Horizons ; Planet Fintech ; mind Fintech ; Tech.eu"),

    ("Founders Future VC", "Founders Future Good", "Neat", "InsurTech (cœur)", "France",
     "Amorçage (seed)", "2022-10", 10, "€", "Investisseur", "Active", "Faible-Moyen",
     "Réinvesti en Série A (2024-09, 50 M€, mené par Hedosophia). Rattachement véhicule (Good vs Fund I) déduit par millésime, non confirmé.",
     "Tech.eu ; fintech.global ; FinSMEs ; Argus de l'Assurance ; TechCrunch ; Le Journal des Entreprises ; Presse-citron ; Finyear"),
    ("Founders Future VC", "Founders Future Good", "Napo", "InsurTech (cœur)", "Royaume-Uni",
     "n/d", None, None, None, "Investisseur (liste cumulative)", "À vérifier", "Faible",
     "Founders Future dans une liste cumulative d'investisseurs (avec DN Capital, Mercia), aucun communiqué daté nommant explicitement Founders Future ni le véhicule. À recouper.",
     "Agrégateur (liste investisseurs Napo)"),

    ("EOS Venture", "EVP I", "Concirrus", "InsurTech (cœur)", "Royaume-Uni",
     "n/d", "2018-04", None, None, "Investisseur (1ère participation du fonds)", "Actif", "Élevé",
     "Nommé explicitement comme 1ère participation d'EVP I.",
     "Insurance Journal ; PE Wire ; Coverager"),
    ("EOS Venture", "EVP I", "Digital Fineprint", "InsurTech (adjacent)", "Royaume-Uni",
     "Seed", "2016-12", None, None, "Investisseur (1ère participation du fonds)", "Acquis par hubb (déc. 2021)", "Élevé",
     "Nommé explicitement comme 1ère participation EVP I. Série A en 2019-08 (avec Pentech).",
     "PE Wire ; Crunchbase"),
    ("EOS Venture", "EVP I", "Westhill", "InsurTech (cœur)", "US/Canada",
     "Série A", "2018-12", None, None, "Lead/co-investisseur", "Actif", "Moyen",
     "Round ultérieur de 13,5 M$ (2023, avec Luge Capital, Nyca Partners) probablement hors EVP I.",
     "Westhill press"),
    ("EOS Venture", "EVP I", "Buckle", "InsurTech (cœur)", "US",
     "Série A", "2020-08", 31, "$", "Co-lead (avec HSCM Bermuda)", "Actif", "Élevé",
     "Cohérent avec la période de déploiement d'EVP I (fonds clos à 85 M$ en 2020).",
     "Reinsurance News ; BusinessWire"),
    ("EOS Venture", "EVP I", "Roadzen", "InsurTech (cœur)", "Inde/US",
     "Série A", "2019", None, None, "Investisseur", "n/d", "Moyen",
     "Dates divergentes selon les sources (mai 2019 vs janv. 2020) ; montant non trouvé.",
     "Eos portfolio page (titre indexé)"),
    ("EOS Venture", "EVP I", "Ticker", "InsurTech (cœur)", "Royaume-Uni",
     "Série B", "2021-06", None, None, "Investisseur (avec Munich Re Ventures)", "Actif", "Moyen",
     "Tour tardif dans le cycle de vie théorique d'EVP I ; attribution possible mais non confirmée explicitement.",
     "Crunchbase/Tracxn"),
    ("EOS Venture", "EVP I", "Neos Insurance", "InsurTech (cœur)", "Royaume-Uni",
     "Seed", "2016-10", 1.22, "$", "Investisseur", "Racheté par Sky (2021-06)", "Faible",
     "Investissement antérieur à la formation formelle du fonds (annoncé 2018) — probable investissement pré-fonds (angel) plutôt qu'EVP I stricto sensu.",
     "Tracxn/Crunchbase"),

    ("Atlantic Vantage Point", "AVP early stage I", "ClimateSecure", "InsurTech (cœur, paramétrique)", "France",
     "n/d", "2014", None, None, "Co-fondateur/investisseur (AXA Seed Factory)", "n/d", "Moyen",
     "Une des 5 sociétés héritées d'AXA Seed Factory (2013), reprises par AXA Strategic Ventures en 2015. Rattachement au véhicule nommé « AVP early stage I » non confirmé littéralement.",
     "Artemis.bm"),
    ("Atlantic Vantage Point", "AVP early stage I", "Particeep", "FinTech liée assurance", "France",
     "Partenariat/participation", "2015-10", None, None, "Investisseur historique (lignée AXA Seed Factory)", "Actif", "Moyen",
     "Lien AXA confirmé (partenariat produit AXA Creditor) ; montant/tour non trouvés.",
     "Crowdfund Insider"),
    ("Atlantic Vantage Point", "AVP early stage I", "FundShop", "FinTech liée assurance", "France",
     "n/d", None, None, None, "AXA cité parmi les actionnaires", "Actif", "Moyen",
     "Rattachement véhicule non confirmé.",
     "Argus de l'Assurance"),
    ("Atlantic Vantage Point", "AVP early stage I", "Widmee", "InsurTech (adjacent, data/marketing)", "France",
     "Seed", "2014-11", 0.21, "€", "Investisseur (AXA Seed Factory)", "n/d (statut incertain)", "Moyen", None,
     "Maddyness"),
    ("Atlantic Vantage Point", "AVP early stage II", "PolicyGenius", "FinTech liée assurance", "US",
     "Série A", "2015-06", 5.3, "$", "Investisseur récurrent (4+ tours, 2015-2022)", "Actif", "Faible",
     "AVP a investi sur au moins 4 tours successifs (A 2015 5,3 M$ ; B 15 M$ ; D 2020 100 M$ ; E 2022 125 M$) chevauchant les deux véhicules (I=2013, II=2019). Rattaché ici à Early Stage II car la majorité des tours (D, E) tombent dans sa fenêtre de déploiement ; impossible de trancher avec certitude.",
     "Global Venturing ; Citybiz"),

    ("Elaia Partners", "Elaia Delta", "Shift Technology", "InsurTech (cœur)", "France",
     "Série B", "2017-10", 28, "$", "Investisseur historique (suiveur, co-mené Accel/General Catalyst)", "Active (privée, >1 Md$ valo 2021)", "Moyen",
     "Elaia investit depuis le seed (déc. 2014, fonds antérieur hors périmètre). Série D 2021 (220 M$, Advent) — Elaia non confirmée participante.",
     "Medium/Elaia ; FinSMEs ; elaia.com/portfolio"),
    ("Elaia Partners", "DV4", "Continuity", "InsurTech (cœur)", "France",
     "Seed", "2021-11", 5, "€", "Co-lead (avec Bpifrance, Kamet Ventures)", "À vérifier (tour suivant 10 M€ annoncé)", "Moyen",
     "Rattachement incertain entre DV4 et Alpha II.",
     "EU-Startups ; Actuia ; Elaia communiqué PDF"),
    ("Elaia Partners", "DV4", "Seyna", "InsurTech (cœur)", "France",
     "Série A", "2022-02", 33, "€", "Co-investisseur (avec White Star Capital)", "Active", "Élevé",
     "Seyna explicitement citée par Elaia comme participation DV4. Tour de suivi ~10 M€ (2024) avec 115K, White Star, Elaia.",
     "EU-Startups (DV4 closing) ; newsassurancespro ; FUSACQ"),

    ("Alven Capital", "Alven VI", "Stoïk", "InsurTech (cœur)", "France",
     "Seed", "2022-01", 3.8, "€", "Co-lead (avec Anthemis Group, Kima Ventures)", "Active", "Élevé",
     "Série B confirmée (2024-10, 25 M€, « Alven Capital Partners a mené », conseil juridique Goodwin) et Série C (2026-01, 20 M€, avec a16z).",
     "Tech.eu ; Goodwin ; Le Monde du Droit ; Tribune de l'Assurance ; Maddyness ; Finyear"),

    ("Serena Capital", "Serena III", "Descartes Underwriting", "InsurTech (cœur)", "France",
     "Série A", "2020-09", 15.7, "€", "Co-investisseur (parité avec Cathay Innovation)", "Active", "Moyen",
     "Rattachement Serena III déduit (millésime 2018 + date tour 2020), non nommé.",
     "mind Fintech ; Cathay Capital ; BlackFin substack"),
    ("Serena Capital", "Serena III", "Acheel", "InsurTech (cœur)", "France",
     "n/d", "2021", 29, "€", "Co-investisseur (avec Xavier Niel, Portag3 Ventures)", "À vérifier", "Faible",
     "Date exacte non confirmée.",
     "newsassurancespro ; LaDN Business ; FUSACQ"),

    ("ISAI", "ISAI Cap Venture", "Zelros", "InsurTech (adjacent)", "France",
     "Série A (extension)", "2021-02", 9, "€", "Co-investisseur (chef de file BGV, avec Plug and Play, HI Inov, 42CAP, astorya.vc)",
     "Exit cédée (rachetée par Earnix, annoncé 2025-05)", "Élevé",
     '"ISAI Cap Venture" nommément cité. Clients Zelros : CNP Assurances, MAIF, BPCE, AXA, Matmut.',
     "Societe.Tech ; Frenchweb ; La Tribune de l'Assurance ; Earnix press release"),

    ("White Star Capital", "WSC II", "Clark", "InsurTech (cœur)", "Allemagne",
     "Série B", "2018-04", 29, "$", "Co-chef de file (avec Portag3 Ventures)", "Active", "Moyen",
     "1er ticket = Série B 2018, cohérent millésime WSC II. Montant incertain (29 M$ vs 20 M€ selon sources).",
     "TechCrunch ; Shackleton Ventures ; talent4boards"),
    ("White Star Capital", "WSC III", "Clark", "InsurTech (cœur)", "Allemagne",
     "Série C", "2021-01", 69, "€", "Participant (réinvestissement)", "Active", "Moyen",
     "Tour mené par Tencent ; White Star investisseur historique augmentant sa participation.",
     "FinTech Global ; AltFi ; InsurTech Digital ; finsmes"),
    ("White Star Capital", "WSC III", "Seyna", "InsurTech (cœur)", "France",
     "Série A", "2022-02", 33, "€", "Co-chef de file (avec Elaia Partners)", "Active", "Moyen",
     "Réinvesti en Série B (2025-09, 10 M€, mené par 115K/CVC La Banque Postale). Total cumulé Seyna : 57 M€.",
     "Medium/WSC billet officiel ; Tech.eu ; Frenchweb ; Maddyness ; mind Fintech ; La Tribune de l'Assurance"),

    ("Accel Partners", "Accel London VI", "Luko", "InsurTech (cœur)", "France",
     "Série A", "2019-11", 20, "€", "Lead", "Exit cédée (rachetée par Admiral Group/Allianz Direct, 2024)", "Moyen",
     "Accel a mené la Série A et participé à la Série B (déc. 2020, 50 M€, menée par EQT Ventures). Rattachement au millésime 2019 déduit, non confirmé nommément.",
     "Sifted ; Tech.eu ; PitchBook newsletter ; Astorya research"),
    ("Accel Partners", "Accel London VII", "Insify", "InsurTech (cœur)", "Pays-Bas",
     "Série A", "2022", 15, "€", "Lead", "Active", "Moyen",
     "Rattachement à Accel London VII déduit du millésime (2021), non confirmé nommément. Adossée à Munich Re.",
     "FFNews ; Accel.com noteworthy"),

    ("Breega Capital", "Breega Venture III", "Coverflex", "FinTech liée assurance / InsurTech adjacent", "Portugal",
     "Pre-seed", "2021-04", 5, "€", "Lead (avec 200M Fund)", "Active", "Faible-Moyen",
     "Lien assurance partiel (gestion assurance santé/prévoyance intégrée à une plateforme d'avantages salariés plus large). Rattachement à Venture III déduit par coïncidence de date.",
     "Tech.eu ; Nordic9 ; fintech.global"),
    ("Breega Capital", "Breega Seed II", "Cuvva", "InsurTech (cœur)", "Royaume-Uni",
     "Série A", "2019-12", 15, "£", "Co-investisseur (avec RTP Global et Digital Horizon)", "À vérifier", "Moyen",
     "Rattachement à Seed II (millésime 2018) déduit par proximité de date, non confirmé nommément.",
     "TechCrunch ; Insurance Times ; Yahoo Finance UK"),

    ("Cathay", "Cathay Innovation II", "Coherent", "InsurTech (cœur)", "Hong Kong",
     "Série A", "2020-11", 14, "$", "Lead (avec Franklin Templeton)", "Active (à vérifier)", "Moyen",
     "Le nom exact du véhicule (Fund II vs III) n'est cité dans aucune source ; chronologie (Fund III clôturé seulement en 2025) rend Fund II largement probable.",
     "Cathay Capital communiqué ; TechCrunch ; DealStreetAsia"),
    ("Cathay", "Cathay Innovation II", "Igloo", "InsurTech (cœur)", "Singapour",
     "Série A", "2020", None, "$", "Lead (Cathay Innovation)", "Active", "Moyen",
     "Série B initiale 19 M$ (mars 2022) étendue à 46 M$ (nov. 2022). Date précise (mois) non trouvée pour la Série A.",
     "DealStreetAsia ; AsiaTechDaily ; TechCrunch ; Igloo press"),
    ("Cathay", "Cathay Innovation II", "Qover", "InsurTech (cœur)", "Belgique",
     "Série B", "2023-01", 25, "$", "Investisseur (participant ; lead = Prime Ventures)", "Active", "Moyen", None,
     "Cathay Capital communiqué"),
    ("Cathay", "Cathay Innovation II", "Coverfy", "InsurTech (adjacent)", "Espagne",
     "n/d", None, None, None, "Investisseur (figure au portefeuille Cathay)", "À vérifier", "Faible",
     "Présence confirmée sur la page portefeuille dédiée de Cathay Innovation, mais date de tour, montant et rôle non trouvés.",
     "Cathay Innovation fiche société ; PitchBook"),

    ("Truffle Capital", "Truffle FinTech & InsurTech Fund II", "MoneyTrack", "FinTech liée assurance", "France",
     "Levées successives", "2020", 2.3, "€", "Investisseur historique (co-créée par Truffle Capital en 2018)",
     "Active (à vérifier)", "Moyen",
     "Co-créée par Truffle Capital dès 2018. Tour complémentaire de 2 M€ ultérieur. Client(s) connu(s) : MGEN, SwissLife (via Owello), Assia, Viamedis. Rattachement au Fund II (140 M€, closing déc. 2019) plausible par chronologie, non confirmé explicitement.",
     "Argus de l'Assurance ; FrenchWeb ; Truffle Capital fiche portefeuille ; LADN"),

    ("Index Venture", "Index Ventures XI", "Coalition", "InsurTech (cœur)", "US",
     "Série D", "2021-03", 175, "$", "Lead (Série D) puis investisseur existant (E, F)", "Actif", "Moyen",
     "Série E (2021-09, 205 M$) et Série F (2022-07, 250 M$, clôture alignée avec millésime Ventures XI). Le round D pourrait relever d'un fonds antérieur.",
     "Insurance Journal ; Coalition blog ; Index Ventures companies page"),
    ("Index Venture", "Index Ventures XII", "Thatch", "FinTech liée assurance / InsurTech adjacent", "US",
     "Série A", "2024-09", 38, "$", "Co-lead (avec General Catalyst)", "Actif", "Élevé",
     "Date (sept. 2024) bien alignée avec le millésime 2024 de Ventures XII, véhicule non nommé explicitement.",
     "PR Newswire ; Coverager"),
    ("Index Venture", "Index Ventures XII", "Alan", "InsurTech (cœur)", "France",
     "Levée de croissance", "2026-03", 100, "€", "Investisseur (avec Greenoaks, Kaaf)", "Actif (~5 Md€ valo)", "Moyen",
     "Index investisseur historique d'Alan depuis Série A 2018/B 2019 (hors périmètre, antérieur aux véhicules ciblés) ; seul le tour de mars 2026 rattachable à Ventures XII.",
     "TechCrunch ; Index Ventures perspectives"),

    ("speedInvest", "Speedinvest IV", "Grace", "InsurTech (adjacent)", "France/Suisse",
     "Seed", "2025-04", 5.9, "€", "Co-investisseur (co-lead avec FinTech Collective)", "Active", "Moyen",
     "Assurance/protection embarquée biens de luxe, partenariat Chubb. Véhicule déduit par date.",
     "TechCrunch ; TradedVC ; FinSMEs"),
    ("speedInvest", "Speedinvest IV", "Bliss (Saúde Bliss)", "InsurTech (adjacent)", "Brésil",
     "Seed", "2023", None, None, "Investisseur (non chef de file)", "Active", "Moyen",
     "Série A ultérieure (2026-03) co-menée par Kfund et Grupo Bradesco. Plateforme IA distribution assurance santé PME.",
     "Speedinvest blog ; K Fund ; Capital-Riesgo.es"),

    ("CommerzVentures", "CommerzVentures Fonds I à III", "Getsafe", "InsurTech (cœur)", "Allemagne",
     "Seed", "2015-10", None, None, "Investisseur", "Active", "Moyen",
     "Reconduit en Série B extension (2021-10, 63 M$ au total).",
     "CommerzVentures communiqué ; Nordic9"),
    ("CommerzVentures", "CommerzVentures Fonds I à III", "By Miles", "InsurTech (cœur)", "Royaume-Uni",
     "Série B", "2020-05", 15, "£", "Chef de file", "Active", "Élevé",
     "Avec Octopus Ventures, Insurtech Gateway, JamJar Investments.",
     "Finextra ; CommerzVentures communiqué ; Insurance Business Mag"),
    ("CommerzVentures", "CommerzVentures Fonds I à III", "Concirrus", "InsurTech (cœur)", "Royaume-Uni",
     "Série B (extension)", "2020-07", 6, "$", "Investisseur (extension menée par CommerzVentures)", "Active", "Élevé",
     "Extension après Série B initiale de 20 M$ menée par AlbionVC.",
     "PE Wire ; Intelligent Insurer ; CommerzVentures communiqué"),
    ("CommerzVentures", "CommerzVentures Fonds I à III", "Afilio", "FinTech liée assurance", "Allemagne",
     "Série A", "2021-08", 13, "$", "Chef de file", "Active", "Élevé",
     "Avec Cherry Ventures, Speedinvest, Cavalry Ventures.",
     "CommerzVentures communiqué ; Life Insurance International ; Nordic9"),
    ("CommerzVentures", "CommerzVentures Fonds I à III", "INSTANDA", "InsurTech (cœur)", "Royaume-Uni",
     "Growth round", "2025-10", 20, "$", "Chef de file", "Active", "Élevé",
     "Avec Toscafund, Dale Ventures. Probablement Fonds III plutôt que I malgré le regroupement de la fiche.",
     "INSTANDA communiqué ; The Insurer ; Fintech.global"),

    ("13books Capital (ex-Element Ventures)", "Fonds I & II", "hepster", "InsurTech (cœur)", "Allemagne",
     "n/d", "2021", 10, "$", "Investisseur", "Active", "Moyen",
     "Déjà en portefeuille au closing du Fonds I (août 2021). Assurance embarquée mobilité/électronique/loisirs, partenariat HDI Embedded. Montant = total levé cumulé, pas le tour spécifique.",
     "13books portfolio ; TechCrunch (lancement fonds) ; Silicon Republic"),
    ("13books Capital (ex-Element Ventures)", "Fonds I & II", "Roadzen", "InsurTech (cœur)", "US (à vérifier)",
     "Série A", None, None, None, "Investisseur", "À vérifier", "Faible",
     "Mention isolée agrégateur, non corroborée par une 2e source.",
     "Superscout"),
]

RE_ANNEE = re.compile(r"(\d{4})")


def annee(date_str) -> float | None:
    if vide(date_str):
        return None
    m = RE_ANNEE.search(str(date_str))
    return float(m.group(1)) if m else None


def main() -> int:
    fonds = pd.read_excel(ENTREE, sheet_name="Fonds")
    part = pd.read_excel(ENTREE, sheet_name="Participations")
    tours = pd.read_excel(ENTREE, sheet_name="Tours_de_table")
    dico = pd.read_excel(ENTREE, sheet_name="Dictionnaire")
    ano = pd.read_excel(ENTREE, sheet_name="Anomalies")

    fonds_idx = fonds.set_index("Fonds_ID")
    part = part.copy()
    existants = {
        (row["Fonds_ID"], row["Start-up"]): idx for idx, row in part.iterrows()
    }

    CHAMPS_COMPLEMENT = [
        "Taille du tour (M€)", "Rôle du véhicule (détail)", "Statut start-up (détail)",
        "Confiance", "Commentaires", "Source(s)",
    ]

    nouvelles_lignes = []
    non_rattachees = []
    mises_a_jour = []

    for (nom_fonds, vehicule, startup, secteur, pays, tour, date, taille, devise,
         role, statut, confiance, commentaires, sources) in ROWS:
        fonds_id = VEHICULE_ID.get((nom_fonds, vehicule))
        if fonds_id is None or fonds_id not in fonds_idx.index:
            non_rattachees.append((nom_fonds, vehicule, startup))
            continue

        cle = (fonds_id, startup)
        if cle in existants:
            # Déjà présente (campagne de collecte initiale) : on complète les
            # colonnes de granularité web sans dupliquer la ligne.
            idx = existants[cle]
            complements = {
                "Taille du tour (M€)": convertir(taille, devise),
                "Rôle du véhicule (détail)": role,
                "Statut start-up (détail)": statut,
                "Confiance": confiance,
                "Commentaires": commentaires,
                "Source(s)": sources,
            }
            for champ, valeur in complements.items():
                if vide(part.at[idx, champ]) and not vide(valeur):
                    part.at[idx, champ] = valeur
            mises_a_jour.append((nom_fonds, vehicule, startup))
            continue

        nouvelles_lignes.append({
            "Fonds_ID": fonds_id,
            "Nom du fonds": nom_fonds,
            "Véhicule": vehicule,
            "Start-up": startup,
            "Secteur": secteur,
            "Stage financé": tour if not vide(tour) else None,
            "Montant investi par le fonds (M€)": None,
            "Total levé par la start-up (M€)": None,
            "Date de création": None,
            "Date d’investissement": annee(date),
            "Pays d’origine": pays,
            "Nb pays d’implantation": None,
            "Positionnement (Leader/Minoritaire)": positionnement(role),
            "Capital social (k€)": None,
            "CA (M€)": None,
            "Valorisation (M€)": None,
            "Nb fonds investisseurs": None,
            "Nb tours (fonds)": None,
            "Nb tours (total)": None,
            "Repositionnement": None,
            "Exit (O/N)": exit_on(statut),
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

    # Anomalies : gaps et cas structurels rencontrés pendant la recherche.
    prochain_id = int(ano["Anomalie_ID"].str.extract(r"(\d+)")[0].astype(int).max()) + 1
    nouvelles_anomalies = [
        {
            "Anomalie_ID": f"ANO-{prochain_id:04d}",
            "Type d'anomalie": "Participation non rattachable à un véhicule suivi",
            "Onglet source": "Participations",
            "Table ou bloc source": "Recherche web véhicules Fermé (Closed-Lot E, 16/09/2026)",
            "Ligne source": None,
            "Entité concernée": "Breega Capital",
            "Champ concerné": "Véhicule",
            "Valeur source": "Mila (déc. 2021, néo-assureur immobilier) financé via « F/I Ventures », véhicule dédié Crédit Agricole antérieur à F/I Venture II",
            "Valeur retenue": "Non intégré (aucun véhicule correspondant dans la base)",
            "Candidats éventuels": "F/I Venture II (millésime 2022, postérieur au tour) — écarté par prudence",
            "Score de similarité": None,
            "Niveau de confiance": "Moyen",
            "Traitement appliqué": "Ligne non ajoutée, signalée pour vérification (véhicule « F/I Ventures » prédécesseur non suivi dans la base Fonds)",
            "Commentaire": "mind Fintech, La Tribune de l'Assurance, CFNews Immo. Si un véhicule « F/I Ventures » (prédécesseur, pré-2022) est ajouté à la base, cette participation devra y être rattachée.",
        },
        {
            "Anomalie_ID": f"ANO-{prochain_id + 1:04d}",
            "Type d'anomalie": "Participations hors périmètre du véhicule suivi",
            "Onglet source": "Participations",
            "Table ou bloc source": "Recherche web véhicules Fermé (Closed-Lot E, 16/09/2026)",
            "Ligne source": None,
            "Entité concernée": "Eurazeo",
            "Champ concerné": "Véhicule",
            "Valeur source": "InsuranceDekho (Série B, oct. 2023, 60 M$) et Igloo (pre-Série C, déc. 2023, 36 M$)",
            "Valeur retenue": "Non intégrées au véhicule « Eurazeo Venture Capital »",
            "Candidats éventuels": "Fonds insurtech dédié Eurazeo/BNP Paribas Cardif (~200 M$, focus Asie) — non suivi dans la base Fonds",
            "Score de similarité": None,
            "Niveau de confiance": "Moyen",
            "Traitement appliqué": "Lignes non ajoutées, signalées pour vérification",
            "Commentaire": "Eurazeo précise gérer ces investissements via un fonds insurtech dédié pour le compte de BNP Paribas Cardif, distinct de la ligne généraliste « Eurazeo Venture Capital » (millésime 2023) suivie dans la base. Newsroom Eurazeo, Business Standard, Reinsurance News, Zee Business, Acnnewswire, IBS Intelligence.",
        },
        {
            "Anomalie_ID": f"ANO-{prochain_id + 2:04d}",
            "Type d'anomalie": "Existence du véhicule non confirmée publiquement",
            "Onglet source": "Fonds",
            "Table ou bloc source": "Recherche web véhicules Fermé (Closed-Lot F, 16/09/2026)",
            "Ligne source": None,
            "Entité concernée": "Truffle Capital",
            "Champ concerné": "Véhicule",
            "Valeur source": "Truffle FinTech & InsurTech Fund III (prévu 2025)",
            "Valeur retenue": "Conservé tel quel (à confirmer)",
            "Candidats éventuels": None,
            "Score de similarité": None,
            "Niveau de confiance": "Faible",
            "Traitement appliqué": "Conservé, signalé pour vérification",
            "Commentaire": "Aucune source publique trouvée confirmant l'existence, le closing ou le portefeuille d'un « Fund III » 2025 sous ce nom — statut à reconfirmer directement auprès de Truffle Capital.",
        },
    ]
    ano_out = pd.concat([ano, pd.DataFrame(nouvelles_anomalies)], ignore_index=True)

    with pd.ExcelWriter(SORTIE, engine="openpyxl") as writer:
        fonds.to_excel(writer, sheet_name="Fonds", index=False)
        part_out.to_excel(writer, sheet_name="Participations", index=False)
        tours.to_excel(writer, sheet_name="Tours_de_table", index=False)
        dico.to_excel(writer, sheet_name="Dictionnaire", index=False)
        ano_out.to_excel(writer, sheet_name="Anomalies", index=False)

    print(f"Classeur produit : {SORTIE}")
    print(f"Participations avant : {len(part)}  →  après : {len(part_out)}  (+{len(nouvelles_lignes)})")
    print(f"Anomalies avant : {len(ano)}  →  après : {len(ano_out)}")
    if mises_a_jour:
        print(f"Lignes existantes complétées (pas dupliquées) : {len(mises_a_jour)}")
        for nf, v, su in mises_a_jour:
            print("  -", nf, "/", v, "/", su)
    if non_rattachees:
        print(f"Lignes non rattachées (véhicule introuvable) : {len(non_rattachees)}")
        for nf, v, su in non_rattachees:
            print("  -", nf, "/", v, "/", su)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
