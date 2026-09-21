#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fusionner_cvc_recherche_externe.py — Intègre la recherche CVC lancée par l'utilisateur
sur un autre outil (tableau1_fonds.csv + tableau2_participations.csv)
========================================================================================

Traite deux tableaux CSV structurés (méthodologie identique : taxonomie stricte,
confiance par ligne, sources, anti-fabrication — cf. base_vc_cvc_assurance.md) :
- Dédoublonne contre la base existante (v16) sur (Fonds_ID normalisé, Start-up normalisé).
- Ajoute 13 nouveaux véhicules CVC (Zurich, Tokio Marine Future Fund, Nationwide Ventures,
  MassMutual Ventures, Guardian Strategic Ventures, New York Life Ventures, Optum Ventures,
  QBE Ventures, IAG Firemark Ventures, Intact Ventures, Liberty Mutual Strategic Ventures,
  Achmea Innovation Fund, Transamerica Ventures [inactif]).
- Enrichit 7 véhicules déjà en base (Allianz X, MS&AD Ventures, Munich Re Ventures, Sompo,
  EOS Venture, Truffle Capital) avec les nouvelles participations documentées.
- Exclut explicitement : 3 lignes MS&AD trop faibles/non vérifiées (Boop, Big Ticket,
  Clarity), 1 ligne MS&AD à risque d'homonymie (Tomorrow), 1 ligne Intact Ventures
  attribuée à la mauvaise entité (Shepherd → Intact Private Capital, pas Intact Ventures),
  2 lignes non rattachables à un véhicule documenté (Index Ventures/Alan, Serena/Descartes
  Underwriting — déjà en base sous Serena III).

Entrée  : VC_Database_Standardisee_v16.xlsx
Sortie  : VC_Database_Standardisee_v17.xlsx

Usage : python3 fusionner_cvc_recherche_externe.py
"""

from __future__ import annotations

import re
import unicodedata
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v16.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v17.xlsx"
UPLOADS = Path("/root/.claude/uploads/963315b8-930e-538e-a15f-800e2ff2b2d1")
CSV_FONDS = UPLOADS / "f6a8589a-tableau1_fonds.csv"
CSV_PART = UPLOADS / "dcf78168-tableau2_participations.csv"

TAUX = {"EUR": 1.0, "USD": 0.92, "GBP": 1.17, "AUD": 0.60, "JPY": 0.0061}
MONTANT_RE = re.compile(r"(\d+(?:[.,]\d+)?)\s*M\s*(EUR|USD|GBP|AUD|JPY)", re.I)


def norm(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def clean(txt):
    if txt is None or (isinstance(txt, float) and pd.isna(txt)):
        return None
    txt = str(txt).replace("**", "")
    txt = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r"\1 (\2)", txt)
    txt = txt.strip()
    return txt or None


_NO_OVERRIDE = object()


def parse_montant(txt, override=_NO_OVERRIDE):
    if override is not _NO_OVERRIDE:
        return override
    if txt is None or (isinstance(txt, float) and pd.isna(txt)):
        return None
    txt = str(txt)
    if "non trouv" in txt.lower() or "s.o." == txt.strip():
        return None
    m = MONTANT_RE.search(txt)
    if not m:
        return None
    val = float(m.group(1).replace(",", "."))
    dev = m.group(2).upper()
    return round(val * TAUX.get(dev, 1.0), 2)


# --- 13 nouveaux véhicules CVC (métadonnées issues de tableau1_fonds.csv) ---
FONDS_NOUVEAUX = [
    {
        "Nom du fonds": "Zurich Insurance Group", "Véhicule": "Investissements stratégiques directs",
        "Fonds_ID": "zurich-insurance-group_investissements-strategiques", "AuM (M€)": None,
        "Statut": "Actif (pas de véhicule CVC formel)", "Phase": "Investissement", "Millésime": None,
        "Géographie": "Europe", "Stratégie": "InsurTech ; prises de participation stratégiques",
        "Stages pratiqués": None, "Pré-Seed": 0.0, "Seed": 0.0, "Pré-Série A": 0.0, "Série A": 1.0,
        "Série B": 0.0, "Série C": 0.0, "Série D": 0.0,
        "Nb participations (déclaré)": None, "Nb participations documentées": 4, "Nb exits": 0.0,
        "Nb InsurTech": 4.0, "Site web": "https://www.zurich.com",
        "Source_AuM": (
            "[Recherche externe utilisateur, 21/09/2026] Société mère : Zurich Insurance Group. "
            "Pas de véhicule de CVC formel identifié — « Zurich Global Ventures » est une unité opérationnelle "
            "(services/distribution), pas un fonds VC ; le « Zurich Innovation Championship » est un programme "
            "d'accélération, pas un véhicule d'investissement. Prises de participation stratégiques directes au "
            "bilan. Sources : Zurich Innovation Championship, Coverager."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "Zurich Insurance Group",
    },
    {
        "Nom du fonds": "Tokio Marine Future Fund", "Véhicule": "Tokio Marine Future Fund",
        "Fonds_ID": "tokio-marine-future-fund_tokio-marine-future-fund", "AuM (M€)": 38.64,
        "Statut": "Actif", "Phase": "Investissement", "Millésime": 2022.0,
        "Géographie": "États-Unis (basé à Palo Alto)", "Stratégie": "InsurTech",
        "Stages pratiqués": "Seed ; Série A", "Pré-Seed": 0.0, "Seed": 1.0, "Pré-Série A": 0.0,
        "Série A": 1.0, "Série B": 0.0, "Série C": 0.0, "Série D": 0.0,
        "Nb participations (déclaré)": 19.0, "Nb participations documentées": 8, "Nb exits": 0.0,
        "Nb InsurTech": 8.0, "Site web": "https://www.tmfuturefund.com",
        "Source_AuM": (
            "[Recherche externe utilisateur, 21/09/2026] Société mère : Tokio Marine Holdings. Dotation "
            "42 M USD annoncée au lancement (avril 2022) → conv. 38,64 M€. 19 sociétés listées sur la page "
            "portefeuille officielle. Sources : tmfuturefund.com, Business Wire (19/04/2022)."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "Tokio Marine Holdings",
    },
    {
        "Nom du fonds": "Nationwide Ventures", "Véhicule": "Nationwide Ventures",
        "Fonds_ID": "nationwide-ventures_nationwide-ventures", "AuM (M€)": 322.0,
        "Statut": "Actif", "Phase": "Investissement", "Millésime": None,
        "Géographie": "États-Unis", "Stratégie": "InsurTech",
        "Stages pratiqués": "Early stage", "Pré-Seed": 0.0, "Seed": 1.0, "Pré-Série A": 0.0,
        "Série A": 1.0, "Série B": 0.0, "Série C": 0.0, "Série D": 0.0,
        "Nb participations (déclaré)": 34.0, "Nb participations documentées": 13, "Nb exits": 1.0,
        "Nb InsurTech": 13.0, "Site web": "https://www.nationwide.com/personal/about-us/ventures",
        "Source_AuM": (
            "[Recherche externe utilisateur, 21/09/2026] Société mère : Nationwide Mutual Insurance Company. "
            "Fonds porté à 350 M USD en février 2021 → conv. 322 M€. 34 sociétés listées sur la page "
            "portefeuille officielle. Sources : Nationwide Ventures, Nationwide News (11/02/2021)."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "Nationwide Mutual Insurance Company",
    },
    {
        "Nom du fonds": "MassMutual Ventures", "Véhicule": "MassMutual Ventures",
        "Fonds_ID": "massmutual-ventures_massmutual-ventures", "AuM (M€)": None,
        "Statut": "Actif", "Phase": "Investissement", "Millésime": None,
        "Géographie": "Amérique du Nord (partenariat Crane Venture Partners : Europe, Inde, APAC)",
        "Stratégie": "InsurTech ; Climate", "Stages pratiqués": "Early stage / multistage",
        "Pré-Seed": 0.0, "Seed": 1.0, "Pré-Série A": 0.0, "Série A": 0.0, "Série B": 1.0,
        "Série C": 0.0, "Série D": 0.0,
        "Nb participations (déclaré)": None, "Nb participations documentées": 1, "Nb exits": None,
        "Nb InsurTech": 1.0, "Site web": "https://www.massmutualventures.com",
        "Source_AuM": (
            "[Recherche externe utilisateur, 21/09/2026] Société mère : Massachusetts Mutual Life Insurance "
            "Co. AuM global non publié ; Climate Technology Fund II = 150 M USD (véhicule sœur, hors "
            "périmètre insurtech strict). Source : massmutualventures.com."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "Massachusetts Mutual Life Insurance Co.",
    },
    {
        "Nom du fonds": "Guardian Strategic Ventures", "Véhicule": "Guardian Strategic Ventures",
        "Fonds_ID": "guardian-strategic-ventures_guardian-strategic-ventures", "AuM (M€)": None,
        "Statut": "Existe mais très peu actif (pic 2021, <2 tours/an)", "Phase": "Investissement",
        "Millésime": None, "Géographie": "États-Unis", "Stratégie": "InsurTech",
        "Stages pratiqués": None, "Pré-Seed": 0.0, "Seed": 0.0, "Pré-Série A": 0.0, "Série A": 0.0,
        "Série B": 1.0, "Série C": 0.0, "Série D": 0.0,
        "Nb participations (déclaré)": 3.0, "Nb participations documentées": 1, "Nb exits": None,
        "Nb InsurTech": 1.0, "Site web": "https://www.guardianventures.com",
        "Source_AuM": (
            "[Recherche externe utilisateur, 21/09/2026] Société mère : The Guardian Life Insurance Company "
            "of America. AuM et tickets non publiés (fourchette 10-50 M USD selon source secondaire, non "
            "confirmée). Sources : Unicorn Nest, CB Insights (fiche publique)."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "The Guardian Life Insurance Company of America",
    },
    {
        "Nom du fonds": "New York Life Ventures", "Véhicule": "NYL Ventures",
        "Fonds_ID": "new-york-life-ventures_nyl-ventures", "AuM (M€)": 920.0,
        "Statut": "Actif", "Phase": "Investissement", "Millésime": 2012.0,
        "Géographie": "États-Unis", "Stratégie": "InsurTech",
        "Stages pratiqués": "Seed à croissance", "Pré-Seed": 0.0, "Seed": 1.0, "Pré-Série A": 0.0,
        "Série A": 1.0, "Série B": 0.0, "Série C": 1.0, "Série D": 0.0,
        "Nb participations (déclaré)": 28.0, "Nb participations documentées": 4, "Nb exits": None,
        "Nb InsurTech": 4.0, "Site web": "https://www.nylventures.com",
        "Source_AuM": (
            "[Recherche externe utilisateur, 21/09/2026] Société mère : New York Life Insurance Company. "
            "AuM >1 Md USD → conv. ~920 M€. 28 sociétés listées sur la page portefeuille, 300+ POC réalisés "
            "depuis 2012. Sources : nylventures.com, New York Life Newsroom (10e anniversaire)."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "New York Life Insurance Company",
    },
    {
        "Nom du fonds": "Optum Ventures", "Véhicule": "Optum Ventures",
        "Fonds_ID": "optum-ventures_optum-ventures", "AuM (M€)": None,
        "Statut": "Actif (majoritairement health tech, quelques lignes assurance)", "Phase": "Investissement",
        "Millésime": None, "Géographie": "États-Unis ; Royaume-Uni",
        "Stratégie": "Santé numérique ; quelques lignes InsurTech/FinTech liée assurance",
        "Stages pratiqués": "Early stage", "Pré-Seed": 0.0, "Seed": 0.0, "Pré-Série A": 0.0,
        "Série A": 0.0, "Série B": 1.0, "Série C": 0.0, "Série D": 0.0,
        "Nb participations (déclaré)": 90.0, "Nb participations documentées": 3, "Nb exits": None,
        "Nb InsurTech": 3.0, "Site web": "https://www.optumventures.com",
        "Source_AuM": (
            "[Recherche externe utilisateur, 21/09/2026] Société mère : UnitedHealth Group / Optum. AuM non "
            "publié. ~90 sociétés au portefeuille, quasi-totalité santé numérique/care delivery (hors "
            "périmètre) — seules 3 lignes retenues comme InsurTech/FinTech liée assurance. Source : "
            "optumventures.com/portfolio."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "UnitedHealth Group / Optum",
    },
    {
        "Nom du fonds": "QBE Ventures", "Véhicule": "QBE Ventures",
        "Fonds_ID": "qbe-ventures_qbe-ventures", "AuM (M€)": None,
        "Statut": "Actif", "Phase": "Investissement", "Millésime": None,
        "Géographie": "Mondiale", "Stratégie": "InsurTech",
        "Stages pratiqués": "Early-mid stage", "Pré-Seed": 0.0, "Seed": 1.0, "Pré-Série A": 0.0,
        "Série A": 0.0, "Série B": 1.0, "Série C": 0.0, "Série D": 0.0,
        "Nb participations (déclaré)": 21.0, "Nb participations documentées": 18, "Nb exits": 7.0,
        "Nb InsurTech": 18.0, "Site web": "https://www.qbe.com/ventures",
        "Source_AuM": (
            "[Recherche externe utilisateur, 21/09/2026] Société mère : QBE Insurance Group. Tickets "
            "2-5 M USD. 14 sociétés actives + 7 exits listés sur la page portefeuille officielle. Source : "
            "qbe.com/ventures/portfolio."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "QBE Insurance Group",
    },
    {
        "Nom du fonds": "IAG Firemark Ventures", "Véhicule": "IAG Firemark Ventures",
        "Fonds_ID": "iag-firemark-ventures_iag-firemark-ventures", "AuM (M€)": None,
        "Statut": "Actif", "Phase": "Investissement", "Millésime": None,
        "Géographie": "Australie ; Nouvelle-Zélande ; États-Unis ; Israël ; Europe",
        "Stratégie": "InsurTech", "Stages pratiqués": "Early stage à croissance",
        "Pré-Seed": 0.0, "Seed": 1.0, "Pré-Série A": 0.0, "Série A": 0.0, "Série B": 1.0,
        "Série C": 0.0, "Série D": 0.0,
        "Nb participations (déclaré)": None, "Nb participations documentées": 10, "Nb exits": None,
        "Nb InsurTech": 10.0, "Site web": "https://iagfiremarkventures.com",
        "Source_AuM": (
            "[Recherche externe utilisateur, 21/09/2026] Société mère : Insurance Australia Group (IAG). "
            "AuM non publié. 30+ sociétés listées sur la page portefeuille officielle (premiers "
            "investissements datés 2016), investissements récents datés 2025-2026. Source : "
            "iagfiremarkventures.com/portfolio."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "Insurance Australia Group",
    },
    {
        "Nom du fonds": "Intact Ventures", "Véhicule": "Intact Ventures",
        "Fonds_ID": "intact-ventures_intact-ventures", "AuM (M€)": 552.0,
        "Statut": "Actif", "Phase": "Investissement", "Millésime": None,
        "Géographie": "Canada ; États-Unis ; Royaume-Uni ; Europe ; Brésil",
        "Stratégie": "InsurTech", "Stages pratiqués": "Série A ; Série B",
        "Pré-Seed": 0.0, "Seed": 0.0, "Pré-Série A": 0.0, "Série A": 1.0, "Série B": 1.0,
        "Série C": 0.0, "Série D": 0.0,
        "Nb participations (déclaré)": 26.0, "Nb participations documentées": 13, "Nb exits": 1.0,
        "Nb InsurTech": 13.0, "Site web": "https://www.intactfc.com/about-us/intact-ventures",
        "Source_AuM": (
            "[Recherche externe utilisateur, 21/09/2026] Société mère : Intact Financial Corporation. "
            "600 M USD/CAD (devise non précisée) sur trois fonds → retenu 552 M€ (conversion USD prudente). "
            "26 investissements directs + 14 fonds (stratégie fund-of-funds distincte, non modélisée ici) ; "
            "« 30+ » sociétés au total. ⚠️ Le tour Shepherd (42 M USD, mars 2026) est mené par « Intact "
            "Private Capital », entité distincte du CVC Intact Ventures — délibérément non attribué ici. "
            "Source : intactfc.com/about-us/intact-ventures."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "Intact Financial Corporation",
    },
    {
        "Nom du fonds": "Liberty Mutual Strategic Ventures", "Véhicule": "Liberty Mutual Strategic Ventures (LMSV)",
        "Fonds_ID": "liberty-mutual-strategic-ventures_lmsv", "AuM (M€)": 184.0,
        "Statut": "Actif", "Phase": "Investissement", "Millésime": 2015.0,
        "Géographie": "États-Unis ; Europe", "Stratégie": "InsurTech",
        "Stages pratiqués": "Seed à Série B", "Pré-Seed": 0.0, "Seed": 1.0, "Pré-Série A": 0.0,
        "Série A": 0.0, "Série B": 1.0, "Série C": 0.0, "Série D": 0.0,
        "Nb participations (déclaré)": 20.0, "Nb participations documentées": 8, "Nb exits": 7.0,
        "Nb InsurTech": 8.0, "Site web": "https://www.lmstrategicventures.com",
        "Source_AuM": (
            "[Recherche externe utilisateur, 21/09/2026] Société mère : Liberty Mutual Insurance. Fonds II = "
            "200 M USD (lancé 2024) → conv. 184 M€. Tickets initiaux 750 k-5 M USD. 20 sociétés actives + "
            "7 exits. Source : lmstrategicventures.com, Agency Checklists (29/07/2024)."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "Liberty Mutual Insurance",
    },
    {
        "Nom du fonds": "Achmea Innovation Fund", "Véhicule": "Achmea Innovation Fund",
        "Fonds_ID": "achmea-innovation-fund_achmea-innovation-fund", "AuM (M€)": None,
        "Statut": "Actif (portefeuille non publié)", "Phase": "Investissement", "Millésime": None,
        "Géographie": "Pays-Bas ; Europe", "Stratégie": "InsurTech",
        "Stages pratiqués": "Pré-seed/seed à Série C+", "Pré-Seed": 1.0, "Seed": 1.0, "Pré-Série A": 0.0,
        "Série A": 0.0, "Série B": 0.0, "Série C": 1.0, "Série D": 0.0,
        "Nb participations (déclaré)": None, "Nb participations documentées": 0, "Nb exits": None,
        "Nb InsurTech": None, "Site web": "https://www.achmeainnovationfund.nl",
        "Source_AuM": (
            "[Recherche externe utilisateur, 21/09/2026] Société mère : Achmea. Tickets 100 k€-1,5 M€. "
            "Existe et actif, mais aucun portefeuille publié sur les sources accessibles (ajouté pour "
            "cartographie concurrentielle, sans participation documentée à ce stade). Source : EU-Startups."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "Achmea",
    },
    {
        "Nom du fonds": "Transamerica Ventures", "Véhicule": "Transamerica Ventures",
        "Fonds_ID": "transamerica-ventures_transamerica-ventures", "AuM (M€)": None,
        "Statut": "Inactif (portefeuille cédé à Montana Capital Partners, véhicule essaimé hors du groupe)",
        "Phase": "Clos", "Millésime": 2013.0, "Géographie": "États-Unis ; Europe",
        "Stratégie": "InsurTech", "Stages pratiqués": None,
        "Pré-Seed": 0.0, "Seed": 0.0, "Pré-Série A": 0.0, "Série A": 0.0, "Série B": 0.0,
        "Série C": 0.0, "Série D": 0.0,
        "Nb participations (déclaré)": None, "Nb participations documentées": 0, "Nb exits": None,
        "Nb InsurTech": None, "Site web": "https://www.transamericaventures.com",
        "Source_AuM": (
            "[Recherche externe utilisateur, 21/09/2026] Société mère : Aegon / Transamerica. CVC arrêté : "
            "portefeuille corporate cédé à Montana Capital Partners, véhicule essaimé hors du groupe. Ajouté "
            "pour cartographie historique (concurrent devenu inactif). Sources : Montana Capital Partners "
            "(deal news), Global Venturing."
        ),
        "Type d'investisseur": "CVC", "Société mère (si CVC)": "Aegon / Transamerica",
    },
]

# Mapping "Fonds/Véhicule" (colonne CSV) -> Fonds_ID (existant ou nouveau)
FONDS_ID_MAP = {
    "allianz x": "allianz-x_allianz-x",
    "ms&ad ventures": "msad-ventures_msad-ventures",
    "munich re ventures / hsb": "munich-re-ventures_munich-re-ventures",
    "aviva ventures": "aviva-ventures_aviva-ventures",
    "sompo holdings (asia)": "sompo_digital-lab-light-vortex",
    "sompo holdings": "sompo_digital-lab-light-vortex",
    "sompo light vortex": "sompo_digital-lab-light-vortex",
    "maif avenir": "maif-avenir_maif-avenir",
    "eos venture partners": "eos-venture_evp-i",
    "truffle capital (vehicule non attribuable)": "truffle-capital_truffle-fintech-et-insurtech-fund-ii",
    "zurich insurance group": "zurich-insurance-group_investissements-strategiques",
    "tokio marine future fund": "tokio-marine-future-fund_tokio-marine-future-fund",
    "nationwide ventures": "nationwide-ventures_nationwide-ventures",
    "massmutual ventures": "massmutual-ventures_massmutual-ventures",
    "guardian strategic ventures": "guardian-strategic-ventures_guardian-strategic-ventures",
    "new york life ventures": "new-york-life-ventures_nyl-ventures",
    "optum ventures": "optum-ventures_optum-ventures",
    "qbe ventures": "qbe-ventures_qbe-ventures",
    "iag firemark ventures": "iag-firemark-ventures_iag-firemark-ventures",
    "intact ventures": "intact-ventures_intact-ventures",
    "liberty mutual strategic ventures": "liberty-mutual-strategic-ventures_lmsv",
}
NOMS_FONDS = {
    "allianz-x_allianz-x": ("Allianz X", "Allianz X"),
    "msad-ventures_msad-ventures": ("MS&AD Ventures", "MS&AD Ventures"),
    "munich-re-ventures_munich-re-ventures": ("Munich Re Ventures", "Munich Re Ventures (incl. HSB Ventures)"),
    "aviva-ventures_aviva-ventures": ("Aviva Ventures", "Aviva Ventures"),
    "sompo_digital-lab-light-vortex": ("Sompo", "Sompo Digital Lab / Sompo Light Vortex"),
    "maif-avenir_maif-avenir": ("MAIF Avenir", "MAIF Avenir"),
    "eos-venture_evp-i": ("EOS Venture", "EVP I"),
    "truffle-capital_truffle-fintech-et-insurtech-fund-ii": ("Truffle Capital", "Truffle FinTech & InsurTech Fund II"),
}
for f in FONDS_NOUVEAUX:
    NOMS_FONDS[f["Fonds_ID"]] = (f["Nom du fonds"], f["Véhicule"])

# Lignes explicitement exclues : (fonds CSV normalisé, start-up normalisée)
EXCLUDE = {
    ("ms&ad ventures", "boop"),
    ("ms&ad ventures", "big ticket"),
    ("ms&ad ventures", "clarity"),
    ("ms&ad ventures", "tomorrow"),
    ("intact ventures", "shepherd"),
    ("index ventures (vehicule non attribuable)", "alan"),
    ("serena (vehicule non attribuable, probablement serena iii)", "descartes underwriting"),
}

# Montants ambigus à ne pas retenir malgré un motif numérique présent dans le texte
MONTANT_OVERRIDE = {
    ("zurich insurance group", "ominimo"): None,
}


def main() -> int:
    fonds = pd.read_excel(ENTREE, sheet_name="Fonds")
    part = pd.read_excel(ENTREE, sheet_name="Participations").copy()
    tours = pd.read_excel(ENTREE, sheet_name="Tours_de_table")
    dico = pd.read_excel(ENTREE, sheet_name="Dictionnaire")
    ano = pd.read_excel(ENTREE, sheet_name="Anomalies")

    fonds_out = pd.concat([fonds, pd.DataFrame(FONDS_NOUVEAUX)], ignore_index=True)

    existants = {(norm(r["Fonds_ID"]), norm(r["Start-up"])) for _, r in part.iterrows()}

    csv = pd.read_csv(CSV_PART, sep=";", quotechar='"', encoding="utf-8")

    nouvelles_lignes = []
    ignorees_doublon = 0
    ignorees_exclusion = 0
    ignorees_non_mappees = 0
    for _, r in csv.iterrows():
        fonds_csv = r["Fonds/Véhicule"]
        startup = clean(r["Start-up"])
        key_exclude = (norm(fonds_csv), norm(startup))
        if key_exclude in EXCLUDE:
            ignorees_exclusion += 1
            continue
        fid = FONDS_ID_MAP.get(norm(fonds_csv))
        if fid is None:
            ignorees_non_mappees += 1
            continue
        if (norm(fid), norm(startup)) in existants:
            ignorees_doublon += 1
            continue
        existants.add((norm(fid), norm(startup)))  # anti-doublon intra-CSV (ex. Snapsheet répétée par fonds)

        nom_fonds, vehicule = NOMS_FONDS[fid]
        statut_csv = (r["Statut start-up (actif/exit)"] or "").strip().lower()
        statut_detail = {
            "actif": "Actif", "exit": "Exit", "à vérifier": "Actif (statut à vérifier — voir anomalie)",
        }.get(statut_csv, clean(r["Statut start-up (actif/exit)"]))
        role = clean(r["Rôle (lead/co-investisseur, avec qui)"])
        position = "Leader" if role and re.match(r"^lead\b", role, re.I) else "Minoritaire"
        override = MONTANT_OVERRIDE.get((norm(fonds_csv), norm(startup)))

        nouvelles_lignes.append({
            "Fonds_ID": fid, "Nom du fonds": nom_fonds, "Véhicule": vehicule,
            "Start-up": startup, "Secteur": clean(r["Secteur (taxonomie stricte)"]),
            "Stage financé": clean(r["Stade du tour"]),
            "Montant investi par le fonds (M€)": parse_montant(r["Montant investi par le fonds si connu"]),
            "Total levé par la start-up (M€)": None, "Date de création": None,
            "Date d’investissement": clean(r["Date (année)"]), "Pays d’origine": clean(r["Pays"]),
            "Nb pays d’implantation": None, "Positionnement (Leader/Minoritaire)": position,
            "Capital social (k€)": None, "CA (M€)": None, "Valorisation (M€)": None,
            "Nb fonds investisseurs": None, "Nb tours (fonds)": None, "Nb tours (total)": None,
            "Repositionnement": None, "Exit (O/N)": "O" if statut_csv == "exit" else "N",
            "MoC exit": None, "MoEP exit": None,
            "Taille du tour (M€)": parse_montant(r["Taille du tour (montant + devise)"], override),
            "% de détention": None, "Rôle du véhicule (détail)": role, "Statut start-up (détail)": statut_detail,
            "Confiance": clean(r["Confiance"]), "Commentaires": clean(r["Commentaire (1-2 phrases)"]),
            "Source(s)": clean(r["Source(s)"]),
        })

    part_out = pd.concat([part, pd.DataFrame(nouvelles_lignes)], ignore_index=True)

    prochain_id = int(ano["Anomalie_ID"].str.extract(r"(\d+)")[0].astype(int).max()) + 1
    ANOMALIES_NOUVELLES = [
        ("Participation à risque de statut", "Participations", "MS&AD Ventures / Tokio Marine Future Fund — ANZEN",
         "Statut start-up",
         "MGA/courtage assurance dirigeants pour PME, listée « actif » par les deux fonds",
         "Statut marqué « à vérifier »",
         "Une source presse (The Insurer, 30/01/2026) rapporte l'arrêt des activités de courtage et une "
         "réduction d'effectifs quelques mois après un Series A de 16 M USD — statut actif maintenu mais "
         "signalé comme fragile.", "Moyen", "The Insurer (30/01/2026), msad.vc/portfolio, tmfuturefund.com"),
        ("Participation non retenue (mauvaise entité)", "Participations", "Intact Ventures", "Start-up",
         "Shepherd (Série B, 42 M USD, mars 2026)", "Non intégrée",
         "Le tour est mené par « Intact Private Capital », entité distincte du CVC Intact Ventures au sein "
         "d'Intact Financial Corporation — attribution à Intact Ventures aurait été une erreur.",
         "Élevé", "PR Newswire (24/03/2026)"),
        ("Participations non retenues (confiance trop faible)", "Participations", "MS&AD Ventures", "Start-up",
         "Boop, Big Ticket, Clarity (taguées « Insurtech and Fintech » par MS&AD, activité non vérifiée)",
         "Non intégrées",
         "Confiance Faible et activité non vérifiée par source indépendante ; Clarity présente en outre un "
         "risque d'homonymie (nom très courant).", "Faible", "msad.vc/portfolio"),
        ("Participation non retenue (risque d'homonymie)", "Participations", "MS&AD Ventures", "Start-up",
         "Tomorrow (marquée « Exit »)", "Non intégrée",
         "Plusieurs sociétés portent ce nom (néobanque allemande, insurtech) ; également listée au "
         "portefeuille d'Intact Ventures sous le même nom — désambiguïsation impossible sans vérification "
         "complémentaire.", "Faible", "msad.vc/portfolio, Intact Ventures"),
        ("Rattachement à un véhicule impossible", "Fonds", "Index Ventures", "Participation Alan (2018)",
         "Rattachement demandé à « Index Ventures XII »",
         "Non intégrée (aucun véhicule Index documenté ne couvre 2018)",
         "L'investissement Alan (Series A, 23 M€, 2018) est antérieur à tous les véhicules Index actuellement "
         "documentés en base (Index Origin I/II : 2021-2022 ; Index Ventures XI/XII : 2022/2024). "
         "L'existence publique d'un véhicule nommé « Index Ventures XII » n'a d'ailleurs pas pu être "
         "confirmée par la recherche externe. Un véhicule de millésime antérieur à 2021 resterait à "
         "documenter séparément.", "Moyen", "Index Ventures, TechCrunch (31/07/2026)"),
        ("Divergence de montant/qualification non tranchée", "Participations", "Zurich Insurance Group — Ominimo",
         "Taille du tour", "10 M€ (Series A) selon une source ; « prise de participation » sans montant "
         "selon d'autres", "Montant non intégré (laissé vide)",
         "Divergence non tranchée par la recherche externe elle-même ; à confirmer sur communiqué Zurich "
         "avant d'intégrer un montant. Ominimo aurait par ailleurs atteint le statut de licorne en 2026 "
         "(20,1 M€ à 1,4 Md€ de valorisation).",
         "Moyen", "Insurance Business, The AI Insider, EU-Startups (2026)"),
    ]
    nouvelles_anomalies = [{
        "Anomalie_ID": f"ANO-{prochain_id + i:04d}", "Type d'anomalie": t, "Onglet source": onglet,
        "Table ou bloc source": "Recherche CVC externe utilisateur (21/09/2026)", "Ligne source": None,
        "Entité concernée": entite, "Champ concerné": champ, "Valeur source": vs, "Valeur retenue": vr,
        "Candidats éventuels": None, "Score de similarité": None, "Niveau de confiance": conf,
        "Traitement appliqué": trait, "Commentaire": src,
    } for i, (t, onglet, entite, champ, vs, vr, trait, conf, src) in enumerate(ANOMALIES_NOUVELLES)]
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
    print(f"Lignes CSV ignorées — doublons : {ignorees_doublon} | exclusions : {ignorees_exclusion} | "
          f"fonds non mappé : {ignorees_non_mappees}")
    print(f"Anomalies avant : {len(ano)}  →  après : {len(ano_out)}  (+{len(nouvelles_anomalies)})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
