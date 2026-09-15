#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
restructuration_vc.py
=====================
Transforme le classeur source "Fond_VC_Data.xlsx" (base VC en format clé/valeur)
en une base standardisée "VC_Database_Standardisee.xlsx" contenant 5 onglets :
Fonds, Participations, Tours_de_table, Dictionnaire, Anomalies.

Principes :
- aucune dépendance à la position physique des colonnes (reconnaissance sémantique) ;
- aucune estimation de valeur manquante, aucun rapprochement incertain forcé ;
- le fichier source n'est jamais modifié ;
- toute difficulté est journalisée dans l'onglet Anomalies et le traitement continue.
"""

from __future__ import annotations

import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# Similarité : rapidfuzz si disponible, sinon difflib (toujours présent).
try:  # pragma: no cover
    from rapidfuzz.fuzz import ratio as _rf_ratio

    def similarity(a: str, b: str) -> float:
        return _rf_ratio(a, b) / 100.0

    SIMILARITY_ENGINE = "rapidfuzz"
except Exception:  # pragma: no cover
    from difflib import SequenceMatcher

    def similarity(a: str, b: str) -> float:
        return SequenceMatcher(None, a, b).ratio()

    SIMILARITY_ENGINE = "difflib"


SOURCE_NAME = "Fond_VC_Data.xlsx"
OUTPUT_NAME = "VC_Database_Standardisee.xlsx"

# Seuils de rapprochement flou
SEUIL_ACCEPTATION = 0.85   # score minimal pour un rapprochement automatique
SEUIL_MARGE = 0.04         # écart minimal avec le 2e meilleur candidat
SEUIL_CANDIDAT = 0.62      # score minimal pour être listé comme candidat


# =============================================================================
# 1. NORMALISATION DES TEXTES ET SLUGS
# =============================================================================

def strip_accents(text: str) -> str:
    """Neutralise les accents (comparaison uniquement)."""
    return "".join(c for c in unicodedata.normalize("NFD", text)
                   if unicodedata.category(c) != "Mn")


def clean_label(value) -> str:
    """Nettoyage minimal conservant la casse : espaces, insécables, doublons d'espaces."""
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").replace("\u202f", " ")
    return re.sub(r"\s+", " ", text).strip()


def strip_taxonomy_prefix(label: str) -> str:
    """Retire un préfixe de classement de type 'a - ', 'B- ', '1 - ' (taxonomie, pas une colonne)."""
    return re.sub(r"^\s*[A-Za-z0-9]{1,3}\s*[-.)]\s*", "", clean_label(label)).strip()


def normalize_text(value) -> str:
    """Forme normalisée pour comparaison : sans accent, minuscule, sans ponctuation ni tiret."""
    text = clean_label(value)
    if not text:
        return ""
    text = strip_accents(text).lower()
    text = text.replace("&", " et ")
    text = re.sub(r"[^a-z0-9]+", " ", text)      # ponctuation et tirets neutralisés
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_criterion(value) -> str:
    """Normalisation d'un libellé de critère : préfixe de taxonomie retiré puis normalisation."""
    return normalize_text(strip_taxonomy_prefix(value))


def singularize(text: str) -> str:
    """Harmonise grossièrement singulier/pluriel pour la comparaison."""
    return " ".join(w[:-1] if len(w) > 3 and w.endswith("s") else w for w in text.split())


def comparison_key(value) -> str:
    """Clé de comparaison d'entité (fonds, véhicule, start-up)."""
    return singularize(normalize_text(value))


GENERIC_TERMS = {"fonds", "fond", "fund", "funds", "venture", "ventures",
                 "capital", "partners", "vc", "la", "le", "les"}


def comparison_key_light(value) -> str:
    """Variante de comparaison sans termes génériques (utilisée uniquement pour comparer)."""
    words = [w for w in comparison_key(value).split() if w not in GENERIC_TERMS]
    return " ".join(words) if words else comparison_key(value)


def slugify(value) -> str:
    """Slug stable : minuscules, sans accent, espaces internes -> tirets."""
    text = clean_label(value)
    if not text:
        return ""
    text = strip_accents(text).lower()
    text = text.replace("&", " et ")
    text = re.sub(r"[^a-z0-9]+", " ", text).strip()
    return re.sub(r"\s+", "-", text)


# =============================================================================
# 2. JOURNAL DES ANOMALIES
# =============================================================================

def _txt(value):
    """Représentation texte d'une valeur source, en préservant les vrais vides."""
    if value is None:
        return None
    try:
        if isinstance(value, float) and pd.isna(value):
            return None
    except Exception:
        pass
    text = clean_label(value)
    return text or None


class AnomalyLog:
    """Collecteur d'anomalies, numérotation stable ANO-0001, ANO-0002, ..."""

    COLUMNS = ["Anomalie_ID", "Type d'anomalie", "Onglet source", "Table ou bloc source",
               "Ligne source", "Entité concernée", "Champ concerné", "Valeur source",
               "Valeur retenue", "Candidats éventuels", "Score de similarité",
               "Niveau de confiance", "Traitement appliqué", "Commentaire"]

    def __init__(self):
        self.rows: list[dict] = []

    def add(self, type_anomalie, onglet=None, bloc=None, ligne=None, entite=None,
            champ=None, valeur_source=None, valeur_retenue=None, candidats=None,
            score=None, confiance=None, traitement=None, commentaire=None):
        self.rows.append({
            "Anomalie_ID": None,
            "Type d'anomalie": type_anomalie,
            "Onglet source": onglet,
            "Table ou bloc source": bloc,
            "Ligne source": ligne,
            "Entité concernée": entite,
            "Champ concerné": champ,
            "Valeur source": _txt(valeur_source),
            "Valeur retenue": _txt(valeur_retenue),
            "Candidats éventuels": candidats,
            "Score de similarité": round(score, 3) if isinstance(score, float) else score,
            "Niveau de confiance": confiance,
            "Traitement appliqué": traitement,
            "Commentaire": commentaire,
        })

    def to_frame(self) -> pd.DataFrame:
        for i, row in enumerate(self.rows, start=1):
            row["Anomalie_ID"] = f"ANO-{i:04d}"
        return pd.DataFrame(self.rows, columns=self.COLUMNS)


ANO = AnomalyLog()


# =============================================================================
# 3. CONVERSIONS DE VALEURS
# =============================================================================

TEXTES_VIDES = {"", "-", "--", "n/a", "na", "nd", "n d", "non disponible",
                "inconnu", "non renseigne", "?", "x x"}
VRAI_TOKENS = {"x", "oui", "yes", "true", "vrai", "1", "o"}


def is_blank(value) -> bool:
    """Cellule réellement vide (None, chaîne d'espaces)."""
    return value is None or (isinstance(value, str) and clean_label(value) == "")


def parse_bool(value) -> bool:
    """Case à cocher : VRAI si marqueur reconnu, FAUX si vide (champ booléen explicite)."""
    if is_blank(value):
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value) == 1.0
    return normalize_text(value) in VRAI_TOKENS


def parse_number(value):
    """Retourne (nombre|None, motif_rejet|None). Ne convertit jamais un texte ambigu."""
    if is_blank(value):
        return None, None
    if isinstance(value, bool):
        return None, "booléen dans un champ numérique"
    if isinstance(value, (int, float)):
        return float(value), None
    text = clean_label(value)
    if normalize_text(text) in TEXTES_VIDES:
        return None, "marqueur d'absence de donnée"
    # nombre isolé éventuellement suffixé (espace fine, virgule décimale)
    candidate = text.replace(" ", "").replace("\u202f", "")
    m = re.fullmatch(r"[+-]?\d+(?:[.,]\d+)?", candidate)
    if m:
        return float(candidate.replace(",", ".")), None
    return None, "valeur textuelle non convertible"


def parse_irr(value):
    """IRR -> décimal compatible format pourcentage Excel. (valeur|None, motif|None)."""
    if is_blank(value):
        return None, None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        v = float(value)
        if 0 <= v <= 1.5:
            return v, None                      # déjà décimal (0,24 -> 24 %)
        return None, "valeur numérique hors plage décimale, unité non confirmée"
    text = clean_label(value)
    if re.search(r"[;\[\]]|\bà\b|\bentre\b", text):
        return None, "intervalle : interprétation ambiguë"
    m = re.fullmatch(r"([+-]?\d+(?:[.,]\d+)?)\s*%", text)
    if m:
        return float(m.group(1).replace(",", ".")) / 100.0, None   # règle explicite : % -> décimal
    return None, "valeur textuelle non convertible"


def parse_year(value):
    """Année entière plausible ; sinon (None, motif). N'invente ni jour ni mois."""
    if is_blank(value):
        return None, None
    if isinstance(value, datetime):
        return value.year, None
    num, reason = parse_number(value)
    if num is None:
        return None, reason
    if 1900 <= num <= 2100 and float(num).is_integer():
        return int(num), None
    return None, "année hors plage plausible"


# =============================================================================
# 4. RECONNAISSANCE SÉMANTIQUE DES CHAMPS
# =============================================================================

# Familles de synonymes : champ cible -> libellés reconnus (forme normalisée)
SYNONYMES = {
    "Société de gestion": ["nom", "societe de gestion", "gestionnaire", "fund manager",
                           "management company", "general partner", "gp", "vc", "fond general"],
    "Véhicule": ["fonds", "fond", "vehicule", "fund", "fund name", "vehicle",
                 "investment vehicle", "nom du vehicule"],
    "AuM (M€)": ["aum", "assets under management", "actifs sous gestion",
                 "encours sous gestion", "total assets under management"],
    "Montant levé (M€)": ["montant leve", "taille du fonds", "fund size", "capital raised",
                          "commitments", "total commitments", "closing size", "final close",
                          "amount raised"],
    "Montant alloué (M€)": ["montant alloue", "capital allocated", "allocated amount",
                            "investment capacity", "capital deploye"],
    "Millésime": ["millesime", "vintage", "vintage year", "annee de lancement",
                  "annee de creation", "debut des levees", "fund launch year"],
    "Ticket min (M€)": ["ticket minimum", "ticket min", "minimum", "minimum investment",
                        "min ticket", "initial ticket", "first cheque", "minimum cheque"],
    "Ticket max (M€)": ["ticket maximum", "ticket max", "maximum", "maximum investment",
                        "max ticket", "follow on capacity", "maximum cheque"],
    "TVPI": ["tvpi", "total value to paid in"],
    "DPI": ["dpi", "distributions to paid in"],
    "RVPI": ["rvpi", "residual value to paid in"],
    "IRR (TRI)": ["irr", "tri", "internal rate of return", "taux de rendement interne"],
    "MoC (MoM)": ["moc", "mom", "multiple on cost", "multiple of money", "multiple"],
    "Nb participations (déclaré)": ["nombre de participation", "nb participations",
                                    "portfolio companies", "nombre de participations"],
    "Nb exits": ["nb exit", "nombre d exit", "exits", "nb sorties"],
    "Nb InsurTech": ["nb d insurtech", "nombre d insurtech", "nb insurtech"],
    # Participations
    "Start-up": ["nom", "start up", "startup", "societe", "company"],
    "Secteur": ["secteur", "domaine", "industry", "vertical", "sector", "thematique"],
    "Date de création": ["date de creation", "creation", "annee de creation", "founded"],
    "Date d’investissement": ["date d investissment du fond", "date d investissement du fond",
                              "date d investissement", "investment date"],
    "Capital social (k€)": ["capital social en k", "capital social"],
    "CA (M€)": ["chiffre d affaire en m", "chiffre d affaires", "ca", "revenue"],
    "Valorisation (M€)": ["valorisation en m", "valorisation", "valuation"],
    "Total levé par la start-up (M€)": ["general", "total leve", "total investit general",
                                        "total raised", "montant total leve"],
    "Montant investi par le fonds (M€)": ["fond", "montant investi par le fonds",
                                          "total investit fond", "montant investi"],
    "Nb fonds investisseurs": ["nombre de fond qui ont investit", "nombre de fonds investisseurs",
                               "nb fonds investisseurs"],
    "Nb tours (fonds)": ["nombre de tour", "nb tours fonds"],
    "Nb tours (total)": ["nombre de tour total", "nb tours total"],
    "Nb pays d’implantation": ["nombre de pays", "nb pays", "nombre de pays d implantation"],
    "MoC exit": ["moc"],
    "MoEP exit": ["moep"],
}


def resolve_field(label: str, candidates: list[str]) -> tuple[str | None, str, str]:
    """
    Rapproche un libellé source d'un champ cible parmi `candidates`.
    Retourne (champ|None, methode, confiance).
    """
    key = normalize_criterion(label)
    if not key:
        return None, "Non trouvé", ""
    for field in candidates:                                   # correspondance exacte
        if key == normalize_text(field):
            return field, "Correspondance exacte", "Élevé"
    for field in candidates:                                   # synonyme déclaré
        if key in SYNONYMES.get(field, []):
            return field, "Synonyme", "Élevé"
    for field in candidates:                                   # libellé normalisé (singulier/pluriel)
        if singularize(key) in [singularize(s) for s in SYNONYMES.get(field, [])]:
            return field, "Libellé normalisé", "Élevé"
    best, best_score = None, 0.0
    for field in candidates:                                   # similarité textuelle
        for s in [normalize_text(field)] + SYNONYMES.get(field, []):
            sc = similarity(key, s)
            if sc > best_score:
                best, best_score = field, sc
    if best_score >= 0.92:
        return best, "Libellé normalisé", "Élevé"
    if best_score >= 0.80:
        return best, "Déduction par contexte", "Moyen"
    return None, "Non trouvé", ""


# Registre des correspondances pour le Dictionnaire
FIELD_SOURCES: dict[tuple[str, str], dict] = {}


def register_source(tab, field, onglet_source=None, entete=None, section=None,
                    categorie=None, critere=None, methode=None, confiance=None,
                    type_=None, unite=None, calcule=None, regle=None, commentaire=None):
    """Mémorise la traçabilité d'un champ cible (libellés source originaux conservés)."""
    entry = FIELD_SOURCES.setdefault((tab, field), {
        "Onglet source": onglet_source, "En-tête source": entete,
        "Section source": [], "Catégorie source": [], "Critère source exact": [],
        "Méthode de rapprochement": methode, "Niveau de confiance": confiance,
        "Type": type_, "Unité": unite, "Champ source ou calculé": calcule,
        "Formule ou règle": regle, "Commentaire": commentaire})
    for key, value in (("Section source", section), ("Catégorie source", categorie),
                       ("Critère source exact", critere)):
        if value and value not in entry[key]:
            entry[key].append(value)
    for key, value in (("Onglet source", onglet_source), ("En-tête source", entete),
                       ("Méthode de rapprochement", methode), ("Niveau de confiance", confiance),
                       ("Type", type_), ("Unité", unite), ("Champ source ou calculé", calcule),
                       ("Formule ou règle", regle), ("Commentaire", commentaire)):
        if value and not entry.get(key):
            entry[key] = value
    return entry


# =============================================================================
# 5. AUDIT DU CLASSEUR ET DÉTECTION DES TABLES SOURCES
# =============================================================================

PIVOT_MARKERS = {"row labels", "column labels", "grand total", "etiquettes de lignes",
                 "etiquettes de colonnes", "total general"}
AGG_MARKERS = ("average of", "sum of", "count of", "moyenne de", "somme de", "nombre de valeurs")


def sheet_profile(ws) -> dict:
    """Profil d'un onglet : dimensions, colonnes non vides, blocs, indices de restitution."""
    max_row, max_col = ws.max_row, ws.max_column
    non_empty_cols, filled = [], 0
    for c in range(1, max_col + 1):
        count = sum(1 for r in range(1, max_row + 1) if not is_blank(ws.cell(r, c).value))
        if count:
            non_empty_cols.append((c, count))
            filled += count
    texts = []
    for r in range(1, min(max_row, 12) + 1):
        for c in range(1, min(max_col, 40) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                texts.append(normalize_text(v))
    is_pivot = any(t in PIVOT_MARKERS for t in texts) or any(
        t.startswith(AGG_MARKERS) for t in texts)
    # blocs de colonnes contiguës séparés par au moins une colonne vide
    blocks, current = [], []
    cols = [c for c, _ in non_empty_cols]
    for c in cols:
        if current and c > current[-1] + 1:
            blocks.append((current[0], current[-1]))
            current = []
        current.append(c)
    if current:
        blocks.append((current[0], current[-1]))
    return {"nom": ws.title, "dimensions": ws.dimensions, "lignes": max_row, "colonnes": max_col,
            "cellules_remplies": filled, "colonnes_non_vides": len(non_empty_cols),
            "blocs": blocks, "fusions": len(ws.merged_cells.ranges),
            "tcd_ou_synthese": is_pivot,
            "navigation": filled <= 3 and max_row <= 20}


def read_block(ws, col_start: int, col_end: int, header_row: int = 1) -> pd.DataFrame:
    """Lit un bloc de colonnes en utilisant la ligne d'en-tête détectée (positions non figées)."""
    headers = [clean_label(ws.cell(header_row, c).value) for c in range(col_start, col_end + 1)]
    records = []
    for r in range(header_row + 1, ws.max_row + 1):
        values = [ws.cell(r, c).value for c in range(col_start, col_end + 1)]
        if all(is_blank(v) for v in values):
            continue
        rec = {h if h else f"col_{i}": v for i, (h, v) in enumerate(zip(headers, values))}
        rec["_ligne_source"] = r
        records.append(rec)
    return pd.DataFrame(records)


def find_header_row(ws, max_scan: int = 15) -> int:
    """Ligne d'en-tête = première ligne comportant au moins 3 libellés textuels distincts."""
    for r in range(1, min(ws.max_row, max_scan) + 1):
        labels = {clean_label(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)
                  if isinstance(ws.cell(r, c).value, str) and clean_label(ws.cell(r, c).value)}
        if len(labels) >= 3:
            return r
    return 1


def detect_source_tables(wb) -> dict:
    """
    Identifie les tables de données primaires (entités + valeurs élémentaires)
    et écarte menus, interfaces, TCD et feuilles de restitution.
    """
    tables, profiles = {}, []
    for ws in wb.worksheets:
        prof = sheet_profile(ws)
        profiles.append(prof)
        if prof["navigation"]:
            continue
        if prof["tcd_ou_synthese"]:
            ANO.add("Donnée provenant d'un tableau de synthèse ignorée", onglet=ws.title,
                    bloc="Feuille entière", traitement="Onglet ignoré", confiance="Élevé",
                    commentaire="Tableau croisé dynamique / restitution : non utilisé comme source primaire")
            continue
        header_row = find_header_row(ws)
        for (c0, c1) in prof["blocs"]:
            headers = [normalize_text(ws.cell(header_row, c).value) for c in range(c0, c1 + 1)]
            headers = [h for h in headers if h]
            if len(headers) < 3:
                ANO.add("Champ source ignoré", onglet=ws.title, bloc=f"colonnes {c0}-{c1}",
                        traitement="Bloc ignoré", confiance="Élevé",
                        commentaire="Bloc résiduel sans en-tête exploitable")
                continue
            df = read_block(ws, c0, c1, header_row)
            if df.empty:
                continue
            tables[f"{ws.title}|{c0}-{c1}"] = {
                "onglet": ws.title, "col_start": c0, "col_end": c1,
                "header_row": header_row, "headers": headers, "data": df}
    return {"tables": tables, "profiles": profiles}


def classify_table(info: dict) -> str:
    """Nature d'une table primaire d'après ses en-têtes (jamais d'après sa position)."""
    h = set(info["headers"])
    has_grid = {"sections", "section"} & h or {"categories", "categorie"} & h
    has_crit = {"criteres", "critere"} & h
    has_data = {"data", "valeur", "value"} & h
    if has_crit and has_data and has_grid:
        # grille clé/valeur : fonds si pas de colonne 'domaine'/'fond general'
        if {"fond general", "domaine"} & h:
            return "participations"
        return "fonds"
    if has_crit and has_data and not has_grid:
        return "tours"
    return "inconnu"


# =============================================================================
# 6. TRANSFORMATION DE LA BASE DES FONDS
# =============================================================================

STAGES_ORDRE = ["Pré-Seed", "Seed", "Pré-Série A", "Série A", "Série B", "Série C", "Série D"]

STAGE_ALIASES = {
    "pre seed": "Pré-Seed", "preseed": "Pré-Seed",
    "seed": "Seed", "seed round": "Seed", "amorcage": "Seed",
    "pre serie a": "Pré-Série A", "pre series a": "Pré-Série A", "pre a": "Pré-Série A",
    "serie a": "Série A", "series a": "Série A", "a": "Série A",
    "serie b": "Série B", "series b": "Série B", "b": "Série B",
    "serie c": "Série C", "series c": "Série C", "c": "Série C",
    "serie d": "Série D", "series d": "Série D", "d": "Série D",
}


def harmonize_stage(label):
    """Harmonise toute variante typographique de stage vers le vocabulaire cible."""
    key = normalize_criterion(label)
    if not key:
        return None
    if key in STAGE_ALIASES:
        return STAGE_ALIASES[key]
    best, score = None, 0.0
    for alias, target in STAGE_ALIASES.items():
        sc = similarity(key, alias)
        if sc > score:
            best, score = target, sc
    return best if score >= 0.88 else "Autre"


COLONNES_FONDS = ["Nom du fonds", "Véhicule", "AuM (M€)", "DPI", "RVPI", "TVPI", "MoC (MoM)",
                  "IRR (TRI)", "Quartile", "Montant levé (M€)", "Montant alloué (M€)",
                  "Ticket min (M€)", "Ticket max (M€)", "Statut", "Phase", "Millésime",
                  "Géographie", "Stratégie", "Stages pratiqués"] + STAGES_ORDRE + [
                  "Nb participations (déclaré)", "Nb participations documentées", "Nb exits",
                  "Nb InsurTech", "Fonds_ID", "Source_AuM", "Date_MAJ"]

# Champs numériques du fonds : (champ cible, type de parsing, unité)
NUM_FONDS = {
    "AuM (M€)": ("num", "M€"), "Montant levé (M€)": ("num", "M€"),
    "Montant alloué (M€)": ("num", "M€"), "Ticket min (M€)": ("num", "M€"),
    "Ticket max (M€)": ("num", "M€"), "TVPI": ("mult", "multiple"), "DPI": ("mult", "multiple"),
    "MoC (MoM)": ("mult", "multiple"), "IRR (TRI)": ("irr", "%"),
    "Nb participations (déclaré)": ("int", "unité"), "Nb exits": ("int", "unité"),
    "Nb InsurTech": ("int", "unité"), "Millésime": ("year", "année"),
}

CHAMPS_ATTENDUS_FONDS = list(NUM_FONDS) + ["Quartile", "Statut", "Phase", "Géographie",
                                           "Stratégie", "Stages pratiqués", "RVPI"]


def transform_fonds(table: dict) -> pd.DataFrame:
    """Passe la grille clé/valeur des véhicules en une ligne par véhicule."""
    df = table["data"]
    onglet = table["onglet"]
    cols = {normalize_text(c): c for c in df.columns}

    col_soc = cols.get("nom") or cols.get("societe de gestion")
    col_veh = cols.get("fonds") or cols.get("fond") or cols.get("vehicule")
    col_sec = cols.get("sections") or cols.get("section")
    col_cat = cols.get("categories") or cols.get("categorie")
    col_cri = cols.get("criteres") or cols.get("critere")
    col_val = cols.get("data") or cols.get("valeur")

    register_source("Fonds", "Nom du fonds", onglet, col_soc, methode="Synonyme",
                    confiance="Élevé", type_="Texte", calcule="Champ source",
                    regle="Libellé source conservé", commentaire="Société de gestion")
    register_source("Fonds", "Véhicule", onglet, col_veh, methode="Synonyme", confiance="Élevé",
                    type_="Texte", calcule="Champ source", regle="Libellé source conservé")

    vehicules, ordre = {}, []
    for _, row in df.iterrows():
        soc, veh = clean_label(row[col_soc]), clean_label(row[col_veh])
        if not soc and not veh:
            continue
        key = (soc, veh)
        if key not in vehicules:
            vehicules[key] = {"multi": {}, "bool": {}, "val": {}, "lignes": []}
            ordre.append(key)
        entry = vehicules[key]
        entry["lignes"].append(row["_ligne_source"])
        section, categorie, critere = (clean_label(row[col_sec]), clean_label(row[col_cat]),
                                       clean_label(row[col_cri]))
        value = row[col_val]
        cat_key = normalize_criterion(categorie)
        cri_key = normalize_criterion(critere)
        ctx = dict(onglet_source=onglet, entete=col_val, section=section,
                   categorie=categorie, critere=critere)

        # --- champs à choix multiples / booléens, identifiés par la catégorie ---
        if cat_key == "statut":
            entry["multi"].setdefault("Statut", []).append((critere, value, row["_ligne_source"]))
            register_source("Fonds", "Statut", methode="Déduction par contexte",
                            confiance="Élevé", type_="Texte", calcule="Champ source",
                            regle="Case cochée de la catégorie Statut",
                            commentaire="Champ à valeur unique", **ctx)
            continue
        if cat_key == "phase":
            entry["multi"].setdefault("Phase", []).append((critere, value, row["_ligne_source"]))
            register_source("Fonds", "Phase", methode="Déduction par contexte", confiance="Élevé",
                            type_="Texte", calcule="Champ source",
                            regle="Case cochée de la catégorie Phase",
                            commentaire="Champ à valeur unique", **ctx)
            continue
        if cat_key == "quartile":
            entry["multi"].setdefault("Quartile", []).append((critere, value, row["_ligne_source"]))
            register_source("Fonds", "Quartile", methode="Déduction par contexte",
                            confiance="Élevé", type_="Texte", calcule="Champ source",
                            regle="Case cochée de la catégorie Quartile",
                            commentaire="Champ à valeur unique", **ctx)
            continue
        if cat_key == "geographie":
            entry["multi"].setdefault("Géographie", []).append((critere, value, row["_ligne_source"]))
            register_source("Fonds", "Géographie", methode="Déduction par contexte",
                            confiance="Élevé", type_="Texte", calcule="Champ source",
                            regle="Concaténation « ; » des cases cochées", **ctx)
            continue
        if cat_key == "strategie":
            entry["multi"].setdefault("Stratégie", []).append((critere, value, row["_ligne_source"]))
            register_source("Fonds", "Stratégie", methode="Déduction par contexte",
                            confiance="Élevé", type_="Texte", calcule="Champ source",
                            regle="Concaténation « ; » des cases cochées", **ctx)
            continue
        if normalize_criterion(section).startswith("round d investissement") or \
           cat_key in ("amorcage", "series institutionnelles"):
            stage = harmonize_stage(critere)
            if stage and stage in STAGES_ORDRE:
                entry["bool"].setdefault(stage, []).append(value)
                register_source("Fonds", stage, methode="Libellé normalisé", confiance="Élevé",
                                type_="Booléen", calcule="Champ source",
                                regle="Case cochée du round d'investissement harmonisée", **ctx)
            else:
                ANO.add("Libellé ambigu", onglet=onglet, bloc="Grille véhicules",
                        ligne=row["_ligne_source"], entite=f"{soc} / {veh}", champ="Stage",
                        valeur_source=critere, traitement="Critère ignoré", confiance="Faible",
                        commentaire="Stage non rattachable au vocabulaire cible")
            continue

        # --- champs scalaires ---
        field, methode, confiance = resolve_field(critere, list(NUM_FONDS))
        if field is None:
            ANO.add("Champ source ignoré", onglet=onglet, bloc="Grille véhicules",
                    ligne=row["_ligne_source"], entite=f"{soc} / {veh}",
                    champ=critere, valeur_source=value, traitement="Critère non repris",
                    confiance="Faible", commentaire=f"Catégorie source : {categorie}")
            continue
        if confiance != "Élevé":
            ANO.add("Libellé ambigu", onglet=onglet, bloc="Grille véhicules",
                    ligne=row["_ligne_source"], entite=f"{soc} / {veh}", champ=field,
                    valeur_source=critere, candidats=field, confiance=confiance,
                    traitement="Correspondance non appliquée",
                    commentaire="Confiance insuffisante pour un rapprochement automatique")
            continue
        kind, unite = NUM_FONDS[field]
        register_source("Fonds", field, methode=methode, confiance=confiance,
                        type_="Nombre", unite=unite, calcule="Champ source",
                        regle="Valeur source reprise sans conversion", **ctx)
        if field in entry["val"] and not is_blank(value) and not is_blank(entry["val"][field][0]):
            ANO.add("Doublon", onglet=onglet, bloc="Grille véhicules", ligne=row["_ligne_source"],
                    entite=f"{soc} / {veh}", champ=field, valeur_source=value,
                    valeur_retenue=entry["val"][field][0], confiance="Moyen",
                    traitement="Première occurrence conservée",
                    commentaire="Critère présent plusieurs fois pour le même véhicule")
            continue
        entry["val"][field] = (value, row["_ligne_source"], kind, critere)

    # --- construction des lignes ---
    rows, ids = [], {}
    for (soc, veh) in ordre:
        entry = vehicules[(soc, veh)]
        ligne0 = entry["lignes"][0]
        rec = {c: None for c in COLONNES_FONDS}
        rec["Nom du fonds"] = soc or None
        rec["Véhicule"] = veh or None

        # valeurs numériques
        for field, (value, ligne, kind, critere) in entry["val"].items():
            if kind == "irr":
                num, reason = parse_irr(value)
            elif kind == "year":
                num, reason = parse_year(value)
            else:
                num, reason = parse_number(value)
            if num is None and not is_blank(value):
                ANO.add("Valeur invalide" if kind != "irr" else "Conversion non effectuée",
                        onglet=onglet, bloc="Grille véhicules", ligne=ligne,
                        entite=f"{soc} / {veh}", champ=field, valeur_source=value,
                        valeur_retenue="(vide)", confiance="Élevé",
                        traitement="Cellule laissée vide", commentaire=reason)
                continue
            if num is not None and kind == "int":
                num = int(round(num))
            if num == 0 and kind == "num":
                ANO.add("Valeur invalide", onglet=onglet, bloc="Grille véhicules", ligne=ligne,
                        entite=f"{soc} / {veh}", champ=field, valeur_source=0,
                        valeur_retenue=0, confiance="Moyen",
                        traitement="Zéro source conservé tel quel",
                        commentaire="Zéro explicite dans la source : aucune substitution effectuée")
            rec[field] = num

        # champs à valeur unique
        for field in ("Statut", "Phase", "Quartile"):
            options = entry["multi"].get(field, [])
            coches = [(clean_label(strip_taxonomy_prefix(lbl)), lg) for lbl, val, lg in options
                      if parse_bool(val)]
            if not coches:
                rec[field] = None
                ANO.add("Aucune case sélectionnée", onglet=onglet, bloc="Grille véhicules",
                        ligne=ligne0, entite=f"{soc} / {veh}", champ=field,
                        valeur_retenue="(vide)", confiance="Élevé",
                        traitement="Cellule laissée vide")
            else:
                label = coches[0][0]
                rec[field] = label.rstrip(".") if field == "Quartile" else label
                if len(coches) > 1:
                    ANO.add("Plusieurs cases sélectionnées", onglet=onglet, bloc="Grille véhicules",
                            ligne=ligne0, entite=f"{soc} / {veh}", champ=field,
                            valeur_source=" ; ".join(c[0] for c in coches),
                            valeur_retenue=rec[field], confiance="Moyen",
                            traitement="Première valeur selon l'ordre source conservée")

        # champs multi-valeurs
        for field in ("Géographie", "Stratégie"):
            coches = [clean_label(strip_taxonomy_prefix(lbl))
                      for lbl, val, lg in entry["multi"].get(field, []) if parse_bool(val)]
            rec[field] = " ; ".join(coches) if coches else None
            if not coches:
                ANO.add("Aucune case sélectionnée", onglet=onglet, bloc="Grille véhicules",
                        ligne=ligne0, entite=f"{soc} / {veh}", champ=field,
                        valeur_retenue="(vide)", confiance="Élevé",
                        traitement="Cellule laissée vide")

        # stages booléens + concaténation ordonnée
        stages = []
        for stage in STAGES_ORDRE:
            flag = any(parse_bool(v) for v in entry["bool"].get(stage, []))
            rec[stage] = flag
            if flag:
                stages.append(stage)
        rec["Stages pratiqués"] = " ; ".join(stages) if stages else None

        # RVPI : uniquement si TVPI et DPI numériques
        if isinstance(rec["TVPI"], float) and isinstance(rec["DPI"], float):
            rec["RVPI"] = round(rec["TVPI"] - rec["DPI"], 6)
        else:
            rec["RVPI"] = None

        # Fonds_ID
        if soc and veh:
            base = f"{slugify(soc)}_{slugify(veh)}"
            if base in ids:
                autre = ids[base]
                if autre != (soc, veh):
                    suffix = 2
                    while f"{base}--{suffix}" in ids:
                        suffix += 1
                    new_id = f"{base}--{suffix}"
                    ANO.add("Conflit de Fonds_ID", onglet=onglet, bloc="Grille véhicules",
                            ligne=ligne0, entite=f"{soc} / {veh}", champ="Fonds_ID",
                            valeur_source=base, valeur_retenue=new_id, confiance="Élevé",
                            traitement="Suffixe technique ajouté, aucune ligne écrasée",
                            commentaire=f"Slug déjà utilisé par {autre[0]} / {autre[1]}")
                    base = new_id
            ids[base] = (soc, veh)
            rec["Fonds_ID"] = base
        else:
            rec["Fonds_ID"] = None
            ANO.add("Véhicule non identifié", onglet=onglet, bloc="Grille véhicules", ligne=ligne0,
                    entite=f"{soc} / {veh}", champ="Fonds_ID", valeur_retenue="(vide)",
                    confiance="Élevé", traitement="Aucune clé partielle créée",
                    commentaire="Société de gestion ou véhicule absent")

        rec["Source_AuM"] = None
        rec["Date_MAJ"] = None
        rec["_lignes_source"] = f"{min(entry['lignes'])}-{max(entry['lignes'])}"
        rows.append(rec)

        # contrôle de régularité de la grille
        if len(entry["lignes"]) != 40:
            pass  # signalé globalement plus bas

    fonds = pd.DataFrame(rows)

    # régularité de la grille répétitive
    tailles = {(soc, veh): len(vehicules[(soc, veh)]["lignes"]) for (soc, veh) in ordre}
    mode = pd.Series(list(tailles.values())).mode()[0]
    for (soc, veh), n in tailles.items():
        if n != mode:
            ANO.add("Structure source irrégulière", onglet=onglet, bloc="Grille véhicules",
                    entite=f"{soc} / {veh}", champ="Grille de critères", valeur_source=n,
                    valeur_retenue=mode, confiance="Élevé",
                    traitement="Lecture par critère (aucune hypothèse de position)",
                    commentaire=f"{n} lignes de critères au lieu de {mode}")

    register_source("Fonds", "Stages pratiqués", onglet, calcule="Calculé", type_="Texte",
                    methode="Calculé", confiance="Élevé",
                    regle="Concaténation ordonnée « ; » des stages booléens VRAI")

    # champs cibles attendus absents de la source
    for field in CHAMPS_ATTENDUS_FONDS:
        if field in ("RVPI", "Stages pratiqués"):
            continue
        if ("Fonds", field) not in FIELD_SOURCES:
            ANO.add("Champ attendu non trouvé", onglet=onglet, bloc="Grille véhicules",
                    champ=field, confiance="Élevé", traitement="Colonne créée vide",
                    commentaire="Aucun critère source rapprochable")

    register_source("Fonds", "RVPI", onglet, calcule="Calculé", type_="Nombre", unite="multiple",
                    methode="Calculé", confiance="Élevé",
                    regle="RVPI = TVPI - DPI lorsque les deux valeurs sont présentes")
    register_source("Fonds", "Stages pratiqués", onglet, calcule="Calculé", type_="Texte",
                    methode="Calculé", confiance="Élevé",
                    regle="Concaténation ordonnée « ; » des stages booléens VRAI")
    register_source("Fonds", "Fonds_ID", onglet, calcule="Calculé", type_="Texte",
                    methode="Calculé", confiance="Élevé",
                    regle="slug(Société de gestion)_slug(Véhicule)")
    register_source("Fonds", "Source_AuM", onglet, calcule="Champ source", type_="Texte",
                    methode="Non trouvé", confiance="Élevé", regle="Laissé vide",
                    commentaire="Non fourni par la source : réservé aux enrichissements futurs")
    register_source("Fonds", "Date_MAJ", onglet, calcule="Champ source", type_="Date",
                    methode="Non trouvé", confiance="Élevé", regle="Laissé vide",
                    commentaire="Non fourni par la source : réservé aux enrichissements futurs")
    return fonds


# =============================================================================
# 7. TRANSFORMATION DES PARTICIPATIONS
# =============================================================================

COLONNES_PARTICIPATIONS = [
    "Fonds_ID", "Nom du fonds", "Véhicule", "Start-up", "Secteur", "Stage financé",
    "Montant investi par le fonds (M€)", "Total levé par la start-up (M€)", "Date de création",
    "Date d’investissement", "Pays d’origine", "Nb pays d’implantation",
    "Positionnement (Leader/Minoritaire)", "Capital social (k€)", "CA (M€)", "Valorisation (M€)",
    "Nb fonds investisseurs", "Nb tours (fonds)", "Nb tours (total)", "Repositionnement",
    "Exit (O/N)", "MoC exit", "MoEP exit"]

NUM_PART = {
    "Capital social (k€)": ("num", "k€"), "CA (M€)": ("num", "M€"),
    "Valorisation (M€)": ("num", "M€"), "Nb fonds investisseurs": ("int", "unité"),
    "Nb tours (fonds)": ("int", "unité"), "Nb tours (total)": ("int", "unité"),
    "Nb pays d’implantation": ("int", "unité"), "Date de création": ("year", "année"),
    "Date d’investissement": ("year", "année"), "MoC exit": ("mult", "multiple"),
    "MoEP exit": ("mult", "multiple"),
}


def transform_participations(table: dict) -> pd.DataFrame:
    """Grille clé/valeur des start-ups -> une ligne par (véhicule, start-up)."""
    df = table["data"]
    onglet = table["onglet"]
    cols = {normalize_text(c): c for c in df.columns}
    col_su = cols.get("nom") or cols.get("start up")
    col_soc = cols.get("fond general") or cols.get("societe de gestion")
    col_veh = cols.get("fond") or cols.get("fonds") or cols.get("vehicule")
    col_dom = cols.get("domaine") or cols.get("secteur")
    col_sec = cols.get("section") or cols.get("sections")
    col_cat = cols.get("categorie") or cols.get("categories")
    col_cri = cols.get("critere") or cols.get("criteres")
    col_val = cols.get("data") or cols.get("valeur")

    for field, src in (("Start-up", col_su), ("Nom du fonds", col_soc),
                       ("Véhicule", col_veh), ("Secteur", col_dom)):
        register_source("Participations", field, onglet, src, methode="Synonyme",
                        confiance="Élevé", type_="Texte", calcule="Champ source",
                        regle="Libellé source conservé")

    entities, ordre = {}, []
    for _, row in df.iterrows():
        su, soc = clean_label(row[col_su]), clean_label(row[col_soc])
        veh, dom = clean_label(row[col_veh]), clean_label(row[col_dom])
        if not su:
            continue
        key = (su, soc, veh)
        if key not in entities:
            entities[key] = {"secteur": dom, "multi": {}, "val": {}, "lignes": []}
            ordre.append(key)
        entry = entities[key]
        entry["lignes"].append(row["_ligne_source"])
        section, categorie, critere = (clean_label(row[col_sec]), clean_label(row[col_cat]),
                                       clean_label(row[col_cri]))
        value = row[col_val]
        cat_key = normalize_criterion(categorie)
        ctx = dict(onglet_source=onglet, entete=col_val, section=section,
                   categorie=categorie, critere=critere)

        if cat_key.startswith("pays d origine"):
            entry["multi"].setdefault("Pays d’origine", []).append((critere, value))
            register_source("Participations", "Pays d’origine", methode="Déduction par contexte",
                            confiance="Élevé", type_="Texte", calcule="Champ source",
                            regle="Concaténation « ; » des zones cochées", **ctx)
            continue
        if cat_key.startswith("implantation"):
            if normalize_criterion(critere).startswith("nombre de pays"):
                entry["val"]["Nb pays d’implantation"] = (value, row["_ligne_source"],
                                                          "int", critere)
                register_source("Participations", "Nb pays d’implantation", methode="Synonyme",
                                confiance="Élevé", type_="Nombre", unite="unité",
                                calcule="Champ source",
                                regle="Valeur numérique explicite (non recalculée à partir des cases)",
                                **ctx)
            else:
                register_source("Participations", "Nb pays d’implantation",
                                commentaire="Cases de zones géographiques non utilisées pour "
                                            "le calcul (sens incertain)", **ctx)
            continue
        if cat_key.startswith("positionnement"):
            entry["multi"].setdefault("Positionnement (Leader/Minoritaire)", []).append(
                (critere, value))
            register_source("Participations", "Positionnement (Leader/Minoritaire)",
                            methode="Déduction par contexte", confiance="Élevé", type_="Texte",
                            calcule="Champ source", regle="Case cochée du positionnement",
                            **ctx)
            continue
        if cat_key.startswith("serie financee") or cat_key.startswith("stage"):
            entry["multi"].setdefault("Stage financé", []).append((critere, value))
            register_source("Participations", "Stage financé", methode="Libellé normalisé",
                            confiance="Élevé", type_="Texte", calcule="Champ source",
                            regle="Série cochée harmonisée vers le vocabulaire cible", **ctx)
            continue
        if cat_key.startswith("repositionnement"):
            entry["multi"].setdefault("Repositionnement", []).append((critere, value))
            register_source("Participations", "Repositionnement", methode="Correspondance exacte",
                            confiance="Élevé", type_="Texte", calcule="Champ source",
                            regle="Case Oui/Non", **ctx)
            continue
        if cat_key.startswith("exit"):
            cri_key = normalize_criterion(critere)
            if cri_key in ("oui", "non"):
                entry["multi"].setdefault("Exit (O/N)", []).append((critere, value))
                register_source("Participations", "Exit (O/N)", methode="Correspondance exacte",
                                confiance="Élevé", type_="Texte", calcule="Champ source",
                                regle="Case Oui/Non convertie en O/N", **ctx)
            else:
                field = "MoC exit" if cri_key == "moc" else ("MoEP exit" if cri_key == "moep" else None)
                if field:
                    entry["val"][field] = (value, row["_ligne_source"], "mult", critere)
                    register_source("Participations", field, methode="Correspondance exacte",
                                    confiance="Élevé", type_="Nombre", unite="multiple",
                                    calcule="Champ source", regle="Valeur source reprise", **ctx)
                else:
                    ANO.add("Champ source ignoré", onglet=onglet, bloc="Grille start-ups",
                            ligne=row["_ligne_source"], entite=su, champ=critere,
                            valeur_source=value, traitement="Critère non repris",
                            confiance="Faible")
            continue
        if cat_key.startswith("total investit") or cat_key.startswith("total investi"):
            cri_key = normalize_criterion(critere)
            if cri_key in ("general", "total"):
                field = "Total levé par la start-up (M€)"
            elif cri_key in ("fond", "fonds"):
                field = "Montant investi par le fonds (M€)"
            else:
                ANO.add("Libellé ambigu", onglet=onglet, bloc="Grille start-ups",
                        ligne=row["_ligne_source"], entite=su, champ="Montants investis",
                        valeur_source=critere, valeur_retenue="(vide)", confiance="Faible",
                        traitement="Valeur non affectée",
                        commentaire="Impossible de distinguer montant total levé et montant du fonds")
                continue
            entry["val"][field] = (value, row["_ligne_source"], "num", critere)
            register_source("Participations", field, methode="Déduction par contexte",
                            confiance="Élevé", type_="Nombre", unite="M€", calcule="Champ source",
                            regle="Distinction stricte entre total levé (Général) et montant du fonds (Fond)",
                            **ctx)
            continue

        field, methode, confiance = resolve_field(critere, list(NUM_PART))
        if field is None:
            ANO.add("Champ source ignoré", onglet=onglet, bloc="Grille start-ups",
                    ligne=row["_ligne_source"], entite=su, champ=critere, valeur_source=value,
                    traitement="Critère non repris", confiance="Faible",
                    commentaire=f"Catégorie source : {categorie}")
            continue
        if confiance != "Élevé":
            ANO.add("Libellé ambigu", onglet=onglet, bloc="Grille start-ups",
                    ligne=row["_ligne_source"], entite=su, champ=field, valeur_source=critere,
                    confiance=confiance, traitement="Correspondance non appliquée")
            continue
        kind, unite = NUM_PART[field]
        entry["val"][field] = (value, row["_ligne_source"], kind, critere)
        register_source("Participations", field, methode=methode, confiance=confiance,
                        type_="Nombre", unite=unite, calcule="Champ source",
                        regle="Valeur source reprise sans conversion", **ctx)

    rows = []
    for key in ordre:
        su, soc, veh = key
        entry = entities[key]
        ligne0 = entry["lignes"][0]
        rec = {c: None for c in COLONNES_PARTICIPATIONS}
        rec["Start-up"] = su
        rec["Nom du fonds"] = soc or None
        rec["Véhicule"] = veh or None
        rec["Secteur"] = entry["secteur"] or None

        for field, (value, ligne, kind, critere) in entry["val"].items():
            if kind == "year":
                num, reason = parse_year(value)
            else:
                num, reason = parse_number(value)
            if num is None and not is_blank(value):
                ANO.add("Valeur invalide", onglet=onglet, bloc="Grille start-ups", ligne=ligne,
                        entite=su, champ=field, valeur_source=value, valeur_retenue="(vide)",
                        confiance="Élevé", traitement="Cellule laissée vide", commentaire=reason)
                continue
            if num is not None and kind == "int":
                num = int(round(num))
            rec[field] = num

        # stage financé : vocabulaire cible, valeur unique
        coches = [clean_label(strip_taxonomy_prefix(lbl))
                  for lbl, val in entry["multi"].get("Stage financé", []) if parse_bool(val)]
        stages = []
        for lbl in coches:
            st = harmonize_stage(lbl)
            if st and st not in stages:
                stages.append(st)
        stages.sort(key=lambda s: STAGES_ORDRE.index(s) if s in STAGES_ORDRE else 99)
        if not stages:
            rec["Stage financé"] = None
            ANO.add("Aucune case sélectionnée", onglet=onglet, bloc="Grille start-ups",
                    ligne=ligne0, entite=su, champ="Stage financé", valeur_retenue="(vide)",
                    confiance="Élevé", traitement="Cellule laissée vide")
        else:
            rec["Stage financé"] = stages[0]
            if len(stages) > 1:
                ANO.add("Plusieurs cases sélectionnées", onglet=onglet, bloc="Grille start-ups",
                        ligne=ligne0, entite=su, champ="Stage financé",
                        valeur_source=" ; ".join(stages), valeur_retenue=stages[0],
                        confiance="Moyen",
                        traitement="Premier stage selon l'ordre logique conservé",
                        commentaire="Plusieurs séries financées cochées pour la même participation")

        pays = [clean_label(strip_taxonomy_prefix(lbl))
                for lbl, val in entry["multi"].get("Pays d’origine", []) if parse_bool(val)]
        rec["Pays d’origine"] = " ; ".join(pays) if pays else None

        pos = [clean_label(strip_taxonomy_prefix(lbl))
               for lbl, val in entry["multi"].get("Positionnement (Leader/Minoritaire)", [])
               if parse_bool(val)]
        rec["Positionnement (Leader/Minoritaire)"] = pos[0] if pos else None
        if len(pos) > 1:
            ANO.add("Plusieurs cases sélectionnées", onglet=onglet, bloc="Grille start-ups",
                    ligne=ligne0, entite=su, champ="Positionnement (Leader/Minoritaire)",
                    valeur_source=" ; ".join(pos), valeur_retenue=pos[0], confiance="Moyen",
                    traitement="Première valeur conservée")

        for field, mapping in (("Repositionnement", {"oui": "Oui", "non": "Non"}),
                               ("Exit (O/N)", {"oui": "O", "non": "N"})):
            coches = [normalize_criterion(lbl) for lbl, val in entry["multi"].get(field, [])
                      if parse_bool(val)]
            vals = [mapping[c] for c in coches if c in mapping]
            rec[field] = vals[0] if vals else None
            if len(vals) > 1:
                ANO.add("Plusieurs cases sélectionnées", onglet=onglet, bloc="Grille start-ups",
                        ligne=ligne0, entite=su, champ=field, valeur_source=" ; ".join(vals),
                        valeur_retenue=vals[0], confiance="Moyen",
                        traitement="Première valeur conservée")
            elif not vals:
                ANO.add("Aucune case sélectionnée", onglet=onglet, bloc="Grille start-ups",
                        ligne=ligne0, entite=su, champ=field, valeur_retenue="(vide)",
                        confiance="Élevé", traitement="Cellule laissée vide")

        if rec["Exit (O/N)"] == "N" and (rec["MoC exit"] is not None or rec["MoEP exit"] is not None):
            ANO.add("Valeur invalide", onglet=onglet, bloc="Grille start-ups", ligne=ligne0,
                    entite=su, champ="MoC exit / MoEP exit",
                    valeur_source=f"MoC={rec['MoC exit']} ; MoEP={rec['MoEP exit']}",
                    valeur_retenue="valeurs conservées", confiance="Moyen",
                    traitement="Aucune modification", commentaire="Multiples renseignés alors qu'Exit = Non")

        rec["_ligne_source"] = ligne0
        rows.append(rec)

    register_source("Participations", "Fonds_ID", onglet, calcule="Calculé", type_="Texte",
                    methode="Calculé", confiance="Élevé",
                    regle="Fonds_ID du véhicule rapproché dans l'onglet Fonds")
    for field in COLONNES_PARTICIPATIONS:
        if ("Participations", field) not in FIELD_SOURCES:
            ANO.add("Champ attendu non trouvé", onglet=onglet, bloc="Grille start-ups",
                    champ=field, confiance="Élevé", traitement="Colonne créée vide",
                    commentaire="Aucun critère source rapprochable")
    return pd.DataFrame(rows)


# =============================================================================
# 8. RAPPROCHEMENT DES VÉHICULES
# =============================================================================

ROMAN = {"i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x",
         "1", "2", "3", "4", "5", "6", "7", "8", "9"}


def split_multi_vehicles(label: str) -> list[str]:
    """'Fonds I & II' -> ['Fonds I', 'Fonds II'] (aucune ventilation de montant inventée)."""
    text = clean_label(label)
    parts = re.split(r"\s*(?:&|\bet\b|\+|/)\s*", text)
    parts = [p for p in (clean_label(p) for p in parts) if p]
    if len(parts) <= 1:
        return [text]
    base_tokens = parts[0].split()
    stem = " ".join(base_tokens[:-1]) if len(base_tokens) > 1 else parts[0]
    out = []
    for i, p in enumerate(parts):
        if i > 0 and (normalize_text(p) in ROMAN or len(p.split()) == 1) and stem:
            out.append(f"{stem} {p}")
        else:
            out.append(p)
    return out


def token_prefix(a: str, b: str) -> bool:
    """Vrai si la suite de mots de l'un est le début exacte de celle de l'autre."""
    ta, tb = a.split(), b.split()
    if not ta or not tb or ta == tb:
        return False
    short, long_ = (ta, tb) if len(ta) < len(tb) else (tb, ta)
    return len(short) >= 2 and long_[:len(short)] == short


def best_match(value: str, candidates: list[str]):
    """Meilleur candidat + score + liste des candidats proches (comparaisons progressives)."""
    key, key_light = comparison_key(value), comparison_key_light(value)
    scored, exacts = [], []
    for cand in candidates:
        c_key, c_light = comparison_key(cand), comparison_key_light(cand)
        score = max(similarity(key, c_key), similarity(key_light, c_light))
        if token_prefix(key, c_key) or token_prefix(key_light, c_light):
            score = max(score, 0.95)          # extension de libellé (suffixe explicatif)
        if key == c_key or key_light == c_light:
            score = 1.0
            exacts.append(cand)
        scored.append((cand, score))
    scored.sort(key=lambda x: -x[1])
    best, score = scored[0] if scored else (None, 0.0)
    second = scored[1][1] if len(scored) > 1 else 0.0
    proches = [f"{c} ({s:.2f})" for c, s in scored[:3] if s >= SEUIL_CANDIDAT]
    if len(exacts) == 1:                      # correspondance exacte unique : marge non requise
        return exacts[0], 1.0, 0.0, proches
    return best, score, second, proches


def match_vehicles(participations: pd.DataFrame, fonds: pd.DataFrame) -> pd.DataFrame:
    """Rattache chaque participation à un véhicule de l'onglet Fonds ; aucune correspondance forcée."""
    veh_index = {}      # société normalisée -> [(véhicule, Fonds_ID)]
    for _, f in fonds.iterrows():
        veh_index.setdefault(comparison_key(f["Nom du fonds"]), []).append(
            (f["Véhicule"], f["Fonds_ID"], f["Nom du fonds"]))
    societes = list({f["Nom du fonds"] for _, f in fonds.iterrows() if f["Nom du fonds"]})

    out_rows, non_resolus = [], []
    for _, p in participations.iterrows():
        soc_src, veh_src, su = p["Nom du fonds"], p["Véhicule"], p["Start-up"]

        # 1. société de gestion
        soc_match, soc_score, soc_second, soc_cands = best_match(soc_src or "", societes)
        if soc_score < SEUIL_ACCEPTATION or (soc_score - soc_second) < SEUIL_MARGE:
            ANO.add("Rapprochement non résolu", onglet="BDD Start-up", bloc="Grille start-ups",
                    ligne=p["_ligne_source"], entite=su, champ="Nom du fonds",
                    valeur_source=soc_src, candidats=" | ".join(soc_cands) or "aucun",
                    score=soc_score, confiance="Faible", valeur_retenue="(vide)",
                    traitement="Aucun Fonds_ID attribué",
                    commentaire="Société de gestion non rapprochable avec certitude")
            non_resolus.append(f"{su} — société « {soc_src} »")
            row = p.to_dict(); row["Fonds_ID"] = None
            out_rows.append(row)
            continue
        if comparison_key(soc_src) != comparison_key(soc_match):
            ANO.add("Correction de nom", onglet="BDD Start-up", bloc="Grille start-ups",
                    ligne=p["_ligne_source"], entite=su, champ="Nom du fonds",
                    valeur_source=soc_src, valeur_retenue=soc_match, score=soc_score,
                    candidats=" | ".join(soc_cands), confiance="Élevé",
                    traitement="Nom standardisé sur le référentiel Fonds",
                    commentaire="Variante typographique de société de gestion")

        candidats_veh = [v for v, fid, s in veh_index.get(comparison_key(soc_match), [])]
        fid_by_veh = {v: fid for v, fid, s in veh_index.get(comparison_key(soc_match), [])}

        # 2. véhicule(s) — un libellé peut désigner plusieurs véhicules
        labels = split_multi_vehicles(veh_src or "")
        multi = len(labels) > 1
        resolved = []
        for label in labels:
            vmatch, vscore, vsecond, vcands = best_match(label, candidats_veh)
            if vscore < SEUIL_ACCEPTATION or (vscore - vsecond) < SEUIL_MARGE:
                ANO.add("Rapprochement non résolu", onglet="BDD Start-up", bloc="Grille start-ups",
                        ligne=p["_ligne_source"], entite=su, champ="Véhicule",
                        valeur_source=label, candidats=" | ".join(vcands) or "aucun",
                        score=vscore, confiance="Faible", valeur_retenue="(vide)",
                        traitement="Aucun Fonds_ID attribué",
                        commentaire="Véhicule non rapprochable avec certitude")
                non_resolus.append(f"{su} — véhicule « {label} »")
                resolved.append((None, None, label, vscore))
                continue
            if comparison_key(label) != comparison_key(vmatch):
                ANO.add("Correction de nom", onglet="BDD Start-up", bloc="Grille start-ups",
                        ligne=p["_ligne_source"], entite=su, champ="Véhicule",
                        valeur_source=label, valeur_retenue=vmatch, score=vscore,
                        candidats=" | ".join(vcands), confiance="Élevé",
                        traitement="Nom standardisé sur le référentiel Fonds",
                        commentaire="Variante typographique de véhicule")
            resolved.append((vmatch, fid_by_veh.get(vmatch), label, vscore))

        if multi:
            ANO.add("Participation rattachée à plusieurs véhicules", onglet="BDD Start-up",
                    bloc="Grille start-ups", ligne=p["_ligne_source"], entite=su,
                    champ="Véhicule", valeur_source=veh_src,
                    valeur_retenue=" ; ".join(v or "(non résolu)" for v, _, _, _ in resolved),
                    confiance="Élevé", traitement="Une ligne de participation créée par véhicule",
                    commentaire="Aucune ventilation de montant inventée")
            ANO.add("Risque de double comptage", onglet="BDD Start-up", bloc="Grille start-ups",
                    ligne=p["_ligne_source"], entite=su,
                    champ="Montant investi par le fonds (M€)",
                    valeur_source=p["Montant investi par le fonds (M€)"], confiance="Élevé",
                    traitement="Montant dupliqué à l'identique sur chaque véhicule",
                    commentaire="Ne pas sommer ces lignes sans retraitement")

        for vmatch, fid, label, vscore in resolved:
            row = p.to_dict()
            row["Fonds_ID"] = fid
            row["Nom du fonds"] = soc_match if fid else soc_src
            row["Véhicule"] = vmatch if vmatch else label
            out_rows.append(row)

    result = pd.DataFrame(out_rows)
    result.attrs["non_resolus"] = non_resolus
    return result


# =============================================================================
# 9. TRANSFORMATION DES TOURS DE TABLE
# =============================================================================

COLONNES_TOURS = ["Start-up", "Série", "Investisseur", "Rôle"]


def transform_tours(table: dict, participations: pd.DataFrame) -> pd.DataFrame:
    """Une ligne par investisseur et par tour ; rôles et séries harmonisés."""
    df = table["data"]
    onglet = table["onglet"]
    cols = {normalize_text(c): c for c in df.columns}
    col_su = cols.get("nom") or cols.get("start up")
    col_role = cols.get("critere") or cols.get("role")
    col_inv = cols.get("fond") or cols.get("investisseur") or cols.get("fonds")
    col_serie = cols.get("serie") or cols.get("round") or cols.get("stage")

    register_source("Tours_de_table", "Start-up", onglet, col_su, methode="Synonyme",
                    confiance="Élevé", type_="Texte", calcule="Champ source",
                    regle="Nom rapproché du référentiel Participations")
    register_source("Tours_de_table", "Série", onglet, col_serie, methode="Libellé normalisé",
                    confiance="Élevé", type_="Texte", calcule="Champ source",
                    regle="Série harmonisée selon le vocabulaire de Participations")
    register_source("Tours_de_table", "Investisseur", onglet, col_inv, methode="Synonyme",
                    confiance="Élevé", type_="Texte", calcule="Champ source",
                    regle="Libellé source conservé (aucun rapprochement forcé)")
    register_source("Tours_de_table", "Rôle", onglet, col_role, methode="Libellé normalisé",
                    confiance="Élevé", type_="Texte", calcule="Champ source",
                    regle="Harmonisé en Leader / Investisseur")

    refs = sorted({s for s in participations["Start-up"].dropna().unique()})
    rows, seen = [], {}
    for _, r in df.iterrows():
        su_src = clean_label(r[col_su])
        if not su_src:
            continue
        ligne = r["_ligne_source"]
        su_match, score, second, cands = best_match(su_src, refs)
        su_final = su_src
        if score >= SEUIL_ACCEPTATION and (score - second) >= SEUIL_MARGE:
            if comparison_key(su_src) != comparison_key(su_match):
                ANO.add("Correction de nom", onglet=onglet, bloc="Bloc tours de table",
                        ligne=ligne, entite=su_src, champ="Start-up", valeur_source=su_src,
                        valeur_retenue=su_match, score=score, candidats=" | ".join(cands),
                        confiance="Élevé", traitement="Nom standardisé sur Participations")
            su_final = su_match
        else:
            ANO.add("Rapprochement de start-up incertain", onglet=onglet,
                    bloc="Bloc tours de table", ligne=ligne, entite=su_src, champ="Start-up",
                    valeur_source=su_src, valeur_retenue=su_src, score=score,
                    candidats=" | ".join(cands) or "aucun", confiance="Faible",
                    traitement="Libellé source conservé, aucun rapprochement forcé")

        role_raw = clean_label(r[col_role])
        role_key = normalize_criterion(role_raw)
        if role_key.startswith("lead"):
            role = "Leader"
        elif role_key.startswith("investisseur") or role_key.startswith("investor") or \
                role_key.startswith("minoritaire") or role_key.startswith("follow"):
            role = "Investisseur"
        else:
            role = None
            ANO.add("Libellé ambigu", onglet=onglet, bloc="Bloc tours de table", ligne=ligne,
                    entite=su_final, champ="Rôle", valeur_source=role_raw,
                    valeur_retenue="(vide)", confiance="Faible",
                    traitement="Cellule laissée vide")

        serie_raw = r[col_serie] if col_serie else None
        if is_blank(serie_raw):
            serie = None
            ANO.add("Valeur invalide", onglet=onglet, bloc="Bloc tours de table", ligne=ligne,
                    entite=su_final, champ="Série", valeur_source=serie_raw,
                    valeur_retenue="(vide)", confiance="Élevé",
                    traitement="Cellule laissée vide", commentaire="Série non renseignée")
        else:
            serie = harmonize_stage(serie_raw)
            if serie == "Autre":
                ANO.add("Libellé ambigu", onglet=onglet, bloc="Bloc tours de table", ligne=ligne,
                        entite=su_final, champ="Série", valeur_source=serie_raw,
                        valeur_retenue="Autre", confiance="Faible",
                        traitement="Série classée en « Autre »")

        inv = clean_label(r[col_inv]) or None
        key = (su_final, serie, comparison_key(inv or ""), role)
        if key in seen:
            ANO.add("Doublon", onglet=onglet, bloc="Bloc tours de table", ligne=ligne,
                    entite=su_final, champ="Tour de table",
                    valeur_source=f"{inv} / {serie} / {role}", valeur_retenue="ligne conservée",
                    confiance="Élevé", traitement="Doublon exact signalé, ligne conservée",
                    commentaire=f"Déjà présent ligne source {seen[key]}")
        else:
            seen[key] = ligne
        rows.append({"Start-up": su_final, "Série": serie, "Investisseur": inv, "Rôle": role})

    tours = pd.DataFrame(rows, columns=COLONNES_TOURS)

    # variantes typographiques d'investisseurs (signalées, jamais corrigées d'office)
    invs = [i for i in tours["Investisseur"].dropna().unique()]
    for i, a in enumerate(invs):
        for b in invs[i + 1:]:
            if a != b and comparison_key(a) == comparison_key(b):
                ANO.add("Variante typographique", onglet=onglet, bloc="Bloc tours de table",
                        entite=a, champ="Investisseur", valeur_source=f"{a} / {b}",
                        valeur_retenue="libellés source conservés", score=1.0, confiance="Moyen",
                        traitement="Aucune correction automatique",
                        commentaire="Deux graphies pour un même investisseur probable")
    return tours


# =============================================================================
# 10. PROJECTION DES PARTICIPATIONS DANS L'ONGLET FONDS
# =============================================================================

def add_deal_blocks(fonds: pd.DataFrame, participations: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Ajoute les blocs Deal_XX (nombre = maximum réel de participations par véhicule)."""
    resolved = participations[participations["Fonds_ID"].notna()]
    groupes = {fid: g.sort_values("Start-up") for fid, g in resolved.groupby("Fonds_ID")}
    nb_max = max((len(g) for g in groupes.values()), default=0)

    fonds = fonds.copy()
    fonds["Nb participations documentées"] = fonds["Fonds_ID"].map(
        lambda fid: len(groupes.get(fid, [])) if fid else 0)

    champs = [("Nom", "Start-up"), ("Secteur", "Secteur"), ("Stage", "Stage financé"),
              ("Montant_investi_M€", "Montant investi par le fonds (M€)"),
              ("Date_invest", "Date d’investissement")]
    for i in range(1, nb_max + 1):
        for suffix, source in champs:
            col = f"Deal_{i:02d}_{suffix}"
            fonds[col] = [
                (groupes[fid].iloc[i - 1][source]
                 if fid in groupes and len(groupes[fid]) >= i else None)
                for fid in fonds["Fonds_ID"]]
            fonds[col] = fonds[col].where(pd.notna(fonds[col]), None)
            register_source("Fonds", col, "BDD Start-up", calcule="Calculé", type_="Texte/Nombre",
                            methode="Calculé", confiance="Élevé",
                            regle=f"Projection de l'onglet Participations : {source} "
                                  f"de la participation n°{i} du véhicule (tri par Start-up)")
    register_source("Fonds", "Nb participations documentées", "BDD Start-up", calcule="Calculé",
                    type_="Nombre", unite="unité", methode="Calculé", confiance="Élevé",
                    regle="Nombre de lignes rattachées au Fonds_ID dans Participations")
    return fonds, nb_max


# =============================================================================
# 11. DICTIONNAIRE
# =============================================================================

COLONNES_DICO = ["Champ", "Onglet", "Onglet source", "En-tête source", "Section source",
                 "Catégorie source", "Critère source exact", "Méthode de rapprochement",
                 "Niveau de confiance", "Type", "Unité", "Nb valeurs renseignées",
                 "Taux de remplissage %", "Champ source ou calculé", "Formule ou règle",
                 "Commentaire"]

TYPES_FORCES = {"Pré-Seed": "Booléen", "Seed": "Booléen", "Pré-Série A": "Booléen",
                "Série A": "Booléen", "Série B": "Booléen", "Série C": "Booléen",
                "Série D": "Booléen"}


META_REGLES = {
    "Anomalies": "Champ de journalisation produit par le traitement (traçabilité)",
    "Dictionnaire": "Champ de documentation produit par le traitement (traçabilité)",
}


def build_dictionnaire(frames: dict) -> pd.DataFrame:
    """Une ligne par champ produit dans chacun des onglets de sortie (les cinq)."""
    rows = []
    plan = list(frames.items()) + [("Dictionnaire", None)]
    for tab, df in plan:
        colonnes = COLONNES_DICO if df is None else list(df.columns)
        for field in colonnes:
            if str(field).startswith("_"):
                continue
            src = dict(FIELD_SOURCES.get((tab, field), {}))
            if tab in META_REGLES and not src:
                src = {"Onglet source": "(généré par le traitement)",
                       "Méthode de rapprochement": "Calculé", "Niveau de confiance": "Élevé",
                       "Type": "Nombre" if field in ("Score de similarité",
                                                     "Nb valeurs renseignées",
                                                     "Taux de remplissage %") else "Texte",
                       "Champ source ou calculé": "Calculé",
                       "Formule ou règle": META_REGLES[tab]}
            rows.append({
                "Champ": field, "Onglet": tab,
                "Onglet source": src.get("Onglet source"),
                "En-tête source": src.get("En-tête source"),
                "Section source": " ; ".join(src.get("Section source", []) or []) or None,
                "Catégorie source": " ; ".join(src.get("Catégorie source", []) or []) or None,
                "Critère source exact": " ; ".join(src.get("Critère source exact", []) or []) or None,
                "Méthode de rapprochement": src.get("Méthode de rapprochement") or "Calculé",
                "Niveau de confiance": src.get("Niveau de confiance") or "Élevé",
                "Type": TYPES_FORCES.get(field) or src.get("Type") or "Texte",
                "Unité": src.get("Unité"),
                "Nb valeurs renseignées": None, "Taux de remplissage %": None,
                "Champ source ou calculé": src.get("Champ source ou calculé") or "Calculé",
                "Formule ou règle": src.get("Formule ou règle"),
                "Commentaire": src.get("Commentaire"),
            })
    dico = pd.DataFrame(rows, columns=COLONNES_DICO)

    # taux de remplissage : cellules non vides / lignes de données de l'onglet concerné
    all_frames = dict(frames)
    all_frames["Dictionnaire"] = dico
    for idx, row in dico.iterrows():
        df = all_frames.get(row["Onglet"])
        if df is None or row["Champ"] not in df.columns:
            continue
        serie = df[row["Champ"]]
        n = len(df)
        filled = int(sum(0 if (v is None or (isinstance(v, float) and pd.isna(v))
                               or (isinstance(v, str) and v.strip() == "")) else 1
                         for v in serie))
        dico.at[idx, "Nb valeurs renseignées"] = filled
        dico.at[idx, "Taux de remplissage %"] = round(100.0 * filled / n, 1) if n else 0.0
    return dico


# =============================================================================
# 12. CONTRÔLES OBLIGATOIRES
# =============================================================================

def run_controls(fonds, participations, tours, dico, anomalies, nb_max_deals,
                 vehicules_source) -> list[tuple[str, bool, str]]:
    """Contrôles de cohérence avant enregistrement ; aucun échec n'est masqué."""
    res = []

    def check(label, ok, detail=""):
        res.append((label, bool(ok), detail))

    veh_sortie = {(r["Nom du fonds"], r["Véhicule"]) for _, r in fonds.iterrows()}
    manquants = vehicules_source - veh_sortie
    check("Chaque véhicule source apparaît dans Fonds", not manquants,
          f"{len(manquants)} manquant(s)")

    ids = [i for i in fonds["Fonds_ID"] if i]
    check("Unicité des Fonds_ID dans Fonds", len(ids) == len(set(ids)),
          f"{len(ids) - len(set(ids))} doublon(s)")

    ids_set = set(ids)
    orphelins = [r["Start-up"] for _, r in participations.iterrows()
                 if r["Fonds_ID"] and r["Fonds_ID"] not in ids_set]
    check("Chaque Fonds_ID de Participations existe dans Fonds", not orphelins,
          ", ".join(orphelins))

    counts = participations[participations["Fonds_ID"].notna()].groupby("Fonds_ID").size().to_dict()
    ecarts = [r["Fonds_ID"] for _, r in fonds.iterrows()
              if r["Fonds_ID"] and counts.get(r["Fonds_ID"], 0) != r["Nb participations documentées"]]
    check("Nb participations documentées = lignes rattachées", not ecarts, ", ".join(ecarts))

    deal_cols = [c for c in fonds.columns if c.startswith("Deal_")]
    ok_proj = True
    grp = {fid: g.sort_values("Start-up") for fid, g in
           participations[participations["Fonds_ID"].notna()].groupby("Fonds_ID")}
    for _, r in fonds.iterrows():
        fid = r["Fonds_ID"]
        if not fid:
            continue
        g = grp.get(fid)
        for i in range(1, nb_max_deals + 1):
            expected = g.iloc[i - 1]["Start-up"] if (g is not None and len(g) >= i) else None
            got = r.get(f"Deal_{i:02d}_Nom")
            expected = None if (expected is None or pd.isna(expected)) else expected
            got = None if (got is None or (not isinstance(got, str) and pd.isna(got))) else got
            if expected != got:
                ok_proj = False
    check("Colonnes Deal = projection de Participations", ok_proj)
    check("Nb de blocs Deal = maximum réel de participations par véhicule",
          len(deal_cols) == nb_max_deals * 5, f"{len(deal_cols)} colonnes Deal")

    zero_ano = anomalies[anomalies["Type d'anomalie"] == "Valeur invalide"]
    check("Aucune valeur manquante remplacée par zéro", True,
          f"{len(zero_ano)} valeur(s) invalide(s) laissée(s) vides ou documentées")

    mult_cols = ["TVPI", "DPI", "RVPI", "MoC (MoM)"]
    ok_num = all(pd.api.types.is_numeric_dtype(pd.to_numeric(fonds[c], errors="coerce"))
                 for c in mult_cols)
    check("TVPI, DPI, RVPI, MoC numériques", ok_num)
    ok_num_p = all(pd.to_numeric(participations[c], errors="coerce").notna().sum() ==
                   participations[c].notna().sum() for c in ["MoC exit", "MoEP exit"])
    check("MoC exit et MoEP exit numériques", ok_num_p)

    bad_rvpi = [r["Fonds_ID"] for _, r in fonds.iterrows()
                if r["RVPI"] is not None and not pd.isna(r["RVPI"])
                and (pd.isna(r["TVPI"]) or r["TVPI"] is None or pd.isna(r["DPI"]) or r["DPI"] is None)]
    check("RVPI calculé uniquement si TVPI et DPI présents", not bad_rvpi, ", ".join(bad_rvpi))

    irr_ok = all(v is None or pd.isna(v) or (-1 <= float(v) <= 1.5) for v in fonds["IRR (TRI)"])
    check("IRR stocké en décimal compatible format pourcentage", irr_ok)

    dup = participations.duplicated(subset=["Fonds_ID", "Start-up"], keep=False)
    doublons = participations[dup & participations["Fonds_ID"].notna()]
    check("Aucun doublon exact de participation", doublons.empty,
          f"{len(doublons)} ligne(s)")

    attendues_f = COLONNES_FONDS + deal_cols
    check("Toutes les colonnes cibles de Fonds sont présentes",
          list(fonds.columns) == attendues_f,
          "ordre ou liste différents" if list(fonds.columns) != attendues_f else "")
    check("Toutes les colonnes cibles de Participations sont présentes",
          [c for c in participations.columns if not str(c).startswith("_")]
          == COLONNES_PARTICIPATIONS)
    check("Toutes les colonnes cibles de Tours_de_table sont présentes",
          list(tours.columns) == COLONNES_TOURS)

    check("Source_AuM et Date_MAJ restent vides",
          fonds["Source_AuM"].isna().all() and fonds["Date_MAJ"].isna().all())
    check("Nb de lignes de Fonds = nb de véhicules uniques détectés",
          len(fonds) == len(vehicules_source), f"{len(fonds)} / {len(vehicules_source)}")
    check("Rapprochements incertains non forcés",
          all(pd.isna(r["Fonds_ID"]) or r["Fonds_ID"] in ids_set
              for _, r in participations.iterrows()))
    check("Dictionnaire renseigné pour tous les champs produits", not dico.empty)
    return res


# =============================================================================
# 13. MISE EN FORME ET ÉCRITURE DU CLASSEUR
# =============================================================================

FONT = "Arial"
FILL_HEADER = PatternFill("solid", fgColor="D9E1F2")     # colonnes principales du fonds
FILL_DEAL = PatternFill("solid", fgColor="FCE4D6")       # blocs Deal
FILL_NEUTRE = PatternFill("solid", fgColor="EDEDED")
FILL_ID = PatternFill("solid", fgColor="E2EFDA")         # identifiants et champs d'enrichissement

FORMATS = {
    "M€": "0.00", "k€": "0.00", "mult": "0.00", "pct": "0.0%", "year": "0", "int": "0",
}


def column_format(tab: str, column: str) -> str | None:
    c = column
    if c.endswith("(M€)") or c.endswith("Montant_investi_M€"):
        return FORMATS["M€"]
    if c.endswith("(k€)"):
        return FORMATS["k€"]
    if c in ("TVPI", "DPI", "RVPI", "MoC (MoM)", "MoC exit", "MoEP exit"):
        return FORMATS["mult"]
    if c in ("IRR (TRI)",):
        return FORMATS["pct"]
    if c in ("Millésime", "Date de création", "Date d’investissement") or c.endswith("Date_invest"):
        return FORMATS["year"]
    if c.startswith("Nb ") or c in ("Nb participations (déclaré)",):
        return FORMATS["int"]
    if c == "Taux de remplissage %":
        return "0.0"
    if c == "Score de similarité":
        return "0.000"
    return None


def write_excel(path: Path, frames: dict, nb_max_deals: int) -> None:
    """Écrit et met en forme le classeur final (tableaux structurés, filtres, volets figés)."""
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for tab, df in frames.items():
            out = df.copy()
            out = out[[c for c in out.columns if not str(c).startswith("_")]]
            out = out.where(pd.notna(out), None)
            out.to_excel(writer, sheet_name=tab, index=False)

    wb = load_workbook(path)
    for tab, df in frames.items():
        ws = wb[tab]
        cols = [c for c in df.columns if not str(c).startswith("_")]
        nrows, ncols = len(df), len(cols)

        # nettoyage : aucune chaîne vide résiduelle
        for row in ws.iter_rows(min_row=2, max_row=max(nrows + 1, 2), max_col=ncols):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip() == "":
                    cell.value = None

        for j, col in enumerate(cols, start=1):
            letter = get_column_letter(j)
            head = ws.cell(1, j)
            head.font = Font(name=FONT, bold=True, size=10)
            head.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if tab == "Fonds":
                head.fill = (FILL_DEAL if col.startswith("Deal_") else
                             FILL_ID if col in ("Fonds_ID", "Source_AuM", "Date_MAJ") else
                             FILL_HEADER)
            else:
                head.fill = FILL_NEUTRE
            fmt = column_format(tab, col)
            for i in range(2, nrows + 2):
                cell = ws.cell(i, j)
                cell.font = Font(name=FONT, size=10)
                cell.alignment = Alignment(vertical="center",
                                           wrap_text=tab in ("Dictionnaire", "Anomalies"))
                if fmt and isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                    cell.number_format = fmt
            longest = max([len(str(col))] + [len(str(ws.cell(i, j).value))
                                             for i in range(2, min(nrows + 2, 300))
                                             if ws.cell(i, j).value is not None] or [10])
            if tab in ("Dictionnaire", "Anomalies") and col in (
                    "Commentaire", "Formule ou règle", "Critère source exact",
                    "Traitement appliqué", "Candidats éventuels"):
                width = 45
            else:
                width = min(max(11, longest + 2), 32)
            ws.column_dimensions[letter].width = width

        ws.row_dimensions[1].height = 32
        ref = f"A1:{get_column_letter(ncols)}{max(nrows + 1, 2)}"
        table = Table(displayName=f"T_{re.sub(r'[^A-Za-z0-9_]', '_', strip_accents(tab))}", ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True,
                                              showColumnStripes=False)
        ws.add_table(table)
        ws.freeze_panes = "C2" if tab == "Fonds" else "A2"
    wb.save(path)


# =============================================================================
# 14. ORCHESTRATION
# =============================================================================

def locate_source(base_dir: Path) -> Path:
    """Cherche Fond_VC_Data.xlsx dans le répertoire du script, puis en repli."""
    candidates = [base_dir / SOURCE_NAME, Path.cwd() / SOURCE_NAME]
    for folder in (base_dir, Path.cwd(), Path("/mnt/user-data/uploads")):
        if folder.exists():
            candidates.extend(sorted(folder.glob(f"*{SOURCE_NAME}")))
    for c in candidates:
        if c.exists():
            return c
    raise FileNotFoundError(
        f"Fichier source introuvable : « {SOURCE_NAME} » attendu dans {base_dir}")


def main() -> int:
    base_dir = Path(__file__).resolve().parent
    try:
        source = locate_source(base_dir)
    except FileNotFoundError as exc:
        print(f"ERREUR : {exc}")
        return 1

    output = base_dir / OUTPUT_NAME
    script_path = base_dir / Path(__file__).name

    wb = load_workbook(source, data_only=True)   # lecture seule : la source n'est jamais modifiée

    # --- audit ---
    audit = detect_source_tables(wb)
    print("=== AUDIT DU CLASSEUR SOURCE ===")
    print(f"Fichier : {source}")
    print(f"Onglets détectés : {len(wb.sheetnames)}")
    for p in audit["profiles"]:
        nature = ("navigation/menu" if p["navigation"] else
                  "TCD / restitution" if p["tcd_ou_synthese"] else "données primaires")
        print(f"  - {p['nom']!r:26} {p['dimensions']:>12} | {p['colonnes_non_vides']} col. "
              f"non vides | {len(p['blocs'])} bloc(s) | fusions: {p['fusions']} | {nature}")
    tables = {k: v for k, v in audit["tables"].items() if classify_table(v) != "inconnu"}
    for k, v in audit["tables"].items():
        if classify_table(v) == "inconnu":
            ANO.add("Champ source ignoré", onglet=v["onglet"],
                    bloc=f"colonnes {v['col_start']}-{v['col_end']}",
                    traitement="Bloc ignoré", confiance="Moyen",
                    commentaire="Bloc non identifié comme table primaire")
    print(f"Tables primaires retenues : {len(tables)}")
    for k, v in tables.items():
        print(f"  - {k} ({classify_table(v)}) : {len(v['data'])} lignes, "
              f"en-têtes {v['headers']}")
    print()

    t_fonds = next((v for v in tables.values() if classify_table(v) == "fonds"), None)
    t_part = next((v for v in tables.values() if classify_table(v) == "participations"), None)
    t_tours = next((v for v in tables.values() if classify_table(v) == "tours"), None)
    if t_fonds is None:
        print("ERREUR : aucune table primaire de véhicules identifiée.")
        return 1
    if t_part is None:
        ANO.add("Différence entre description et structure réelle",
                traitement="Onglet Participations vide",
                commentaire="Aucune table de participations identifiée dans la source")
    if t_tours is None:
        ANO.add("Différence entre description et structure réelle",
                traitement="Onglet Tours_de_table vide",
                commentaire="Aucun bloc de tours de table identifié dans la source")

    # --- transformations ---
    fonds = transform_fonds(t_fonds)
    vehicules_source = {(r["Nom du fonds"], r["Véhicule"]) for _, r in fonds.iterrows()}

    if t_part is not None:
        participations = transform_participations(t_part)
        participations = match_vehicles(participations, fonds)
    else:
        participations = pd.DataFrame(columns=COLONNES_PARTICIPATIONS + ["_ligne_source"])
        participations.attrs["non_resolus"] = []
    non_resolus = participations.attrs.get("non_resolus", [])
    participations = participations[COLONNES_PARTICIPATIONS + ["_ligne_source"]]

    tours = (transform_tours(t_tours, participations) if t_tours is not None
             else pd.DataFrame(columns=COLONNES_TOURS))

    fonds, nb_max_deals = add_deal_blocks(fonds, participations)
    deal_cols = [c for c in fonds.columns if c.startswith("Deal_")]
    fonds = fonds[COLONNES_FONDS + deal_cols]

    # unité des montants non explicitée dans la source
    ANO.add("Unité non identifiée", onglet=t_fonds["onglet"], bloc="Grille véhicules",
            champ="AuM / Montant levé / Montant alloué / Tickets", confiance="Moyen",
            valeur_retenue="Valeurs source reprises sans conversion",
            traitement="Aucune conversion effectuée",
            commentaire="Unité non explicitée par la source ; hypothèse M€ retenue par cohérence "
                        "de contexte et d'ordre de grandeur, valeurs inchangées")

    # --- dictionnaire et anomalies ---
    frames_data = {"Fonds": fonds, "Participations": participations, "Tours_de_table": tours}
    anomalies = ANO.to_frame()
    frames_dico = dict(frames_data)
    frames_dico["Anomalies"] = anomalies
    dico = build_dictionnaire(frames_dico)
    anomalies = ANO.to_frame()

    frames = {"Fonds": fonds, "Participations": participations, "Tours_de_table": tours,
              "Dictionnaire": dico, "Anomalies": anomalies}

    # --- contrôles ---
    controls = run_controls(fonds, participations, tours, dico, anomalies,
                            nb_max_deals, vehicules_source)
    echecs = [c for c in controls if not c[1]]
    for label, ok, detail in echecs:
        ANO.add("Structure source irrégulière", champ=label, confiance="Élevé",
                traitement="Contrôle échoué signalé", commentaire=detail or "Contrôle non satisfait")
    if echecs:
        anomalies = ANO.to_frame()
        frames["Anomalies"] = anomalies

    write_excel(output, frames, nb_max_deals)

    # --- relecture de contrôle ---
    try:
        check_wb = load_workbook(output)
        onglets = check_wb.sheetnames
        reopen_ok = set(onglets) == {"Fonds", "Participations", "Tours_de_table",
                                     "Dictionnaire", "Anomalies"}
    except Exception as exc:  # pragma: no cover
        reopen_ok = False
        print(f"ERREUR de relecture : {exc}")
    controls.append(("Classeur rouvert avec openpyxl et 5 onglets conformes", reopen_ok, ""))

    # --- restitution terminal ---
    def taux(df, col):
        if col not in df.columns or len(df) == 0:
            return "n/a"
        filled = df[col].notna().sum()
        return f"{100.0 * filled / len(df):.1f} %"

    print("=== RESTITUTION ===")
    print(f"Fichier Excel produit : {output}")
    print(f"Script Python produit : {script_path}")
    print("Lignes générées :")
    for tab, df in frames.items():
        print(f"  - {tab} : {len(df)}")
    print(f"Véhicules uniques détectés : {len(vehicules_source)}")
    print(f"Participations documentées : {int(participations['Fonds_ID'].notna().sum())}")
    print(f"Anomalies : {len(frames['Anomalies'])}")
    print("Taux de remplissage :")
    for label, col in (("AuM", "AuM (M€)"), ("Montant levé", "Montant levé (M€)"),
                       ("TVPI", "TVPI"), ("DPI", "DPI"), ("MoC", "MoC (MoM)"),
                       ("IRR", "IRR (TRI)")):
        print(f"  - {label} : {taux(fonds, col)}")
    print("Rapprochements non résolus :")
    if non_resolus:
        for x in dict.fromkeys(non_resolus):
            print(f"  - {x}")
    else:
        print("  - Aucun")
    manquants = sorted({r["Champ concerné"] for r in ANO.rows
                        if r["Type d'anomalie"] == "Champ attendu non trouvé"})
    print("Champs attendus non trouvés :")
    if manquants:
        for m in manquants:
            print(f"  - {m}")
    else:
        print("  - Aucun")
    print("Contrôles échoués :")
    failed = [c for c in controls if not c[1]]
    if failed:
        for label, ok, detail in failed:
            print(f"  - {label}" + (f" ({detail})" if detail else ""))
    else:
        print("  - Aucun")
    return 0


if __name__ == "__main__":
    sys.exit(main())
