#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
collecte_perf_vc.py — Passe de recherche ciblée sur les indicateurs de performance (source 4)
================================================================================================

Entrée  : VC_Database_Standardisee_v4.xlsx
Sortie  : VC_Database_Standardisee_v5.xlsx

Passe complémentaire ciblée sur MoC (MoM) et IRR (TRI), à partir de rapports annuels,
brochures investisseurs et communiqués de presse citant la performance historique d'un
fonds (souvent repris lors du closing du fonds suivant). Recherche menée par 12 lots
(mêmes regroupements que la passe précédente) sur les 111 véhicules.

Constat de cette passe : la performance d'un fonds de capital-risque privé n'est presque
jamais publique (confirmé systématiquement sur 111 véhicules) ; les rares chiffres trouvés
proviennent soit de LPs institutionnels publics, soit de communiqués de closing citant la
performance du fonds précédent. Chaque figure trouvée est explicitement qualifiée Net ou
Gross (jamais supposée) ; les objectifs prévisionnels ("expected to generate...") sont
exclus des cellules de données et documentés en Anomalies uniquement.

Règles appliquées : identiques à collecte_web_vc.py (cellules vides uniquement, aucune
estimation, traçabilité Source_AuM/Date_MAJ, numérotation Anomalies séquentielle reprise
après la dernière ligne existante).
"""

from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import restructuration_vc as rvc  # réutilisation des fonctions de mise en forme

FICHIER_ENTREE = "VC_Database_Standardisee_v4.xlsx"
FICHIER_SORTIE = "VC_Database_Standardisee_v5.xlsx"
DATE_MAJ = date(2026, 9, 15)
ONGLETS = ["Fonds", "Participations", "Tours_de_table", "Dictionnaire", "Anomalies"]

SOURCE_LABEL = ("Passe ciblée MoC/IRR par lots (rapports annuels, brochures investisseurs, "
                "communiqués de closing citant la performance du fonds précédent) — 15/09/2026")


# =============================================================================
# 1. TABLE DE COLLECTE — les seules données de performance jugées assez fiables pour être
#    écrites dans une cellule vide (net/gross toujours explicite)
# =============================================================================

COLLECTE_PERF = [
    dict(
        fid="committed-capital_committed-capital-eis-fund",
        vals={"MoC (MoM)": 3.0},
        net_gross={"MoC (MoM)": "non précisé"},
        confiance="Moyen",
        src=("Committed Capital, page officielle 'Growth EIS Fund' (committedcapital.co.uk/"
             "growth-eis-fund/) et IFA Magazine — ROI de 3,0x (hors avantages fiscaux) ; "
             "le libellé net/gross n'est pas précisé par la source. Complète le TVPI=3,0 et "
             "l'IRR=36,8 % déjà renseignés lors de la passe de collecte web précédente pour "
             "ce même fonds."),
    ),
]


# =============================================================================
# 2. DIVERGENCES — la base a déjà une valeur, la source web la contredit ou la nuance
# =============================================================================

DIVERGENCES = [
    dict(fid="white-star-capital_wsc-ii", champ="TVPI", base="1.81",
         web="TVPI ≈ 2,5x selon des propos du managing partner Eric Martineau-Fortin "
             "rapportés par Unquote.com (closing du Fund III, ~oct. 2021) ; net/gross non "
             "précisé par la source",
         com="Valeur base conservée (1,81) : la source web est une déclaration orale "
             "approximative et non datée avec précision, sans label net/gross, à une date "
             "antérieure à la valeur de la base actuelle — non arbitrée."),
]


# =============================================================================
# 3. OBJECTIFS PRÉVISIONNELS ÉCARTÉS — cibles annoncées par un fonds, jamais des
#    performances réalisées : documentées, jamais écrites dans une cellule
# =============================================================================

CIBLES_NON_RETENUES = [
    dict(fid="truffle-capital_truffle-fintech-et-insurtech-fund-ii",
         champ="MoC (MoM) / IRR (TRI)",
         candidats=("Objectif annoncé par Philippe Pouletty (CEO) lors du closing (~140 M€, "
                    "2019) : multiple de 2,5x–3,5x et IRR d'environ 20 % ('expected to "
                    "generate') — Unquote.com, communiqué Truffle Capital"),
         com="Chiffres explicitement prévisionnels (cible annoncée par le fonds lui-même lors "
             "de sa levée), pas une performance réalisée mesurée a posteriori : cellule "
             "laissée vide conformément à la règle anti-estimation."),
]


# =============================================================================
# 4. RAPPROCHEMENTS NON RÉSOLUS COMPLÉMENTAIRES — identifiés lors de cette passe
# =============================================================================

IDENTITY_RISKS = [
    dict(fids=["xange-capital_xange-digital-3"],
         commentaire="Un article de blog XAnge (nov. 2021, 'VC Performance Explained') cite "
             "un IRR net > 19 % et un DPI net de 0,23x pour un fonds désigné uniquement "
             "'XAnge 3' (vintage 2017) — rapprochement avec XAnge Digital 3 plausible sur le "
             "millésime mais non confirmé nominalement ; non retenu comme donnée fiable."),
    dict(fids=["eurazeo_eurazeo-venture-capital"],
         commentaire="Les rapports trimestriels Eurazeo ('Trading Update', PDF) publient un "
             "IRR/MOIC Gross par véhicule de la série 'Digital' (Digital II : Gross IRR 18 %, "
             "Gross MOIC 2,3x au FY2023 ; Digital III : Gross IRR en baisse de 17 % à 6 % entre "
             "FY2023 et 9M2025, Gross MOIC 1,3x) — mais aucun véhicule ne porte le nom exact "
             "'Eurazeo Venture Capital' (qui désigne la ligne métier, pas un fonds). "
             "Rapprochement non confirmé ; aucune valeur reportée sur cette ligne."),
]


# =============================================================================
# 5. OUTILS (identiques à collecte_web_vc.py)
# =============================================================================

def localiser_entree(base_dir: Path) -> Path:
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
            "Table ou bloc source": "Rapports annuels / brochures investisseurs / presse spécialisée",
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
# 6. APPLICATION
# =============================================================================

def appliquer_collecte(fonds: pd.DataFrame, journal: JournalAnomalies) -> tuple[int, list[str]]:
    remplies, touchees = 0, []
    index = {fid: i for i, fid in enumerate(fonds["Fonds_ID"])}
    for maj in COLLECTE_PERF:
        fid = maj["fid"]
        if fid not in index:
            journal.ajouter("Rapprochement non résolu", entite=fid, champ="Fonds_ID",
                            valeur_source=fid, confiance="Élevé",
                            traitement="Mise à jour non appliquée",
                            commentaire="Fonds_ID absent de la base : vérifier le référentiel")
            continue
        i = index[fid]
        entite = f"{fonds.at[i, 'Nom du fonds']} / {fonds.at[i, 'Véhicule']}"
        ecrit = []
        for champ, valeur in maj["vals"].items():
            if not est_vide(fonds.at[i, champ]):
                continue  # cellule déjà renseignée entre-temps : on ne touche à rien
            fonds.at[i, champ] = valeur
            ecrit.append(champ)
        if not ecrit:
            continue
        remplies += len(ecrit)
        touchees.append(entite)
        ng = "; ".join(f"{c}={maj['net_gross'].get(c, 'non précisé')}" for c in ecrit)
        note = f"[Collecte perf {DATE_MAJ:%d/%m/%Y}] {maj['src']} (net/gross : {ng})"
        actuelle_src = fonds.at[i, "Source_AuM"]
        fonds.at[i, "Source_AuM"] = note if est_vide(actuelle_src) else f"{actuelle_src} | {note}"
        fonds.at[i, "Date_MAJ"] = DATE_MAJ
        journal.ajouter("Mise à jour web", entite=entite, champ=" ; ".join(ecrit),
                        valeur_retenue=" ; ".join(f"{c}={maj['vals'][c]}" for c in ecrit),
                        confiance=maj["confiance"],
                        traitement="Cellules vides complétées à partir de la source web "
                                   "(indicateur de performance, net/gross précisé)",
                        commentaire=maj["src"])
    return remplies, touchees


def journaliser_divergences(fonds: pd.DataFrame, journal: JournalAnomalies) -> None:
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


def journaliser_cibles_non_retenues(fonds: pd.DataFrame, journal: JournalAnomalies) -> None:
    index = {fid: i for i, fid in enumerate(fonds["Fonds_ID"])}
    for c in CIBLES_NON_RETENUES:
        i = index.get(c["fid"])
        entite = (f"{fonds.at[i, 'Nom du fonds']} / {fonds.at[i, 'Véhicule']}"
                  if i is not None else c["fid"])
        journal.ajouter("Valeur en fourchette non retenue", entite=entite, champ=c["champ"],
                        candidats=c["candidats"], confiance="Faible",
                        traitement="Cellule laissée vide : chiffre prévisionnel/cible, pas une "
                                   "performance réalisée",
                        commentaire=c["com"])


def journaliser_identity_risks(fonds: pd.DataFrame, journal: JournalAnomalies) -> None:
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


def recalculer_dictionnaire(dico: pd.DataFrame, frames: dict) -> pd.DataFrame:
    for idx, row in dico.iterrows():
        df = frames.get(row["Onglet"])
        if df is None or row["Champ"] not in df.columns:
            continue
        n = len(df)
        remplies = int(sum(0 if est_vide(v) else 1 for v in df[row["Champ"]]))
        dico.at[idx, "Nb valeurs renseignées"] = remplies
        dico.at[idx, "Taux de remplissage %"] = round(100.0 * remplies / n, 1) if n else 0.0
    return dico


# =============================================================================
# 7. ORCHESTRATION
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
    remplies, touchees = appliquer_collecte(fonds, journal)
    journaliser_divergences(fonds, journal)
    journaliser_cibles_non_retenues(fonds, journal)
    journaliser_identity_risks(fonds, journal)
    frames["Fonds"] = fonds

    frames["Anomalies"] = journal.resultat()
    frames["Dictionnaire"] = recalculer_dictionnaire(frames["Dictionnaire"], frames)

    nb_deals = len([c for c in fonds.columns if str(c).startswith("Deal_")]) // 5
    rvc.write_excel(sortie, frames, nb_deals)

    apres = {k: int(sum(0 if est_vide(x) else 1 for col in v.columns for x in v[col]))
             for k, v in frames.items()}
    print(f"Fichier produit : {sortie}")
    print(f"Entrée          : {entree}")
    print()
    print(f"Cellules vides complétées par la passe performance : {remplies}")
    print(f"Véhicules mis à jour : {len(touchees)}")
    for t in touchees:
        print(f"  - {t}")
    print()
    print(f"Divergences journalisées : {len(DIVERGENCES)}")
    print(f"Cibles prévisionnelles non retenues : {len(CIBLES_NON_RETENUES)}")
    print(f"Rapprochements non résolus (complémentaires) : {len(IDENTITY_RISKS)}")
    print()
    print("Taux de remplissage (indicateurs de performance, onglet Fonds) :")
    for col in ("TVPI", "DPI", "RVPI", "MoC (MoM)", "IRR (TRI)", "Quartile"):
        n = sum(0 if est_vide(v) else 1 for v in fonds[col])
        print(f"  - {col} : {100.0 * n / len(fonds):.1f} %  ({n}/{len(fonds)})")
    print()
    print("Lignes par onglet :", {k: len(v) for k, v in frames.items()})
    print("Cellules renseignées avant / après :",
          {k: f"{avant[k]} -> {apres[k]}" for k in frames})
    return 0


if __name__ == "__main__":
    sys.exit(main())
