#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fusionner_participations_ouverts.py — Fusionne le fichier de vérification web
================================================================================

Intègre Participations_Assurance_Vehicules_Ouverts.xlsx (recherche dédiée aux
véhicules "Ouvert"/statut vide, 116 lignes sourcées et notées en confiance)
dans VC_Database_Standardisee_v5.xlsx, produit VC_Database_Standardisee_v6.xlsx.

- Ajoute au Participations les start-ups absentes, avec de nouvelles colonnes
  (Taille du tour (M€), % de détention, Rôle du véhicule (détail), Statut
  start-up (détail), Confiance, Commentaires, Source(s)) pour ne pas perdre
  la granularité de la recherche.
- Corrige le doublon identifié (Astorya.vc / "Embrea" -> nom exact "Embea").
- Journalise une anomalie pour The Family (statut "Ouvert" en base alors que
  le fonds serait en gestion extinctive d'après la recherche).

Usage : python3 fusionner_participations_ouverts.py
"""

from __future__ import annotations

import math
import re
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE_BASE = BASE_DIR / "VC_Database_Standardisee_v5.xlsx"
ENTREE_FICHIER = Path(
    "/root/.claude/uploads/963315b8-930e-538e-a15f-800e2ff2b2d1/"
    "e0de41e7-Participations_Assurance_Vehicules_Ouverts.xlsx"
)
SORTIE = BASE_DIR / "VC_Database_Standardisee_v6.xlsx"

TAUX_CHANGE = {"€": 1.0, "$": 0.92, "£": 1.17, "CHF": 1.07, "AUD": 0.60}

RE_ANNEE = re.compile(r"(\d{4})")
RE_NOTE_BASE = re.compile(r"\s*\(base\s*:\s*[«\"][^»\"]*[»\"]\)\s*$")


def nettoyer_startup(nom: str) -> str:
    """Retire les notes de correction type ' (base : « Embrea »)' du nom."""
    return RE_NOTE_BASE.sub("", str(nom)).strip()


def vide(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    if isinstance(v, str) and v.strip().lower() in ("", "n.d.", "n.d", "nan"):
        return True
    return False


def convertir(montant, devise) -> float | None:
    if vide(montant):
        return None
    taux = TAUX_CHANGE.get(str(devise).strip()) if not vide(devise) else 1.0
    if taux is None:
        return None
    return round(float(montant) * taux, 2)


def annee(date_str) -> float | None:
    if vide(date_str):
        return None
    m = RE_ANNEE.search(str(date_str))
    return float(m.group(1)) if m else None


def positionnement(role: str) -> str | None:
    if vide(role):
        return None
    r = role.lower()
    if any(k in r for k in ("actionnaire de référence", "actionnaire principal", "investisseur unique")):
        return "Leader"
    if r.startswith("lead") or "lead" in r.split(" (")[0]:
        return "Leader"
    if r.startswith("co-lead") or r.startswith("co-investisseur") or "investisseur existant" in r or "incubateur" in r:
        return "Minoritaire"
    return None


def exit_on(statut: str) -> str | None:
    if vide(statut):
        return None
    s = statut.lower()
    if "exit" in s:
        return "O"
    if "arrêt" in s or "défaillance" in s:
        return "N"
    if "active" in s:
        return "N"
    return None


def main() -> int:
    fonds = pd.read_excel(ENTREE_BASE, sheet_name="Fonds")
    part = pd.read_excel(ENTREE_BASE, sheet_name="Participations")
    tours = pd.read_excel(ENTREE_BASE, sheet_name="Tours_de_table")
    dico = pd.read_excel(ENTREE_BASE, sheet_name="Dictionnaire")
    ano = pd.read_excel(ENTREE_BASE, sheet_name="Anomalies")

    detail = pd.read_excel(ENTREE_FICHIER, sheet_name="Détail participations")

    lookup = (
        fonds[["Nom du fonds", "Fonds_ID", "Véhicule"]]
        .drop_duplicates(subset=["Nom du fonds"])
        .set_index("Nom du fonds")
    )

    # Nouvelles colonnes, absentes du schéma v5 : on les ajoute (vides pour les
    # 10 lignes existantes) pour ne pas perdre la granularité de la recherche.
    nouvelles_colonnes = [
        "Taille du tour (M€)", "% de détention", "Rôle du véhicule (détail)",
        "Statut start-up (détail)", "Confiance", "Commentaires", "Source(s)",
    ]
    for c in nouvelles_colonnes:
        if c not in part.columns:
            part[c] = None

    # Correction du doublon identifié : Astorya.vc / "Embrea" -> nom exact "Embea"
    masque_embrea = (part["Véhicule"] == "Astorya.vc") & (part["Start-up"] == "Embrea")
    if masque_embrea.any():
        part.loc[masque_embrea, "Start-up"] = "Embea"
        part.loc[masque_embrea, "Confiance"] = "Moyen"
        part.loc[masque_embrea, "Source(s)"] = "https://www.teaserclub.com/investors/astorya-vc?sector=InsurTech"
        part.loc[masque_embrea, "Commentaires"] = "Nom exact « Embea » (corrigé depuis « Embrea ») ; vérification web 16/09/2026"
        deja_couvertes = {("Astorya.vc", "Embea")}
    else:
        deja_couvertes = set()

    nouvelles_lignes = []
    ignorees_sans_fonds = []

    for _, r in detail.iterrows():
        veh_fichier = r["Véhicule"]
        startup = nettoyer_startup(r["Start-up"])
        if (veh_fichier, startup) in deja_couvertes:
            continue
        nom_fonds = r["Société de gestion"]
        if nom_fonds not in lookup.index:
            ignorees_sans_fonds.append((nom_fonds, startup))
            continue
        info_fonds = lookup.loc[nom_fonds]

        nouvelles_lignes.append({
            "Fonds_ID": info_fonds["Fonds_ID"],
            "Nom du fonds": nom_fonds,
            "Véhicule": info_fonds["Véhicule"],
            "Start-up": startup,
            "Secteur": r.get("Segment"),
            "Stage financé": None if vide(r.get("Tour financé")) else r.get("Tour financé"),
            "Montant investi par le fonds (M€)": convertir(r.get("Ticket du véhicule (M)"), r.get("Devise ticket")),
            "Total levé par la start-up (M€)": None,
            "Date de création": None,
            "Date d’investissement": annee(r.get("Date (AAAA-MM)")),
            "Pays d’origine": None if vide(r.get("Pays")) else r.get("Pays"),
            "Nb pays d’implantation": None,
            "Positionnement (Leader/Minoritaire)": positionnement(r.get("Rôle du véhicule")),
            "Capital social (k€)": None,
            "CA (M€)": None,
            "Valorisation (M€)": None,
            "Nb fonds investisseurs": None,
            "Nb tours (fonds)": None,
            "Nb tours (total)": None,
            "Repositionnement": None,
            "Exit (O/N)": exit_on(r.get("Statut start-up")),
            "MoC exit": None,
            "MoEP exit": None,
            "Taille du tour (M€)": convertir(r.get("Taille du tour (M)"), r.get("Devise tour")),
            "% de détention": None if vide(r.get("% de détention")) else r.get("% de détention"),
            "Rôle du véhicule (détail)": None if vide(r.get("Rôle du véhicule")) else r.get("Rôle du véhicule"),
            "Statut start-up (détail)": None if vide(r.get("Statut start-up")) else r.get("Statut start-up"),
            "Confiance": None if vide(r.get("Confiance")) else r.get("Confiance"),
            "Commentaires": None if vide(r.get("Commentaires")) else r.get("Commentaires"),
            "Source(s)": None if vide(r.get("Source(s)")) else r.get("Source(s)"),
        })

    part_out = pd.concat([part, pd.DataFrame(nouvelles_lignes)], ignore_index=True)

    # Anomalie : The Family listé "Ouvert" en base alors que la recherche web
    # indique une gestion extinctive (litiges, plus d'investissements actifs).
    prochain_id = int(ano["Anomalie_ID"].str.extract(r"(\d+)")[0].astype(int).max()) + 1
    nouvelle_anomalie = {
        "Anomalie_ID": f"ANO-{prochain_id:04d}",
        "Type d'anomalie": "Statut fonds potentiellement obsolète",
        "Onglet source": "Fonds",
        "Table ou bloc source": "Participations_Assurance_Vehicules_Ouverts.xlsx (vérification web 16/09/2026)",
        "Ligne source": None,
        "Entité concernée": "The Family",
        "Champ concerné": "Statut",
        "Valeur source": "Ouvert",
        "Valeur retenue": "Ouvert (à confirmer)",
        "Candidats éventuels": "Fermé / gestion extinctive",
        "Score de similarité": None,
        "Niveau de confiance": "Moyen",
        "Traitement appliqué": "Conservé, signalé pour vérification",
        "Commentaire": (
            "Sifted 08/2024 et Crunchbase indiquent une gestion extinctive du "
            "portefeuille et des litiges avec le cofondateur O. Ammar — le statut "
            "« Ouvert » de la base n'est peut-être plus à jour."
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
    print(f"Participations avant : {len(part)}  →  après : {len(part_out)}  (+{len(nouvelles_lignes)})")
    print(f"Anomalies avant : {len(ano)}  →  après : {len(ano_out)}")
    if ignorees_sans_fonds:
        print(f"Lignes ignorées (fonds non trouvé dans la base) : {len(ignorees_sans_fonds)}")
        for nom, su in ignorees_sans_fonds:
            print("  -", nom, "/", su)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
