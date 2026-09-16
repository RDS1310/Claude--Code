#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generer_page_gvi.py — Génère la page « Stratégie & Pipe GVI »
================================================================

Lit gvi_data/GVI_Status_Report_Startups.xlsx (onglet Start-up) et produit
gvi_strategie_pipe.html : la stratégie d'investissement de GVI (synthèse
fournie par l'utilisateur) + le pipe de start-ups suivies.

Exclut délibérément les colonnes personnelles/confidentielles du fichier
source (contacts, numéros, mails, deck, notes internes de négociation) :
seules les données de niveau « fiche entreprise » sont publiées.

Usage : python3 generer_page_gvi.py
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "gvi_data" / "GVI_Status_Report_Startups.xlsx"
SORTIE = BASE_DIR / "gvi_strategie_pipe.html"


def est_vide(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


MARQUEURS_VIDES = {"n.a", "n.a.", "na", "n/a", "nd", "n.d.", "non disponible", "-", "--"}


def s(v):
    if est_vide(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.strftime("%d/%m/%Y")
    txt = str(v).strip()
    if txt.lower() in MARQUEURS_VIDES:
        return None
    if re.fullmatch(r"\d+\.0", txt):
        txt = txt[:-2]
    return txt or None


def statut_de(row) -> str:
    if str(row.get("Volt'terre \nGVI", "")).strip().lower() == "oui":
        return "portefeuille"
    interet = row.get("Intérêt GVI")
    if est_vide(interet):
        return "non_note"
    interet = str(interet).strip()
    return {"3": "priorite", "2": "suivi", "1": "sourcing", "0": "ecarte"}.get(interet, "a_evaluer")


def main() -> int:
    if not ENTREE.exists():
        print(f"ERREUR : {ENTREE} introuvable.")
        return 1

    su = pd.read_excel(ENTREE, "Start-up")
    su = su.astype(object).where(pd.notna(su), None)

    # Base des fonds déjà suivis dans Radar Insurtech VC (pour le rapprochement croisé)
    fonds_connus = set()
    radar = BASE_DIR / "radar_insurtech_vc.html"
    if radar.exists():
        m = re.search(r"window\.__DATA__ = (.*?);\s*</script>", radar.read_text(encoding="utf-8"), re.S)
        if m:
            data = json.loads(m.group(1))
            fonds_connus = {f["nom"].strip().lower() for f in data["fonds"]}

    def connu(nom: str) -> bool:
        if not nom:
            return False
        n = nom.strip().lower()
        return any(n in fc or fc in n for fc in fonds_connus)

    startups = []
    for _, r in su.iterrows():
        nom = s(r.get("Start-up"))
        if not nom:
            continue
        fonds_vc = [s(r.get(f"Fonds de VC {i}")) for i in range(1, 6)]
        fonds_vc = [{"nom": f, "connu": connu(f)} for f in fonds_vc if f]
        startups.append({
            "nom": nom,
            "statut": statut_de(r),
            "interet": s(r.get("Intérêt GVI")),
            "type": s(r.get("Type ")),
            "maturite": s(r.get("Maturité")),
            "creation": s(r.get("Création")),
            "pays": s(r.get("Pays\nd'origine")),
            "theme": s(r.get("Thème")),
            "secteur": s(r.get("Secteur")),
            "sousSecteur1": s(r.get("Sous-Secteur\n 1")),
            "sousSecteur2": s(r.get("Sous-secteur\n 2")),
            "pitch": s(r.get("Pitch")),
            "implantation": s(r.get("Implantation")),
            "concurrents": s(r.get("Concurrents")),
            "ca": s(r.get("CA")),
            "etp": s(r.get("ETP")),
            "totalLeve": s(r.get("Total  levé")),
            "derniereLevee": s(r.get("Dernière levée")),
            "fondsVC": fonds_vc,
            "voltInno": s(r.get("Volt'terre \nInno.")),
            "voltGVI": s(r.get("Volt'terre \nGVI")),
            "maj": s(r.get("MAJ")),
        })

    data = {"genere": pd.Timestamp.now().strftime("%d/%m/%Y"), "startups": startups}

    template = (BASE_DIR / "_gvi_template.html").read_text(encoding="utf-8")
    html = template.replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False))
    SORTIE.write_text(html, encoding="utf-8")

    print(f"Page produite : {SORTIE}")
    print(f"Start-ups     : {len(startups)}")
    from collections import Counter
    print("Répartition   :", Counter(x["statut"] for x in startups))
    return 0


if __name__ == "__main__":
    sys.exit(main())
