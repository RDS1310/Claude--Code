#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generer_page_gvi.py — Génère la page « Stratégie & Pipe GVI »
================================================================

Lit gvi_data/GVI_Status_Report_Startups.xlsx (onglet Start-up) et produit
gvi_strategie_pipe.html : la stratégie d'investissement de GVI (synthèse
fournie par l'utilisateur) + le pipe nominatif des start-ups suivies, avec
moteur de recherche (même logique que Radar Insurtech VC).

Exclut délibérément les colonnes de contact personnel du fichier source
(Contact Startup, Numéro, Mail, Deck) : seules les données de niveau
« fiche entreprise » sont publiées (nom, secteur, statut, financier,
fonds VC, pitch, contact Groupama interne, commentaires).

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


STATUT_LABEL = {
    "priorite": "Priorité", "suivi": "Suivi actif", "sourcing": "Sourcing",
    "portefeuille": "Portefeuille", "ecarte": "Écarté", "a_evaluer": "À évaluer",
    "non_note": "Non noté",
}
ORDRE_STATUT = ["portefeuille", "priorite", "suivi", "sourcing", "a_evaluer", "ecarte", "non_note"]
ORDRE_MATURITE = ["Pré-seed", "Pré-Seed", "Seed", "Pré-série A", "Pré-Série A", "Série A", "Série B",
                   "Série C", "Série D", "Série E", "Série F", "Scale up"]


def canon_maturite(v):
    """Normalise les variantes de casse/accent présentes dans le fichier source (ex. 'Serie B')."""
    if not v:
        return v
    for ref in ORDRE_MATURITE:
        if v.strip().lower().replace("é", "e") == ref.strip().lower().replace("é", "e"):
            return ref
    return v


def compte(vals, ordre=None):
    from collections import Counter
    c = Counter(v for v in vals if v)
    items = list(c.items())
    if ordre:
        rang = {v: i for i, v in enumerate(ordre)}
        items.sort(key=lambda kv: (rang.get(kv[0], len(ordre)), -kv[1]))
    else:
        items.sort(key=lambda kv: -kv[1])
    return [{"label": k, "n": v} for k, v in items]


def main() -> int:
    if not ENTREE.exists():
        print(f"ERREUR : {ENTREE} introuvable.")
        return 1

    su = pd.read_excel(ENTREE, "Start-up")
    su = su.astype(object).where(pd.notna(su), None)

    # Base des fonds déjà suivis dans Radar Insurtech VC (pour un comptage agrégé uniquement)
    fonds_connus = set()
    radar = BASE_DIR / "radar_insurtech_vc.html"
    if radar.exists():
        m = re.search(r"window\.__DATA__ = (.*?);\s*</script>", radar.read_text(encoding="utf-8"), re.S)
        if m:
            data_radar = json.loads(m.group(1))
            fonds_connus = {f["nom"].strip().lower() for f in data_radar["fonds"]}

    def connu(nom: str) -> bool:
        n = (nom or "").strip().lower()
        return bool(n) and any(n in fc or fc in n for fc in fonds_connus)

    statuts, secteurs, maturites, volt_inno, volt_gvi = [], [], [], [], []
    fonds_vc_mentions, fonds_vc_connus = 0, 0
    total = 0
    pipe = []
    for _, r in su.iterrows():
        nom = s(r.get("Start-up"))
        if not nom:
            continue
        total += 1
        statut_key = statut_de(r)
        statut_lbl = STATUT_LABEL[statut_key]
        secteur = s(r.get("Secteur"))
        maturite = canon_maturite(s(r.get("Maturité")))
        statuts.append(statut_lbl)
        secteurs.append(secteur)
        maturites.append(maturite)
        vi = s(r.get("Volt'terre \nInno.")) or "Non renseigné"
        vg = s(r.get("Volt'terre \nGVI")) or "Non renseigné"
        volt_inno.append(vi)
        volt_gvi.append(vg)

        fonds_vc = []
        for i in range(1, 6):
            f = s(r.get(f"Fonds de VC {i}"))
            if f:
                fonds_vc_mentions += 1
                est_connu = connu(f)
                if est_connu:
                    fonds_vc_connus += 1
                fonds_vc.append({"nom": f, "connu": est_connu})

        pipe.append({
            "nom": nom,
            "statutKey": statut_key,
            "statut": statut_lbl,
            "secteur": secteur,
            "sousSecteur1": s(r.get("Sous-Secteur\n 1")),
            "sousSecteur2": s(r.get("Sous-secteur\n 2")),
            "type": s(r.get("Type ")),
            "creation": s(r.get("Création")),
            "maturite": maturite,
            "pays": s(r.get("Pays\nd'origine")),
            "fondateurs": s(r.get("Fondateurs ")),
            "theme": s(r.get("Thème")),
            "pitch": s(r.get("Pitch")),
            "implantation": s(r.get("Implantation")),
            "concurrents": s(r.get("Concurrents")),
            "ca": s(r.get("CA")),
            "etp": s(r.get("ETP")),
            "totalLeve": s(r.get("Total  levé")),
            "derniereLevee": s(r.get("Dernière levée")),
            "fondsVc": fonds_vc,
            "contactGroupama": s(r.get("Contact Groupama")),
            "via": s(r.get("Via")),
            "commentaires": s(r.get("Commentaires")),
            "derniereRencontre": s(r.get("Denière rencontre")),
            "maj": s(r.get("MAJ")),
            "voltInno": vi,
            "voltGVI": vg,
        })

    pipe.sort(key=lambda p: (ORDRE_STATUT.index(p["statutKey"]), p["nom"].lower()))

    data = {
        "genere": pd.Timestamp.now().strftime("%d/%m/%Y"),
        "total": total,
        "parStatut": sorted(compte(statuts), key=lambda kv: ORDRE_STATUT.index(
            [k for k, v in STATUT_LABEL.items() if v == kv["label"]][0])),
        "parSecteur": compte(secteurs)[:12],
        "parMaturite": compte(maturites, ORDRE_MATURITE),
        "voltInno": compte(volt_inno),
        "voltGVI": compte(volt_gvi),
        "fondsVcMentions": fonds_vc_mentions,
        "fondsVcConnus": fonds_vc_connus,
        "pipe": pipe,
    }

    template = (BASE_DIR / "_gvi_template.html").read_text(encoding="utf-8")
    html = template.replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False))
    SORTIE.write_text(html, encoding="utf-8")

    print(f"Page produite (pipe nominatif + agrégats) : {SORTIE}")
    print(f"Start-ups comptabilisées : {total}")
    print("Par statut :", {x['label']: x['n'] for x in data['parStatut']})
    return 0


if __name__ == "__main__":
    sys.exit(main())
