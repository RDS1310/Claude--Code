#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generer_site_vc.py — Génère la page HTML « Radar Insurtech VC » à partir du classeur
======================================================================================

Détecte automatiquement le dernier classeur VC_Database_Standardisee_v*.xlsx du dossier,
en extrait les onglets Fonds et Participations, et régénère radar_insurtech_vc.html : une
page autonome (données embarquées en JSON, pas de backend) avec navigation par société de
gestion, fiche détaillée par véhicule, vue Start-ups et comparateur.

À rejouer à chaque nouvelle version du classeur (v6, v7, ...) : le fichier produit est
toujours à la même adresse, son contenu change avec les données sources. La page elle-même
propose aussi un rechargement direct d'un classeur .xlsx dans le navigateur (bouton
« Recharger l'Excel », via SheetJS), pour les mises à jour ponctuelles sans repasser par
Python.

Usage : python3 generer_site_vc.py
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
SORTIE = BASE_DIR / "radar_insurtech_vc.html"

STAGES = ["Pré-Seed", "Seed", "Pré-Série A", "Série A", "Série B", "Série C", "Série D"]
STAGE_COLS = {
    "Pré-Seed": "Pré-Seed", "Seed": "Seed", "Pré-Série A": "Pré-Série A",
    "Série A": "Série A", "Série B": "Série B", "Série C": "Série C", "Série D": "Série D",
}


def localiser_dernier_classeur() -> Path:
    candidats = sorted(
        BASE_DIR.glob("VC_Database_Standardisee_v*.xlsx"),
        key=lambda p: int(re.search(r"_v(\d+)\.xlsx$", p.name).group(1)),
    )
    if not candidats:
        raise FileNotFoundError("Aucun classeur VC_Database_Standardisee_v*.xlsx trouvé.")
    return candidats[-1]


def est_vide(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


def j(v):
    """Convertit une valeur pandas en valeur JSON-safe (None, str, int, float)."""
    if est_vide(v):
        return None
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, (int, float, str)):
        return v
    return str(v)


def liste(v) -> list[str]:
    if est_vide(v):
        return []
    return [x.strip() for x in str(v).split(";") if x.strip()]


def construire_fonds(fonds: pd.DataFrame) -> list[dict]:
    out = []
    for _, r in fonds.iterrows():
        deals = []
        for n in ("01", "02", "03"):
            nom = r.get(f"Deal_{n}_Nom")
            if est_vide(nom):
                continue
            deals.append({
                "nom": j(nom),
                "secteur": j(r.get(f"Deal_{n}_Secteur")),
                "stage": j(r.get(f"Deal_{n}_Stage")),
                "montant": j(r.get(f"Deal_{n}_Montant_investi_M€")),
                "date": j(r.get(f"Deal_{n}_Date_invest")),
            })
        out.append({
            "fid": j(r["Fonds_ID"]),
            "nom": j(r["Nom du fonds"]),
            "vehicule": j(r["Véhicule"]),
            "site": j(r.get("Site web")),
            "statut": j(r.get("Statut")),
            "phase": j(r.get("Phase")),
            "millesime": j(r.get("Millésime")),
            "geo": liste(r.get("Géographie")),
            "strategie": liste(r.get("Stratégie")),
            "stages": {s: bool(r.get(STAGE_COLS[s]) == 1.0 or r.get(STAGE_COLS[s]) == 1)
                       for s in STAGES},
            "aum": j(r.get("AuM (M€)")),
            "leve": j(r.get("Montant levé (M€)")),
            "alloue": j(r.get("Montant alloué (M€)")),
            "ticketMin": j(r.get("Ticket min (M€)")),
            "ticketMax": j(r.get("Ticket max (M€)")),
            "participations": j(r.get("Nb participations (déclaré)")),
            "exits": j(r.get("Nb exits")),
            "insurtech": j(r.get("Nb InsurTech")),
            "tvpi": j(r.get("TVPI")),
            "dpi": j(r.get("DPI")),
            "rvpi": j(r.get("RVPI")),
            "moc": j(r.get("MoC (MoM)")),
            "irr": j(r.get("IRR (TRI)")),
            "quartile": j(r.get("Quartile")),
            "deals": deals,
            "source": j(r.get("Source_AuM")),
            "maj": j(r.get("Date_MAJ")),
            "typeInvestisseur": j(r.get("Type d'investisseur")) or "VC indépendant",
            "societeMere": j(r.get("Société mère (si CVC)")),
        })
    return out


SUFFIXES_FONDS = [
    "capital partners", "venture capital", "asset management",
    "capital", "ventures", "venture", "partners", "management",
    "investissement", "investments", "group", "vc",
]
CORE_MIN_LEN = 5
CORE_DENYLIST = {
    "index", "start", "open", "good", "step", "cash", "seed", "team", "next", "deal",
    "group", "asset", "trust", "bank", "life", "home", "insurtech", "assurtech",
    "fintech", "tech", "digital", "innovation", "capital", "ventures",
}


def coeur_nom(nom: str) -> str:
    """Réduit un nom de fonds à son cœur distinctif (retire les suffixes génériques
    type 'Capital'/'Ventures'/'VC') pour le rapprochement avec les actus."""
    n = nom.strip()
    changed = True
    while changed:
        changed = False
        low = n.lower()
        for suf in SUFFIXES_FONDS:
            if low.endswith(" " + suf):
                n = n[: -(len(suf) + 1)].strip()
                changed = True
                break
    return n


def rattacher_actus(fonds: list[dict]) -> None:
    """Croise chaque fonds avec veille_data/entries.json et ma_data/deals.json : si le
    cœur du nom du fonds apparaît dans une actu, l'ajoute à f['actus'] (lien cliquable
    vers la source, affiché sur la fiche du fonds dans Radar Insurtech VC)."""
    def charger(rel):
        p = BASE_DIR / rel
        if not p.exists():
            return None
        return json.loads(p.read_text(encoding="utf-8"))

    veille = charger("veille_data/entries.json") or {"entries": []}
    ma = charger("ma_data/deals.json") or {"confirmees": [], "rumeurs": []}

    noms = sorted({f["nom"] for f in fonds})
    coeurs = {nom: coeur_nom(nom) for nom in noms}
    actus_par_nom: dict[str, list[dict]] = {nom: [] for nom in noms}

    def matche(nom: str, texte: str) -> bool:
        texte_low = texte.lower()
        if nom.lower() in texte_low:
            return True
        coeur = coeurs[nom]
        if len(coeur) < CORE_MIN_LEN or coeur.lower() in CORE_DENYLIST:
            return False
        return coeur.lower() in texte_low

    URL_VEILLE = "https://claude.ai/artifact/5FXTpkCo6KvGu6YyeYAHiw"
    URL_MA = "https://claude.ai/artifact/BPSeTnKnWY2PkAGcH5s2ai"

    for e in veille.get("entries", []):
        texte = f"{e.get('titre','')} {e.get('resume','')}"
        for nom in noms:
            if matche(nom, texte):
                actus_par_nom[nom].append({
                    "type": "veille", "titre": e.get("titre"), "date": e.get("date"),
                    "url": e.get("url"), "source": e.get("source"), "page": URL_VEILLE,
                })

    for d in ma.get("confirmees", []):
        texte = f"{d.get('acquereur','')} {d.get('cible','')} {d.get('resume','')}"
        for nom in noms:
            if matche(nom, texte):
                actus_par_nom[nom].append({
                    "type": "ma", "titre": f"{d.get('acquereur')} → {d.get('cible')}",
                    "date": d.get("date"), "url": d.get("url"), "source": d.get("source"), "page": URL_MA,
                })

    for d in ma.get("rumeurs", []):
        texte = f"{d.get('acteurs_pressentis','')} {d.get('cible','')} {d.get('resume','')}"
        for nom in noms:
            if matche(nom, texte):
                actus_par_nom[nom].append({
                    "type": "ma", "titre": f"{d.get('acteurs_pressentis')} → {d.get('cible')} (rumeur)",
                    "date": d.get("date"), "url": d.get("url"), "source": d.get("source"), "page": URL_MA,
                })

    for f in fonds:
        f["actus"] = actus_par_nom.get(f["nom"], [])


def construire_startups(part: pd.DataFrame) -> list[dict]:
    out = []
    for _, r in part.iterrows():
        out.append({
            "fid": j(r.get("Fonds_ID")),
            "fondsNom": j(r.get("Nom du fonds")),
            "vehicule": j(r.get("Véhicule")),
            "nom": j(r.get("Start-up")),
            "secteur": j(r.get("Secteur")),
            "stage": j(r.get("Stage financé")),
            "montantFonds": j(r.get("Montant investi par le fonds (M€)")),
            "totalLeve": j(r.get("Total levé par la start-up (M€)")),
            "dateCreation": j(r.get("Date de création")),
            "dateInvest": j(r.get("Date d’investissement")),
            "pays": j(r.get("Pays d’origine")),
            "nbPays": j(r.get("Nb pays d’implantation")),
            "position": j(r.get("Positionnement (Leader/Minoritaire)")),
            "capitalSocial": j(r.get("Capital social (k€)")),
            "ca": j(r.get("CA (M€)")),
            "valorisation": j(r.get("Valorisation (M€)")),
            "nbFonds": j(r.get("Nb fonds investisseurs")),
            "nbToursFonds": j(r.get("Nb tours (fonds)")),
            "nbToursTotal": j(r.get("Nb tours (total)")),
            "repositionnement": j(r.get("Repositionnement")),
            "exit": j(r.get("Exit (O/N)")),
            "mocExit": j(r.get("MoC exit")),
            "moepExit": j(r.get("MoEP exit")),
            "tailleDuTour": j(r.get("Taille du tour (M€)")),
            "pctDetention": j(r.get("% de détention")),
            "roleDetail": j(r.get("Rôle du véhicule (détail)")),
            "statutDetail": j(r.get("Statut start-up (détail)")),
            "confiance": j(r.get("Confiance")),
            "commentaires": j(r.get("Commentaires")),
            "sources": j(r.get("Source(s)")),
        })
    return out


def main() -> int:
    entree = localiser_dernier_classeur()
    version = re.search(r"_v(\d+)\.xlsx$", entree.name).group(1)
    print(f"Classeur source : {entree.name}")

    fonds_df = pd.read_excel(entree, "Fonds")
    fonds_df = fonds_df.astype(object).where(pd.notna(fonds_df), None)
    part_df = pd.read_excel(entree, "Participations")
    part_df = part_df.astype(object).where(pd.notna(part_df), None)
    ano_df = pd.read_excel(entree, "Anomalies")

    fonds = construire_fonds(fonds_df)
    rattacher_actus(fonds)
    startups = construire_startups(part_df)

    data = {
        "version": version,
        "genere": pd.Timestamp.now().strftime("%d/%m/%Y"),
        "nbAnomalies": int(len(ano_df)),
        "fonds": fonds,
        "startups": startups,
    }

    template = (BASE_DIR / "_site_template.html").read_text(encoding="utf-8")
    html = template.replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False))
    SORTIE.write_text(html, encoding="utf-8")

    print(f"Page produite   : {SORTIE}")
    print(f"Véhicules       : {len(fonds)}")
    print(f"Sociétés (Participations) : {len(startups)}")
    print(f"Sociétés de gestion : {len(set(f['nom'] for f in fonds))}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
