#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
maj_radar_v25.py — Emails des contacts clés (Apollo.io)
=======================================================

Entrée  : VC_Database_Standardisee_v24.xlsx + radar_refresh/emails_apollo_2026_10.json
Sortie  : VC_Database_Standardisee_v25.xlsx

Les 85 contacts clés de l'onglet Fonds ont été passés dans Apollo.io par R. De Sousa
(révélation manuelle dans l'application, le plan gratuit bloquant l'API). Le fichier
renvoyé est stocké tel quel dans radar_refresh/emails_apollo_2026_10.json avec le statut
Apollo de chaque ligne : verified (V), guessed (M, email non vérifié), unavailable (x).

Ce script :
- remplit Contact_0x_Email avec l'email renvoyé (jamais d'email déduit d'un format type) ;
- ajoute Contact_0x_Email_Statut (statut brut Apollo), lu par generer_site_vc.py pour le
  code couleur des fiches (vert = vérifié, gris = à confirmer, rouge = non disponible) ;
- applique la valeur à toutes les lignes (véhicules) du fonds portant ce contact ;
- trace l'opération dans Anomalies (une ligne de synthèse + une ligne par email non vérifié)
  et documente la nouvelle colonne dans le Dictionnaire.

Usage : python3 maj_radar_v25.py
"""
from __future__ import annotations

import json
import re
import unicodedata
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v24.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v25.xlsx"
EMAILS = BASE_DIR / "radar_refresh" / "emails_apollo_2026_10.json"
DATE = "15/10/2026"


def cle(texte) -> str:
    """Nom normalisé : sans accents, titres (Dr., CFA) ni ponctuation."""
    t = unicodedata.normalize("NFKD", str(texte or "")).encode("ascii", "ignore").decode().lower()
    t = re.sub(r"\b(dr|cfa)\b\.?", " ", t)
    return re.sub(r"[^a-z]+", " ", t).strip()


def main() -> int:
    data = json.loads(EMAILS.read_text(encoding="utf-8"))
    contacts = {(c["fonds"], cle(f"{c['prenom']} {c['nom']}")): c for c in data["contacts"]}

    feuilles = pd.read_excel(ENTREE, sheet_name=None)
    fonds = feuilles["Fonds"].astype(object)
    fonds = fonds.where(pd.notna(fonds), None)

    # Nouvelles colonnes de statut, placées juste après chaque colonne Email
    for n in ("01", "02"):
        col = f"Contact_{n}_Email_Statut"
        if col not in fonds.columns:
            fonds.insert(fonds.columns.get_loc(f"Contact_{n}_Email") + 1, col, None)

    trouves = set()
    for i, r in fonds.iterrows():
        for n in ("01", "02"):
            c = contacts.get((r["Nom du fonds"], cle(r[f"Contact_{n}_Nom"])))
            if not c:
                continue
            fonds.at[i, f"Contact_{n}_Email"] = c["email"]
            fonds.at[i, f"Contact_{n}_Email_Statut"] = c["statut_apollo"]
            trouves.add((c["fonds"], cle(f"{c['prenom']} {c['nom']}")))

    manquants = [c for k, c in contacts.items() if k not in trouves]
    if manquants:
        raise SystemExit(f"Contacts introuvables dans Fonds : {[(c['fonds'], c['nom']) for c in manquants]}")

    ano = feuilles["Anomalies"]
    suivant = int(ano["Anomalie_ID"].str.extract(r"(\d+)")[0].astype(int).max()) + 1
    compte = {s: sum(c["statut_apollo"] == s for c in data["contacts"]) for s in ("verified", "guessed", "unavailable")}
    lignes = [{
        "Anomalie_ID": f"ANO-{suivant:04d}", "Type d'anomalie": "Enrichissement emails (Apollo)",
        "Onglet source": "Fonds", "Table ou bloc source": f"Apollo.io — {DATE}", "Ligne source": None,
        "Entité concernée": f"{len(data['contacts'])} contacts clés", "Champ concerné": "Contact_0x_Email / Contact_0x_Email_Statut",
        "Valeur source": data["_meta"]["source"],
        "Valeur retenue": f"{compte['verified']} vérifiés, {compte['guessed']} non vérifiés, {compte['unavailable']} non disponibles",
        "Candidats éventuels": None, "Score de similarité": None, "Niveau de confiance": "Élevé",
        "Traitement appliqué": "Emails repris tels que renvoyés par Apollo, statut conservé ; aucun email reconstitué à partir d'un format type",
        "Commentaire": data["_meta"]["legende"],
    }]
    for c in data["contacts"]:
        if c["statut_apollo"] == "guessed":
            suivant += 1
            lignes.append({
                "Anomalie_ID": f"ANO-{suivant:04d}", "Type d'anomalie": "Valeur approximative conservée",
                "Onglet source": "Fonds", "Table ou bloc source": f"Apollo.io — {DATE}", "Ligne source": None,
                "Entité concernée": f"{c['fonds']} / {c['prenom']} {c['nom']}", "Champ concerné": "Contact_0x_Email",
                "Valeur source": c["email"], "Valeur retenue": c["email"], "Candidats éventuels": None,
                "Score de similarité": None, "Niveau de confiance": "Faible",
                "Traitement appliqué": "Email non vérifié par Apollo : affiché en gris « à confirmer »",
                "Commentaire": "À confirmer avant tout envoi (domaine ou orthographe à vérifier).",
            })
    feuilles["Anomalies"] = pd.concat([ano, pd.DataFrame(lignes)], ignore_index=True)

    dico = feuilles["Dictionnaire"]
    if "Contact_01_Email_Statut" not in set(dico["Champ"]):
        nouv = [{"Champ": f"Contact_{n}_Email_Statut", "Onglet": "Fonds", "Méthode de rapprochement": "Apollo.io",
                 "Niveau de confiance": "Élevé", "Type": "Texte", "Champ source ou calculé": "Champ source",
                 "Formule ou règle": "Statut brut Apollo : verified / guessed / unavailable",
                 "Commentaire": "Code couleur des fiches : vert = vérifié, gris = à confirmer, rouge = non disponible"}
                for n in ("01", "02")]
        feuilles["Dictionnaire"] = pd.concat([dico, pd.DataFrame(nouv)], ignore_index=True)

    feuilles["Fonds"] = fonds
    with pd.ExcelWriter(SORTIE, engine="openpyxl") as w:
        for nom in ("Fonds", "Participations", "Tours_de_table", "Dictionnaire", "Anomalies"):
            feuilles[nom].to_excel(w, sheet_name=nom, index=False)
    print(f"{SORTIE.name} : {compte} — anomalies jusqu'à ANO-{suivant:04d}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
