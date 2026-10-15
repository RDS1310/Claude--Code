#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
maj_radar_v26.py — Contacts société et contacts à remplacer
===========================================================

Entrée  : VC_Database_Standardisee_v25.xlsx + radar_refresh/contacts_societe_2026_10.json
Sortie  : VC_Database_Standardisee_v26.xlsx

Suite de l'enrichissement Apollo (v25), pour les fonds dont aucun contact n'a d'email :
- « Contact société » : adresse générique (contact@, hello@…) ou formulaire publiés par le
  fonds sur son propre site, trouvés par recherche restreinte au domaine officiel. Nouvelles
  colonnes Contact_Societe_Email / _Libelle / _Formulaire / _Source, identiques sur toutes
  les lignes (véhicules) du fonds ;
- emails de partenaires publiés sur le site du fonds (Marc Menasé, Ralf Molitor, Will
  Thorne) : remplacent le statut « unavailable » d'Apollo, statut « published » (vert) ou
  « published_old » (document ancien, gris à confirmer) ;
- contacts à remplacer (départ, fonds fermé, marque dissoute) : nouvelle colonne
  Contact_0x_A_Remplacer avec le motif, affichée sur la fiche ;
- chaque décision est tracée dans Anomalies ; aucune adresse n'est reconstituée.

Usage : python3 maj_radar_v26.py
"""
from __future__ import annotations

import json
import re
import unicodedata
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v25.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v26.xlsx"
DONNEES = BASE_DIR / "radar_refresh" / "contacts_societe_2026_10.json"
DATE = "15/10/2026"
COLS_SOCIETE = ["Contact_Societe_Email", "Contact_Societe_Libelle", "Contact_Societe_Formulaire", "Contact_Societe_Source"]


def cle(texte) -> str:
    t = unicodedata.normalize("NFKD", str(texte or "")).encode("ascii", "ignore").decode().lower()
    t = re.sub(r"\b(dr|cfa)\b\.?", " ", t)
    return re.sub(r"[^a-z]+", " ", t).strip()


class Journal:
    def __init__(self, ano):
        self.ano, self.lignes = ano, []
        self.suivant = int(ano["Anomalie_ID"].str.extract(r"(\d+)")[0].astype(int).max()) + 1

    def add(self, type_, entite, champ, source, retenue, confiance, traitement, commentaire):
        self.lignes.append({
            "Anomalie_ID": f"ANO-{self.suivant:04d}", "Type d'anomalie": type_, "Onglet source": "Fonds",
            "Table ou bloc source": f"Contacts — {DATE}", "Ligne source": None, "Entité concernée": entite,
            "Champ concerné": champ, "Valeur source": source, "Valeur retenue": retenue,
            "Candidats éventuels": None, "Score de similarité": None, "Niveau de confiance": confiance,
            "Traitement appliqué": traitement, "Commentaire": commentaire})
        self.suivant += 1


def lignes_contact(fonds, nom_fonds, contact):
    """(index, n) des lignes du fonds où le contact apparaît en position n."""
    out = []
    for i, r in fonds[fonds["Nom du fonds"] == nom_fonds].iterrows():
        for n in ("01", "02"):
            if cle(r[f"Contact_{n}_Nom"]) == cle(contact):
                out.append((i, n))
    if not out:
        raise SystemExit(f"Contact introuvable : {nom_fonds} / {contact}")
    return out


def main() -> int:
    d = json.loads(DONNEES.read_text(encoding="utf-8"))
    feuilles = pd.read_excel(ENTREE, sheet_name=None)
    fonds = feuilles["Fonds"].astype(object)
    fonds = fonds.where(pd.notna(fonds), None)
    j = Journal(feuilles["Anomalies"])

    for n in ("01", "02"):
        col = f"Contact_{n}_A_Remplacer"
        if col not in fonds.columns:
            fonds.insert(fonds.columns.get_loc(f"Contact_{n}_Email_Statut") + 1, col, None)
    for col in COLS_SOCIETE:
        if col not in fonds.columns:
            fonds[col] = None

    # 1. Contacts société
    for c in d["contacts_societe"]:
        masque = fonds["Nom du fonds"] == c["fonds"]
        if not masque.any():
            raise SystemExit(f"Fonds inconnu : {c['fonds']}")
        for col, cle_json in zip(COLS_SOCIETE, ("email", "libelle", "formulaire", "source")):
            fonds.loc[masque, col] = c.get(cle_json)
    avec = [c for c in d["contacts_societe"] if c["email"]]
    j.add("Enrichissement contacts société", f"{len(d['contacts_societe'])} fonds", "Contact_Societe_*",
          d["_meta"]["methode"], f"{len(avec)} adresses génériques, {len(d['contacts_societe']) - len(avec)} formulaires",
          "Moyen", "Adresse ou formulaire publié sur le site officiel du fonds, source conservée",
          "Sans adresse publiée : " + " ; ".join(d["sans_adresse_publiee"]))

    # 2. Emails de partenaires publiés par le fonds
    for e in d["emails_partenaires_publies"]:
        for i, n in lignes_contact(fonds, e["fonds"], e["contact"]):
            fonds.at[i, f"Contact_{n}_Email"] = e["email"]
            fonds.at[i, f"Contact_{n}_Email_Statut"] = e["statut"]
        j.add("Mise à jour web — email publié", f"{e['fonds']} / {e['contact']}", "Contact_0x_Email",
              "Apollo : unavailable", e["email"], "Élevé" if e["statut"] == "published" else "Faible",
              "Email publié par le fonds lui-même : remplace l'absence de résultat Apollo",
              f"{e['commentaire']}. Source : {e['source']}")

    # 3. Contacts à remplacer
    for r in d["contacts_a_remplacer"]:
        for i, n in lignes_contact(fonds, r["fonds"], r["contact"]):
            fonds.at[i, f"Contact_{n}_A_Remplacer"] = r["motif"]
        j.add("Rapprochement non résolu", f"{r['fonds']} / {r['contact']}", "Contact_0x_Nom", r["contact"],
              "Conservé, signalé « à remplacer »", "Moyen", "Contact obsolète signalé sur la fiche", r["motif"])

    for a in d.get("a_verifier", []):
        j.add("Information non confirmée", f"{a['fonds']} / {a['contact']}", "Contact_0x_Email", None,
              "Inchangé", "Faible", "Non modifié — à vérifier", a["point"])

    feuilles["Fonds"] = fonds
    feuilles["Anomalies"] = pd.concat([j.ano, pd.DataFrame(j.lignes)], ignore_index=True)
    dico = feuilles["Dictionnaire"]
    nouv = [{"Champ": c, "Onglet": "Fonds", "Type": "Texte", "Champ source ou calculé": "Champ source",
             "Commentaire": "Contact générique publié par le fonds (site officiel)"} for c in COLS_SOCIETE] + \
           [{"Champ": f"Contact_{n}_A_Remplacer", "Onglet": "Fonds", "Type": "Texte",
             "Champ source ou calculé": "Champ source", "Commentaire": "Motif si le contact n'est plus pertinent"}
            for n in ("01", "02")]
    feuilles["Dictionnaire"] = pd.concat([dico, pd.DataFrame([x for x in nouv if x["Champ"] not in set(dico["Champ"])])],
                                         ignore_index=True)

    with pd.ExcelWriter(SORTIE, engine="openpyxl") as w:
        for nom in ("Fonds", "Participations", "Tours_de_table", "Dictionnaire", "Anomalies"):
            feuilles[nom].to_excel(w, sheet_name=nom, index=False)
    print(f"{SORTIE.name} : {len(avec)} emails société, {len(d['emails_partenaires_publies'])} emails partenaires, "
          f"{len(d['contacts_a_remplacer'])} contacts à remplacer — anomalies jusqu'à ANO-{j.suivant - 1:04d}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
