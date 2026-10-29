#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
maj_radar_v29.py — Enrichissement CVC : Allianz X, AVP (ex-AXA Venture Partners), MAIF Avenir (Ternel)
=====================================================================================================

Entrée  : VC_Database_Standardisee_v28.xlsx + radar_refresh/cvc_2026_10.json
Sortie  : VC_Database_Standardisee_v29.xlsx

Sources : pages portefeuille officielles récupérées le 29/10/2026 par le workflow
recuperer_pages.yml (allianzx.com/ourcompanies, axavp.com/companies, maif-avenir.fr/portefeuille),
complétées par des recherches web citées dans le JSON (taille, AuM, tickets, MBO d'AVP).

- Allianz X, MAIF Avenir : champs du fonds complétés (valeur précédente conservée en Anomalies),
  participations manquantes ajoutées ;
- AVP : société de gestion absente du Radar → deux véhicules créés (Venture & Growth, qui porte
  les participations ; Late Growth), classée « VC indépendant » depuis le MBO de 2024 ;
- secteur attribué d'après la description publiée ; hors assurance → « FinTech (hors assurance) »
  ou « Hors périmètre assurance » (exclus du graphique Tendances) ; sorties marquées Exit = O ;
- doublons ignorés après normalisation du nom (liens_croises.norm) ; aucune date ni montant inventé.

Usage : python3 maj_radar_v29.py
"""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

from liens_croises import norm
from maj_radar_v27 import Journal, slugify

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v28.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v29.xlsx"
CAMPAGNE = BASE_DIR / "radar_refresh" / "cvc_2026_10.json"
DATE_MAJ = pd.Timestamp("2026-10-29")
ETIQUETTE = "[Enrichissement CVC 29/10/2026]"
STAGES_COLS = ["Pré-Seed", "Seed", "Pré-Série A", "Série A", "Série B", "Série C", "Série D"]


def main() -> int:
    c = json.loads(CAMPAGNE.read_text(encoding="utf-8"))
    feuilles = pd.read_excel(ENTREE, sheet_name=None)
    fonds = feuilles["Fonds"].astype(object)
    part = feuilles["Participations"].astype(object)
    fonds = fonds.where(pd.notna(fonds), None)
    part = part.where(pd.notna(part), None)
    j = Journal(feuilles["Anomalies"])

    def ligne_fonds(nom, veh):
        m = (fonds["Nom du fonds"] == nom) & (fonds["Véhicule"] == veh)
        return fonds.index[m][0] if m.any() else None

    # --- 1. Fonds existants : champs complétés ------------------------------------------
    for nom, f in c["fonds"].items():
        i = ligne_fonds(nom, f["vehicule"])
        if i is None:
            raise SystemExit(f"Fonds introuvable : {nom}")
        for champ, valeur in f["maj"].items():
            avant = fonds.at[i, champ]
            if avant == valeur:
                continue
            fonds.at[i, champ] = valeur
            j.add("Mise à jour web — champ fonds", nom, champ, None if avant is None else str(avant), str(valeur),
                  "Élevé" if avant is None else "Moyen", "Champ complété" if avant is None else "Valeur remplacée (ancienne conservée ici)",
                  f"{f['commentaire']} Sources : {' ; '.join(f['sources'])}", onglet="Fonds")
        fonds.at[i, "Source_AuM"] = f"{ETIQUETTE} {f['commentaire']} Sources : {' ; '.join(f['sources'])}"
        fonds.at[i, "Date_MAJ"] = DATE_MAJ

    # --- 2. Nouvelle société de gestion : AVP -------------------------------------------
    nf = c["nouveau_fonds"]
    for v in nf["vehicules"]:
        if ligne_fonds(nf["nom"], v["vehicule"]) is not None:
            continue
        ligne = {col: None for col in fonds.columns}
        ligne.update(nf["commun"])
        ligne.update({"Nom du fonds": nf["nom"], "Véhicule": v["vehicule"], "Stages pratiqués": v["Stages pratiqués"],
                      "Fonds_ID": f"{slugify(nf['nom'])}_{slugify(v['vehicule'])}", "Nb participations documentées": 0,
                      "Source_AuM": f"{ETIQUETTE} {nf['commentaire']} {v.get('commentaire_vehicule', '')} Sources : {' ; '.join(nf['sources'])}".strip(),
                      "Date_MAJ": DATE_MAJ})
        for col in STAGES_COLS:
            ligne[col] = 1 if col in v["Stages pratiqués"] else 0
        fonds = pd.concat([fonds, pd.DataFrame([ligne])], ignore_index=True)
    j.add("Mise à jour web — nouvelle société de gestion", nf["nom"], "Lignes Fonds", "AXA Venture Partners (absent du Radar)",
          ", ".join(v["vehicule"] for v in nf["vehicules"]), "Élevé", "Créée ; type « VC indépendant » depuis le MBO de 2024",
          nf["commentaire"], onglet="Fonds")
    j.add("Divergence de sources", nf["nom"], "AuM (M€)", "> 2 000 (communiqué du MBO, 2024)", "1 360 (déclaration ADEME, exercice 2024)",
          "Moyen", "Valeur réglementaire datée retenue", "Les deux chiffres sont cités dans Source_AuM.", onglet="Fonds")
    j.add("Règle d'interprétation", nf["nom"], "Type d'investisseur", "Ex-CVC d'AXA", "VC indépendant", "Élevé",
          "AXA reste investisseur mais ne contrôle plus la société depuis le MBO",
          "Suivi demandé avec les CVC (29/10/2026) ; reste dans le périmètre (siège européen, Paris).", onglet="Fonds")

    # --- 3. Participations ----------------------------------------------------------------
    porteur = {"Allianz X": "Allianz X", "MAIF Avenir": "MAIF Avenir",
               nf["nom"]: next(v["vehicule"] for v in nf["vehicules"] if v.get("porte_participations"))}
    existants = {(r["Nom du fonds"], norm(r["Start-up"])) for _, r in part.iterrows()}
    alias = {"nauto empiric earth": "nauto", "american well amwell": "american well"}
    bilan = {}
    for nom_fonds, liste in c["participations"].items():
        veh = porteur[nom_fonds]
        i = ligne_fonds(nom_fonds, veh)
        fid = fonds.at[i, "Fonds_ID"]
        ajout, deja = 0, []
        for p in liste:
            n = norm(p["startup"])
            if (nom_fonds, n) in existants or (nom_fonds, alias.get(n, n)) in existants:
                deja.append(p["startup"])
                continue
            ligne = {col: None for col in part.columns}
            ligne.update({
                "Fonds_ID": fid, "Nom du fonds": nom_fonds, "Véhicule": veh, "Start-up": p["startup"],
                "Secteur": p["secteur"], "Stage financé": p.get("stage"),
                "Exit (O/N)": "O" if p.get("exit") else "N",
                "Statut start-up (détail)": "Exit (sortie indiquée par le fonds)" if p.get("exit") else "Active",
                "Confiance": "Élevé",
                "Commentaires": f"{ETIQUETTE} Listée sur la page portefeuille officielle du fonds. Description publiée : {p['description']}",
                "Source(s)": f"{p['url']} (consulté le 29/10/2026)",
            })
            ligne.update(c.get("complements", {}).get(f"{nom_fonds}|{p['startup']}", {}))
            part = pd.concat([part, pd.DataFrame([ligne])], ignore_index=True)
            existants.add((nom_fonds, n))
            ajout += 1
        fonds.at[i, "Nb participations documentées"] = int((part["Fonds_ID"] == fid).sum())
        bilan[nom_fonds] = (ajout, deja)
        j.add("Mise à jour web — participations", nom_fonds, "Participations", f"{len(liste)} sociétés sur la page officielle",
              f"+{ajout} lignes", "Élevé", "Ajoutées ; doublons ignorés : " + (", ".join(deja) or "aucun"),
              "Hors assurance étiquetées « FinTech (hors assurance) » / « Hors périmètre assurance » : "
              + ", ".join(p["startup"] for p in liste if "hors" in (p["secteur"] or "").lower() and p["startup"] not in deja),
              onglet="Participations")

    with pd.ExcelWriter(SORTIE, engine="openpyxl") as writer:
        for nom, df in feuilles.items():
            df = {"Fonds": fonds, "Participations": part, "Anomalies": j.table()}.get(nom, df)
            df.to_excel(writer, sheet_name=nom, index=False)
    print(f"Classeur produit : {SORTIE.name} · Fonds {len(fonds)} · Participations {len(part)} · "
          f"Anomalies +{len(j.lignes)} (jusqu'à ANO-{j.suivant - 1:04d})")
    for k, (a, d) in bilan.items():
        print(f"  {k} : +{a} (déjà présentes : {', '.join(d) or 'aucune'})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
