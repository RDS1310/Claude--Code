#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
maj_radar_v27.py — Portefeuille Portage (Sagard) complet
========================================================

Entrée  : VC_Database_Standardisee_v26.xlsx + radar_refresh/portefeuille_sagard_2026_10.json
Sortie  : VC_Database_Standardisee_v27.xlsx

Le portefeuille officiel de Sagard (sagard.com, chargé par l'API WordPress, récupéré le 20/10/2026
par le workflow .github/workflows/recuperer_pages.yml car le site est bloqué en session) attribue
chaque société à un fonds (« PV I » … « PV IV », « PCS I »). Ce script ajoute les participations
Portage manquantes au Radar :

- PV I/II/III/IV → véhicules existants Portag3 Venture I/II/III et Portage Ventures IV ;
- PCS I (Portage Capital Solutions, stratégie lancée en 2022) → nouveau véhicule ;
- une ligne par couple (véhicule, société) ; les lignes déjà présentes (Clark / PV II) sont laissées ;
- secteur attribué d'après la description publiée par Sagard ; les sociétés hors assurance sont
  gardées mais étiquetées « FinTech (hors assurance) » ou « Hors périmètre assurance », que le
  graphique Tendances exclut (cf. secteur_bucket dans generer_site_vc.py) ;
- pays seulement quand la description le dit ; ni date, ni montant, ni stade (non publiés) ;
- « previous » chez Sagard = la société n'est plus en portefeuille → Exit (O/N) = O.

Chaque décision de périmètre est journalisée dans l'onglet Anomalies.

Usage : python3 maj_radar_v27.py
"""
from __future__ import annotations

import json
import re
import unicodedata
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v26.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v27.xlsx"
CAMPAGNE = BASE_DIR / "radar_refresh" / "portefeuille_sagard_2026_10.json"

DATE_MAJ = pd.Timestamp("2026-10-20")
ETIQUETTE = "[Portefeuille Sagard 20/10/2026]"
FONDS = "Portage Venture"


def slugify(value) -> str:
    """Même règle que restructuration_vc.slugify (Fonds_ID stables)."""
    text = str(value or "").strip()
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode().lower()
    text = text.replace("&", " et ")
    text = re.sub(r"[^a-z0-9]+", " ", text).strip()
    return re.sub(r"\s+", "-", text)


class Journal:
    """Lignes d'anomalies numérotées à la suite de l'existant."""

    def __init__(self, ano: pd.DataFrame):
        self.ano = ano
        self.suivant = int(ano["Anomalie_ID"].str.extract(r"(\d+)")[0].astype(int).max()) + 1
        self.lignes: list[dict] = []

    def add(self, type_, entite, champ, source, retenue, confiance, traitement, commentaire,
            onglet="Participations"):
        self.lignes.append({
            "Anomalie_ID": f"ANO-{self.suivant:04d}", "Type d'anomalie": type_,
            "Onglet source": onglet, "Table ou bloc source": f"Portefeuille Sagard {DATE_MAJ:%d/%m/%Y}",
            "Ligne source": None, "Entité concernée": entite, "Champ concerné": champ,
            "Valeur source": source, "Valeur retenue": retenue, "Candidats éventuels": None,
            "Score de similarité": None, "Niveau de confiance": confiance,
            "Traitement appliqué": traitement, "Commentaire": commentaire,
        })
        self.suivant += 1

    def table(self) -> pd.DataFrame:
        return pd.concat([self.ano, pd.DataFrame(self.lignes)], ignore_index=True)


def main() -> int:
    c = json.loads(CAMPAGNE.read_text(encoding="utf-8"))
    feuilles = pd.read_excel(ENTREE, sheet_name=None)
    fonds = feuilles["Fonds"].astype(object)
    part = feuilles["Participations"].astype(object)
    fonds = fonds.where(pd.notna(fonds), None)
    part = part.where(pd.notna(part), None)
    j = Journal(feuilles["Anomalies"])

    # --- 1. Nouveau véhicule Portage Capital Solutions I ---------------------------------
    nv = c["nouveau_vehicule"]
    if not ((fonds["Nom du fonds"] == FONDS) & (fonds["Véhicule"] == nv["vehicule"])).any():
        modele = fonds[(fonds["Nom du fonds"] == FONDS) & (fonds["Véhicule"] == "Portage Ventures IV")].iloc[0]
        ligne = {col: None for col in fonds.columns}
        ligne.update({
            "Nom du fonds": FONDS, "Véhicule": nv["vehicule"], "Géographie": "International",
            "Stratégie": "FinTech ; InsurTech", "Type d'investisseur": modele.get("Type d'investisseur"),
            "Site web": nv["site"], "Nb participations documentées": 0,
            "Fonds_ID": f"{slugify(FONDS)}_{slugify(nv['vehicule'])}",
            "Source_AuM": f"{ETIQUETTE} {nv['commentaire']} Sources : {' ; '.join(nv['sources'])}",
            "Date_MAJ": DATE_MAJ,
        })
        fonds = pd.concat([fonds, pd.DataFrame([ligne])], ignore_index=True)
        j.add("Mise à jour web — nouveau véhicule", f"{FONDS} / {nv['vehicule']}", "Ligne Fonds", "PCS I (filtre Sagard)",
              nv["vehicule"], "Élevé", "Véhicule créé (taille, millésime, stades non publiés : laissés vides)",
              nv["commentaire"], onglet="Fonds")

    # --- 2. Participations ---------------------------------------------------------------
    ajoutees, deja = 0, []
    for p in c["participations"]:
        masque = (fonds["Nom du fonds"] == FONDS) & (fonds["Véhicule"] == p["vehicule"])
        if masque.sum() != 1:
            raise SystemExit(f"Véhicule introuvable : {p['vehicule']}")
        fid = fonds.loc[masque, "Fonds_ID"].iloc[0]
        if ((part["Fonds_ID"] == fid) & (part["Start-up"].astype(str).str.lower() == p["startup"].lower())).any():
            deja.append(f"{p['startup']} ({p['vehicule']})")
            continue
        sortie = p["statut_sagard"] == "previous"
        desc = p["description_sagard"] or "aucune description publiée"
        ligne = {col: None for col in part.columns}
        ligne.update({
            "Fonds_ID": fid, "Nom du fonds": FONDS, "Véhicule": p["vehicule"], "Start-up": p["startup"],
            "Secteur": p["secteur"], "Pays d’origine": p["pays"], "Exit (O/N)": "O" if sortie else "N",
            "Statut start-up (détail)": "Exit (sortie du portefeuille Sagard)" if sortie else "Active",
            "Confiance": "Élevé",
            "Commentaires": f"{ETIQUETTE} Listée par Sagard dans le fonds « {p['etiquette_sagard'].upper()} » "
                            f"({'ancienne participation' if sortie else 'participation en cours'}). "
                            f"Description Sagard : {desc} Date, montant et stade non publiés.",
            "Source(s)": f"Sagard — portefeuille officiel ({p['url']}), consulté le 20/10/2026",
        })
        part = pd.concat([part, pd.DataFrame([ligne])], ignore_index=True)
        i = fonds.index[masque][0]
        fonds.at[i, "Nb participations documentées"] = int(fonds.at[i, "Nb participations documentées"] or 0) + 1
        fonds.at[i, "Date_MAJ"] = DATE_MAJ
        ajoutees += 1

    # --- 3. Décisions journalisées -----------------------------------------------------
    parts = c["participations"]
    def noms(cond):
        return ", ".join(sorted({p["startup"] for p in parts if cond(p)}))
    j.add("Périmètre — participations hors assurance", f"{FONDS} (portefeuille Sagard)", "Secteur",
          "Sociétés fintech non assurance", "FinTech (hors assurance)", "Élevé",
          "Ajoutées mais exclues du graphique Tendances", noms(lambda p: p["secteur"] == "FinTech (hors assurance)"))
    j.add("Périmètre — participations hors assurance et hors fintech", f"{FONDS} (portefeuille Sagard)", "Secteur",
          "Description Sagard sans lien avec l'assurance ni les services financiers", "Hors périmètre assurance", "Élevé",
          "Ajoutées mais exclues du graphique Tendances", noms(lambda p: p["secteur"] == "Hors périmètre assurance"))
    j.add("Information manquante", f"{FONDS} (portefeuille Sagard)", "Secteur", "Aucune description publiée par Sagard",
          "Vide", "Moyen", "Secteur laissé vide (non trouvé)", noms(lambda p: p["secteur"] is None))
    j.add("Règle d'interprétation", f"{FONDS} (portefeuille Sagard)", "Exit (O/N)", "Statut Sagard « previous »",
          "Exit = O ; statut « Exit (sortie du portefeuille Sagard) »", "Moyen",
          "Appliqué ; le mode de sortie (cession, IPO, radiation) n'est pas publié",
          noms(lambda p: p["statut_sagard"] == "previous"))
    j.add("Source partielle", f"{FONDS} (portefeuille Sagard)", "Participations",
          c["avertissement_sagard"], "Lignes existantes conservées", "Élevé",
          "Aucune suppression ; doublons ignorés : " + (", ".join(deja) or "aucun"),
          "Pays renseigné seulement si la description Sagard le dit ; dates, montants et stades non publiés.")

    with pd.ExcelWriter(SORTIE, engine="openpyxl") as writer:
        for nom, df in feuilles.items():
            df = {"Fonds": fonds, "Participations": part, "Anomalies": j.table()}.get(nom, df)
            df.to_excel(writer, sheet_name=nom, index=False)

    print(f"Classeur produit : {SORTIE.name}")
    print(f"Fonds : {len(fonds)} lignes · Participations : {len(part)} (+{ajoutees}) · "
          f"Anomalies : +{len(j.lignes)} (jusqu'à ANO-{j.suivant - 1:04d}) · déjà présentes : {', '.join(deja) or 'aucune'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
