#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
maj_radar_v24.py — Rafraîchissement complet du Radar (campagne web d'octobre 2026)
==================================================================================

Entrée  : VC_Database_Standardisee_v23.xlsx + radar_refresh/maj_2026_10.json
Sortie  : VC_Database_Standardisee_v24.xlsx (+ actualités ajoutées à veille_data/entries.json)

La campagne (02/10/2026) a passé en revue les 75 sociétés de gestion exploitables de
l'onglet Fonds (recherche WebSearch par société ; WebFetch bloqué par le proxy réseau, d'où
des sources citées d'après les extraits de recherche). Les résultats vérifiés sont stockés
dans radar_refresh/maj_2026_10.json — ce script se contente de les appliquer :

- nouveaux véhicules (closings 2025-2026 absents de la base) → nouvelles lignes Fonds ;
- changements de statut documentés (Munich Re Ventures) → cellule mise à jour, ancienne
  valeur conservée dans l'onglet Anomalies ;
- nouvelles participations assurance → nouvelles lignes Participations ;
- mise à jour d'attributs de start-up (valorisation d'Alan) sur toutes ses lignes ;
- pistes écartées et points à vérifier → journalisés en Anomalies, sans toucher aux données ;
- actualités datées (≥ 01/06/2026) → veille_data/entries.json (dédoublonnage sur l'URL),
  ce qui alimente aussi les fiches fonds du Radar via rattacher_actus().

Règles : aucune estimation (un champ non publié reste vide), devises converties aux taux
imposés ($ × 0,92), confiance Élevé = ≥ 2 sources concordantes, chaque modification tracée
(Source_AuM / Date_MAJ / ligne Anomalies). Le script est idempotent sur entries.json.

Usage : python3 maj_radar_v24.py
"""
from __future__ import annotations

import json
import re
import unicodedata
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v23.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v24.xlsx"
CAMPAGNE = BASE_DIR / "radar_refresh" / "maj_2026_10.json"
VEILLE = BASE_DIR / "veille_data" / "entries.json"

DATE_MAJ = pd.Timestamp("2026-10-02")
ETIQUETTE = "[Rafraîchissement Radar 02/10/2026]"
STAGES = ["Pré-Seed", "Seed", "Pré-Série A", "Série A", "Série B", "Série C", "Série D"]

# Champs qualitatifs des nouveaux véhicules, tirés des sources citées dans le JSON
# (vocabulaires fermés du Dictionnaire ; un champ non documenté reste vide).
QUALIF_VEHICULES = {
    "Portage Ventures IV": {"Géographie": "International", "Stratégie": "FinTech ; InsurTech"},
    "Serena IV": {"Géographie": "Europe", "Stratégie": "Innovation ; ESG"},
    "Accel London IX": {"Géographie": "Europe", "Stratégie": "Généraliste"},
    "ISAI Venture IV": {"Géographie": "Europe ; Amérique", "Stratégie": "Généraliste"},
    "WSC IV": {"Géographie": "International", "Stratégie": "Généraliste",
               "Stages pratiqués": "Série A ; Série B"},
    "Samaipata III": {"Géographie": "Europe", "Stratégie": "Digital ; Innovation"},
}


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
            onglet="Fonds"):
        self.lignes.append({
            "Anomalie_ID": f"ANO-{self.suivant:04d}", "Type d'anomalie": type_,
            "Onglet source": onglet, "Table ou bloc source": f"Campagne web {DATE_MAJ:%d/%m/%Y}",
            "Ligne source": None, "Entité concernée": entite, "Champ concerné": champ,
            "Valeur source": source, "Valeur retenue": retenue, "Candidats éventuels": None,
            "Score de similarité": None, "Niveau de confiance": confiance,
            "Traitement appliqué": traitement, "Commentaire": commentaire,
        })
        self.suivant += 1

    def table(self) -> pd.DataFrame:
        return pd.concat([self.ano, pd.DataFrame(self.lignes)], ignore_index=True)


def sources_txt(sources: list[str]) -> str:
    return " ; ".join(sources)


def main() -> int:
    c = json.loads(CAMPAGNE.read_text(encoding="utf-8"))
    fonds = pd.read_excel(ENTREE, sheet_name="Fonds")
    part = pd.read_excel(ENTREE, sheet_name="Participations")
    tours = pd.read_excel(ENTREE, sheet_name="Tours_de_table")
    dico = pd.read_excel(ENTREE, sheet_name="Dictionnaire")
    j = Journal(pd.read_excel(ENTREE, sheet_name="Anomalies"))
    fonds = fonds.astype(object).where(pd.notna(fonds), None)
    part = part.astype(object).where(pd.notna(part), None)

    # --- 1. Nouveaux véhicules ---------------------------------------------------------
    for v in c["nouveaux_vehicules"]:
        modele = fonds[fonds["Nom du fonds"] == v["fonds"]]
        if modele.empty:
            raise SystemExit(f"Société de gestion inconnue : {v['fonds']}")
        if (modele["Véhicule"] == v["vehicule"]).any():
            print(f"Déjà présent, ignoré : {v['fonds']} / {v['vehicule']}")
            continue
        ref = modele.iloc[-1]
        ligne = {col: None for col in fonds.columns}
        ligne.update({
            "Nom du fonds": v["fonds"], "Véhicule": v["vehicule"],
            "Montant levé (M€)": v["montant_leve_meur"], "Millésime": v["millesime"],
            "Statut": v["statut"], "Phase": v["phase"],
            "Ticket min (M€)": v.get("ticket_min_meur"), "Ticket max (M€)": v.get("ticket_max_meur"),
            "Fonds_ID": f"{slugify(v['fonds'])}_{slugify(v['vehicule'])}",
            "Source_AuM": f"{ETIQUETTE} {v['commentaire']} Sources : {sources_txt(v['sources'])}",
            "Date_MAJ": DATE_MAJ,
            # Attributs de la société de gestion, identiques pour tous ses véhicules
            "Site web": ref["Site web"], "Type d'investisseur": ref["Type d'investisseur"],
            "Société mère (si CVC)": ref["Société mère (si CVC)"],
            "Nb participations documentées": 0,
        })
        ligne.update(QUALIF_VEHICULES.get(v["vehicule"], {}))
        stages = [s.strip() for s in (ligne.get("Stages pratiqués") or "").split(";") if s.strip()]
        if stages:
            for s in STAGES:
                ligne[s] = 1 if s in stages else 0
        pos = modele.index[-1] + 1
        fonds = pd.concat([fonds.iloc[:pos], pd.DataFrame([ligne]), fonds.iloc[pos:]], ignore_index=True)
        j.add("Mise à jour web — nouveau véhicule", f"{v['fonds']} / {v['vehicule']}",
              "Ligne Fonds", v["montant_leve_devise"], f"{v['montant_leve_meur']} M€",
              v["confiance"], "Véhicule ajouté (closing absent de la base)",
              f"{v['commentaire']} Sources : {sources_txt(v['sources'])}")

    # --- 2. Mises à jour de champs (changement de statut documenté) ---------------------
    for m in c["maj_champs"]:
        masque = (fonds["Nom du fonds"] == m["fonds"]) & (fonds["Véhicule"] == m["vehicule"])
        if masque.sum() != 1:
            raise SystemExit(f"Véhicule introuvable : {m['fonds']} / {m['vehicule']}")
        i = fonds.index[masque][0]
        for champ, valeur in m["champs"].items():
            ancien = fonds.at[i, champ]
            fonds.at[i, champ] = valeur
            j.add("Mise à jour web — changement de statut", f"{m['fonds']} / {m['vehicule']}",
                  champ, ancien, valeur, m["confiance"],
                  "Valeur existante remplacée (fait daté et sourcé ; ancienne valeur conservée ici)",
                  f"{m['commentaire']} Sources : {sources_txt(m['sources'])}")
        fonds.at[i, "Source_AuM"] = f"{fonds.at[i, 'Source_AuM'] or ''} | {ETIQUETTE} {m['commentaire']}".strip(" |")
        fonds.at[i, "Date_MAJ"] = DATE_MAJ

    # --- 3. Nouvelles participations ------------------------------------------------------
    for p in c["participations"]:
        masque = (fonds["Nom du fonds"] == p["fonds"]) & (fonds["Véhicule"] == p["vehicule"])
        if masque.sum() != 1:
            raise SystemExit(f"Véhicule introuvable : {p['fonds']} / {p['vehicule']}")
        fid = fonds.loc[masque, "Fonds_ID"].iloc[0]
        if ((part["Fonds_ID"] == fid) & (part["Start-up"] == p["startup"])).any():
            print(f"Participation déjà présente, ignorée : {p['fonds']} / {p['startup']}")
            continue
        ligne = {col: None for col in part.columns}
        ligne.update({
            "Fonds_ID": fid, "Nom du fonds": p["fonds"], "Véhicule": p["vehicule"],
            "Start-up": p["startup"], "Secteur": p["secteur"], "Stage financé": p["stage"],
            "Date d’investissement": p["date_invest"], "Pays d’origine": p["pays"],
            "Positionnement (Leader/Minoritaire)": "Leader" if p["role"].startswith("Leader") else "Minoritaire",
            "Exit (O/N)": "N", "Taille du tour (M€)": p["taille_tour_meur"],
            "Rôle du véhicule (détail)": p["role"], "Statut start-up (détail)": p["statut"],
            "Confiance": p["confiance"], "Commentaires": f"{ETIQUETTE} {p['commentaire']}",
            "Source(s)": sources_txt(p["sources"]),
        })
        part = pd.concat([part, pd.DataFrame([ligne])], ignore_index=True)
        i = fonds.index[masque][0]
        fonds.at[i, "Nb participations documentées"] = (fonds.at[i, "Nb participations documentées"] or 0) + 1
        fonds.at[i, "Date_MAJ"] = DATE_MAJ
        j.add("Mise à jour web — nouvelle participation", f"{p['fonds']} / {p['startup']}",
              "Ligne Participations", None, f"{p['stage']} {p['date_invest']}", p["confiance"],
              "Participation ajoutée", f"{p['commentaire']} Sources : {sources_txt(p['sources'])}",
              onglet="Participations")

    # --- 4. Attributs de start-up mis à jour sur toutes leurs lignes --------------------
    for m in c.get("maj_participations", []):
        masque = part["Start-up"] == m["startup"]
        for champ, valeur in m["champs"].items():
            anciens = sorted({str(x) for x in part.loc[masque, champ] if x is not None}) or ["vide"]
            part.loc[masque, champ] = valeur
            j.add("Mise à jour web — attribut start-up", m["startup"], champ, ", ".join(anciens),
                  valeur, m["confiance"], f"Appliqué aux {int(masque.sum())} lignes de la start-up",
                  f"{m['commentaire']} Sources : {sources_txt(m['sources'])}", onglet="Participations")

    # --- 5. Points à vérifier et pistes écartées (aucune donnée modifiée) -----------------
    for a in c.get("a_verifier", []):
        j.add("Information non confirmée", a["fonds"], "Statut", a["piste"], "Inchangé", "Faible",
              "Non appliqué — à vérifier manuellement", f"{a['raison']}. Source : {a['source']} ({a['url']})")
    j.add("Pistes écartées (campagne web)", f"{len(c['ecartes'])} pistes", "Divers", None, "Aucune modification",
          "Élevé", "Pistes examinées puis écartées",
          " | ".join(f"{e['fonds']} — {e['piste']} : {e['raison']}" for e in c["ecartes"]))
    j.add("Campagne web — bilan", f"{len(set(c['rien_de_neuf']))} sociétés sans nouveauté", "Divers", None,
          "Aucune modification", "Élevé", "Recherche menée, rien de nouveau depuis la passe du 15/09/2026",
          ", ".join(sorted(set(c["rien_de_neuf"])))
          + f". Non recherchés (business angels non exploitables) : {', '.join(c.get('non_recherches', []))}.")

    with pd.ExcelWriter(SORTIE, engine="openpyxl") as writer:
        fonds.to_excel(writer, sheet_name="Fonds", index=False)
        part.to_excel(writer, sheet_name="Participations", index=False)
        tours.to_excel(writer, sheet_name="Tours_de_table", index=False)
        dico.to_excel(writer, sheet_name="Dictionnaire", index=False)
        j.table().to_excel(writer, sheet_name="Anomalies", index=False)

    # --- 6. Actualités → veille_data/entries.json ------------------------------------------
    veille = json.loads(VEILLE.read_text(encoding="utf-8"))
    urls = {e.get("url") for e in veille["entries"]}
    ajouts = [a for a in c["actus"] if a["url"] not in urls]
    veille["entries"].extend(ajouts)
    VEILLE.write_text(json.dumps(veille, ensure_ascii=False, indent=1), encoding="utf-8")

    print(f"Classeur produit : {SORTIE.name}")
    print(f"Fonds : {len(fonds)} lignes · Participations : {len(part)} · Anomalies : +{len(j.lignes)} "
          f"(jusqu'à ANO-{j.suivant - 1:04d})")
    print(f"Actualités ajoutées à la veille : {len(ajouts)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
