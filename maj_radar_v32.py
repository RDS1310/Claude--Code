#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
maj_radar_v32.py — Rafraîchissement des fiches start-ups et des portefeuilles (31/10/2026)
========================================================================================

Entrée  : VC_Database_Standardisee_v31.xlsx
          radar_refresh/societes_2026_10.json      (recherches web du 31/10/2026, une fiche par société)
          radar_refresh/portefeuilles_2026_10.json (relevé des sites des fonds, s'il existe)
Sortie  : VC_Database_Standardisee_v32.xlsx

Demande de l'utilisateur (31/10/2026) : « refaire un scrapping … les participations ainsi que toutes les
infos que tu puisses pour les participations (valo, exit etc.) ».

1. Données « société » (valables pour toutes les lignes de participation d'une même start-up) :
   total levé, dernier tour connu (stade, montant, année, chef de file), valorisation et son année,
   sortie (rachat, introduction en bourse, liquidation), pays et année de création manquants, CA.
   Une valeur déjà saisie et différente n'est pas écrasée : l'écart est journalisé.
   Une sortie « totale » (rachat à 100 %, cession d'actifs, liquidation, faillite) passe Exit (O/N) à O ;
   une introduction en bourse ou une prise de contrôle partielle ne renseigne que le statut.
2. Portefeuilles relevés sur les sites des fonds : participations nouvelles et sorties signalées par le
   fonds lui-même (voir portefeuilles_2026_10.json), avec la page du fonds comme source.

Montants en M€ ($ × 0,92 ; £ × 1,17). Années seules, jamais de mois inventé.
Usage : python3 maj_radar_v32.py
"""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

from maj_radar_v27 import Journal

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v31.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v32.xlsx"
SOCIETES = BASE_DIR / "radar_refresh" / "societes_2026_10.json"
PORTEFEUILLES = BASE_DIR / "radar_refresh" / "portefeuilles_2026_10.json"
ETIQUETTE = "[Recherche 31/10/2026]"
NOUVELLES_COLONNES = {
    "Dernier tour (stade)": ("Texte", None, "Dernier tour de financement connu de la start-up (tous investisseurs)"),
    "Dernier tour (M€)": ("Nombre", "M€", "Montant du dernier tour connu"),
    "Dernier tour (année)": ("Entier", "année", "Année du dernier tour connu"),
    "Dernier tour (chef de file)": ("Texte", None, "Investisseur(s) chef(s) de file du dernier tour"),
    "Valorisation (année)": ("Entier", "année", "Année de la valorisation indiquée"),
}
SORTIES_TOTALES = ("Acquisition", "Cession d'actifs", "Liquidation", "Faillite")
REMPLACABLES = {"Total levé par la start-up (M€)", "Dernier tour (stade)", "Dernier tour (M€)", "Dernier tour (année)",
                "Dernier tour (chef de file)", "Valorisation (M€)", "Valorisation (année)"}


class JournalV32(Journal):
    def add(self, *args, bloc="Rafraîchissement 31/10/2026", **kw):
        super().add(*args, **kw)
        self.lignes[-1]["Table ou bloc source"] = bloc


def vide(v) -> bool:
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() in ("", "nan", "non trouvé")


def ajouter(df, i, col, texte, sep=" "):
    if texte and texte not in str(df.at[i, col] or ""):
        df.at[i, col] = texte if vide(df.at[i, col]) else f"{df.at[i, col]}{sep}{texte}"


def fmt(v) -> str:
    if not isinstance(v, (int, float)):
        return str(v)
    return f"{v:,.0f}".replace(",", " ") if v >= 1000 else f"{v:g}".replace(".", ",")


def appliquer_societes(part: pd.DataFrame, j: JournalV32) -> tuple[int, int, int]:
    data = json.loads(SOCIETES.read_text(encoding="utf-8"))["societes"]
    nb_soc = nb_lignes = nb_sorties = 0
    for nom, e in data.items():
        noms = [nom] + e.get("alias", [])
        lignes = part.index[part["Start-up"].isin(noms)]
        if not len(lignes):
            j.add("Valeur non trouvée", nom, "Start-up", nom, "Non appliqué", "Faible",
                  "Fiche de recherche sans ligne correspondante", "Libellé à rapprocher à la main.")
            continue
        nb_soc += 1
        valeurs = {
            "Total levé par la start-up (M€)": e.get("total_leve"),
            "Valorisation (M€)": e.get("valorisation"),
            "Valorisation (année)": e.get("valo_annee") if e.get("valorisation") is not None else None,
            "Pays d’origine": e.get("pays"),
            "Date de création": e.get("creation"),
        }
        t = e.get("dernier_tour") or {}
        valeurs.update({"Dernier tour (stade)": t.get("stage"), "Dernier tour (M€)": t.get("montant"),
                        "Dernier tour (année)": t.get("annee"), "Dernier tour (chef de file)": t.get("lead")})
        if e.get("ca"):
            valeurs.update({"CA (M€)": e["ca"]["montant"], "Exercice CA": e["ca"]["annee"]})
        notes = [x for x in (e.get("total_note") and f"Total levé : {e['total_note']}.",
                             e.get("valo_note") and f"Valorisation : {e['valo_note']}.",
                             e.get("statut_note") and f"Statut : {e['statut_note']}.",
                             t.get("note") and f"Dernier tour : {t['note']}.") if x]
        p = e.get("tour_precedent")
        if p:
            notes.append("Tour précédent : " + ", ".join(fmt(x) for x in (p.get("stage"), p.get("annee"),
                         p.get("montant") and f"{fmt(p['montant'])} M€", p.get("lead")) if x) + ".")
        sortie = e.get("exit")
        statut = None
        if sortie:
            statut = " ".join(x for x in (sortie["type"], f"par {sortie['acquereur']}" if sortie.get("acquereur") else "",
                                           f"({sortie['annee']})" if sortie.get("annee") else "") if x)
            if sortie.get("montant"):
                statut += f", {fmt(sortie['montant'])} M€"
            if sortie.get("note"):
                notes.append(f"Sortie : {sortie['note']}.")
        for i in lignes:
            nb_lignes += 1
            for col, val in valeurs.items():
                if val is None:
                    continue
                avant = part.at[i, col]
                if vide(avant):
                    part.at[i, col] = val
                elif str(avant).strip() != str(val) and col not in ("Pays d’origine",):
                    try:
                        proche = abs(float(avant) - float(val)) <= 0.05 * max(abs(float(val)), 1)
                    except (TypeError, ValueError):
                        proche = False
                    if proche:
                        continue
                    if col in REMPLACABLES:  # un total levé ou un dernier tour ne fait que croître : la source datée la plus récente l'emporte
                        part.at[i, col] = val
                        j.add("Valeur remplacée par une source plus récente", f"{part.at[i, 'Nom du fonds']} / {nom}", col,
                              str(avant), str(val), e.get("confiance", "Moyen"), "Remplacée", e["sources"][0])
                    else:
                        j.add("Divergence entre sources", f"{part.at[i, 'Nom du fonds']} / {nom}", col, str(avant),
                              str(avant), "Moyen", "Valeur existante conservée",
                              f"Recherche du 31/10/2026 : {val} ({e['sources'][0]}).")
            if statut:
                part.at[i, "Statut start-up (détail)"] = statut
                avant = part.at[i, "Exit (O/N)"]
                if sortie["type"].startswith(SORTIES_TOTALES) and str(avant) != "O":
                    part.at[i, "Exit (O/N)"] = "O"
                    nb_sorties += 1
                    j.add("Reclassement", f"{part.at[i, 'Nom du fonds']} / {nom}", "Exit (O/N)",
                          "vide" if vide(avant) else str(avant), "O",
                          e.get("confiance", "Moyen"), "Sortie déduite du rachat de la société",
                          f"{statut} — {e['sources'][0]}")
            ajouter(part, i, "Commentaires", f"{ETIQUETTE} " + " ".join(notes) if notes else "")
            for s in e["sources"]:
                ajouter(part, i, "Source(s)", s, sep=" ; ")
    j.add("Enrichissement source publique", "Start-ups", "Total levé, dernier tour, valorisation, sortie, pays",
          "Recherches web du 31/10/2026 (presse spécialisée, communiqués, agrégateurs)",
          f"{nb_soc} sociétés, {nb_lignes} lignes", "Moyen",
          "Données de niveau société reportées sur chaque participation",
          "Les totaux issus d'agrégateurs (CB Insights, Dealroom…) sont signalés comme tels en commentaire.")
    return nb_soc, nb_lignes, nb_sorties


def appliquer_portefeuilles(fonds: pd.DataFrame, part: pd.DataFrame, j: JournalV32) -> tuple[pd.DataFrame, int, int]:
    """Participations nouvelles et sorties publiées par les fonds eux-mêmes."""
    if not PORTEFEUILLES.exists():
        return part, 0, 0
    data = json.loads(PORTEFEUILLES.read_text(encoding="utf-8"))
    part = part.copy()
    nouvelles, sorties = [], 0
    for f in data["fonds"]:
        fid = f["fonds_id"]
        ref = fonds[fonds["Fonds_ID"] == fid]
        assert len(ref) == 1, fid
        ref = ref.iloc[0]
        for s in f.get("ajouts", []):
            if ((part["Fonds_ID"] == fid) & (part["Start-up"].str.lower() == s["nom"].lower())).any():
                continue
            nouvelles.append({"Fonds_ID": fid, "Nom du fonds": ref["Nom du fonds"], "Véhicule": ref["Véhicule"],
                              "Start-up": s["nom"], "Secteur": s.get("secteur"), "Pays d’origine": s.get("pays"),
                              "Date d’investissement": s.get("annee"),
                              "Exit (O/N)": "O" if s.get("sortie") else "N", "Statut start-up (détail)": s.get("statut"),
                              "Confiance": "Moyen", "Source(s)": f["page"],
                              "Rôle du véhicule (détail)": f.get("note") if "Véhicule non précisé" in f.get("note", "") else None,
                              "Commentaires": f"[Portefeuille du fonds 31/10/2026] Listée sur la page portefeuille officielle."
                                              + (f" Description publiée : {s['description']}" if s.get("description") else "")})
        for s in f.get("sorties", []):
            m = (part["Fonds_ID"] == fid) & (part["Start-up"].str.lower() == s.lower())
            for i in part.index[m]:
                avant = part.at[i, "Exit (O/N)"]
                if str(avant) != "O":
                    part.at[i, "Exit (O/N)"] = "O"
                    sorties += 1
                    j.add("Reclassement", f"{ref['Nom du fonds']} / {s}", "Exit (O/N)", "vide" if vide(avant) else str(avant),
                          "O", "Élevé", "Marquée « exit / realized » sur la page portefeuille du fonds", f["page"])
        if f.get("ajouts"):
            j.add("Enrichissement source publique", ref["Nom du fonds"], "Participations", f["page"],
                  f"+{sum(1 for n in nouvelles if n['Fonds_ID'] == fid)} participations", "Moyen",
                  "Ajout depuis la page portefeuille officielle (date et montant non publiés)", f.get("note", ""))
    for m in data.get("maj", []):  # année d'investissement publiée par le fonds
        for i in part.index[(part["Fonds_ID"] == m["fonds_id"]) & (part["Start-up"] == m["startup"])]:
            avant = part.at[i, "Date d’investissement"]
            if vide(avant):
                part.at[i, "Date d’investissement"] = m["annee"]
                ajouter(part, i, "Source(s)", m["page"], sep=" ; ")
            elif str(avant)[:4] != str(m["annee"]):
                j.add("Divergence entre sources", m["startup"], "Date d’investissement", str(avant), str(avant), "Moyen",
                      "Valeur existante conservée", f"La page du fonds indique {m['annee']} : {m['page']}")
    for c in data.get("corrections", []):
        for i in part.index[(part["Nom du fonds"] == c["fonds"]) & (part["Start-up"] == c["startup"])]:
            avant = part.at[i, c["champ"]]
            if str(avant) != c["valeur"]:
                part.at[i, c["champ"]] = c["valeur"]
                j.add("Correction", f"{c['fonds']} / {c['startup']}", c["champ"], str(avant), c["valeur"], "Moyen",
                      "Corrigé d'après la page portefeuille du fonds", f"{c['motif']} {c['page']}")
    if nouvelles:
        part = pd.concat([part, pd.DataFrame(nouvelles)], ignore_index=True)
    return part, len(nouvelles), sorties


def main() -> int:
    feuilles = pd.read_excel(ENTREE, sheet_name=None)
    fonds = feuilles["Fonds"].astype(object)
    part = feuilles["Participations"].astype(object)
    part = part.where(pd.notna(part), None)
    j = JournalV32(feuilles["Anomalies"])
    for col in NOUVELLES_COLONNES:
        if col not in part.columns:
            apres = "Valorisation (M€)" if col == "Valorisation (année)" else "Total levé par la start-up (M€)"
            decalage = list(NOUVELLES_COLONNES).index(col) if col.startswith("Dernier") else 0
            part.insert(part.columns.get_loc(apres) + 1 + decalage, col, None)
    part, nb_ajouts, nb_sorties_fonds = appliquer_portefeuilles(fonds, part, j)  # d'abord les lignes, puis les données société
    nb_soc, nb_lignes, nb_sorties = appliquer_societes(part, j)

    dico = feuilles["Dictionnaire"]
    for col, (type_, unite, comm) in NOUVELLES_COLONNES.items():
        if not (dico["Champ"] == col).any():
            dico = pd.concat([dico, pd.DataFrame([{"Champ": col, "Onglet": "Participations", "Type": type_, "Unité": unite,
                                                   "Champ source ou calculé": "Champ source", "Commentaire": comm}])],
                             ignore_index=True)
    feuilles.update({"Participations": part, "Anomalies": j.table(), "Dictionnaire": dico})
    with pd.ExcelWriter(SORTIE, engine="openpyxl") as writer:
        for nom, df in feuilles.items():
            df.to_excel(writer, sheet_name=nom, index=False)
    print(f"Classeur produit : {SORTIE.name} · {nb_soc} sociétés / {nb_lignes} lignes enrichies · "
          f"sorties {nb_sorties} (rachats) + {nb_sorties_fonds} (pages des fonds) · +{nb_ajouts} participations · "
          f"Anomalies +{len(j.lignes)} (jusqu'à ANO-{j.suivant - 1:04d})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
