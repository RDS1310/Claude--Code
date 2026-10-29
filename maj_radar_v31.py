#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
maj_radar_v31.py — Fiches start-ups : tours datés des CVC, AlloBrain, données publiques françaises
=================================================================================================

Entrée  : VC_Database_Standardisee_v30.xlsx
          radar_refresh/startups_2026_10.json   (recherches web du 29/10/2026, sources citées)
          radar_refresh/entreprises_fr.json     (API Recherche d'entreprises, via GitHub Actions)
Sortie  : VC_Database_Standardisee_v31.xlsx

Demande de l'utilisateur (29/10/2026) : compléter années et montants des participations assurance,
et alimenter les fiches start-ups (levée d'AlloBrain, CA public…).

1. Participations assurance d'Allianz X, AVP et MAIF Avenir + AlloBrain (Groupama) : date, stage,
   taille du tour, montant, rôle — uniquement ce que les sources publient. Une valeur déjà saisie et
   différente est remplacée seulement si la source est plus précise, et l'écart est journalisé.
2. Doublon Allianz X « Coalition » / « Coalition, Inc. » : les deux lignes fusionnées.
3. Sociétés françaises rapprochées d'un SIREN (statut « trouvé ») : SIREN, date de création,
   tranche d'effectif INSEE, dernier CA publié (exercice indiqué). Garde-fou : une création
   postérieure au premier investissement connu signale un homonyme → rien n'est écrit, journalisé.
   Un CA à 0 correspond à des comptes déposés sans chiffre d'affaires publié (confidentialité) :
   ignoré. Les statuts « ambigu » et « non trouvé » sont journalisés une fois, en bloc.

Usage : python3 maj_radar_v31.py
"""
from __future__ import annotations

import json
from datetime import date
from pathlib import Path

import pandas as pd

from liens_croises import DENY, norm
from maj_radar_v27 import Journal

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v30.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v31.xlsx"
TROUVAILLES = BASE_DIR / "radar_refresh" / "startups_2026_10.json"
PUBLIQUES = BASE_DIR / "radar_refresh" / "entreprises_fr.json"
ETIQUETTE = "[Recherche 29/10/2026]"
ETIQUETTE_PUB = "[Données publiques 29/10/2026]"
NOUVELLES_COLONNES = {
    "SIREN": ("Texte", None, "Identifiant INSEE (sociétés françaises), API Recherche d'entreprises"),
    "Effectif (tranche INSEE)": ("Texte", "salariés", "Tranche d'effectif salarié publiée par l'INSEE, avec son année"),
    "Exercice CA": ("Entier", "année", "Exercice du CA (M€) quand il provient des comptes publiés"),
}
BLOCS = [(348, 350, "Libellés v28"), (351, 368, "CVC 10/2026"), (369, 372, "Cohérence 29/10/2026")]


class JournalV31(Journal):
    def add(self, *args, bloc="Fiches start-ups 29/10/2026", **kw):
        super().add(*args, **kw)
        self.lignes[-1]["Table ou bloc source"] = bloc


def vide(v) -> bool:
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() in ("", "nan", "non trouvé")


def ajouter(df, i, col, texte, sep=" "):
    df.at[i, col] = texte if vide(df.at[i, col]) else f"{df.at[i, col]}{sep}{texte}"


def main() -> int:
    feuilles = pd.read_excel(ENTREE, sheet_name=None)
    part = feuilles["Participations"].astype(object)
    part = part.where(pd.notna(part), None)
    ano = feuilles["Anomalies"]
    # Les lignes v28–v30 avaient hérité du libellé de bloc « Portefeuille Sagard » de la classe Journal
    num = ano["Anomalie_ID"].str.extract(r"(\d+)")[0].astype(int)
    for a, b, lib in BLOCS:
        ano.loc[num.between(a, b), "Table ou bloc source"] = lib
    j = JournalV31(ano)

    for col in NOUVELLES_COLONNES:
        if col not in part.columns:
            part.insert(part.columns.get_loc("CA (M€)") + 1, col, None)

    def ligne(fonds, startup):
        m = part["Nom du fonds"].astype(str).str.startswith(fonds) & (part["Start-up"] == startup)
        assert m.sum() == 1, (fonds, startup, int(m.sum()))
        return part.index[m][0]

    # 1. Tours datés / montants issus des sources
    t = json.loads(TROUVAILLES.read_text(encoding="utf-8"))
    for f in t["maj"]:
        i = ligne(f["fonds"], f["startup"])
        for col, val in f["champs"].items():
            avant = part.at[i, col]
            if not vide(avant) and str(avant) != str(val):
                j.add("Valeur remplacée par une source plus précise", f"{f['fonds']} / {f['startup']}", col,
                      str(avant), str(val), f["confiance"], "Remplacée", f["commentaire"])
            part.at[i, col] = val
        part.at[i, "Confiance"] = f["confiance"]
        ajouter(part, i, "Commentaires", f"{ETIQUETTE} {f['commentaire']}")
        ajouter(part, i, "Source(s)", " ; ".join(f["sources"]), sep=" ; ")

    # 2. Doublons au sein d'un même véhicule
    for d in t["doublons"]:
        g, s = ligne(d["fonds"], d["garder"]), ligne(d["fonds"], d["supprimer"])
        for col in part.columns:
            if vide(part.at[g, col]) and not vide(part.at[s, col]):
                part.at[g, col] = part.at[s, col]
        part = part.drop(index=s)
        j.add("Doublon", f"{d['fonds']} / {d['garder']}", "Start-up", f"{d['garder']} + {d['supprimer']}",
              d["garder"], "Élevé", "Lignes fusionnées (champs vides complétés)", "Même société sous deux libellés.")

    # 3. Données publiques des sociétés françaises
    pub = json.loads(PUBLIQUES.read_text(encoding="utf-8"))["societes"]
    fr = part["Pays d’origine"].astype(str).str.contains("France", na=False)
    ecrits, homonymes = 0, []
    for nom, e in sorted(pub.items()):
        lignes = part.index[fr & (part["Start-up"] == nom)]
        if e["statut"] != "trouvé" or not len(lignes):
            continue
        if norm(nom) in DENY:  # noms génériques déjà jugés ambigus pour les liens croisés
            homonymes.append(f"{nom} (nom générique)")
            continue
        creation = int(e["date_creation"][:4]) if e.get("date_creation") else None
        annees = pd.to_numeric(part.loc[lignes, "Date d’investissement"], errors="coerce").dropna()
        if creation and len(annees) and creation > annees.min():
            homonymes.append(f"{nom} (SIREN {e['siren']} créé en {creation}, investi en {int(annees.min())})")
            continue
        ca = [(an, v["ca"]) for an, v in e["finances"].items() if v.get("ca")]
        for i in lignes:
            part.at[i, "SIREN"] = e["siren"]
            if creation:
                avant = part.at[i, "Date de création"]
                if vide(avant):
                    part.at[i, "Date de création"] = creation
                elif str(avant)[:4] != str(creation):
                    j.add("Divergence entre sources", f"{part.at[i, 'Nom du fonds']} / {nom}", "Date de création",
                          str(avant), str(avant), "Moyen", "Valeur existante conservée",
                          f"Registre INSEE : {creation} (SIREN {e['siren']}) — peut dater une société antérieure au pivot.")
            if e.get("tranche_effectif"):
                part.at[i, "Effectif (tranche INSEE)"] = f"{e['tranche_effectif']} ({e['annee_effectif']})"
            texte = f"{ETIQUETTE_PUB} SIREN {e['siren']} ({e['raison_sociale']})"
            if ca:
                an, v = ca[-1]
                if vide(part.at[i, "CA (M€)"]):
                    part.at[i, "CA (M€)"] = round(v / 1e6, 2)
                    part.at[i, "Exercice CA"] = int(an)
                texte += f" ; CA {an} publié : {v / 1e6:.2f} M€".replace(".", ",")
            ajouter(part, i, "Commentaires", texte)
            ajouter(part, i, "Source(s)", e["fiche"], sep=" ; ")
        ecrits += 1
    j.add("Enrichissement source publique", "Sociétés françaises", "SIREN, Date de création, Effectif, CA",
          "recherche-entreprises.api.gouv.fr", f"{ecrits} sociétés complétées", "Élevé",
          "Données du registre (INSEE/INPI), CA uniquement quand les comptes le publient",
          "Rapprochement par nom exact ou quasi exact, activité plausible et société active.")
    if homonymes:
        j.add("Risque d'homonymie", "Sociétés françaises", "SIREN", " ; ".join(homonymes), "Non appliqué", "Faible",
              "Rien écrit : création postérieure au premier investissement", "Homonyme probable, SIREN à identifier à la main.")
    for statut, type_ in (("ambigu", "Risque d'homonymie"), ("non trouvé", "Valeur non trouvée")):
        noms = sorted(n for n, e in pub.items() if e["statut"] == statut)
        if noms:
            j.add(type_, "Sociétés françaises", "SIREN", " ; ".join(noms), "Non appliqué", "Faible",
                  f"Statut « {statut} » au rapprochement", "À rapprocher à la main (raison sociale différente du nom commercial, homonymes).")

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
    print(f"Classeur produit : {SORTIE.name} · {ecrits} sociétés FR complétées · homonymes écartés {len(homonymes)} · "
          f"Anomalies +{len(j.lignes)} (jusqu'à ANO-{j.suivant - 1:04d}) · {date.today():%d/%m/%Y}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
