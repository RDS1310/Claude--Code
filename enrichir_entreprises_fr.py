#!/usr/bin/env python3
"""
enrichir_entreprises_fr.py — Données publiques des start-ups françaises (gratuit, sans clé)
==========================================================================================

Interroge l'API officielle « Recherche d'entreprises » (recherche-entreprises.api.gouv.fr, données
INSEE/INPI) pour chaque société française du Radar (Participations, pays France) et du pipe GVI
(pays France), et enregistre : SIREN, raison sociale, date de création, tranche d'effectif, siège,
activité, état, et les comptes publiés (CA, résultat net par exercice) quand la société les publie.

Le rapprochement nom → SIREN est prudent : un candidat n'est retenu que si son nom (raison sociale,
sigle ou enseigne) correspond au nom suivi (score ≥ 0,85), s'il est actif et si son activité (NAF) est
plausible pour une start-up ; entre homonymes, la plus grande tranche d'effectif l'emporte, et une
égalité donne le statut « ambigu ». Sinon la société est « non trouvée » avec ses candidats.
Le script d'application (maj_radar_vN) vérifie en plus que la création précède l'investissement. Rien n'est écrit dans le Radar :
le résultat (radar_refresh/entreprises_fr.json) est appliqué ensuite par un script maj_radar_vN.

Lancé par GitHub Actions (.github/workflows/enrichir_entreprises.yml) : l'API n'est pas joignable
depuis les sessions Claude. Limite de l'API : 7 requêtes/s → pause de 0,25 s entre deux appels.
"""
from __future__ import annotations

import difflib
import json
import re
import time
import urllib.parse
import urllib.request
from datetime import date
from pathlib import Path

import pandas as pd

from liens_croises import norm

BASE_DIR = Path(__file__).resolve().parent
SORTIE = BASE_DIR / "radar_refresh" / "entreprises_fr.json"
API = "https://recherche-entreprises.api.gouv.fr/search?"
SEUIL = 0.85
TRANCHES = {"NN": None, "00": "0", "01": "1-2", "02": "3-5", "03": "6-9", "11": "10-19", "12": "20-49", "21": "50-99",
            "22": "100-199", "31": "200-249", "32": "250-499", "41": "500-999", "42": "1 000-1 999", "51": "2 000-4 999",
            "52": "5 000-9 999", "53": "10 000+"}
FORMES = r"\b(sas|sasu|sa|sarl|eurl|sci|snc|scop)\b"
# Activités plausibles pour une start-up suivie (logiciel, données, finance/assurance, conseil, R&D, santé) :
# sert à départager les homonymes (une « ALAN » créée en 2018 avec 6 salariés n'est pas Alan SA).
NAF_PLAUSIBLES = ("58.2", "62.", "63.", "64.", "65.", "66.", "70.2", "71.12", "72.", "73.", "74.9", "82.", "85.", "86.")
RANG_EFFECTIF = {k: i for i, k in enumerate(TRANCHES)}


def dernier_classeur() -> Path:
    return sorted(BASE_DIR.glob("VC_Database_Standardisee_v*.xlsx"),
                  key=lambda p: int(re.search(r"_v(\d+)\.xlsx$", p.name).group(1)))[-1]


def cle(nom: str) -> str:
    return re.sub(r"\s+", " ", re.sub(FORMES, "", norm(re.sub(r"\(.*?\)", "", str(nom))))).strip()


def appel(q: str) -> list[dict]:
    url = API + urllib.parse.urlencode({"q": q, "per_page": 5, "page": 1})
    for essai in range(3):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "radar-insurtech"}), timeout=20) as r:
                return json.loads(r.read().decode()).get("results", [])
        except Exception as e:  # 429 / réseau : on réessaie après une pause
            time.sleep(2 + 3 * essai)
            err = e
    raise RuntimeError(f"API injoignable pour {q!r} : {err}")


def score(nom: str, r: dict) -> float:
    cible = cle(nom)
    noms = [r.get("nom_raison_sociale"), r.get("nom_complet"), r.get("sigle")]
    noms += (r.get("siege") or {}).get("liste_enseignes") or []
    best = 0.0
    for n in filter(None, noms):
        c = cle(n)
        if not c:
            continue
        s = 1.0 if c == cible else difflib.SequenceMatcher(None, c, cible).ratio()
        if c.startswith(cible + " ") or cible.startswith(c + " "):
            s = max(s, 0.9)
        best = max(best, s)
    return round(best, 3)


def resume(r: dict) -> dict:
    fin = r.get("finances") or {}
    exercices = {an: {"ca": v.get("ca"), "resultat_net": v.get("resultat_net")} for an, v in sorted(fin.items())}
    siege = r.get("siege") or {}
    return {
        "siren": r.get("siren"), "raison_sociale": r.get("nom_raison_sociale") or r.get("nom_complet"),
        "date_creation": r.get("date_creation"), "etat": r.get("etat_administratif"),
        "tranche_effectif": TRANCHES.get(str(r.get("tranche_effectif_salarie") or "NN")),
        "annee_effectif": r.get("annee_tranche_effectif_salarie"),
        "activite": r.get("activite_principale"), "commune": siege.get("libelle_commune"),
        "nature_juridique": r.get("nature_juridique"), "categorie": r.get("categorie_entreprise"),
        "finances": exercices,
        "fiche": f"https://annuaire-entreprises.data.gouv.fr/entreprise/{r.get('siren')}",
    }


def societes() -> dict[str, set]:
    noms: dict[str, set] = {}
    part = pd.read_excel(dernier_classeur(), sheet_name="Participations")
    for _, r in part[part["Pays d’origine"].astype(str).str.contains("France", na=False)].iterrows():
        noms.setdefault(str(r["Start-up"]).strip(), set()).add("radar")
    pipe = pd.read_excel(BASE_DIR / "gvi_data" / "GVI_Status_Report_Startups.xlsx", sheet_name="Start-up")
    col_pays = next(c for c in pipe.columns if str(c).startswith("Pays"))
    for _, r in pipe[pipe[col_pays].astype(str).str.contains("France", case=False, na=False)].iterrows():
        if pd.notna(r["Start-up"]):
            noms.setdefault(str(r["Start-up"]).strip(), set()).add("pipe")
    return noms


def plausible(r: dict) -> bool:
    return str(r.get("activite_principale") or "").startswith(NAF_PLAUSIBLES) and r.get("etat_administratif") == "A"


def main() -> int:
    resultats, trouves = {}, 0
    for nom, origines in sorted(societes().items()):
        candidats = appel(cle(nom) or nom)
        time.sleep(0.25)
        if not candidats and cle(nom) != nom:
            candidats = appel(nom)
            time.sleep(0.25)
        notes = sorted(((score(nom, c), c) for c in candidats), key=lambda x: -x[0])
        retenus = [c for s, c in notes if s >= SEUIL]
        # Homonymes : on ne garde que les activités plausibles, puis la plus grande tranche d'effectif.
        bons = sorted((c for c in retenus if plausible(c)),
                      key=lambda c: -RANG_EFFECTIF.get(str(c.get("tranche_effectif_salarie") or "NN"), 0))
        if bons:
            ambigu = len(bons) > 1 and (RANG_EFFECTIF.get(str(bons[0].get("tranche_effectif_salarie") or "NN"), 0)
                                        == RANG_EFFECTIF.get(str(bons[1].get("tranche_effectif_salarie") or "NN"), 0))
            resultats[nom] = {"statut": "ambigu" if ambigu else "trouvé", "score": score(nom, bons[0]),
                              "origines": sorted(origines), **resume(bons[0]),
                              "autres_homonymes": [{"siren": c.get("siren"), "nom": c.get("nom_complet"),
                                                    "creation": c.get("date_creation"),
                                                    "activite": c.get("activite_principale")} for c in retenus if c is not bons[0]]}
            trouves += not ambigu
        else:
            resultats[nom] = {"statut": "non trouvé", "origines": sorted(origines),
                              "candidats": [{"score": s, "siren": c.get("siren"), "nom": c.get("nom_complet")} for s, c in notes[:3]]}
    SORTIE.write_text(json.dumps({"date": date.today().isoformat(), "source": "recherche-entreprises.api.gouv.fr",
                                  "societes": resultats}, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{trouves}/{len(resultats)} sociétés rapprochées → {SORTIE.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
