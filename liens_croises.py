#!/usr/bin/env python3
"""
liens_croises.py — Index commun qui relie les 5 pages entre elles
==================================================================

Chaque générateur l'importe pour que la même société ou le même fonds renvoie vers les
autres pages, avec les mêmes actualités :

  Radar (fiche start-up)  ← actualités Veille / M&A + liens Pipe GVI et Analyse GVI
  Pipe GVI (ligne)        ← investisseurs suivis dans le Radar + actualités + Analyse GVI
  Veille / M&A            → fiche Radar du fonds ou de la start-up citée (jamais le pipe :
                            la Veille ne doit pas révéler le statut pipe d'une société)
  Analyses GVI            ← index par nom (les fiches créées sur la page sont reliées côté navigateur)

Liens profonds : Radar #s-<slug> (start-up) et #f-<Fonds_ID> (fonds) ; Pipe #p-<slug> ;
Veille #v-<clé> ; Analyses #<id>. Les noms sont rapprochés après normalisation (casse,
accents, ponctuation, ™) ; les actus sont rattachées à une start-up si son nom apparaît
comme mot entier dans le titre ou le résumé (noms ambigus exclus : DENY).
"""
from __future__ import annotations

import hashlib
import json
import re
import unicodedata
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
URLS = {
    "radar": "https://claude.ai/artifact/X6o5LvDnYe6Qo5fPYFkX7g",
    "pipe": "https://claude.ai/artifact/58n6yRVTLgDuhavqhxihGn",
    "veille": "https://claude.ai/artifact/5FXTpkCo6KvGu6YyeYAHiw",
    "ma": "https://claude.ai/artifact/BPSeTnKnWY2PkAGcH5s2ai",
    "analyses": "https://claude.ai/artifact/Mor6DTNTSbekT4irHSLyJe",
}
# Noms trop communs ou homonymes connus : pas de rattachement automatique d'actualités.
DENY = {"grace", "ticker", "mash", "the family", "portage", "cathay", "serena", "casey", "continuity",
        "asterias", "seeds", "albert", "atomic", "tilt", "midas", "conduit", "croissant", "happl", "allocate"}
NOM_MIN = 4


def norm(nom) -> str:
    t = unicodedata.normalize("NFKD", str(nom or "")).encode("ascii", "ignore").decode().lower()
    t = re.sub(r"[™®]", "", t)
    t = re.sub(r"[^a-z0-9]+", " ", t).strip()
    t = re.sub(r"\b(sas|sa|ltd|gmbh|inc|ag|bv|srl)$", "", t).strip()
    return t


def slug(nom) -> str:
    return norm(nom).replace(" ", "-")


def cle_url(url) -> str:
    return "v-" + hashlib.md5(str(url or "").encode()).hexdigest()[:8]


def _json(rel, defaut):
    p = BASE_DIR / rel
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else defaut


def _dernier_classeur() -> Path:
    fichiers = sorted(BASE_DIR.glob("VC_Database_Standardisee_v*.xlsx"),
                      key=lambda p: int(re.search(r"_v(\d+)\.xlsx$", p.name).group(1)))
    return fichiers[-1]


def _mot(nom_n: str, texte_n: str) -> bool:
    return bool(nom_n) and re.search(rf"(?<![a-z0-9]){re.escape(nom_n)}(?![a-z0-9])", texte_n) is not None


def construire_index() -> dict:
    """{societes: {norm: {...}}, fonds: {norm: {nom, fid}}} — une seule source pour toutes les pages."""
    classeur = _dernier_classeur()
    part = pd.read_excel(classeur, sheet_name="Participations")
    fonds_df = pd.read_excel(classeur, sheet_name="Fonds")
    pipe_df = pd.read_excel(BASE_DIR / "gvi_data" / "GVI_Status_Report_Startups.xlsx", sheet_name="Start-up")
    veille = _json("veille_data/entries.json", {"entries": []})
    ma = _json("ma_data/deals.json", {"confirmees": [], "rumeurs": []})
    analyses = _json("analyses_data/analyses.json", {"analyses": []})

    soc: dict[str, dict] = {}

    def entree(nom):
        n = norm(nom)
        if not n:
            return None
        return soc.setdefault(n, {"nom": str(nom).replace("™", "").strip(), "slug": slug(nom), "radar": [],
                                  "pipe": False, "analyse": None, "actus": []})

    for _, r in part.iterrows():
        e = entree(r.get("Start-up"))
        if e is not None and pd.notna(r.get("Fonds_ID")):
            e["radar"].append({"fid": r["Fonds_ID"], "fonds": r.get("Nom du fonds"), "vehicule": r.get("Véhicule")})
    for nom in pipe_df["Start-up"].dropna():
        e = entree(nom)
        if e is not None:
            e["pipe"] = True
    for a in analyses.get("analyses", []):
        e = entree(a.get("nom"))
        if e is not None:
            e["analyse"] = a.get("id")

    fonds = {}
    for _, r in fonds_df.iterrows():
        n = norm(r.get("Nom du fonds"))
        if n and n not in fonds:
            fonds[n] = {"nom": r.get("Nom du fonds"), "fid": r.get("Fonds_ID")}

    # Actualités : Veille puis M&A, rattachées aux sociétés citées
    items = [{"type": "veille", "titre": x.get("titre"), "date": x.get("date"), "source": x.get("source"),
              "url": x.get("url"), "page": f"{URLS['veille']}#{cle_url(x.get('url'))}",
              "texte": f"{x.get('titre', '')} {x.get('resume', '')}"} for x in veille.get("entries", [])]
    for d in ma.get("confirmees", []):
        items.append({"type": "ma", "titre": f"{d.get('acquereur')} → {d.get('cible')}", "date": d.get("date"),
                      "source": d.get("source"), "url": d.get("url"), "page": URLS["ma"],
                      "texte": f"{d.get('acquereur', '')} {d.get('cible', '')} {d.get('resume', '')}"})
    for d in ma.get("rumeurs", []):
        items.append({"type": "ma", "titre": f"{d.get('acteurs_pressentis')} → {d.get('cible')} (rumeur)",
                      "date": d.get("date"), "source": d.get("source"), "url": d.get("url"), "page": URLS["ma"],
                      "texte": f"{d.get('acteurs_pressentis', '')} {d.get('cible', '')} {d.get('resume', '')}"})
    for it in items:
        texte_n = norm(it.pop("texte"))
        for n, e in soc.items():
            if len(n) >= NOM_MIN and n not in DENY and _mot(n, texte_n):
                e["actus"].append(it)
    return {"societes": soc, "fonds": fonds}


def liens_texte(index: dict, texte: str) -> list[dict]:
    """Fonds et start-ups du Radar cités dans un texte (Veille, M&A) — jamais le pipe."""
    texte_n = norm(texte)
    out, vus = [], set()
    for n, f in index["fonds"].items():
        if len(n) >= NOM_MIN and n not in DENY and _mot(n, texte_n) and n not in vus:
            vus.add(n)
            out.append({"type": "fonds", "nom": f["nom"], "url": f"{URLS['radar']}#f-{f['fid']}"})
    for n, e in index["societes"].items():
        if e["radar"] and len(n) >= NOM_MIN and n not in DENY and _mot(n, texte_n) and n not in vus:
            vus.add(n)
            out.append({"type": "startup", "nom": e["nom"], "url": f"{URLS['radar']}#s-{e['slug']}"})
    return out


def dates_desc(actus: list[dict]) -> list[dict]:
    def k(a):
        try:
            j, m, y = (int(x) for x in str(a.get("date", "")).split("/"))
            return (y, m, j)
        except ValueError:
            return (0, 0, 0)
    return sorted(actus, key=k, reverse=True)
