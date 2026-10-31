#!/usr/bin/env python3
"""
analyse_portefeuilles.py — Confronte le relevé des sites (outils/portefeuilles/releve.json) au Radar
====================================================================================================

Pour chaque société de gestion : noms candidats extraits des pages portefeuille (textes de liens courts,
textes alternatifs des logos, intitulés), comparés aux participations du dernier classeur avec un
rapprochement souple (sans parenthèses, préfixe, similarité de bigrammes ≥ 0,85). Signale aussi les
indices de sortie (« exit », « realized », « acquired by »…) autour des noms connus.

Sortie : outils/portefeuilles/analyse.json — à relire avant toute écriture dans le Radar (maj_radar_vN).
Usage : python3 outils/analyse_portefeuilles.py
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
RELEVE = BASE_DIR / "outils" / "portefeuilles" / "releve.json"
SORTIE = BASE_DIR / "outils" / "portefeuilles" / "analyse.json"
BRUIT = re.compile(r"^(home|accueil|about|team|équipe|contact|news|blog|portfolio|portefeuille|companies|read more|"
                   r"en savoir plus|learn more|voir|view|all|tous|toutes|filter|menu|logo|linkedin|twitter|x|instagram|"
                   r"facebook|youtube|cookies?|privacy|legal|mentions légales|imprint|impressum|careers?|jobs|search|"
                   r"next|previous|back|retour|fr|en|de|close|open|subscribe|newsletter|investments?|insights?|press|"
                   r"events?|approach|thesis|founders|apply|pitch|login|sign in|load more|show more|voir plus)$", re.I)
SORTIE_MOTS = re.compile(r"\b(exit(ed)?|realized|realised|acquired|acquisition|ipo|cédée?s?|vendue?s?|former|past investments?|alumni|sold)\b", re.I)
ASSURANCE = re.compile(r"insur|assur|versicher|reinsur|réassur|claims?|sinistre|underwrit|mga\b|broker|courtier|"
                       r"actuar|parametric|paramétrique|risk|risque|cyber|telematic|mobility|health|santé|pension|"
                       r"prévoyance|benefit|warranty|garantie|embedded", re.I)


def norm(s: str) -> str:
    s = re.sub(r"\(.*?\)", "", str(s)).lower()
    s = re.sub(r"\b(inc|ltd|gmbh|sas|sa|ag|bv|limited|group|technologies|technology|insurance|ai|io|app|hq)\b\.?", "", s)
    return re.sub(r"[^a-z0-9]+", "", s)


def bigrammes(s: str) -> set:
    return {s[i:i + 2] for i in range(len(s) - 1)}


def meme(a: str, b: str) -> bool:
    a, b = norm(a), norm(b)
    if not a or not b:
        return False
    if a == b or (len(a) > 3 and len(b) > 3 and (a.startswith(b) or b.startswith(a))):
        return True
    x, y = bigrammes(a), bigrammes(b)
    return bool(x and y) and 2 * len(x & y) / (len(x) + len(y)) >= 0.85


def candidats(page: dict, domaine: str) -> list[tuple[str, str]]:
    """(nom, contexte) : liens courts vers une sous-page portefeuille ou un site externe, logos, intitulés."""
    out = []
    for l in page.get("liens", []):
        t = re.sub(r"\s+", " ", l["t"]).strip()
        if not t or len(t) > 40 or BRUIT.match(t) or t.count(" ") > 4:
            continue
        h = l["h"]
        externe = domaine not in h
        sous_page = re.search(r"/(portfolio|portefeuille|companies|company|participations?|investments?|startups?)/[^/?#]+", h)
        if externe or sous_page:
            out.append((t.split("\n")[0], h))
    for a in page.get("alts", []):
        a = re.sub(r"\s*(logo|icon|image)\s*", " ", a, flags=re.I).strip()
        if a and len(a) <= 40 and not BRUIT.match(a):
            out.append((a, "logo"))
    return out


def main() -> int:
    releve = json.loads(RELEVE.read_text(encoding="utf-8"))["fonds"] if RELEVE.exists() else {}
    for f in sorted((RELEVE.parent / "fonds").glob("*.json")):  # relevé par lots (une fiche par société de gestion)
        x = json.loads(f.read_text(encoding="utf-8"))
        releve[x["firme"]] = x
    classeur = sorted(BASE_DIR.glob("VC_Database_Standardisee_v*.xlsx"), key=lambda p: int(re.search(r"_v(\d+)", p.name).group(1)))[-1]
    part = pd.read_excel(classeur, sheet_name="Participations")
    res = {}
    for firme, r in releve.items():
        connues = sorted(set(part.loc[part["Nom du fonds"] == firme, "Start-up"].dropna().astype(str)))
        domaine = re.sub(r"^https?://(www\.)?", "", (r.get("accueil") or {}).get("final") or r["site"]).split("/")[0]
        pages = [r.get("accueil") or {}] + r.get("pages", [])
        vus, noms = set(), []
        for p in pages:
            for n, ctx in candidats(p, domaine):
                k = norm(n)
                if k and k not in vus:
                    vus.add(k)
                    noms.append({"nom": n, "contexte": ctx, "page": p.get("url")})
        slugs = [u.rstrip("/").split("/")[-1].replace("-", " ") for u in r.get("sitemap", [])]
        textes = "\n".join(p.get("texte", "") for p in pages)
        trouvees = [c for c in connues if any(meme(c, n["nom"]) for n in noms) or norm(c) in norm(textes)]
        nouvelles = [n for n in noms if not any(meme(c, n["nom"]) for c in connues)]
        indices_sortie = []
        for c in connues:
            for m in re.finditer(re.escape(c.split(" (")[0]), textes, re.I):
                fenetre = textes[max(0, m.start() - 120): m.end() + 120]
                if SORTIE_MOTS.search(fenetre):
                    indices_sortie.append({"nom": c, "extrait": re.sub(r"\s+", " ", fenetre)})
                    break
        res[firme] = {
            "site": r["site"], "pages": [p.get("url") for p in r.get("pages", [])], "nb_pages": len(r.get("pages", [])),
            "texte_total": len(textes), "connues": len(connues), "retrouvees": trouvees,
            "non_retrouvees": [c for c in connues if c not in trouvees],
            "nouvelles_candidates": nouvelles, "slugs_sitemap": slugs[:200], "indices_sortie": indices_sortie,
            "mentions_assurance": [n["nom"] for n in nouvelles if ASSURANCE.search(n["nom"] + " " + n["contexte"])],
        }
    SORTIE.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
    for f, x in sorted(res.items()):
        print(f"{f[:32]:32} pages {x['nb_pages']} · texte {x['texte_total']:>6} · connues {x['connues']:>3} · "
              f"retrouvées {len(x['retrouvees']):>3} · candidates {len(x['nouvelles_candidates']):>3} · sorties {len(x['indices_sortie'])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
