#!/usr/bin/env python3
"""
generer_page_analyses.py — Régénère la page « Analyses GVI »
============================================================

Lit analyses_data/analyses.json (grille des 6 critères GVI + fiches de pré-analyse)
et produit analyses_gvi.html à partir de _analyses_template.html.

Ajouter une analyse : compléter analyses_data/analyses.json (une entrée par société,
la plus récente en premier), relancer ce script, publier sur l'artefact existant.

Usage : python3 generer_page_analyses.py
"""
from __future__ import annotations

import json
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "analyses_data" / "analyses.json"
TEMPLATE = BASE_DIR / "_analyses_template.html"
SORTIE = BASE_DIR / "analyses_gvi.html"
VERDICTS = {"etudier", "surveiller", "ecarter"}


def main() -> int:
    data = json.loads(ENTREE.read_text(encoding="utf-8"))
    cles = {g["cle"] for g in data["grille"]}
    for a in data["analyses"]:
        assert a["verdict"] in VERDICTS, f"{a['nom']} : verdict inconnu {a['verdict']}"
        inconnues = set(a.get("notes", {})) - cles
        assert not inconnues, f"{a['nom']} : critères inconnus {inconnues}"
        for c, n in a.get("notes", {}).items():
            assert n.get("note") in (None, 1, 2, 3), f"{a['nom']} / {c} : note hors échelle"
    # Index des liens vers Radar / Pipe / actualités, par nom normalisé (les fiches créées sur la
    # page sont rapprochées dans le navigateur avec la même normalisation que liens_croises.norm)
    from liens_croises import construire_index, URLS, dates_desc
    data["liens"] = {
        n: {"radar": f"{URLS['radar']}#s-{e['slug']}" if e["radar"] else None,
            "pipe": f"{URLS['pipe']}#p-{e['slug']}" if e["pipe"] else None,
            "actus": dates_desc(e["actus"])[:5]}
        for n, e in construire_index()["societes"].items() if e["radar"] or e["pipe"] or e["actus"]
    }
    html = TEMPLATE.read_text(encoding="utf-8").replace(
        "__DATA_JSON__", json.dumps(data, ensure_ascii=False).replace("</", "<\\/"))
    SORTIE.write_text(html, encoding="utf-8")
    print(f"Page produite : {SORTIE.name} — {len(data['analyses'])} analyse(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
