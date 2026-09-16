#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generer_page_ma.py — Régénère la page « Étude de Marché M&A Assurance »
=========================================================================

Lit ma_data/deals.json (opérations confirmées + rumeurs de marché) et
produit etude_marche_ma.html.

Usage : python3 generer_page_ma.py
"""

from __future__ import annotations

import json
import sys
from collections import Counter
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "ma_data" / "deals.json"
SORTIE = BASE_DIR / "etude_marche_ma.html"


def calculer_tendances(data: dict) -> list[dict]:
    """Tally des thèmes (confirmées + rumeurs) pour le panneau Tendances."""
    themes = [d.get("theme") for d in data.get("confirmees", []) if d.get("theme")]
    themes += [d.get("theme") for d in data.get("rumeurs", []) if d.get("theme")]
    return [{"theme": t, "n": n} for t, n in Counter(themes).most_common()]


def calculer_comps(data: dict) -> list[dict]:
    """Comparables de valorisation : opérations confirmées avec valorisation chiffrée."""
    comps = [
        {
            "cible": d.get("cible"), "secteur": d.get("secteur"),
            "valorisationM": d.get("valorisationM"), "multiple": d.get("multiple"),
            "date": d.get("date"), "url": d.get("url"),
        }
        for d in data.get("confirmees", []) if d.get("valorisationM")
    ]
    comps.sort(key=lambda c: c["valorisationM"], reverse=True)
    return comps


def main() -> int:
    data = json.loads(ENTREE.read_text(encoding="utf-8"))
    data["tendances"] = calculer_tendances(data)
    data["comps"] = calculer_comps(data)
    template = (BASE_DIR / "_ma_template.html").read_text(encoding="utf-8")
    html = template.replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False))
    SORTIE.write_text(html, encoding="utf-8")

    print(f"Page produite : {SORTIE}")
    print(f"Confirmées    : {len(data.get('confirmees', []))}")
    print(f"Rumeurs       : {len(data.get('rumeurs', []))}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
