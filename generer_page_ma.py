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
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "ma_data" / "deals.json"
SORTIE = BASE_DIR / "etude_marche_ma.html"


def main() -> int:
    data = json.loads(ENTREE.read_text(encoding="utf-8"))
    template = (BASE_DIR / "_ma_template.html").read_text(encoding="utf-8")
    html = template.replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False))
    SORTIE.write_text(html, encoding="utf-8")

    print(f"Page produite : {SORTIE}")
    print(f"Confirmées    : {len(data.get('confirmees', []))}")
    print(f"Rumeurs       : {len(data.get('rumeurs', []))}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
