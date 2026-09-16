#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generer_page_veille.py — Régénère la page « Veille Insurtech »
================================================================

Lit veille_data/entries.json (le journal des actualités détectées) et
produit veille_insurtech.html. C'est ce script que la Routine hebdomadaire
relance après avoir ajouté ses nouvelles entrées au journal.

Usage : python3 generer_page_veille.py
"""

from __future__ import annotations

import json
import sys
from collections import Counter
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "veille_data" / "entries.json"
SORTIE = BASE_DIR / "veille_insurtech.html"


def calculer_tendances(entries: list[dict]) -> list[dict]:
    """Tally des thèmes des entrées pour le panneau Tendances."""
    themes = [e.get("theme") for e in entries if e.get("theme")]
    return [{"theme": t, "n": n} for t, n in Counter(themes).most_common()]


def main() -> int:
    if not ENTREE.exists():
        ENTREE.parent.mkdir(exist_ok=True)
        ENTREE.write_text(json.dumps({
            "entries": [], "derniereMaj": None, "prochaine": None,
            "frequence": "Hebdomadaire (jour à définir)"
        }, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"Journal initialisé : {ENTREE}")

    data = json.loads(ENTREE.read_text(encoding="utf-8"))
    data["entries"] = sorted(data.get("entries", []), key=lambda e: e.get("date", ""), reverse=True)
    data["tendances"] = calculer_tendances(data["entries"])

    template = (BASE_DIR / "_veille_template.html").read_text(encoding="utf-8")
    html = template.replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False))
    SORTIE.write_text(html, encoding="utf-8")

    print(f"Page produite : {SORTIE}")
    print(f"Entrées       : {len(data['entries'])}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
