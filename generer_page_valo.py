#!/usr/bin/env python3
"""
generer_page_valo.py — Page « Valorisation du portefeuille » (démonstration anonymisée)
=====================================================================================

Assemble _valo_template.html + le module partagé valo/ (moteur.js, styles.css) en
valorisation_portefeuille.html. La page n'embarque aucune donnée : les participations
(anonymisées) sont dans la base de l'artefact, collection « valorisations ».
Le même module sert l'onglet privé de la page Analyses GVI (generer_page_analyses.py).

Usage : python3 generer_page_valo.py
"""
from __future__ import annotations

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE = BASE_DIR / "_valo_template.html"
SORTIE = BASE_DIR / "valorisation_portefeuille.html"


def injecter_valo(html: str) -> str:
    """Remplace les marqueurs du module de valorisation (CSS et JS partagés)."""
    css = (BASE_DIR / "valo" / "styles.css").read_text(encoding="utf-8")
    js = (BASE_DIR / "valo" / "moteur.js").read_text(encoding="utf-8")
    assert "/*__VALO_CSS__*/" in html and "/*__VALO_JS__*/" in html
    return html.replace("/*__VALO_CSS__*/", css).replace("/*__VALO_JS__*/", js.replace("</script", "<\\/script"))


def main() -> int:
    SORTIE.write_text(injecter_valo(TEMPLATE.read_text(encoding="utf-8")), encoding="utf-8")
    print(f"Page produite : {SORTIE.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
