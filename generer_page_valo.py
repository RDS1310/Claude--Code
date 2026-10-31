#!/usr/bin/env python3
"""
generer_page_valo.py — Page « Valorisation du portefeuille » (démonstration anonymisée)
=====================================================================================

Assemble _valo_template.html + le module partagé valo/ (moteur.js, styles.css) en
valorisation_portefeuille.html. Les participations anonymisées sont embarquées dans la page publiée (lisible sans compte
claude.ai) depuis valo_data/demo.json — fichier HORS DÉPÔT (.gitignore) : les chiffres restent
confidentiels. Sans ce fichier, la page est produite vide et lit la base de l'artefact
(collection « valorisations »). Pour recréer valo_data/demo.json dans une nouvelle session :
ArtifactData list de la collection sur l'artefact de démonstration.
La page produite avec données (valorisation_portefeuille.html) n'est pas suivie par git.
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


DONNEES = BASE_DIR / "valo_data" / "demo.json"


def main() -> int:
    import json
    docs = json.loads(DONNEES.read_text(encoding="utf-8")) if DONNEES.exists() else []
    html = injecter_valo(TEMPLATE.read_text(encoding="utf-8")).replace(
        "__VALO_JSON__", json.dumps(docs, ensure_ascii=False).replace("</", "<\\/"))
    SORTIE.write_text(html, encoding="utf-8")
    print(f"Page produite : {SORTIE.name} — {len(docs)} participation(s) embarquée(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
