#!/usr/bin/env python3
"""
generer_page_valo.py — Module « Valorisation du portefeuille » (injection du moteur partagé)
=========================================================================================

injecter_valo(html) remplace les marqueurs /*__VALO_CSS__*/ et /*__VALO_JS__*/ par valo/styles.css
et valo/moteur.js. Utilisé par generer_page_analyses.py (onglet « Valorisation du portefeuille »).
La page de démonstration anonymisée a été supprimée le 01/11/2026 à la demande de l'utilisateur.
"""
from __future__ import annotations

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent


def injecter_valo(html: str) -> str:
    """Remplace les marqueurs du module de valorisation (CSS et JS partagés)."""
    css = (BASE_DIR / "valo" / "styles.css").read_text(encoding="utf-8")
    js = (BASE_DIR / "valo" / "moteur.js").read_text(encoding="utf-8")
    assert "/*__VALO_CSS__*/" in html and "/*__VALO_JS__*/" in html
    return html.replace("/*__VALO_CSS__*/", css).replace("/*__VALO_JS__*/", js.replace("</script", "<\\/script"))
