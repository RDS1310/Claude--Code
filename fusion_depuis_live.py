#!/usr/bin/env python3
"""
fusion_depuis_live.py — Recale les données du dépôt sur la page en ligne
========================================================================

Les routines Claude (veille hebdomadaire, M&A mensuelle) publient la page sans pouvoir
pousser sur GitHub. La page en ligne est donc la version de référence : avant toute mise
à jour, on lit l'artefact (outil Artifact, action « read ») et on fusionne ses données ici.

Règle : la version dont `derniereMaj` est la plus récente sert de base.
  - veille : on y ajoute les entrées de l'autre version absentes (même url) — aucune perte ;
  - ma     : la base est reprise telle quelle (les statuts de deals évoluent : rumeur → confirmé).
Les champs calculés par les générateurs (tendances, comps) sont retirés.

Usage : python3 fusion_depuis_live.py veille|ma <page_en_ligne.html>
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
FICHIERS = {"veille": BASE_DIR / "veille_data" / "entries.json", "ma": BASE_DIR / "ma_data" / "deals.json"}
CALCULES = {"tendances", "comps"}


def cle_date(jjmmaaaa: str | None) -> tuple:
    try:
        j, m, a = (int(x) for x in (jjmmaaaa or "").split("/"))
        return (a, m, j)
    except ValueError:
        return (0, 0, 0)


def donnees_page(html: str) -> dict:
    m = re.search(r"window\.__DATA__\s*=\s*(\{.*?\});\s*</script>", html, re.S)
    if not m:
        raise SystemExit("window.__DATA__ introuvable dans la page en ligne")
    data = json.loads(m.group(1).replace("<\\/", "</"))
    return {k: v for k, v in data.items() if k not in CALCULES}


def main() -> int:
    if len(sys.argv) != 3 or sys.argv[1] not in FICHIERS:
        print(__doc__)
        return 2
    page, chemin_live = sys.argv[1], Path(sys.argv[2])
    depot = json.loads(FICHIERS[page].read_text(encoding="utf-8"))
    live = donnees_page(chemin_live.read_text(encoding="utf-8"))

    base, autre = (live, depot) if cle_date(live.get("derniereMaj")) >= cle_date(depot.get("derniereMaj")) else (depot, live)
    ajoutees = 0
    if page == "veille":
        urls = {e.get("url") for e in base.get("entries", [])}
        for e in autre.get("entries", []):
            if e.get("url") not in urls:
                base.setdefault("entries", []).append(e)
                ajoutees += 1
        base["entries"].sort(key=lambda e: cle_date(e.get("date")), reverse=True)

    FICHIERS[page].write_text(json.dumps(base, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"Base : {'page en ligne' if base is live else 'dépôt'} (derniereMaj {base.get('derniereMaj')}) ; "
          f"dépôt {depot.get('derniereMaj')} / en ligne {live.get('derniereMaj')} ; entrées ajoutées depuis l'autre version : {ajoutees}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
