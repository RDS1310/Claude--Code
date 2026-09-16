#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
nettoyer_doublons_participations.py — Corrige des doublons/erreurs historiques
====================================================================================

En examinant les 6 lignes de Participations sans Confiance renseignée (reliquats
de la toute première campagne de collecte, antérieure à la colonne Confiance),
plusieurs anomalies de qualité de données sont apparues :

- "Friss" (BlackfinTech 1) est un doublon de casse de "FRISS", déjà présent et
  vérifié avec Confiance Élevé — fusion, le "Montant investi par le fonds" est
  conservé sur la ligne canonique.
- "Descartes UnderWritting" (faute de frappe) apparaît deux fois en double d'une
  ligne "Descartes Underwriting" déjà vérifiée : sur BlackfinTech 1 (Seed), le
  montant associé (12 M€) est incohérent avec la taille du tour (2,3 M€ total) —
  écarté ; sur BlackfinTech II (Série B, tour de 110 M€), le montant (12 M€) est
  plausible et conservé.
- "Orus" (Portag3 Venture III) n'est pas un doublon : la vérification Partech
  (Vérif-Lot V1) a confirmé que "Frst, Partech, Portage Ventures + business
  angels" ont co-investi dans le seed d'Orus (juillet 2022) — la ligne est
  légitime, seulement sous-documentée (Confiance vide). Upgrade en Moyen.
- "Zeffy" (Ring Capital / Mission I) est une plateforme de dons pour associations
  sans aucun lien avec l'assurance — hors périmètre du radar insurtech. Retirée.
- "Neat" (Founders Future Fund I, millésime 2018) est un doublon erroné : Neat
  (seed 2022) est déjà correctement rattachée à Founders Future Good (millésime
  2021, vérifiée Moyen-Élevé) dans Vérif-Lot V1. Retirée.

Entrée  : VC_Database_Standardisee_v8.xlsx
Sortie  : VC_Database_Standardisee_v9.xlsx

Usage : python3 nettoyer_doublons_participations.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v8.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v9.xlsx"


def main() -> int:
    fonds = pd.read_excel(ENTREE, sheet_name="Fonds")
    part = pd.read_excel(ENTREE, sheet_name="Participations").copy()
    tours = pd.read_excel(ENTREE, sheet_name="Tours_de_table")
    dico = pd.read_excel(ENTREE, sheet_name="Dictionnaire")
    ano = pd.read_excel(ENTREE, sheet_name="Anomalies")

    idx = {
        (r["Fonds_ID"], r["Start-up"]): i for i, r in part.iterrows()
    }

    # 1) Friss -> fusion dans FRISS (BlackfinTech 1)
    fid_bf1 = "blackfin-capital-partners_blackfintech-1"
    i_friss_minuscule = idx.get((fid_bf1, "Friss"))
    i_friss_canon = idx.get((fid_bf1, "FRISS"))
    a_supprimer = []
    if i_friss_minuscule is not None and i_friss_canon is not None:
        part.at[i_friss_canon, "Montant investi par le fonds (M€)"] = part.at[
            i_friss_minuscule, "Montant investi par le fonds (M€)"
        ]
        a_supprimer.append(i_friss_minuscule)

    # 2) Descartes UnderWritting -> fusion dans Descartes Underwriting (BlackfinTech 1, Seed)
    #    Montant 12M€ incohérent avec un tour de 2,3M€ : écarté (pas de report).
    i_dup_bf1 = idx.get((fid_bf1, "Descartes UnderWritting"))
    if i_dup_bf1 is not None:
        a_supprimer.append(i_dup_bf1)

    # 3) Descartes UnderWritting -> fusion dans Descartes Underwriting (BlackfinTech II, Série B)
    fid_bf2 = "blackfin-capital-partners_blackfintech-ii"
    i_dup_bf2 = idx.get((fid_bf2, "Descartes UnderWritting"))
    i_canon_bf2 = idx.get((fid_bf2, "Descartes Underwriting"))
    if i_dup_bf2 is not None and i_canon_bf2 is not None:
        part.at[i_canon_bf2, "Montant investi par le fonds (M€)"] = part.at[
            i_dup_bf2, "Montant investi par le fonds (M€)"
        ]
        a_supprimer.append(i_dup_bf2)

    # 4) Orus / Portag3 Venture III : ligne légitime, sous-documentée -> upgrade
    i_orus = idx.get(("portage-venture_portag3-venture-iii", "Orus"))
    if i_orus is not None:
        part.at[i_orus, "Confiance"] = "Moyen"
        part.at[i_orus, "Stage financé"] = "Seed"
        part.at[i_orus, "Date d’investissement"] = 2022.0
        part.at[i_orus, "Taille du tour (M€)"] = 5.0
        part.at[i_orus, "Commentaires"] = (
            "Vérif-Lot V1 (verification Partech/Orus) confirme que le seed "
            "d'Orus (juillet 2022, 5M€) a réuni Frst, Partech, Portage Ventures "
            "et des business angels — Portage Ventures est donc un co-investisseur "
            "légitime, ligne auparavant sous-documentée."
        )
        part.at[i_orus, "Source(s)"] = "Partech (annonce Orus) ; Silicon Canals ; EU-Startups"

    # 5) Zeffy : hors périmètre (pas de lien assurance) -> retrait
    i_zeffy = idx.get(("ring-capital_mission-i", "Zeffy"))
    if i_zeffy is not None:
        a_supprimer.append(i_zeffy)

    # 6) Neat / Founders Future Fund I : doublon erroné (déjà sous Founders Future Good) -> retrait
    i_neat_dup = idx.get(("founders-future-vc_founders-future-fund-i", "Neat"))
    if i_neat_dup is not None:
        a_supprimer.append(i_neat_dup)

    part = part.drop(index=a_supprimer)
    part_out = part.reset_index(drop=True)

    prochain_id = int(ano["Anomalie_ID"].str.extract(r"(\d+)")[0].astype(int).max()) + 1
    nouvelles_anomalies = [
        {
            "Anomalie_ID": f"ANO-{prochain_id:04d}",
            "Type d'anomalie": "Participation hors périmètre (pas de lien assurance)",
            "Onglet source": "Participations",
            "Table ou bloc source": "Nettoyage doublons (16/09/2026)",
            "Ligne source": None,
            "Entité concernée": "Ring Capital",
            "Champ concerné": "Start-up",
            "Valeur source": "Zeffy (Mission I, Seed, 4,4M€)",
            "Valeur retenue": "Ligne retirée",
            "Candidats éventuels": None,
            "Score de similarité": None,
            "Niveau de confiance": "Élevé",
            "Traitement appliqué": "Retirée : Zeffy est une plateforme de dons pour associations, sans aucun lien avec l'assurance — reliquat de la toute première campagne de collecte (préexistante à la colonne Confiance), non filtrée pour le périmètre insurtech.",
            "Commentaire": "Détectée en examinant les lignes sans Confiance renseignée.",
        },
        {
            "Anomalie_ID": f"ANO-{prochain_id + 1:04d}",
            "Type d'anomalie": "Doublon (mauvais véhicule)",
            "Onglet source": "Participations",
            "Table ou bloc source": "Nettoyage doublons (16/09/2026)",
            "Ligne source": None,
            "Entité concernée": "Founders Future VC",
            "Champ concerné": "Véhicule",
            "Valeur source": "Neat rattachée à « Founders Future Fund I » (millésime 2018, sans Confiance)",
            "Valeur retenue": "Ligne retirée (doublon)",
            "Candidats éventuels": "Founders Future Good (millésime 2021) — déjà correctement rattachée, Confiance Moyen-Élevé",
            "Score de similarité": None,
            "Niveau de confiance": "Élevé",
            "Traitement appliqué": "Retirée : Neat (seed 2022) est déjà rattachée à Founders Future Good via Vérif-Lot V1, incompatible avec le millésime 2018 de Fund I.",
            "Commentaire": "Reliquat de la toute première campagne de collecte, antérieur à la recherche dédiée sur Neat.",
        },
    ]
    ano_out = pd.concat([ano, pd.DataFrame(nouvelles_anomalies)], ignore_index=True)

    with pd.ExcelWriter(SORTIE, engine="openpyxl") as writer:
        fonds.to_excel(writer, sheet_name="Fonds", index=False)
        part_out.to_excel(writer, sheet_name="Participations", index=False)
        tours.to_excel(writer, sheet_name="Tours_de_table", index=False)
        dico.to_excel(writer, sheet_name="Dictionnaire", index=False)
        ano_out.to_excel(writer, sheet_name="Anomalies", index=False)

    print(f"Classeur produit : {SORTIE}")
    print(f"Participations avant : {len(pd.read_excel(ENTREE, sheet_name='Participations'))}  →  après : {len(part_out)}")
    print(f"Lignes supprimées (doublons/hors périmètre) : {len(a_supprimer)}")
    print(f"Anomalies avant : {len(ano)}  →  après : {len(ano_out)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
