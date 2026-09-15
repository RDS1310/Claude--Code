#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
maj_web_vc.py — Mise à jour de la base standardisée à partir des sites des fonds (source 3)
===========================================================================================

Entrée  : VC_Database_Standardisee_v2.xlsx
Sortie  : VC_Database_Standardisee_v3.xlsx

Ce script applique une passe de vérification web (sites officiels des sociétés de gestion,
communiqués de closing, presse spécialisée) sur les champs restés vides côté source 1.

Règles :
- on ne remplit que des cellules VIDES (aucune valeur existante n'est écrasée sans trace) ;
- toute valeur retenue est horodatée (Date_MAJ) et sourcée (Source_AuM / Anomalies) ;
- les divergences entre la base et la source web sont journalisées, jamais arbitrées en silence ;
- les marqueurs « n.a » sont remplacés par de vraies cellules vides (règle du chantier initial :
  une valeur manquante reste vide, jamais « NA »).
"""

from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import restructuration_vc as rvc  # réutilisation des fonctions de mise en forme

FICHIER_ENTREE = "VC_Database_Standardisee_v2.xlsx"
FICHIER_SORTIE = "VC_Database_Standardisee_v3.xlsx"
DATE_MAJ = date(2026, 9, 14)
ONGLETS = ["Fonds", "Participations", "Tours_de_table", "Dictionnaire", "Anomalies"]

# Marqueurs d'absence de donnée à neutraliser (doivent rester des cellules vides)
MARQUEURS_VIDES = {"n.a", "n.a.", "na", "n/a", "nd", "n.d.", "non disponible", "inconnu", "-", "--"}


# =============================================================================
# 1. RÉSULTATS DE LA PASSE WEB
# =============================================================================
# Chaque entrée : Fonds_ID -> valeurs vérifiées + source. `ecrase` liste les champs
# pour lesquels une correction est assumée (valeur par défaut manifestement fausse).

MAJ_WEB = [
    dict(fid="alven-capital_alven-iv", vals={"AuM (M€)": 2000},
         src="alven.co (communiqué Alven VI, 2022) : « €2B in assets under management » ; "
             "Climate Transparency Hub / ADEME (2025) : Alven Capital Partners, 6 fonds"),
    dict(fid="alven-capital_alven-v", vals={"AuM (M€)": 2000, "Montant levé (M€)": 250},
         src="alven.co, communiqué « Alven Capital closes its fifth generation fund at €250M "
             "hard cap » (first and final close, cible initiale €200M) ; AuM €2Md (alven.co, 2022)"),
    dict(fid="alven-capital_alven-vi", vals={"AuM (M€)": 2000},
         src="alven.co, communiqué Alven VI (hard cap €350M, cible initiale €300M) — "
             "« €2B in assets under management »"),
    dict(fid="speedinvest_speedinvest-iv",
         vals={"AuM (M€)": 1000, "Montant levé (M€)": 350, "Millésime": 2022,
               "Statut": "Fermé", "Géographie": "Europe", "Stratégie": "Généraliste",
               "Pré-Seed": True, "Seed": True, "Stages pratiqués": "Pré-Seed ; Seed"},
         ecrase=["Pré-Seed", "Seed"],
         src="EU-Startups / tech.eu / TechCrunch (31/01/2024) : closing final de Speedinvest 4 "
             "à €350M (premier closing €300M en déc. 2022, millésime 2022) ; AuM du groupe "
             "> €1Md ; stratégie pré-seed / seed multi-verticales (deeptech, fintech, health, "
             "marketplaces, climate, SaaS)"),
    dict(fid="eos-venture_evp-i", vals={"AuM (M€)": 78.2},
         src="Insurance Insider / sonr.global (mai 2020) : closing d'EVP I à $85M, « total funds "
             "under management now reaching $85m » — converti à 1 USD ≈ 0,92 EUR (règle de "
             "conversion déjà retenue dans la base)"),
    dict(fid="elaia-partners_dv4", vals={"AuM (M€)": 700},
         src="CFNEWS / Agefi Asset Management (avr. 2024) : Elaia Partners gère 700 M€ d'actifs "
             "(niveau société de gestion, avant partenariat Lazard Elaia Capital)"),
    dict(fid="xange-capital_xange-4", vals={"AuM (M€)": 700},
         src="Maddyness (fiche fonds, dernier investissement 31/03/2025) et SuperVenture/Informa : "
             "XAnge, « €700 million under management », marque innovation du groupe Siparex"),
    dict(fid="eurazeo_eurazeo-venture-capital", vals={"AuM (M€)": 3200},
         src="Maddyness (mai 2024, interview Matthieu Baret) : l'expertise Venture Capital "
             "d'Eurazeo représente 3,2 Md€ d'actifs sous gestion au 31/12/2023 (AuM groupe : "
             "39 Md€ fin 2025, communiqué Eurazeo 11/03/2026)"),
]

# Divergences constatées entre la base et la source web : signalées, non arbitrées
DIVERGENCES = [
    dict(fid="alven-capital_alven-v", champ="Millésime", base="2016",
         web="closing unique annoncé en 2017 (alven.co)",
         com="Millésime base conservé : le communiqué ne précise pas la date de lancement de la levée"),
    dict(fid="elaia-partners_elaia-delta", champ="AuM (M€)", base="805",
         web="700 M€ au niveau société de gestion (CFNEWS, avr. 2024)",
         com="AuM société renseigné à des dates différentes selon les véhicules (805 / 850 / 700) : "
             "à harmoniser avec une date de référence unique"),
    dict(fid="elaia-partners_dv4", champ="Montant levé (M€)", base="200",
         web="premier closing à 120 M€ en mai 2021 (FrenchWeb, communiqué Elaia DV4)",
         com="La base retient 200 M€ (taille finale probable) : écart non arbitré, valeur base conservée"),
    dict(fid="eos-venture_evp-i", champ="Montant levé (M€)", base="85",
         web="$85M (Insurance Insider, mai 2020) ≈ 78,2 M€",
         com="Montant exprimé en USD dans la source 1 mais stocké dans une colonne M€ : "
             "valeur base conservée, conversion appliquée uniquement à l'AuM"),
    dict(fid="bpi-france_large-ventures", champ="AuM (M€)", base="(vide)",
         web="fonds doté de 1,75 Md€ (Bpifrance / L'Info Durable, 2021) après 500 M€ en 2013 "
             "puis 600 M€ complémentaires",
         com="AuM laissé vide : incohérent avec le montant levé de 2 500 M€ figurant en base, "
             "aucune source primaire 2025-2026 trouvée pour trancher"),
    dict(fid="xange-capital_xange-4", champ="Millésime", base="2021",
         web="levée de 220 M€ annoncée le 07/07/2022 (Journal du Net) ; premier closing à 125 M€ "
             "annoncé en 2021 (Maddyness)",
         com="Millésime base conservé (premier closing) : cohérent avec la règle « début des levées »"),
]

# Sites officiels identifiés lors de la passe web (société de gestion)
SITES = {
    "Alven Capital": "https://alven.co",
    "Elaia Partners": "https://www.elaia.com",
    "Eurazeo": "https://www.eurazeo.com",
    "XAnge Capital": "https://www.xange.vc",
    "BPI France": "https://www.bpifrance.fr",
    "Atlantic Vantage Point": "https://www.avpcap.com",
    "115K": "https://www.115k.fr",
    "365.fintech": "https://www.365fintech.sk",
    "Cadence Growth Capital": "https://www.cadencegrowthcapital.com",
    "CommerzVentures": "https://www.commerzventures.com",
    "Committed Capital": "https://www.committedcapital.co.uk",
    "Concentric": "https://www.concentric.vc",
    "13books Capital (ex-Element Ventures)": "https://www.13bookscapital.com",
    "Helsana HealthInvest": "https://www.helsana.ch",
    "Howden Ventures": "https://www.howdengroupholdings.com",
    "Insurtech Gateway": "https://www.insurtechgateway.com",
    "Open CNP": "https://open.cnp.fr",
    "SCOR Ventures": "https://www.scor.com",
    "Schumpeter Ventures": "https://www.schumpeter.vc",
    "Seed X Liechtenstein": "https://www.seedx.li",
    "Tenity": "https://www.tenity.com",
    "Venpace": "https://www.venpace.com",
    "Step Venture": "https://www.stepventure.eu",
    "Kickstart Innovation": "https://www.kickstart-innovation.com",
    "Mash VC": "https://www.mash.vc",
    "Seraphim Space": "https://www.seraphim.vc",
    "South East Angels": "https://www.southeastangels.co.uk",
    "The Net Street Capital": "https://www.thenetstreet.com",
    "Blast.Club": "https://www.blast.club",
    "speedInvest": "https://www.speedinvest.com",
}

SOURCE_LABEL = "Passe web (sites officiels des fonds + communiqués) — 14/09/2026"


# =============================================================================
# 2. OUTILS
# =============================================================================

def localiser_entree(base_dir: Path) -> Path:
    """Cherche le classeur d'entrée dans le répertoire du script puis dans les dépôts connus."""
    for dossier in (base_dir, Path.cwd(), Path("/mnt/user-data/uploads")):
        chemin = dossier / FICHIER_ENTREE
        if chemin.exists():
            return chemin
        if dossier.exists():
            trouves = sorted(dossier.glob(f"*{FICHIER_ENTREE}"))
            if trouves:
                return trouves[0]
    raise FileNotFoundError(f"Classeur d'entrée introuvable : {FICHIER_ENTREE}")


def est_marqueur_vide(valeur) -> bool:
    """Vrai si la cellule contient un marqueur d'absence de donnée (à neutraliser)."""
    if not isinstance(valeur, str):
        return False
    return valeur.strip().lower() in MARQUEURS_VIDES


def est_vide(valeur) -> bool:
    if valeur is None:
        return True
    if isinstance(valeur, float) and pd.isna(valeur):
        return True
    if isinstance(valeur, str) and valeur.strip() == "":
        return True
    return False


def nettoyer_marqueurs(frames: dict) -> int:
    """Remplace « n.a » & co par de vraies cellules vides. Retourne le nombre de cellules."""
    total = 0
    for nom, df in frames.items():
        for col in df.columns:
            masque = df[col].map(est_marqueur_vide)
            total += int(masque.sum())
            if masque.any():
                df.loc[masque, col] = None
        # recast des colonnes numériques redevenues exploitables
        frames[nom] = df
    return total


class JournalAnomalies:
    """Ajoute des lignes à l'onglet Anomalies en poursuivant la numérotation existante."""

    def __init__(self, anomalies: pd.DataFrame):
        self.colonnes = list(anomalies.columns)
        self.df = anomalies
        derniers = [int(str(a).split("-")[-1]) for a in anomalies["Anomalie_ID"].dropna()
                    if str(a).startswith("ANO-")]
        self.compteur = max(derniers) if derniers else 0
        self.nouvelles = []

    def ajouter(self, type_ano, entite=None, champ=None, valeur_source=None, valeur_retenue=None,
                candidats=None, confiance="Élevé", traitement=None, commentaire=None):
        self.compteur += 1
        ligne = {c: None for c in self.colonnes}
        ligne.update({
            "Anomalie_ID": f"ANO-{self.compteur:04d}",
            "Type d'anomalie": type_ano,
            "Onglet source": SOURCE_LABEL,
            "Table ou bloc source": "Sites officiels / communiqués / presse spécialisée",
            "Entité concernée": entite,
            "Champ concerné": champ,
            "Valeur source": valeur_source,
            "Valeur retenue": valeur_retenue,
            "Candidats éventuels": candidats,
            "Niveau de confiance": confiance,
            "Traitement appliqué": traitement,
            "Commentaire": commentaire,
        })
        self.nouvelles.append(ligne)

    def resultat(self) -> pd.DataFrame:
        if not self.nouvelles:
            return self.df
        return pd.concat([self.df, pd.DataFrame(self.nouvelles, columns=self.colonnes)],
                         ignore_index=True)


# =============================================================================
# 3. APPLICATION DE LA PASSE WEB
# =============================================================================

def appliquer_maj(fonds: pd.DataFrame, journal: JournalAnomalies) -> tuple[int, int, list[str]]:
    """Remplit les cellules vides avec les valeurs vérifiées ; journalise chaque écriture."""
    remplies, corrigees, touchees = 0, 0, []
    index = {fid: i for i, fid in enumerate(fonds["Fonds_ID"])}
    for maj in MAJ_WEB:
        fid = maj["fid"]
        if fid not in index:
            journal.ajouter("Rapprochement non résolu", entite=fid, champ="Fonds_ID",
                            valeur_source=fid, confiance="Élevé",
                            traitement="Mise à jour non appliquée",
                            commentaire="Fonds_ID absent de la base : vérifier le référentiel")
            continue
        i = index[fid]
        entite = f"{fonds.at[i, 'Nom du fonds']} / {fonds.at[i, 'Véhicule']}"
        ecrase = set(maj.get("ecrase", []))
        ecrit = []
        for champ, valeur in maj["vals"].items():
            actuelle = fonds.at[i, champ]
            if not est_vide(actuelle) and champ not in ecrase:
                if str(actuelle) != str(valeur):
                    journal.ajouter("Variante typographique", entite=entite, champ=champ,
                                    valeur_source=f"web : {valeur}", valeur_retenue=actuelle,
                                    confiance="Moyen", traitement="Valeur existante conservée",
                                    commentaire="Écart entre la base et la source web : non arbitré")
                continue
            if champ in ecrase and not est_vide(actuelle):
                journal.ajouter("Correction de nom", entite=entite, champ=champ,
                                valeur_source=actuelle, valeur_retenue=valeur, confiance="Élevé",
                                traitement="Valeur corrigée à partir de la source web",
                                commentaire="Valeur par défaut (FAUX) corrigée : stade documenté "
                                            "par le fonds")
                corrigees += 1
            else:
                remplies += 1
            fonds.at[i, champ] = valeur
            ecrit.append(champ)
        if not ecrit:
            continue
        touchees.append(entite)
        # traçabilité : source + date de mise à jour
        note = f"[MAJ web {DATE_MAJ:%d/%m/%Y}] {maj['src']}"
        actuelle_src = fonds.at[i, "Source_AuM"]
        fonds.at[i, "Source_AuM"] = note if est_vide(actuelle_src) else f"{actuelle_src} | {note}"
        fonds.at[i, "Date_MAJ"] = DATE_MAJ
        journal.ajouter("Mise à jour web", entite=entite, champ=" ; ".join(ecrit),
                        valeur_retenue=" ; ".join(f"{c}={maj['vals'][c]}" for c in ecrit),
                        confiance="Élevé",
                        traitement="Cellules vides complétées à partir de la source web",
                        commentaire=maj["src"])
    return remplies, corrigees, touchees


def journaliser_divergences(fonds: pd.DataFrame, journal: JournalAnomalies) -> None:
    """Consigne les écarts base / web sans les arbitrer."""
    index = {fid: i for i, fid in enumerate(fonds["Fonds_ID"])}
    for d in DIVERGENCES:
        i = index.get(d["fid"])
        entite = (f"{fonds.at[i, 'Nom du fonds']} / {fonds.at[i, 'Véhicule']}"
                  if i is not None else d["fid"])
        journal.ajouter("Différence entre description et structure réelle", entite=entite,
                        champ=d["champ"], valeur_source=f"web : {d['web']}",
                        valeur_retenue=f"base : {d['base']}", confiance="Moyen",
                        traitement="Valeur de la base conservée, écart documenté",
                        commentaire=d["com"])


def ajouter_site_web(fonds: pd.DataFrame, journal: JournalAnomalies) -> int:
    """Ajoute la colonne « Site web » (niveau société de gestion)."""
    if "Site web" not in fonds.columns:
        fonds["Site web"] = None
    n = 0
    for i, nom in enumerate(fonds["Nom du fonds"]):
        url = SITES.get(str(nom).strip())
        if url and est_vide(fonds.at[i, "Site web"]):
            fonds.at[i, "Site web"] = url
            n += 1
    journal.ajouter("Champ ajouté au modèle", champ="Site web",
                    valeur_retenue=f"{n} lignes renseignées", confiance="Élevé",
                    traitement="Colonne ajoutée en fin de tableau",
                    commentaire="Site officiel de la société de gestion, relevé lors de la passe "
                                "web ; vide lorsque le site n'a pas été confirmé")
    return n


def recalculer_dictionnaire(dico: pd.DataFrame, frames: dict) -> pd.DataFrame:
    """Recalcule les taux de remplissage et ajoute la ligne du nouveau champ."""
    if not ((dico["Champ"] == "Site web") & (dico["Onglet"] == "Fonds")).any():
        ligne = {c: None for c in dico.columns}
        ligne.update({
            "Champ": "Site web", "Onglet": "Fonds", "Onglet source": SOURCE_LABEL,
            "En-tête source": "Site officiel de la société de gestion",
            "Méthode de rapprochement": "Déduction par contexte", "Niveau de confiance": "Élevé",
            "Type": "Texte", "Champ source ou calculé": "Champ source",
            "Formule ou règle": "URL du site officiel relevée lors de la passe web",
            "Commentaire": "Colonne ajoutée au modèle cible (après les blocs Deal) pour préparer "
                           "les prochaines passes de collecte",
        })
        dico = pd.concat([dico, pd.DataFrame([ligne])], ignore_index=True)
    for idx, row in dico.iterrows():
        df = frames.get(row["Onglet"])
        if df is None or row["Champ"] not in df.columns:
            continue
        n = len(df)
        remplies = int(sum(0 if est_vide(v) else 1 for v in df[row["Champ"]]))
        dico.at[idx, "Nb valeurs renseignées"] = remplies
        dico.at[idx, "Taux de remplissage %"] = round(100.0 * remplies / n, 1) if n else 0.0
    return dico


# =============================================================================
# 4. ORCHESTRATION
# =============================================================================

def main() -> int:
    base_dir = Path(__file__).resolve().parent
    try:
        entree = localiser_entree(base_dir)
    except FileNotFoundError as exc:
        print(f"ERREUR : {exc}")
        return 1
    sortie = base_dir / FICHIER_SORTIE

    frames = {nom: pd.read_excel(entree, nom) for nom in ONGLETS}
    frames = {k: v.astype(object).where(pd.notna(v), None) for k, v in frames.items()}
    avant = {k: int(sum(0 if est_vide(x) else 1 for col in v.columns for x in v[col]))
             for k, v in frames.items()}

    journal = JournalAnomalies(frames["Anomalies"])

    # 1. neutralisation des marqueurs « n.a »
    nettoyees = nettoyer_marqueurs(frames)
    if nettoyees:
        journal.ajouter("Valeur invalide", champ="Toutes colonnes",
                        valeur_source="n.a", valeur_retenue="(vide)", confiance="Élevé",
                        traitement=f"{nettoyees} cellules « n.a » remplacées par des cellules vides",
                        commentaire="Règle du modèle : une valeur manquante reste vide et n'est "
                                    "jamais représentée par un marqueur textuel (NA, n.a, -, ...)")

    # 2. passe web
    fonds = frames["Fonds"]
    remplies, corrigees, touchees = appliquer_maj(fonds, journal)
    journaliser_divergences(fonds, journal)
    n_sites = ajouter_site_web(fonds, journal)
    frames["Fonds"] = fonds

    # 3. dictionnaire + anomalies
    frames["Anomalies"] = journal.resultat()
    frames["Dictionnaire"] = recalculer_dictionnaire(frames["Dictionnaire"], frames)

    # 4. écriture et mise en forme
    nb_deals = len([c for c in fonds.columns if str(c).startswith("Deal_")]) // 5
    rvc.write_excel(sortie, frames, nb_deals)

    # 5. restitution
    apres = {k: int(sum(0 if est_vide(x) else 1 for col in v.columns for x in v[col]))
             for k, v in frames.items()}
    print(f"Fichier produit : {sortie}")
    print(f"Entrée          : {entree}")
    print()
    print(f"Marqueurs « n.a » neutralisés : {nettoyees} cellules")
    print(f"Cellules vides complétées par la passe web : {remplies}")
    print(f"Valeurs corrigées : {corrigees}")
    print(f"Sites officiels renseignés : {n_sites}")
    print(f"Véhicules mis à jour : {len(touchees)}")
    for t in touchees:
        print(f"  - {t}")
    print()
    print("Taux de remplissage (onglet Fonds) :")
    for col in ("AuM (M€)", "Montant levé (M€)", "Millésime", "Ticket min (M€)", "TVPI", "DPI",
                "IRR (TRI)", "Site web"):
        n = sum(0 if est_vide(v) else 1 for v in fonds[col])
        print(f"  - {col} : {100.0 * n / len(fonds):.1f} %  ({n}/{len(fonds)})")
    print()
    print("Lignes par onglet :", {k: len(v) for k, v in frames.items()})
    print("Cellules renseignées avant / après :",
          {k: f"{avant[k]} -> {apres[k]}" for k in frames})
    return 0


if __name__ == "__main__":
    sys.exit(main())
