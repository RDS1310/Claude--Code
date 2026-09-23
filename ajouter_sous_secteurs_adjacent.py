"""
Documente dans Anomalies le résultat de la sous-classification du bucket
"InsurTech (adjacent)" (graphique Tendances, 23/09/2026) : la logique de
classement elle-même vit dans generer_site_vc.py (ADJACENT_SOUS_SECTEUR),
ce script ne touche que la traçabilité côté classeur source (Anomalies),
sans modifier les champs "Secteur" existants dans Participations.
"""
import shutil
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
SRC = BASE_DIR / "VC_Database_Standardisee_v22.xlsx"
DST = BASE_DIR / "VC_Database_Standardisee_v23.xlsx"

shutil.copyfile(SRC, DST)

an = pd.read_excel(DST, sheet_name="Anomalies")
nums = an["Anomalie_ID"].str.extract(r"(\d+)")[0]
next_num = int(nums.max()) + 1
next_id = f"ANO-{next_num:04d}"

row = {
    "Anomalie_ID": next_id,
    "Type d'anomalie": "Sous-classification sectorielle (graphique Tendances)",
    "Onglet source": "Participations",
    "Table ou bloc source": "Sous-classement manuel + WebSearch ciblé (23/09/2026)",
    "Ligne source": None,
    "Entité concernée": "57 start-ups taguées « InsurTech (adjacent) »",
    "Champ concerné": "Verticale assurance (nouveau, usage graphique uniquement — Secteur non modifié)",
    "Valeur source": "InsurTech (adjacent) — bucket unique, sans sous-verticale",
    "Valeur retenue": (
        "Réparties en 6 sous-verticales (Cyber, Climat & risques catastrophes, Santé & prévoyance, "
        "Télématique/mobilité & prévention IoT, Emploi & avantages sociaux, Data/IA & distribution) "
        "pour le graphique Tendances de Radar Insurtech VC ; mapping en dur dans generer_site_vc.py "
        "(ADJACENT_SOUS_SECTEUR), fondé sur les Commentaires déjà collectés, complété par une recherche "
        "WebSearch ciblée pour 7 entrées insuffisamment documentées."
    ),
    "Candidats éventuels": None,
    "Score de similarité": None,
    "Niveau de confiance": "Moyen",
    "Traitement appliqué": (
        "4 des 58 lignes examinées (MuchBetter.ai, Gretel, hypt., Value Factory) se révèlent, à la lecture "
        "de leur description réelle, être des sociétés horizontales sans verticale assurance identifiable "
        "(EdTech/RH, data synthétique générique rachetée par Nvidia, SaaS avis clients horizontal, société "
        "non identifiée) : le tag « adjacent » existant semble provenir du fonds investisseur plutôt que de "
        "l'activité de la société elle-même. Exclues du graphique Tendances (comptabilisées dans les "
        "participations exclues, motif affiché) plutôt que forcées dans une sous-verticale inexacte. Le champ "
        "« Secteur » d'origine n'a pas été modifié dans Participations — seule la couche d'agrégation du "
        "graphique en tient compte ; une revue ultérieure du tag source de ces 4 lignes reste à faire."
    ),
    "Commentaire": (
        "Sous-verticales par volume (lignes) : Data/IA & distribution 19, Télématique/mobilité 11, "
        "Santé & prévoyance 9, Cyber 6, Emploi & avantages sociaux 5, Climat & risques catastrophes 4, "
        "hors périmètre (exclues) 4."
    ),
}

an = pd.concat([an, pd.DataFrame([row])], ignore_index=True)

with pd.ExcelWriter(DST, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    an.to_excel(writer, sheet_name="Anomalies", index=False)

print(f"OK -> {DST.name} ; Anomalie ajoutée : {next_id} ; lignes Anomalies : {len(an)}")
