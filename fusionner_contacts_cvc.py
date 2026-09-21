#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fusionner_contacts_cvc.py — Ajoute les contacts clés (Partner/MD) des 32 CVC
================================================================================

Intègre la recherche (WebSearch, 3 lots parallèles, méthodologie anti-fabrication
identique) sur les Partner/Managing Director/Head of Ventures des 32 véhicules CVC
déjà suivis dans la base. Aucune adresse email n'a été recherchée (consigne
explicite) — le champ Contact_XX_Email reste vide, à renseigner par l'utilisateur
(Apollo ou autre) dans un second temps. Là où le contact identifié a quitté son
poste ou où la situation est en transition, c'est noté explicitement dans le titre
plutôt que tu (l'agent) ne choisisse un successeur non confirmé.

2 véhicules sur 32 n'ont aucun contact fiable et actuel identifié (Guardian
Strategic Ventures/GIS — successeur du MD parti en 2018-19 non trouvé ;
Transamerica Ventures — confirmé inactif, portefeuille cédé en 2021) : laissés
sans contact plutôt que d'inventer un nom.

Entrée  : VC_Database_Standardisee_v17.xlsx
Sortie  : VC_Database_Standardisee_v18.xlsx

Usage : python3 fusionner_contacts_cvc.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v17.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v18.xlsx"

# Fonds_ID -> liste de (nom, titre, linkedin_ou_None) — jusqu'à 2 par véhicule.
CONTACTS: dict[str, list[tuple[str, str, str | None]]] = {
    "commerzventures_commerzventures-fonds-i-a-iii": [
        ("Patrick Meisberger", "Managing Partner & Co-fondateur", "https://de.linkedin.com/in/patrickmeisberger"),
        ("Stefan Tirtey", "Managing Partner & Co-fondateur", "https://de.linkedin.com/in/stirtey"),
    ],
    "ethias-ventures_ethias-ventures": [
        ("Marius Declerck", "Head of Ethias Ventures (titre exact incertain selon les sources)",
         "https://be.linkedin.com/in/marius-declerck"),
    ],
    "helsana-healthinvest_helsana-healthinvest": [
        ("Ralf Molitor", "Managing Director", "https://ch.linkedin.com/in/ralf-molitor-38152722"),
        ("Dietrich Aumann", "Investment Manager", None),
    ],
    "howden-ventures_howden-ventures": [
        ("Tom Hoad", "Head of Howden Ventures", "https://www.linkedin.com/in/tom-hoad-9b538a38/"),
        ("Luke Hakes", "Partner (rejoint mars 2025, ex-Octopus Ventures)", "https://uk.linkedin.com/in/lukehakes"),
    ],
    "insurtech-capital_insurtech-capital-groupe-apicil": [
        ("Minh Q. Tran", "Fondateur & Managing Partner, Mandalore Partners / Odysseus Alternative Ventures "
         "(société de gestion du fonds, pas salarié direct Apicil)", "https://fr.linkedin.com/in/minhtran"),
    ],
    "macif-innovation_macif-innovation": [
        ("François-Xavier Marchand", "VC Investment @ Macif Innovation",
         "https://fr.linkedin.com/in/fran%C3%A7ois-xavier-marchand-2b2b2826"),
        ("Mira Le Lay", "Directrice Stratégie et Performance (supervise l'Innovation depuis le 01/03/2026)",
         "https://fr.linkedin.com/in/mira-le-lay"),
    ],
    "open-cnp_open-cnp": [
        ("Alexandra Pailhes, CFA", "Head of Investments @ Open CNP",
         "https://fr.linkedin.com/in/alexandra-pailhes-cfa-1718a9"),
        ("Anas Tazlaoui", "Startup investor @ Open CNP", "https://fr.linkedin.com/in/anastazla"),
    ],
    "scor-ventures_scor-ventures": [
        ("Will Thorne", "Managing Partner", "https://uk.linkedin.com/in/willthorne"),
    ],
    "uniqa-ventures_uniqa-ventures": [
        ("Christian Huber", "Head of Private Markets, UNIQA Capital Markets (reprise intérimaire depuis la "
         "dissolution de la marque UNIQA Ventures, nov. 2024)", "https://at.linkedin.com/in/christian-huber-b2161a5b"),
    ],
    "nca-next-commerce-accelerator_nca-next-commerce-accelerator": [
        ("Thorsten Wittmütz", "Managing Partner / Co-fondateur", "https://de.linkedin.com/in/thorstenwittmuetz"),
        ("Christoph Schepan", "Managing Partner", "https://de.linkedin.com/in/schepan"),
    ],
    "kickstart-innovation_kickstart-innovation": [
        ("Katka Letzing", "CEO & Co-Founder", "https://ch.linkedin.com/in/katka-letzing-b59955382"),
    ],
    "allianz-x_allianz-x": [
        ("Dr. Nazim Cetin", "CEO, Allianz X (depuis 2017)", "https://de.linkedin.com/in/dr-nazim-cetin-b602a557"),
    ],
    "maif-avenir_maif-avenir": [
        ("Mohamed Abdesslam", "Managing Partner, Ternel (ex-MAIF Avenir)",
         "https://www.linkedin.com/in/mohamed-abdesslam-51b4984/"),
        ("Timothée Poulain", "Partner, Ternel", "https://fr.linkedin.com/in/timotheepoulain"),
    ],
    "munich-re-ventures_munich-re-ventures": [
        ("Jacqueline LeSage Krause", "Managing Director", None),
        ("Oshri Kaplan", "Managing Director, Insurtech, Cybersecurity & Privacy",
         "https://www.linkedin.com/in/kaplanoshri"),
    ],
    "msad-ventures_msad-ventures": [
        ("Jon Soberg", "CEO & Managing Partner", "https://www.linkedin.com/in/jonsoberg/"),
        ("Jack (Tasuku) Toyama", "President, Managing Partner & Director (depuis mars 2021)",
         "https://www.linkedin.com/in/jacktoyama/"),
    ],
    "sompo_digital-lab-light-vortex": [
        ("Albert B. Chu", "CEO, Sompo Digital Lab", "https://www.linkedin.com/in/albertbchu/"),
        ("Koichi Narasaki", "CEO, Sompo Light Vortex (transition de leadership possible vers Atsushi Miya, "
         "non confirmée — à vérifier)", None),
    ],
    "generali-ventures_generali-ventures": [
        ("Danilo Raponi", "Group Head of Innovation, Generali ; Managing Director du Generali Innovation Fund",
         None),
    ],
    "aviva-ventures_aviva-ventures": [
        ("Ben Luckett", "Managing Director – Venture & Strategic Capital, Aviva Investors ; fondateur d'Aviva "
         "Ventures (2015)", "https://uk.linkedin.com/in/benluckett1"),
    ],
    "zurich-insurance-group_investissements-strategiques": [
        ("Cara Morton", "Ex-CEO, Zurich Global Ventures (a quitté ce poste en janvier 2026 ; successeur non "
         "identifié)", None),
    ],
    "tokio-marine-future-fund_tokio-marine-future-fund": [
        ("Yoshi Yoshida", "Corporate Venturing Lead (Silicon Valley), Tokio Marine Holdings",
         "https://www.linkedin.com/in/yoshi-yoshida-80a7aa171/"),
        ("Steve Pretre", "Managing Partner, World Innovation Lab (WiL) — gère le fonds pour le compte de "
         "Tokio Marine", None),
    ],
    "nationwide-ventures_nationwide-ventures": [
        ("Erik Ross", "Managing Partner, Nationwide Ventures (fondateur, depuis 2016)",
         "https://www.linkedin.com/in/esross/"),
    ],
    "massmutual-ventures_massmutual-ventures": [
        ("Doug Russell", "Managing Partner & Head of MassMutual Ventures (depuis 2014)",
         "https://www.linkedin.com/in/douglas-russell-9217a018/"),
    ],
    "new-york-life-ventures_nyl-ventures": [
        ("Joel Albarella", "Senior Vice President & Head of NYL Ventures (fondateur, 2012)",
         "https://www.linkedin.com/in/joelalbarella/"),
    ],
    "optum-ventures_optum-ventures": [
        ("Laura Veroneau", "Managing Partner", "https://www.linkedin.com/in/laura-veroneau/"),
        ("Larry Renfro", "Managing Partner (ex-CEO Optum, Vice Chairman UnitedHealth Group)", None),
    ],
    "qbe-ventures_qbe-ventures": [
        ("James Orchard", "Chief Executive Officer", "https://au.linkedin.com/in/james-orchard-10839612"),
        ("Daniel Wypler", "Partner & Global Head of Investments", "https://www.linkedin.com/in/wypler/"),
    ],
    "iag-firemark-ventures_iag-firemark-ventures": [
        ("Scott Gunther", "General Partner", "https://www.linkedin.com/in/scottdgunther/"),
        ("Scott Bishop", "EGM Innovation & Ventures, Firemark Collective", None),
    ],
    "intact-ventures_intact-ventures": [
        ("Justin Smith-Lorenzetti", "Founder & Managing Director",
         "https://www.linkedin.com/in/justin-smith-lorenzetti/"),
    ],
    "liberty-mutual-strategic-ventures_lmsv": [
        ("Russ MacTough", "Managing Director / Managing Partner", "https://www.linkedin.com/in/rmactough/"),
    ],
    "achmea-innovation-fund_achmea-innovation-fund": [
        ("Katharina Maass", "Head of Achmea Innovation Fund / Fund Manager (depuis mars 2020)",
         "https://www.linkedin.com/in/katharina-maass-809401/"),
        ("Henrieke Hoftijzer", "Investment Director", "https://nl.linkedin.com/in/henriekehoftijzer"),
    ],
    # Guardian Strategic Ventures (GIS Strategic Ventures LLC) et Transamerica Ventures :
    # aucun contact actuel fiable identifié — volontairement absents de ce dict.
}


def main() -> int:
    fonds = pd.read_excel(ENTREE, sheet_name="Fonds")
    part = pd.read_excel(ENTREE, sheet_name="Participations")
    tours = pd.read_excel(ENTREE, sheet_name="Tours_de_table")
    dico = pd.read_excel(ENTREE, sheet_name="Dictionnaire")
    ano = pd.read_excel(ENTREE, sheet_name="Anomalies")

    for col in ("Contact_01_Nom", "Contact_01_Titre", "Contact_01_LinkedIn", "Contact_01_Email",
                "Contact_02_Nom", "Contact_02_Titre", "Contact_02_LinkedIn", "Contact_02_Email"):
        fonds[col] = None

    n_fonds_avec_contact = 0
    n_contacts = 0
    for fid, contacts in CONTACTS.items():
        mask = fonds["Fonds_ID"] == fid
        if not mask.any():
            print(f"AVERTISSEMENT : Fonds_ID introuvable dans la base : {fid}")
            continue
        n_fonds_avec_contact += 1
        for i, (nom, titre, linkedin) in enumerate(contacts[:2], start=1):
            fonds.loc[mask, f"Contact_{i:02d}_Nom"] = nom
            fonds.loc[mask, f"Contact_{i:02d}_Titre"] = titre
            fonds.loc[mask, f"Contact_{i:02d}_LinkedIn"] = linkedin
            n_contacts += 1

    nouvelle_ligne_dico = {
        "Colonne": "Contact_01/02_Nom / Titre / LinkedIn / Email",
        "Description": (
            "Contact(s) clé(s) du véhicule (Partner/Managing Director/Head of Ventures), échantillon des 32 "
            "CVC. Recherche WebSearch, aucune adresse email recherchée — à renseigner séparément (Apollo ou "
            "autre)."
        ),
        "Valeurs possibles": "Texte libre",
    }
    if {"Colonne", "Description", "Valeurs possibles"} <= set(dico.columns):
        dico = pd.concat([dico, pd.DataFrame([nouvelle_ligne_dico])], ignore_index=True)

    with pd.ExcelWriter(SORTIE, engine="openpyxl") as writer:
        fonds.to_excel(writer, sheet_name="Fonds", index=False)
        part.to_excel(writer, sheet_name="Participations", index=False)
        tours.to_excel(writer, sheet_name="Tours_de_table", index=False)
        dico.to_excel(writer, sheet_name="Dictionnaire", index=False)
        ano.to_excel(writer, sheet_name="Anomalies", index=False)

    print(f"Classeur produit : {SORTIE}")
    print(f"Véhicules avec au moins un contact : {n_fonds_avec_contact} / {len(CONTACTS)} tentés (32 CVC ciblés)")
    print(f"Contacts ajoutés : {n_contacts}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
