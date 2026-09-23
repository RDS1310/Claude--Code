#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fusionner_contacts_vc.py — Adds key contacts for the 56 traditional VC firms
================================================================================

Same pattern as fusionner_contacts_cvc.py, applied to the non-CVC firms of the
base (VC indépendant). Lean research (Sonnet, 3 background agents, compact JSON
output, one contact per firm — fintech/insurtech partner when identifiable, else
Managing/General Partner or Founder). No email searched (left for the user to
fill in separately). 4 firms out of 56 came back with no reliable contact
(BPI France, Committed Capital, Mash VC, The49) and are left without one rather
than guessed. The 3 individual "fund" entries (business angels investing in
their own name) get themselves as the contact.

Input  : VC_Database_Standardisee_v19.xlsx
Output : VC_Database_Standardisee_v20.xlsx

Usage: python3 fusionner_contacts_vc.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ENTREE = BASE_DIR / "VC_Database_Standardisee_v19.xlsx"
SORTIE = BASE_DIR / "VC_Database_Standardisee_v20.xlsx"

# "Nom du fonds" (as it appears in the Fonds sheet) -> (name, title, linkedin_or_None)
CONTACTS: dict[str, tuple[str, str, str | None]] = {
    "115K": ("Damien Launoy", "Managing Director / Partner", "https://www.linkedin.com/in/damien-launoy/"),
    "13books Capital (ex-Element Ventures)": ("Steve Gibson", "Co-Founder & Partner", None),
    "360 Capital Partners": ("Fausto Boni", "Founder & General Partner", None),
    "365.fintech": ("Rudolf Vrabel", "Managing Director", None),
    "Accel Partners": ("Luca Bocchio", "Partner (Fintech/Insurtech, Europe)", None),
    "Alven Capital": ("Charles Letourneur", "Co-Founder & Managing Partner", None),
    "Aris Occitanie VC": ("Romain Fonade", "Director", "https://fr.linkedin.com/in/romain-fonade-65853841"),
    "Astorya.vc": ("Florian Graillot", "Founding Partner (InsurTech)",
                   "https://www.linkedin.com/in/florian-graillot-56a1aba/"),
    "Atlantic Vantage Point": ("François Robinet", "Founder & Managing Partner", None),
    "BlackFin Capital Partners": ("Michele Foradori", "Managing Director, BlackFin Tech (Fintech/Insurtech)",
                                   "https://fr.linkedin.com/in/micheleforadori"),
    "Blast.Club": ("Anthony Bourbon", "Founder", None),
    "Breega Capital": ("Ben Marrel", "Co-Founder & Managing Partner", None),
    "Cadence Growth Capital": ("Leonard Clemens", "Managing Partner & Co-Founder (pertinence insurtech non "
                               "établie — société généraliste growth-equity)", None),
    "Cathay": ("Nicolas du Cray", "Partner (Fintech/Insurtech), Cathay Innovation",
               "https://sg.linkedin.com/in/nducray"),
    "Concentric": ("Kjartan Rist", "Founding Partner", None),
    "EOS Venture": ("Sam Evans", "Founding Partner (InsurTech)", "https://uk.linkedin.com/in/sam-evans-1b429237"),
    "Elaia Partners": ("Alexis Frentz", "Partner (Fintech)", "https://fr.linkedin.com/in/alexisfrentz"),
    "Elevation Capital Partners": ("Benjamin Cohen", "Co-founder & Managing Partner", None),
    "Eurazeo": ("Matthieu Baret", "Managing Partner, Venture (fonds Insurtech)", None),
    "Evolem Start": ("Thomas Rival", "Head of Evolem Start", None),
    "Founders Future VC": ("Marc Menasé", "Founder & Managing Partner", None),
    "Goldsmith Ventures": ("James Pringle", "Co-Founder", None),
    "ISAI": ("Christophe Raynaud", "Co-Founder & Venture Partner", None),
    "Index Venture": ("Jan Hammer", "Partner, Fintech/Insurtech (a mené le deal Alan)", None),
    "Insurtech Gateway": ("Stephen Brittain", "Co-Founder & Director", None),
    "Kima Venture Capital": ("Jean de La Rochebrochard", "Partner", None),
    "NewAlpha Asset Management": ("Laurent Langlais", "Managing Partner (VC fintech/insurtech)", None),
    "NewFund Capital": ("François Véron", "Co-Founder & Managing Partner", None),
    "Partech": ("Reza Malekzadeh", "Partner, Fintech & Insurtech", None),
    "Portage Venture": ("Hélène Falchier", "General Partner, Fintech", None),
    "Ring Capital": ("Nicolas Celier", "Co-Founder & Partner", None),
    "Samaipata Ventures": ("José del Barrio", "Co-Founder & CEO", None),
    "Schumpeter Ventures": ("Udo Bröskamp", "Co-Founder & Managing Partner", None),
    "Seed X Liechtenstein": ("Mathias Jaeggi", "CEO & Partner", None),
    "Seraphim Space": ("Mark Boggett", "CEO & Co-Founder / Managing Partner", None),
    "Serena Capital": ("Sébastien Le Roy", "Partner (FinTech & Insurtech)", None),
    "Sharpstone Capitale": ("Germain Gaschet", "Founding Partner", "https://fr.linkedin.com/in/germain-gaschet"),
    "South East Angels": ("Kristina Pereckaite", "Founder", "https://www.linkedin.com/in/kristinapereckaite/"),
    "Start Venture": ("João Menano", "Venture Partner", "https://www.linkedin.com/in/joaomenano/"),
    "Step Venture": ("Michele Novelli", "Partner (ex-Zest Fintech/Insurtech)", None),
    "TSP Ventures": ("Chris Smith", "Founder & CEO", "https://uk.linkedin.com/in/chris-smith-a38554a"),
    "Tenity": ("Andreas Iten", "CEO & Co-Founder (Managing Partner)", None),
    "The Family": ("Alice Zagury", "Co-Founder & CEO", None),
    "The Moon Venture": ("Sébastien Fertier", "Founder", None),
    "The Net Street Capital": ("Jorge Blasco", "Managing Partner", "https://www.linkedin.com/in/jorgeblasco/"),
    "TomCat": ("Thomas Gruederich", "Co-Founder & Managing Director", None),
    "TrueSight Ventures": ("Hampus Monthan Nordenskjöld", "Founding Partner", None),
    "Truffle Capital": ("Bernard-Louis Roques", "General Partner & Co-Founder (Fintech/Insurtech)",
                        "https://www.linkedin.com/in/blroques/"),
    "Venpace": ("Ingo Küpper", "Co-Founder / Managing Partner", None),
    "White Star Capital": ("Eric Martineau-Fortin", "Founder & Managing Partner", None),
    "XAnge Capital": ("Guillaume Meulle", "Managing Partner (FinTech)", None),
    "speedInvest": ("Stefan Klestil", "Partner (Fintech & Insurtech)", None),
    # Business angels investing in their own name: they are the contact.
    "David Semmens": ("David Semmens", "Business angel (investissement personnel)", None),
    "Mohammad Hossein Tavangar": ("Mohammad Hossein Tavangar", "Business angel (investissement personnel)", None),
    "Patrice Fleurquin": ("Patrice Fleurquin", "Business angel (investissement personnel)", None),
    # No reliable contact found: BPI France, Committed Capital, Mash VC, The49 — left absent on purpose.
}


def main() -> int:
    fonds = pd.read_excel(ENTREE, sheet_name="Fonds")
    part = pd.read_excel(ENTREE, sheet_name="Participations")
    tours = pd.read_excel(ENTREE, sheet_name="Tours_de_table")
    dico = pd.read_excel(ENTREE, sheet_name="Dictionnaire")
    ano = pd.read_excel(ENTREE, sheet_name="Anomalies")

    if "Contact_01_Nom" not in fonds.columns:
        for col in ("Contact_01_Nom", "Contact_01_Titre", "Contact_01_LinkedIn", "Contact_01_Email",
                    "Contact_02_Nom", "Contact_02_Titre", "Contact_02_LinkedIn", "Contact_02_Email"):
            fonds[col] = None

    n_firms, n_rows = 0, 0
    for nom_fonds, (nom, titre, linkedin) in CONTACTS.items():
        mask = (fonds["Nom du fonds"] == nom_fonds) & fonds["Contact_01_Nom"].isna()
        if not mask.any():
            print(f"AVERTISSEMENT : « {nom_fonds} » introuvable (ou déjà pourvu) dans la base.")
            continue
        n_firms += 1
        n_rows += int(mask.sum())
        fonds.loc[mask, "Contact_01_Nom"] = nom
        fonds.loc[mask, "Contact_01_Titre"] = titre
        fonds.loc[mask, "Contact_01_LinkedIn"] = linkedin

    with pd.ExcelWriter(SORTIE, engine="openpyxl") as writer:
        fonds.to_excel(writer, sheet_name="Fonds", index=False)
        part.to_excel(writer, sheet_name="Participations", index=False)
        tours.to_excel(writer, sheet_name="Tours_de_table", index=False)
        dico.to_excel(writer, sheet_name="Dictionnaire", index=False)
        ano.to_excel(writer, sheet_name="Anomalies", index=False)

    print(f"Classeur produit : {SORTIE}")
    print(f"Sociétés de gestion pourvues : {n_firms} / {len(CONTACTS)}  (lignes Fonds mises à jour : {n_rows})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
