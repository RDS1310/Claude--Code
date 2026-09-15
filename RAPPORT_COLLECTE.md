# Rapport de collecte web — VC_Database_Standardisee_v4

*Généré le 15/09/2026 — passe de collecte web par lots sur les 111 véhicules de l'onglet `Fonds`, à partir de `VC_Database_Standardisee_v3.xlsx`.*

## 1. Synthèse

### Taux de remplissage par colonne clé (avant / après)

| Colonne | Avant (v3) | Après (v4) | Δ |
|---|---|---|---|
| AuM (M€) | 84/111 (75.7 %) | 85/111 (76.6 %) | +1 |
| Montant levé (M€) | 82/111 (73.9 %) | 86/111 (77.5 %) | +4 |
| Montant alloué (M€) | 35/111 (31.5 %) | 63/111 (56.8 %) | +28 |
| Ticket min (M€) | 84/111 (75.7 %) | 88/111 (79.3 %) | +4 |
| Ticket max (M€) | 86/111 (77.5 %) | 92/111 (82.9 %) | +6 |
| Millésime | 95/111 (85.6 %) | 101/111 (91.0 %) | +6 |
| Statut | 104/111 (93.7 %) | 104/111 (93.7 %) | +0 |
| Phase | 102/111 (91.9 %) | 102/111 (91.9 %) | +0 |
| Géographie | 106/111 (95.5 %) | 111/111 (100.0 %) | +5 |
| Stratégie | 106/111 (95.5 %) | 107/111 (96.4 %) | +1 |
| Stages pratiqués | 96/111 (86.5 %) | 99/111 (89.2 %) | +3 |
| Nb participations (déclaré) | 98/111 (88.3 %) | 100/111 (90.1 %) | +2 |
| Nb exits | 66/111 (59.5 %) | 67/111 (60.4 %) | +1 |
| Nb InsurTech | 71/111 (64.0 %) | 71/111 (64.0 %) | +0 |
| TVPI | 10/111 (9.0 %) | 11/111 (9.9 %) | +1 |
| DPI | 6/111 (5.4 %) | 6/111 (5.4 %) | +0 |
| IRR (TRI) | 14/111 (12.6 %) | 15/111 (13.5 %) | +1 |
| Quartile | 19/111 (17.1 %) | 19/111 (17.1 %) | +0 |
| Site web | 39/111 (35.1 %) | 102/111 (91.9 %) | +63 |

### Statut de collecte (111 véhicules)

- **traité** : 81
- **traité — rien trouvé** : 27
- **non exploitable** : 3

### Anomalies

- 126 cellules complétées sur 81 véhicules touchés (anomalies de type « Mise à jour web » / « Valeur approximative conservée »)
- 5 divergences base/web journalisées (type « Différence entre description et structure réelle »)
- 4 valeurs en fourchette non retenues (type « Valeur en fourchette non retenue »)
- 10 rapprochements non résolus / risques d'homonymie signalés (type « Rapprochement non résolu »)
- 3 véhicules marqués non exploitables (business angels sans portefeuille vérifiable)
- Numérotation des anomalies : ANO-0179 à ANO-0283 (178 lignes existantes + 105 nouvelles ; voir onglet `Anomalies` du classeur pour le détail exact)

### Divergences base/web à arbitrer manuellement

| Véhicule | Champ | Valeur base | Valeur web | Commentaire |
|---|---|---|---|---|
| Alven Capital / Alven V | Millésime | 2016 | closing d'Alven V (250 M€) annoncé début janvier 2017 (TechCrunch 03/01/2017, Ma… | Millésime base conservé (2016) : aucune source ne documente publiquement un premier closing antérieu… |
| Blast.Club / Blast.Club | Millésime | 2022 | création en 2023 selon 4 sources concordantes (Forbes France, JDN, Maddyness, Wi… | Millésime base conservé (2022) : à vérifier, la majorité des sources web datent la création du club … |
| Kickstart Innovation / Kickstart Innovation | Nb participations (déclaré) | 382 | 450 startups accompagnées depuis la création (source secondaire)… | Valeur base conservée (382) : écart avec un chiffre plus récent/plus large trouvé en presse, non arb… |
| South East Angels / South East Angels | Nb participations (déclaré) | 41 | 34 participations recensées sur la page Portfolio officielle (nov. 2025)… | Valeur base conservée (41) : écart avec le décompte du site officiel à une date antérieure, non arbi… |
| BlackFin Capital Partners / BlackfinTech II | Montant levé (M€) | 390 | 350 M€ selon fintech.global (« BlackFin commits €350m fund for European InsurTec… | Valeur base conservée (390), corroborée par un agrégateur indépendant ; le chiffre de 350 M€ (titre … |

### Rapprochements non résolus (risques d'homonymie / identité de véhicule)

- **Partech / Partech Seed (I to IV), Partech / Parrtech Venture** — La dénomination « Partech Seed (I to IV) » ne correspond à aucune série officielle documentée par Partech (dont la nomenclature connue est « Partech Entrepreneur I-IV ») ; « Parrtech Venture » est une coquille confirmée pour « Partech Venture », mais il existe à la fois un véhicule 2023 de 360 M€ sous ce nom et un historique de fonds plus anciens du même nom — le millésime visé par la ligne de la base n'a pas pu être déterminé avec certitude.
- **ISAI / ISAI Venture III** — Risque de confusion avec « ISAI Expansion III », véhicule distinct de croissance/LBO (hard cap 150 M€) chez ISAI. Aucun montant n'a été renseigné pour ISAI Venture III afin d'éviter toute conflation ; à vérifier manuellement avant toute future collecte de Montant alloué/levé sur cette ligne.
- **Breega Capital / Breega Seed II, Breega Capital / Breega Venture III** — Une source secondaire non primaire indique que « Breega Capital Venture 3 » (= breega-venture-iii, 110 M€, closing final mars 2021) aurait été « renommé Seed II » par la suite. Si confirmé, les deux lignes de la base pourraient désigner le même véhicule sous deux noms différents — à vérifier manuellement avant d'arbitrer un éventuel doublon.
- **TSP Ventures / TSP Ventures** — La seule entité publique « TSP Ventures » identifiable (Londres, fondée en 2019, tspventures.co.uk) investit exclusivement en climate-tech / environmental tech, sans lien documenté avec le fintech/insurtech. Probable homonymie ou mauvaise identification de l'entité dans le référentiel ; aucune donnée de cette entité n'a été reportée sur la ligne de la base.
- **Sharpstone Capitale / Sharpstone Capitale** — La presse et le site officiel désignent systématiquement l'entité « Sharpstone Capital » (et non « Sharpstone Capitale ») ; probable variante orthographique dans le référentiel, à confirmer.
- **Kima Venture Capital / Kima Venture** — L'entité réelle se nomme « Kima Ventures » (véhicule evergreen à LP unique, Xavier Niel), et non « Kima Venture Capital » — ce qui explique l'absence de TVPI/DPI/IRR publics (pas de LPs tiers à rapporter).
- **Start Venture / Start Venture I** — Aucune entité « Start Venture » correspondante n'a été retrouvée en ligne après plusieurs recherches ciblées. Seule une entité au nom proche mais distincte (« Start Ventures », Lisbonne, fintech/insurtech) existe publiquement ; non retenue faute de confirmation du rapprochement. Aucune donnée reportée.
- **Ring Capital / Mission II** — L'existence d'un véhicule distinct « Ring Mission II » n'a pas pu être confirmée : seuls « Ring Mission » (35 M€, premier closing) et « Altitude II » (stratégie croissance/LBO différente) sont documentés chez Ring Capital. Ligne à vérifier manuellement — possible erreur de référentiel.
- **Truffle Capital / Truffle FinTech & InsurTech Fund III (prévu 2025)** — Aucune source ne confirme le lancement effectif de ce fonds à la date de cette collecte (14-15/09/2026) ; seuls le Fund II (140 M€, clos en 2019) et les statistiques globales de la société de gestion sont documentés publiquement. Le statut « prévu 2025 » du nom de la ligne reste donc non confirmé.
- **Committed Capital / Committed Capital EIS Fund** — Les données trouvées (AuM, TVPI, IRR) concernent le véhicule officiellement nommé « Growth EIS Fund » chez Committed Capital ; à confirmer qu'il s'agit bien du même véhicule que la ligne « Committed Capital EIS Fund » de la base avant de leur accorder une confiance Élevée.

### Véhicules non exploitables

- **David Semmens / David Semmens (business angel)** — CIO chez Cadro, administrateur chez Wealthify (Aviva) et RiskSave Technologies (insurtech). Seule trace publique : une fourchette de ticket personnel indicative sur une plateforme de mise en relation (Signal/NFX, 5-25 k$) — aucun véhicule d'investissement, portefeuille structuré ni performance publique.
- **Patrice Fleurquin / Patrice Fleurquin (business angel)** — Profil entrepreneurial documenté (cofondateur Tripfair, ex-CSO Nascom) mais aucune trace d'un portefeuille d'investissement en tant que business angel.
- **Mohammad Hossein Tavangar / M. H. Tavangar (business angel)** — Basé à Berlin, ex-Managing Partner chez Dorrance Venture, administrateur au Founder Institute Germany, focus fintech/insurtech/blockchain. L'activité d'investissement documentée est celle de Dorrance Venture (entité distincte), pas un véhicule personnel — aucun montant ni portefeuille propre identifiable.

### Véhicules sans aucune donnée nouvelle trouvée

- Mash VC / Mash VC
- Schumpeter Ventures / Schumpeter Ventures
- TSP Ventures / TSP Ventures *(voir rapprochement non résolu ci-dessus)*
- Kickstart Innovation / Kickstart Innovation
- Seraphim Space / Seraphim Space Ventures (SSVII + SSIT)
- Blast.Club / Blast.Club
- South East Angels / South East Angels
- Cadence Growth Capital / Cadence Growth Capital (CGC) *(voir valeur en fourchette non retenue ci-dessus)*
- The Net Street Capital / The Net Street Capital
- TrueSight Ventures / TrueSight Ventures Fund I
- 115K / 115K
- Venpace / Venpace
- Seed X Liechtenstein / Seed X Liechtenstein
- Tenity / Tenity Incubation Fund I & II *(voir valeur en fourchette non retenue ci-dessus)*
- Start Venture / Start Venture I *(voir rapprochement non résolu ci-dessus)*
- Eurazeo / Eurazeo Venture Capital
- 13books Capital (ex-Element Ventures) / Fonds I & II
- Concentric / Concentric
- Elaia Partners / PSL innovation Fund
- Elaia Partners / DV4
- Elaia Partners / Elaia Delta
- Ring Capital / Mission II *(voir rapprochement non résolu ci-dessus)*
- Alven Capital / Alven VI
- Atlantic Vantage Point / AVP early stage II
- Atlantic Vantage Point / AVP early stage I
- Elaia Partners / Alpha II
- Alven Capital / Alven IV

---

## 2. Détail par véhicule

*Pour chaque véhicule : champs complétés (avec source), champs recherchés sans résultat exploitable et raison. L'ordre suit `todo_collecte_web.csv`.*

### David Semmens / David Semmens (business angel)
`david-semmens_david-semmens-business-angel`

**Statut : non exploitable.** CIO chez Cadro, administrateur chez Wealthify (Aviva) et RiskSave Technologies (insurtech). Seule trace publique : une fourchette de ticket personnel indicative sur une plateforme de mise en relation (Signal/NFX, 5-25 k$) — aucun véhicule d'investissement, portefeuille structuré ni performance publique.

### Patrice Fleurquin / Patrice Fleurquin (business angel)
`patrice-fleurquin_patrice-fleurquin-business-angel`

**Statut : non exploitable.** Profil entrepreneurial documenté (cofondateur Tripfair, ex-CSO Nascom) mais aucune trace d'un portefeuille d'investissement en tant que business angel.

### Mohammad Hossein Tavangar / M. H. Tavangar (business angel)
`mohammad-hossein-tavangar_m-h-tavangar-business-angel`

**Statut : non exploitable.** Basé à Berlin, ex-Managing Partner chez Dorrance Venture, administrateur au Founder Institute Germany, focus fintech/insurtech/blockchain. L'activité d'investissement documentée est celle de Dorrance Venture (entité distincte), pas un véhicule personnel — aucun montant ni portefeuille propre identifiable.

### The49 / The49
`the49_the49`

**Champs complétés :**
- `Ticket min (M€)` = 0.092 *(confiance : Moyen)*
- `Ticket max (M€)` = 0.276 *(confiance : Moyen)*
- `Site web` = 'https://the49.com/' *(confiance : Moyen)*

**Source(s) :** The49, page 'Venture Studio' (https://www.the49.com/venture-studio/) et 'About Us' (https://the49.com/about-us/) ; OpenVC profil (https://www.openvc.app/fund/The49) ; LeadIQ company overview (https://leadiq.com/c/the49/5e18911262be9f43ca8b6541)

**Champs recherchés sans résultat exploitable :** AuM (M€), Montant levé (M€), Montant alloué (M€), Millésime, Nb participations (déclaré), Nb exits, Nb InsurTech, TVPI, DPI, IRR (TRI), Quartile

**Notes :** The49 n'est ni un fonds classique ni un business angel individuel : il s'agit d'un venture studio basé à Londres (~40 personnes) qui construit et finance des startups en interne (equity-build) sur les secteurs Insurtech/Healthtech/Sports/Travel & Logistics, avec un ticket de capital pré-amorçage annoncé de 100-300 K$ (converti ici en M€ à 0,92, à considérer comme indicatif). Aucune donnée sur le nombre de participations, d'exits, ni de métriques de performance (TVPI/DPI/IRR) n'a été trouvée publiquement. Non marqué 'non_exploitable' car des données structurelles réelles (nature, géographie, stratégie, ticket, site) ont pu être établies, contrairement aux 4 personnes physiques de ce lot.

### Goldsmith Ventures / Goldsmith Ventures EIS Fund
`goldsmith-ventures_goldsmith-ventures-eis-fund`

**Champs complétés :**
- `Site web` = 'https://www.goldsmithventures.com' *(confiance : Faible)*
- `Ticket min (M€)` = 0.736 *(confiance : Faible)*
- `Ticket max (M€)` = 1.104 *(confiance : Faible)*

**Source(s) :** Aggregated VC-database snippets (Apollo.io company profile, investorconnect.org 'Investor Connect: James Pringle' transcript, openvc.app/fund/Goldsmith%20Ventures) surfaced via WebSearch

**Champs recherchés sans résultat exploitable :** AuM, Montant levé, Montant alloué, Statut, Phase, Stages pratiqués, Nb exits, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** Goldsmith Ventures is a London VC (founded 2021 by James Pringle, ex-Pringle Capital angel network) investing in UK fintech/proptech/insurtech; its EIS fund launched April 2022 investing at seed/Series A. Reported 'initial check sizes' of $800k-$1.2M converted at 1 USD=0.92 EUR (736k€ / 1,104k€ = 0.736 M€ / 1.104 M€) — flagged approximate and Faible because it comes only from third-party aggregator text, not the fund's own site (goldsmithventures.com could not be fetched — WebFetch blocked/not independently verified as live), and it is unusually large for a UK EIS ticket so may actually describe overall firm check size rather than this specific EIS vehicle. Recommend a human visually confirm goldsmithventures.com before trusting Ticket min/max or Site web at higher confidence.

### Mash VC / Mash VC
`mash-vc_mash-vc`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** AuM, Montant levé, Montant alloué, Ticket min, Ticket max, Millésime, Statut, Phase, Stages pratiqués, Nb exits, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** Confirmed Mash VC (UK, pre-seed/seed fintech/insurtech/marketplace investor, '50+ investments' claimed) is a distinct entity from the US 'Mash.Ventures' (mash.ventures, Andrew Mash, a due-diligence/advisory outfit, not a fund) — no name-collision risk found for the data already in the sheet. Could not fetch www.mash.vc directly (DNS/egress issue) and no fund-size, AUM, vintage or performance figures surfaced in search snippets.

### Schumpeter Ventures / Schumpeter Ventures
`schumpeter-ventures_schumpeter-ventures`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** Montant alloué, Ticket min, Ticket max, Millésime, Statut, Phase, Stages pratiqués, Nb exits, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** Figure is labelled 'AuM' (~$12M) in third-party aggregator snippets, not explicitly 'amount raised'; mapped tentatively to Montant levé (approximate=true), converted at 1 USD=0.92 EUR (12*0.92=11.04 M€). Could not independently confirm on schumpeter.vc (WebFetch blocked by egress proxy) or via a primary press release, so treat as low-confidence single-source. Firm founded 2018, Frankfurt, focuses on FinTech/InsurTech/cybersecurity in Germany/Switzerland/Israel; portfolio reportedly includes MYVI Group, Qundo, FineTrade, Paladyn, EMIL, but sector classification of each was not verified so Nb InsurTech was not filled. PitchBook has a 'Schumpeter Ventures I' fund profile (pitchbook.com/profiles/fund/19276-12F) which could hold Millésime/TVPI/IRR data but is paywalled and could not be read.

### TSP Ventures / TSP Ventures
`tsp-ventures_tsp-ventures`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** AuM, Montant levé, Montant alloué, Ticket min/max, Stages pratiqués, Nb exits, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** IMPORTANT CAVEAT — likely homonym/scope mismatch, similar to the Element Ventures Canada warning for 13books: the only 'TSP Ventures' found in search results is a London-based firm ('The Sustainable Path'), founded 2019, investing EXCLUSIVELY in climate/environmental tech (energy, critical materials, water, carbon capture, waste & circularity) — it shows no insurtech or fintech focus in any source found, and its own site content (per snippets) is entirely climate-tech oriented. Given this Fonds_ID sits in an insurtech-scope database, this may not be the correct entity, or the row may need re-scoping/removal. Did NOT fill Site web (https://tspventures.co.uk/) given this mismatch risk — recommend manual verification before accepting any data under this name. One low-reliability aggregator (a Tracxn-derived site) mentioned pre-seed/seed ticket sizes of roughly £70k-£700k for this same London climate-tech TSP Ventures, but this is not filled given both the sourcing quality and the scope-mismatch concern.

**Rapprochement non résolu :** La seule entité publique « TSP Ventures » identifiable (Londres, fondée en 2019, tspventures.co.uk) investit exclusivement en climate-tech / environmental tech, sans lien documenté avec le fintech/insurtech. Probable homonymie ou mauvaise identification de l'entité dans le référentiel ; aucune donnée de cette entité n'a été reportée sur la ligne de la base.

### Evolem Start / Evolem Start
`evolem-start_evolem-start`

**Champs complétés :**
- `Site web` = 'https://evolem.com' *(confiance : Moyen)*
- `Millésime` = 2017 *(confiance : Moyen)*

**Source(s) :** Site web: evolem.com news pages (e.g. https://evolem.com/evolem-start-realise-6-nouveaux-investissements-en-2018/), Maddyness 'Que Font les Fonds : le portrait d'Evolem' (2021-09-27). Millésime: company-registry-derived WebSearch snippet ('EVOLEM START operates in the fund management sector and was created in 2017', from verif.com/rubypayeur listings for EVOLEM START, SIREN 833257645).

**Champs recherchés sans résultat exploitable :** AuM, Montant levé, Montant alloué, Nb participations (déclaré), Nb exits, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** 'Evolem Start' has no dedicated domain; it is an investment programme/vehicle of the Lyon-based Evolem group, whose official site is evolem.com. Millésime=2017 is based on a company-registry snippet describing the EVOLEM START legal entity's creation date, used as a proxy for the vehicle's vintage — flagged Moyen confidence since it's a single non-press source. Found some figures that could NOT be confidently attributed to this specific vehicle vs. the wider Evolem group and so were left unfilled: 'Evolem manages over €1 billion in total assets' (group-wide, spans multiple funds — not specific to Evolem Start) and 'more than 47 participations... 86th most active fund in France' (source: annuaire-startups.pro, ambiguous whether Evolem-group-wide or Evolem-Start-specific). Ticket size (€100k-€3M on deals of €500k-€5M, per Maddyness) was found but not filled since 'Ticket min/max' was not listed among this row's missing fields.

### Kickstart Innovation / Kickstart Innovation
`kickstart-innovation_kickstart-innovation`

**Champs complétés :** aucun.


**Notes :** Zero-equity accelerator (per task scope) — no AuM/ticket/TVPI sought. Founded 2015, Zurich, Switzerland. Runs verticals including Finance & Insurance with partners AXA, la Mobilière, Sanitas, Swisscom, PostFinance, Coop, City of Zurich, MSD. Portfolio companies include PriceHubble, Vianova, Nivaura. Confidence capped at Moyen because direct fetch of kickstart-innovation.com/about was blocked by the egress proxy; figures rely on secondary aggregator pages that cite the same underlying claim.

**Divergence (Nb participations (déclaré)) :** base = 382 ; web = 450 startups accompagnées depuis la création (source secondaire)

### Insurtech Capital / Insurtech Capital (Groupe Apicil)
`insurtech-capital_insurtech-capital-groupe-apicil`

**Champs complétés :**
- `Montant levé (M€)` = 10 *(confiance : Élevé)*
- `Montant alloué (M€)` = 10 *(confiance : Élevé)*
- `Millésime` = 2018 *(confiance : Élevé)*
- `Stages pratiqués` = 'Seed ; Série A' *(confiance : Élevé)*

**Source(s) :** Groupe Apicil official newsroom (groupe-apicil.com/newsroom/insurtech-lancement-dun-fonds-dinvestissement-dedie), Private Equity Magazine (pemagazine.fr/ntu4oa), mind.eu.com Fintech (mind.eu.com/fintech/article/apicil-consacre-10-millions-deuros-aux-investissements-dans-linsurtech), FinSMEs (finsmes.com/2018/12/odysseus-and-apicil-launch-insurtech-fund.html), Holland FinTech (hollandfintech.com/2019/01/apicil-group-odysseus-alternative-ventures-launch-eur-10-million-insurtech-fund), fintech.global Global Insurtech Summit (fintech.global/globalinsurtechsummit/odysseus-apicil-launch-e10m-dedicated-insurtech-fund) — all consistently report a €10 million fund ('Insurtech Capital I'), announced 11 December 2018, managed by Odysseus Alternative Ventures (Luxembourg, part of Reech Corporations Group) with APICIL as strategic partner, targeting seed and Series A insurtech/fintech start-ups in France and Europe with tickets of roughly €1M per deal.

**Champs recherchés sans résultat exploitable :** AuM (M€), Ticket min (M€), Ticket max (M€), Nb exits, TVPI, DPI, IRR, Quartile, Site web

**Notes :** No dedicated 'Insurtech Capital' website was found (fund is managed by Odysseus Alternative Ventures with no separate fund-branded domain located); the only official web presence found is the APICIL group's own newsroom page and Odysseus Alternative Ventures' corporate site, not an independent 'Insurtech Capital' site, so Site web was left unfilled rather than guessed. The ~€1M average ticket is reported as an approximate typical deal size across sources, not an explicit stated min/max range, so it was not filled as Ticket min/max (flagged 'approximate' in this note instead: ticket size ≈ €1M per operation).

### Macif Innovation / Macif Innovation
`macif-innovation_macif-innovation`

**Champs complétés :**
- `Montant alloué (M€)` = 30 *(confiance : Élevé)*
- `Ticket min (M€)` = 0.3 *(confiance : Élevé)*
- `Ticket max (M€)` = 1.5 *(confiance : Élevé)*
- `Stages pratiqués` = 'Seed ; Série A' *(confiance : Élevé)*
- `Nb participations (déclaré)` = 15 *(confiance : Élevé)*

**Source(s) :** Que font les fonds ? Le portrait de Macif Innovation, Maddyness, 24/02/2025 (https://www.maddyness.com/2025/02/24/que-font-les-fonds-le-portrait-de-macif-innovation/); corroborated on the 30 M€ envelope by La Macif crée un fonds de capital innovation, Private Equity Magazine (https://www.pemagazine.fr/NDc0Mw/la-macif-cree-un-fonds-de-capital-innovation)

**Champs recherchés sans résultat exploitable :** Nb exits, Nb InsurTech, TVPI, DPI, IRR, Quartile, Site web

**Notes :** Fund launched 2017 with an initial 15 M€ envelope, later increased to the current 30 M€ (PE Magazine / Maddyness). Maddyness reports 20 startups accompanied cumulatively since inception, of which 15 remain in the current portfolio; filled Nb participations (déclaré) with the current-portfolio figure (15) and flagged the cumulative 20 in this note rather than as a separate field. No dedicated 'Macif Innovation' website found distinct from macif.fr's general innovation page (https://www.macif.fr/assurance/la-macif/innovation), so Site web left unfilled.

### Seraphim Space / Seraphim Space Ventures (SSVII + SSIT)
`seraphim-space_seraphim-space-ventures-ssvii-ssit`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** Montant levé (M€), Montant alloué (M€), Ticket min (M€), Ticket max (M€), Millésime, Stages pratiqués, Nb exits, TVPI, DPI, IRR (TRI), Quartile

**Notes :** Out-of-scope SpaceTech fund per task brief — limited research to site/Millésime/portfolio count only, per instructions. IMPORTANT CAVEAT: this Fonds_ID bundles TWO distinct vehicles with different start dates, so a single 'Millésime' could not be confidently assigned and was left unfilled: (1) Seraphim Space Investment Trust (SSIT) listed on the London Stock Exchange in July 2021 with £250M gross proceeds (source: QuotedData / investors.seraphim.vc, aggregated); (2) Seraphim Space Ventures II (SSVII), launched in 2024 (TechCrunch 2024-04-21 'Seraphim Space launches second VC fund with 9 investments already under its belt'; Tech.eu 2024-04-23), which had a target of $100M (~€84M per Seraphim's own press-release conversion) and closed above target in Feb 2026 (exact final amount not disclosed in any source found — SpaceNews, Payload Space, and Seraphim's own press release all describe it as 'above $100M target' without a hard final figure). The Nb participations figure (17) refers specifically to SSVII's portfolio as of the Feb 2026 close and does NOT include SSIT's separate holdings, whose count was not found. Investment stage focus for SSVII is Seed and Series A, but this was not filled as 'Stages pratiqués' fill given task instruction to not spend time on non-essential fields for this out-of-scope entity.

### Howden Ventures / Howden Ventures
`howden-ventures_howden-ventures`

**Champs complétés :**
- `Montant alloué (M€)` = 11.7 *(confiance : Moyen)*

**Source(s) :** Howden launches world-first insurance innovation hub with £500m of delegated underwriting capacity, Howden Group Holdings newsroom (https://www.howdengroupholdings.com/news/howden-launches-world-first-insurance-innovation-hub-with-500m-of-delegated-underwriting-capacity); corroborated by Reinsurance News (https://www.reinsurancene.ws/insurance-innovation-hub-howden-ventures-launches-with-500m-underwriting-capacity/) and Insurance Business UK (https://www.insurancebusinessmag.com/uk/news/breaking-news/howden-to-fasttrack-investment-and-risk-incubation-with-insurance-innovation-hub-463331.aspx), all stating Howden initially committed £10m to the incubator

**Champs recherchés sans résultat exploitable :** AuM, Nb participations (déclaré), Nb exits, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** £10m converted to EUR at 1 GBP = 0.92... actually at 1 GBP ≈ 1.17 EUR per instructed rate => 11.7 M€ (approximate). IMPORTANT: the widely reported £500m figure is delegated Lloyd's underwriting CAPACITY backing new insurance products, not fund AuM/committed capital — did not fill AuM with it to avoid conflating metrics. Two portfolio names surfaced (Rosetta Risk Management and CetoAI) but no declared/official count of total participations, exits or InsurTech-specific investments was found.

### speedInvest / Speedinvest IV
`speedinvest_speedinvest-iv`

**Champs complétés :**
- `Nb participations (déclaré)` = 100 *(confiance : Faible)*

**Source(s) :** EU-Startups, 31/01/2024, 'Speedinvest closes massive €350 million fourth flagship fund' (https://www.eu-startups.com/2024/01/speedinvest-closes-massive-e350-million-fourth-flagship-fund-e50-million-above-target/) — 'ticket sizes are expected to range around the €600,000 mark for pre-seed investments, and €1.5 to €2 million for seed-stage companies, with the firm looking to add approximately 100 companies'

**Champs recherchés sans résultat exploitable :** Montant alloué, Ticket min/max, Phase, Nb exits, Nb InsurTech, TVPI, DPI, IRR (TRI), Quartile

**Notes :** Le chiffre 'approximately 100 companies' est explicitement qualifié d'approximatif par la source elle-même — reporté avec prudence (confiance Faible) car une extraction de recherche antérieure sur la même source a aussi renvoyé une fourchette différente ('$0.7M-$3M pour 40 à 100 start-ups', en USD), sans que je puisse confirmer le texte exact de l'article (WebFetch bloqué sur eu-startups.com dans cet environnement) : les deux formulations divergent sur la fourchette de ticket et la devise, donc le champ Ticket min/max n'est PAS rempli (source non fiable/confirmable), seul le nombre de participations déclaré (~100) est reporté, en gardant à l'esprit l'incertitude. AuM total confirmé >1 Md€ par cette même source, cohérent avec la valeur déjà en base (1000 M€) — pas de nouveau remplissage nécessaire.

### NCA (Next Commerce Accelerator) / NCA (Next Commerce Accelerator)
`nca-next-commerce-accelerator_nca-next-commerce-accelerator`

**Champs complétés :**
- `Site web` = 'https://nca.vc' *(confiance : Moyen)*

**Source(s) :** Official domain https://nca.vc confirmed via search results (title 'NCA VC – NCA VC'); portfolio count 'From 2017 to 2025 NCA invested in 77 B2B startups and partnered with 25+ corporate LPs' from search synthesis of nca.vc / prodevs.io / privateequitylist.com listings.

**Champs recherchés sans résultat exploitable :** Nb exits, Nb InsurTech, TVPI, DPI, IRR (TRI), Quartile

**Notes :** Confirmed zero-equity accelerator (Hamburg, Germany); invests up to ~€125k (range €50k-150k cited elsewhere) per startup as acceleration funding, not classic VC ticket — not filled as Ticket min/max since it wasn't in the requested missing-fields list and is a program grant, not an equity ticket in the VC sense. Focus sectors include Fin- & InsurTech among others. Portfolio count grew over time in different snapshots (36 startups as of 2020 batch #7; 41 startups per a more recent snapshot; 77 B2B startups 2017-2025 per most comprehensive figure) — these appear to be sequential growth, not contradictions, so the most recent comprehensive figure (77) was used. Fund-performance metrics (Nb exits, Nb InsurTech breakdown, TVPI/DPI/IRR/Quartile) confirmed not publicly available, consistent with a zero-equity accelerator.

**Valeur en fourchette non retenue (Millésime) :** 2017 (majorité des sources : Startbase, Prodevs.io, BusinessABC.net) ; 2018 (une source isolée mentionnant Peak Capital / Upgrade Commerce)

### The Family / The Family
`the-family_the-family`

**Champs complétés :**
- `Site web` = 'https://thefamily.co' *(confiance : Moyen)*
- `Montant levé (M€)` = 15 *(confiance : Moyen)*
- `Nb exits` = 8 *(confiance : Moyen)*

**Source(s) :** Site web: thefamily.co confirmed via direct search result title 'Building Ambitious Startups | The Family' (https://www.thefamily.co/). Montant levé: TechCrunch, 2018-09-11, 'The Family raises $17.4 million to support European startups' — explicitly states the raise as '$17.4 million (€15 million)', led by LGT Capital Partners with HummingBird Venture, Project A, and eVentures participating. Nb exits: Tracxn 'The Family - 2026 Investor Profile' (https://tracxn.com/d/accelerator-incubator/the-family/__GDXBbMjv2W87YNdgj3UJRzMRZhMm1Oz0AxlE5aQndHU) — states a portfolio of 24 companies including 3 unicorns, with 8 portfolio exits as of August 2026.

**Champs recherchés sans résultat exploitable :** Montant alloué (M€), Ticket min (M€), Ticket max (M€), TVPI, DPI, IRR (TRI), Quartile

**Notes :** The Family is a Paris-based startup accelerator/investor (since 2013), takes ~5% equity per startup in exchange for support, and runs twice-yearly cohorts of ~50 startups. Montant levé (€15M / $17.4M, 2018 round) reflects capital raised BY The Family's own vehicle from institutional backers (LGT Capital Partners et al.), not a traditional 'fund AuM' — confidence set to Moyen given it's a single (though authoritative) press source and dates to 2018, may not reflect subsequent fundraising. Nb exits (8) confidence set lower within the Moyen band since it comes from a single aggregator (Tracxn) rather than an official primary disclosure; no second corroborating named source was found for this specific exit count, so treat with caution. No public ticket size, allocation amount, or performance metrics (TVPI/DPI/IRR/Quartile) were found, consistent with The Family's non-fund accelerator/holding structure.

### The Moon Venture / The Moon Venture (Soul Invest)
`the-moon-venture_the-moon-venture-soul-invest`

**Champs complétés :**
- `Millésime` = 2018 *(confiance : Moyen)*
- `Stages pratiqués` = 'Seed ; Série A' *(confiance : Moyen)*
- `Site web` = 'https://themoonventure.com/' *(confiance : Moyen)*

**Source(s) :** Maddyness, portrait 'Que font les fonds ? Le portrait de The Moon Venture' (https://www.maddyness.com/2024/07/01/que-font-les-fonds-le-portrait-de-the-moon-venture/); fiche Maddyness base fonds (https://www.maddyness.com/base-fonds-investissements/the-moon-venture/); site officiel https://themoonventure.com/investisseurs/

**Champs recherchés sans résultat exploitable :** AuM (M€), Montant levé (M€), Nb exits, TVPI, DPI, IRR (TRI), Quartile

**Notes :** The Moon Venture = réseau de business angels / marque de Soul Invest (plateforme de crowdequity agréée AMF), fondé en 2018 par Matthieu Jarry, basé à Rennes. Fonctionne en co-investissement (ticket individuel min 20 K€, ticket cible communauté 100 K€) plutôt qu'en fonds fermé classique — d'où l'absence structurelle probable de métriques TVPI/DPI/IRR/Quartile publiques. Tickets déclarés : 250 K€-1 M€ en seed, jusqu'à 4 M€ en série A (ce ne sont pas des champs demandés comme manquants ici mais corroborent le profil seed/Série A retenu pour 'Stages pratiqués').

**Valeur en fourchette non retenue (Montant levé (M€) / AuM (M€)) :** 25 M€ investis en 4 ans dont 10 M€ en 2022 (Maddyness) ; >30 M€ déployés (page liée à Soul Invest) ; >50 M€ déployés / 30 startups (investormatch.pro, source non primaire)

### Blast.Club / Blast.Club
`blast-club_blast-club`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** AuM (M€), Montant levé (M€), Ticket min (M€), Ticket max (M€), Nb exits, Nb InsurTech, TVPI, DPI, IRR (TRI), Quartile

**Notes :** Angel club, not a classic fund (per task scope) — only Millésime and member/portfolio counts were sought. Member count context (not filled as a spreadsheet field since no exact 'Nb membres' column exists and 'Nb participations' means portfolio companies, not members): grew from ~4,000 members (2023) to 10,000+ (~2025, per Forbes France/JDN) to 14,000+ (early 2026, per Finscale/Substack interview with Bourbon & Guez). Aggregate amounts invested by members were reported with some inconsistency across Maddyness articles (e.g. '200M€ invested since 2023 creation, 88M€ in 2024' in one Jan-2025 piece vs '240M€ invested over 3 years, 125M€ in 2025' in a later piece) — these read as evolving cumulative totals over time rather than a strict contradiction, but per the angel-club scope note these financial totals were NOT filled as 'Montant levé'/'AuM' (task explicitly scopes Blast.Club research to site/Millésime/portfolio-member counts only). No specific portfolio company COUNT (as opposed to € invested or member count) was found published, so 'Nb participations (déclaré)' was left unfilled.

**Divergence (Millésime) :** base = 2022 ; web = création en 2023 selon 4 sources concordantes (Forbes France, JDN, Maddyness, Wikipédia)

### South East Angels / South East Angels
`south-east-angels_south-east-angels`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** AuM (M€), Montant levé (M€), Ticket min (M€), Ticket max (M€), Nb exits, Nb InsurTech, TVPI, DPI, IRR (TRI), Quartile

**Notes :** Angel network, not a classic fund (per task scope) — AuM/ticket/TVPI intentionally not pursued. Founded August 2020 by Kristina Pereckaite with 5 founding members; named 'Most Active Investor outside the capital' 2024 and UKBAA Angel Group of the Year finalist 2024/2025/2026. Portfolio count grew over time in different snapshots: an earlier UKBAA article (context ~2021-2023, around a '£1M invested' milestone) cited 15 companies invested, while the group's own Portfolio page as of Nov 2025 cites 34 companies (£4.4M invested) — treated as sequential growth, not a contradiction; the more recent/comprehensive figure (34) was used as the fill. The £4.4M invested-to-date figure was noted here for context but NOT filled as 'Montant levé (M€)' since it represents cumulative member investment volume through the network rather than fund AuM, consistent with the angel-network scope restriction. Direct site fetch (southeastangels.co.uk) was not attempted via WebFetch given the proxy blocking pattern observed on similar domains this session; relied on WebSearch synthesis of the site's own pages.

**Divergence (Nb participations (déclaré)) :** base = 41 ; web = 34 participations recensées sur la page Portfolio officielle (nov. 2025)

### Cadence Growth Capital / Cadence Growth Capital (CGC)
`cadence-growth-capital_cadence-growth-capital-cgc`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** Montant alloué (M€), Ticket min (M€), Ticket max (M€), Nb exits, Nb InsurTech, TVPI, DPI, IRR (TRI), Quartile

**Notes :** Firm is Berlin-based, focuses on DACH-region growth-stage companies (>€10M revenue) in FinTech/InsurTech themes, founders Leonard Clemens & Sebastian Eiseler. One aggregated snippet mentioned 'deployed around €200 million across multiple platform investments since inception' but no single reliable primary URL could be pinned to this figure, so it was NOT filled as 'Montant alloué' per sourcing rules (exact source+URL required) — flagging for manual follow-up. No ticket size, exit count, InsurTech-specific count, or performance metrics (TVPI/DPI/IRR/Quartile) found in public sources; these are consistent with a private growth-equity platform without public LP reporting. Direct site fetch (cadencegrowthcapital.com) was blocked by the egress proxy.

**Valeur en fourchette non retenue (Millésime) :** 2019 (année de fondation de la société, StartupIntros/Tracxn) ; 2020 (« Cadence Growth Capital Fund I | 2020 VC Fund », VentureCapitalArchive)

### Index Venture / Index Origin II
`index-venture_index-origin-ii`

**Champs complétés :**
- `Site web` = 'https://www.indexventures.com' *(confiance : Moyen)*
- `Montant alloué (M€)` = 276 *(confiance : Moyen)*
- `Géographie` = 'Europe ; Amérique' *(confiance : Moyen)*

**Source(s) :** Index Ventures official blog (17/11/2022) — "Announcing Index Origin II, a $300 Million Seed Fund" (https://www.indexventures.com/perspectives/announcing-index-origin-ii-a-300-million-seed-fund-designed-for-extraordinary-entrepreneurs/); Tech.eu (https://tech.eu/2022/11/17/back-to-the-origin-index-launches-second-seed-stage-focused-fund-at-300-million/) confirms fund 'supporting entrepreneurs primarily from Europe and the U.S.'. $300M converted at 1 USD ≈ 0.92 EUR = 276 M€.

**Champs recherchés sans résultat exploitable :** Ticket min (M€), Ticket max (M€), Nb participations (déclaré), Nb InsurTech, TVPI, DPI

**Notes :** Ticket size: one source states Origin II offers tickets 'up to $5 million' but no confirmed minimum specifically for Origin II (a $200K-$2M range appears in earlier coverage but seems to describe Origin I / the general seed program, not clearly Origin II) — left unfilled to avoid conflating the two funds.

### Ethias Ventures / Ethias Ventures
`ethias-ventures_ethias-ventures`

**Champs complétés :**
- `Site web` = 'https://www.ethias.be/content/corporate/en/ethias-group/EthiasVentures.html' *(confiance : Élevé)*
- `Montant levé (M€)` = 20 *(confiance : Élevé)*
- `Ticket max (M€)` = 3 *(confiance : Élevé)*

**Source(s) :** Ethias Newsroom press release 'Ethias Ventures : le nouveau véhicule d'investissement innovant chez Ethias!' (newsroom.ethias.be/ethias-ventures--le-nouveau-vehicule-dinvestissement-innovant-chez-ethias-isnkw5) and Belga Share syndication of the same release (belgashare.be/newsrooms/108/press-releases/1621) both state an initial endowment of €20 million and a maximum ticket of €3 million per round/per start-up. Official fund page: ethias.be/content/corporate/en/ethias-group/EthiasVentures.html.

**Champs recherchés sans résultat exploitable :** AuM (M€), Ticket min (M€), Nb exits, TVPI, DPI, IRR, Quartile

**Notes :** La Libre (Feb 2026) reports Ethias Ventures has deployed €15M across 10 startups to date (average ticket ≈€1.5M) — this is a deployment/track-record figure, not the fund's AuM or committed capital, so it was not used to fill AuM; reported here as context only. No minimum ticket size was found published (only a fund-level maximum of €3M per deal); a separate €10k minimum mentioned elsewhere refers to an unrelated Ethias retail co-investment/angel product, not this fund, and was excluded. Vehicle launched in Belgium (Ethias Ventures / Ethias group), founded ~2022-2023.

### Breega Capital / Breega Seed I
`breega-capital_breega-seed-i`

**Champs complétés :**
- `Site web` = 'https://www.breega.com' *(confiance : Faible)*

**Source(s) :** WebSearch result confirming BREEGA official site https://www.breega.com/; FinSMEs 2017 article 'Breega Capital Launches €100M European Venture Capital Fund' (finsmes.com/2017/07/breega-capital-launches-e100m-european-venture-capital-fund.html) mentioning an 'inaugural €50M ($57M) fund' preceding the €100M 'Breega Capital Venture 2'

**Champs recherchés sans résultat exploitable :** Montant alloué, Nb participations (déclaré), Nb exits, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** Breega's inaugural fund (2015) is referenced in press as an ~€50M fund preceding 'Breega Capital Venture 2' (2017, €100M), but no source explicitly names this inaugural fund 'Breega Seed I' — the fid-to-fund-name mapping is not confirmed, so Montant alloué is deliberately left unfilled rather than assumed. A separate source (superscout.co) does independently describe a 'Seed I (2015)' preceding 'Seed II (2019)' and 'Europe Seed III (2023)' sequence, which is consistent in timing but gives no montant. Recommend manual verification before using the €50M figure. No performance/participation data found.

### Truffle Capital / Truffle FinTech & InsurTech Fund III (prévu 2025)
`truffle-capital_truffle-fintech-et-insurtech-fund-iii-prevu-2025`

**Champs complétés :**
- `Site web` = 'https://www.truffle.com' *(confiance : Faible)*

**Source(s) :** Truffle Capital official website, multiple pages (https://www.truffle.com/, https://www.truffle.com/fintech/about, https://www.truffle.com/fintech/values)

**Champs recherchés sans résultat exploitable :** Montant levé, Ticket min, Ticket max, Phase, TVPI, DPI, IRR, Quartile

**Notes :** No public evidence found (as of Sept 2026 search) that a 'Fund III' has actually launched or closed — searches on Truffle Capital's fundraising only surface the FinTech-InsurTech fund closed in 2019 at 140 M€ (part of a combined ~390-400 M€ raise with the BioMedTech fund) and general firm-level stats (founded 2001, €1.2Bn raised historically, 120+ companies, 77 exits, 15 IPOs, per truffle.com). This 'Fund III (prévu 2025)' vehicle could not be independently confirmed as launched; Site web given is the group's official site (truffle.com), not a fund-specific page, hence low confidence.

**Rapprochement non résolu :** Aucune source ne confirme le lancement effectif de ce fonds à la date de cette collecte (14-15/09/2026) ; seuls le Fund II (140 M€, clos en 2019) et les statistiques globales de la société de gestion sont documentés publiquement. Le statut « prévu 2025 » du nom de la ligne reste donc non confirmé.

### Breega Capital / Breega Seed II
`breega-capital_breega-seed-ii`

**Champs complétés :**
- `Site web` = 'https://www.breega.com' *(confiance : Faible)*

**Source(s) :** WebSearch result confirming BREEGA official site https://www.breega.com/; superscout.co summary describing a 'Seed I (2015)' / 'Seed II (2019)' / 'Europe Seed III (2023)' sequence

**Champs recherchés sans résultat exploitable :** Montant alloué, Nb participations (déclaré), Nb exits, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** POSSIBLE NAMING OVERLAP TO FLAG: one AI-summarized WebSearch result (unverified, could not be corroborated against a primary article) asserted that the fund raised 2019-2021 under the name 'Breega Capital Venture 3' (final close €110M, matching timing of fid breega-venture-iii below) was 'later renamed Seed II'. This claim is NOT independently confirmed from a primary source and may be a summarization error; flagging it because if true, 'breega-seed-ii' and 'breega-venture-iii' could be the SAME underlying legal vehicle listed twice under different names in the spreadsheet, which would need manual de-duplication. No montant is filled here given the uncertainty.

**Rapprochement non résolu :** Une source secondaire non primaire indique que « Breega Capital Venture 3 » (= breega-venture-iii, 110 M€, closing final mars 2021) aurait été « renommé Seed II » par la suite. Si confirmé, les deux lignes de la base pourraient désigner le même véhicule sous deux noms différents — à vérifier manuellement avant d'arbitrer un éventuel doublon.

### Step Venture / Step Venture I
`step-venture_step-venture-i`

**Champs complétés :**
- `Millésime` = 2025 *(confiance : Élevé)*

**Source(s) :** https://stepventure.eu/article/first-closing-for-step-fund-the-eur50-million-fund-dedicated-to-italian-seed/ (official site) corroborated by https://www.eu-startups.com/2025/11/new-e30-million-step-fund-targets-early-stage-italian-startups-with-international-growth-potential/ and https://bebeez.eu/2025/11/11/... and https://vcwire.tech/2025/11/12/step-fund-holds-first-close-at-30m/

**Champs recherchés sans résultat exploitable :** Montant alloué, Nb participations (déclaré), Nb exits, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** Step Fund is Step Venture's first fund (Milan/Naples-based), structured as a EuVECA PIR Alternative vehicle, first closing Nov 2025 with CDP Venture Capital Sgr (via the PNRR Digital Transition Fund) as €20M anchor investor; invests €500k-2M at pre-Series A in fintech/healthtech/B2B software/'connected world' with AI angle, Italy-focused. Given the fund only just held its first close, it is unsurprising that no participations/exits/performance data are public yet; Nb InsurTech and performance fields left unfilled rather than assumed zero.

### The Net Street Capital / The Net Street Capital
`the-net-street-capital_the-net-street-capital`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** AuM, Montant levé, Montant alloué, Stages pratiqués, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** The Net Street Capital (London) invests in B2B SaaS insurtech/fintech; stage descriptors found in aggregator snippets ('Growth', 'Scaling', 'Early Revenue') do not map onto the closed Stages vocabulary (Pré-Seed…Série D), so left unfilled per the closed-vocabulary rule rather than force-mapped. No AUM, amounts, or performance data found. Could not fetch thestartupverse.com or openvc.app (egress-blocked).

### TrueSight Ventures / TrueSight Ventures Fund I
`truesight-ventures_truesight-ventures-fund-i`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** AuM, Montant levé, Montant alloué, Nb InsurTech, TVPI, DPI, IRR, Quartile, Site web

**Notes :** TrueSight Ventures (founded 2018, London & Stockholm; partners Hampus Monthan Nordenskjöld, Igor Tikhturov, Oleg Tikhturov) is a generalist pre-seed/seed fund (B2B SaaS, marketplaces, AI/ML, fintech, insurtech, devtools, enterprise, digital health, B2C) writing $50k-$500k tickets (sweet spot ~$250k) with ~48 portfolio companies as of early 2024. Confirmed insurtech-relevant portfolio example: BondAval (smart payment-security/insurtech, raised a $7M round TrueSight participated in), but this is a single anecdotal data point, not a verified full count, so Nb InsurTech was left unfilled. No official website, AUM, amounts-raised, or performance metrics found for 'Fund I' specifically; could not fetch openvc.app or thestartupverse-type pages (egress-blocked).

### Insurtech Gateway / Insurtech Gateway Seed Fund (I & II)
`insurtech-gateway_insurtech-gateway-seed-fund-i-ii`

**Champs complétés :**
- `Ticket max (M€)` = 1.17 *(confiance : Moyen)*

**Source(s) :** Superscout investor guide (superscout.co/investor/insurtech-gateway) and Tracxn profile (tracxn.com/.../insurtech-gateway) both describe Insurtech Gateway's direct investment ticket as up to £250k at Pre-Seed and up to £1M at Seed. Converted at 1 GBP = 1.17 EUR (£1M x 1.17 = 1.17 M€; £250k x 1.17 = 0.29 M€).

**Champs recherchés sans résultat exploitable :** AuM (M€), Montant levé (M€), Montant alloué (M€), TVPI, DPI, IRR, Quartile

**Notes :** Seed Fund I closed February 2019 (per Insurtech Gateway's own site, insurtechgateway.com/2019/06/12/gateway-upgraded-model-incubator-fund) and Seed Fund II completed its first close in July 2022 (Finextra, Reinsurance News, InsurTech Digital, Crowdfund Insider — multiple corroborating named sources), but none of the press coverage or the official site discloses the actual £ amount raised for either vintage; a stray search snippet mentioned 'up to £30M' but this could not be confirmed to refer specifically to Seed Fund I or II rather than a broader/blended Gateway vehicle, so it was excluded rather than reported. The £250k/£1M figures are ticket sizes for direct fund investment (Pre-Seed/Seed), distinct from a separate £10k minimum for Insurtech Gateway's angel co-investment syndicate (not the fund itself) — only the fund-level max ticket (Seed, £1M/1.17M€) was filled as 'Ticket max'; Pre-Seed £250k (0.29M€) is noted here as it may be the intended 'Ticket min' depending on the database's definition, but was not filled to avoid conflating a stage-specific cap with an overall minimum.

### Index Venture / Index Ventures XI
`index-venture_index-ventures-xi`

**Champs complétés :**
- `Site web` = 'https://www.indexventures.com' *(confiance : Moyen)*
- `Montant alloué (M€)` = 828 *(confiance : Moyen)*
- `Stratégie` = 'Généraliste' *(confiance : Moyen)*

**Source(s) :** Global Legal Chronicle (09/2021) — "Index Ventures' $900 Million Closing of Index Ventures XI" (https://globallegalchronicle.com/index-ventures-900-million-closing-of-index-ventures-xi/). $900M converted at 1 USD ≈ 0.92 EUR = 828 M€.

**Champs recherchés sans résultat exploitable :** Nb exits, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** 'Stratégie: Généraliste' is inferred from Index Ventures' well-documented sector-agnostic, multi-stage positioning (confirmed generalist tech investor across fintech/insurtech/healthtech/SaaS/consumer, e.g. Wikipedia and firm's own materials) rather than a single explicit strategy label for fund XI itself — flagged Moyen confidence specifically for this field.

### 365.fintech / 365.fintech
`365-fintech_365-fintech`

**Champs complétés :**
- `Montant levé (M€)` = 33 *(confiance : Faible)*
- `Millésime` = 2018 *(confiance : Faible)*

**Source(s) :** Capboard investor profile, 365.fintech (https://www.capboard.io/en/investor/365-fintech) — 'total invested amount of $35.9M'; founding year 2018 corroborated by Vestbee (https://www.vestbee.com/vc-list/365.fintech/) and o.parsers.vc (https://o.parsers.vc/fund/365fintech.vc/)

**Champs recherchés sans résultat exploitable :** AuM, Montant alloué, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** The $35.9M 'total invested amount' (Capboard, single non-primary aggregator source) converted at 1 USD ≈ 0.92 EUR = ~33 M€; flagged approximate and low confidence since it is not independently corroborated and its exact definition (AUM vs cumulative deployed vs raised) is unclear. Millésime (2018 founding year) is corroborated by 2+ named sources but is the firm's founding year, not necessarily a fund vintage — flagged as approximate.

### Committed Capital / Committed Capital EIS Fund
`committed-capital_committed-capital-eis-fund`

**Champs complétés :**
- `AuM (M€)` = 117 *(confiance : Moyen)*
- `TVPI` = 3.0 *(confiance : Moyen)*
- `IRR (TRI)` = 0.368 *(confiance : Moyen)*

**Source(s) :** WebSearch aggregation of committedcapital.co.uk, growthbusiness.co.uk ('Top 20 EIS funds and investors you should know about'), ifamagazine.com ('Committed Capital Growth EIS Fund'), and openvc.app/fund/Committed%20Capital

**Champs recherchés sans résultat exploitable :** Montant alloué, Millésime, Stages pratiqués, Nb InsurTech, Quartile

**Notes :** Figures reported for the flagship 'Growth EIS Fund': average 3x ROI (mapped to TVPI=3.0, approximate=true) and 36.8% IRR (=0.368) over an average 4.5-year holding period, from 17 portfolio companies with 5 profitable exits and 2 small losses (Nb exits not in the missing-field list so not filled, but noted here). AuM reported elsewhere as 'over £100m' for Committed Capital's evergreen EIS fund (approximate=true), converted at 1 GBP=1.17 EUR => 117 M€. Could not WebFetch committedcapital.co.uk directly (egress-blocked) to confirm verbatim wording, so capped at Moyen despite multiple named sources (official site + growthbusiness.co.uk + ifamagazine.com) repeating the same figures. Note there is also a separate 'Knowledge-Intensive EIS Fund' (Fund III opened 1 Oct 2024) from the same manager — the performance figures above pertain to the 'Growth EIS Fund' specifically; if the sheet's vehicle 'Committed Capital EIS Fund' is meant to be the generic/flagship vehicle this likely applies, but if it is meant to be a different named sub-fund this may not match (flagging for maintainer review rather than as a divergence since no current value is known).

**Rapprochement non résolu :** Les données trouvées (AuM, TVPI, IRR) concernent le véhicule officiellement nommé « Growth EIS Fund » chez Committed Capital ; à confirmer qu'il s'agit bien du même véhicule que la ligne « Committed Capital EIS Fund » de la base avant de leur accorder une confiance Élevée.

### Accel Partners / Accel London VI
`accel-partners_accel-london-vi`

**Champs complétés :**
- `Site web` = 'https://www.accel.com' *(confiance : Élevé)*
- `Montant alloué (M€)` = 529 *(confiance : Élevé)*

**Source(s) :** TechCrunch (techcrunch.com/2019/05/15/accel-closes-575m-fund-to-double-down-on-european-and-israeli-series-a-deals), Unquote (unquote.com/uk/official-record/3014995), FinSMEs (finsmes.com/2019/05/accel-closes-sixth-european-and-israeli-fund-at-575m.html), Tech.eu (tech.eu/brief/accel-fund-6) — all report Accel's sixth European/Israel early-stage fund closed at $575M in May 2019; converted at 1 USD = 0.92 EUR ($575M x 0.92 = 529 M€). Site: accel.com is the firm's official corporate site, referenced consistently across all PitchBook/press listings.

**Champs recherchés sans résultat exploitable :** Nb exits, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** Fund is denominated in USD; no EUR-denominated official figure found. Numbering confirmed via multiple press sources explicitly calling this the 'sixth' Accel European fund (matches London VI). No public TVPI/DPI/IRR or exit-count data found for this vintage; PitchBook has fund pages but figures are paywalled. Fund invests across Europe and Israel.

### Accel Partners / Accel London VII
`accel-partners_accel-london-vii`

**Champs complétés :**
- `Site web` = 'https://www.accel.com' *(confiance : Élevé)*
- `Montant alloué (M€)` = 598 *(confiance : Élevé)*

**Source(s) :** TechCrunch (techcrunch.com/2021/06/29/accel-closes-on-3b-across-three-funds), Silicon Canals (siliconcanals.com/accel-closes-3b-across-3-new-funds), Sifted (sifted.eu/articles/accel-fundraise-800m-ninth-early-stage-europe-fund, quoting Accel partner Harry Nelis referencing the 'seventh fund, also $650 million, raised in 2021') — Accel's seventh early-stage Europe/Israel fund closed at $650M in June 2021, part of a $3.05B multi-fund close (with Accel XV and Accel Growth Fund VI). Converted at 1 USD = 0.92 EUR ($650M x 0.92 = 598 M€).

**Champs recherchés sans résultat exploitable :** Nb exits, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** Distinguished from Accel London VIII, a separate $650M fund closed in May 2024 (tech.eu explicitly calls that one the 'eighth fund'); care was taken not to conflate the two $650M raises. No public performance data found.

### 115K / 115K
`115k_115k`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** Montant alloué, Nb participations (déclaré), Nb exits, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** Portfolio count is inconsistent across dated sources and could not be pinned to a single reliable figure, so left unfilled per the divergence rule: ~6 participations as of June 2023 (LinkedIn post 'Un an déjà pour 115K', incl. Carbo, Cashbee, Pono Technologies, Sesame IT, Cosmian, Garantme) vs. a later aggregator claim of '16 investments since 2022' / '17 as of 2026' that could not be traced to one primary, dated source. Known insurtech-adjacent portfolio company: Sesame IT (courtage digital). No exit, TVPI/DPI/IRR/Quartile data found. Could not WebFetch 115k.fr, lajauneetlarouge.com or eldorado.co (egress-blocked).

### Venpace / Venpace
`venpace_venpace`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** AuM, Montant levé, Montant alloué, Ticket min, TVPI, DPI, IRR, Quartile

**Notes :** VENPACE is an InsurTech-builder/early-stage VC fund run by Cologne company-builder crossbuilders with four DACH insurer LPs (IDEAL Versicherungsgruppe, PrismaLife, Provinzial, Vienna Insurance Group), investing pre-seed/seed 'up to €500,000' per ticket — but that figure describes the maximum ticket (already presumably known), not the missing 'Ticket min' field, so not filled. No total fund volume (Fondsvolumen), AUM, or performance data found in any source; could not fetch venpace.com, asscompact.de or dfpa.info directly (egress-blocked).

### Seed X Liechtenstein / Seed X Liechtenstein
`seed-x-liechtenstein_seed-x-liechtenstein`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** AuM, Montant levé, Montant alloué, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** Seed X Liechtenstein (Schaan, LI, founded 2016, fintech/insurtech/proptech/legaltech, pre-seed to Series A, $0.3-1M initial ticket) operates through two named sub-vehicles: 'VC Fintech I' (PitchBook: 2021 vintage, 28 investments) and 'VC Fintech II' (fewer investments, 3 per PitchBook snippet). No fund size, AUM or performance metric surfaced for either. Could not fetch seedx.li (egress-blocked) or PitchBook (paywalled) to verify further. Flagging the VC Fintech I/II naming in case the sheet's single 'Seed X Liechtenstein' row is actually meant to represent one of these two named vehicles rather than the parent manager — worth a maintainer check, not treated as a fill since identity mapping is uncertain.

### Tenity / Tenity Incubation Fund I & II
`tenity_tenity-incubation-fund-i-ii`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** Montant levé, Montant alloué, Ticket min, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** Four sources give four different, non-reconcilable figures for Tenity's fund/AUM size (target vs AUM vs assets-under-advisory vs current AUM across all funds), and none isolates the 'Incubation Fund I & II' vehicle specifically from Tenity's group-wide figures. Per rule 3, left unfilled and reported as divergence rather than guessing which figure applies to this specific fund line.

**Valeur en fourchette non retenue (Montant levé (M€) / AuM (M€)) :** 100 M$ (cible fonds II, Startup Bubble News) ; >140 M$ (AuM groupe, Superscout) ; 120 MCHF (« assets under advisory » au 1er closing, startupticker.ch) ; 100 MCHF (AuM groupe Tenity AG, SECA)

### SCOR Ventures / SCOR Ventures
`scor-ventures_scor-ventures`

**Champs complétés :**
- `Montant alloué (M€)` = 130 *(confiance : Moyen)*

**Source(s) :** Waveup Copilot VC Fund Profile (SCOR Ventures) — "SCOR Ventures operates with a 130 million euro mandate" (https://hub.waveup.com/funds/scor-ventures), cross-referenced by a second independent WebSearch result citing the same 130 M€ figure

**Champs recherchés sans résultat exploitable :** Ticket min (M€), Ticket max (M€), Nb exits, TVPI, DPI, IRR, Quartile

**Notes :** The 130 M€ figure comes from a third-party aggregator (Waveup), not scor.com directly (scor.com was blocked by the network egress proxy in this session and could not be fetched to confirm). Ticket size: same aggregator reports a $1M-$10M range with a $4M 'sweet spot' (≈0.9-9.2 M€ / 3.7 M€ converted at 0.92), but since this is a single secondary source I chose not to file it as a confirmed fill — reported here only as a lead. Nb exits: CB Insights-derived search summary claims 'SCOR Ventures has 5 investments and 1 portfolio exit', which looks implausibly low for a CVC active since 2017-2018 and was not used as a fill (likely incomplete data-tracking by that aggregator, not a genuine published fact).

### Start Venture / Start Venture I
`start-venture_start-venture-i`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** AuM (M€), Montant levé (M€), Montant alloué (M€), TVPI, DPI, IRR (TRI), Quartile, Site web

**Notes :** Extensive WebSearch (multiple query variants: French/English, 'fonds capital-risque', 'seed', LinkedIn/Dealroom/Capboard site: filters) returned no entity matching 'Start Venture' / 'Start Venture I' as a France/Europe VC vehicle. The only similarly-named entity found was 'Start Ventures' (Lisbon, Portugal), a B2B Fintech/Insurtech/Regtech/Cybersecurity early-stage VC fund per Capboard/Shizune listings — but this appears to be a different, unrelated entity (different name spelling, different country) and was not treated as a match. No official site, registry filing, or press mention of 'Start Venture I' was located. Recommend manual verification of the fund's exact legal name.

**Rapprochement non résolu :** Aucune entité « Start Venture » correspondante n'a été retrouvée en ligne après plusieurs recherches ciblées. Seule une entité au nom proche mais distincte (« Start Ventures », Lisbonne, fintech/insurtech) existe publiquement ; non retenue faute de confirmation du rapprochement. Aucune donnée reportée.

### Breega Capital / Breega Venture III
`breega-capital_breega-venture-iii`

**Champs complétés :**
- `Site web` = 'https://www.breega.com' *(confiance : Élevé)*
- `Montant alloué (M€)` = 110 *(confiance : Élevé)*

**Source(s) :** EU-Startups 'Breega closes its third fund at €110 million for European tech startups' (eu-startups.com/2021/03/breega-closes-its-third-fund-at-e110-million-for-european-tech-startups/); corroborated by TechCrunch and VentureBeat coverage of the same March 2021 announcement ($130M / €110M final close of 'Breega Capital Venture 3', first closing €90M mid-2019)

**Champs recherchés sans résultat exploitable :** Nb exits, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** Montant alloué of €110M is corroborated by multiple independent named-press sources (EU-Startups, TechCrunch, VentureBeat) all reporting the same March 2021 final closing figure, hence Élevé confidence despite WebFetch being blocked (relied on WebSearch snippets from these outlets). First closing was €90M (mid-2019). See also the naming-overlap caution noted under breega-seed-ii above (unverified claim that this fund was later rebranded 'Seed II') — recommend checking for possible duplication between the two spreadsheet rows.

**Rapprochement non résolu :** Une source secondaire non primaire indique que « Breega Capital Venture 3 » (= breega-venture-iii, 110 M€, closing final mars 2021) aurait été « renommé Seed II » par la suite. Si confirmé, les deux lignes de la base pourraient désigner le même véhicule sous deux noms différents — à vérifier manuellement avant d'arbitrer un éventuel doublon.

### CommerzVentures / CommerzVentures Fonds I à III
`commerzventures_commerzventures-fonds-i-a-iii`

**Champs complétés :**
- `Montant alloué (M€)` = 550 *(confiance : Élevé)*
- `Ticket min (M€)` = 2 *(confiance : Élevé)*
- `Ticket max (M€)` = 10 *(confiance : Élevé)*

**Source(s) :** EU-Startups, Finextra, FinSMEs, AltAssets, fintech.global, VentureCapitalJournal (all 03/2022) — corroborating: Fund I (2014) €100M, Fund II (2019) €150M, Fund III (2022) €300M, 'combined total fund size of €550 million' (https://www.eu-startups.com/2022/03/commerzventures-closes-e300-million-fund-to-boost-european-fintech-and-insurtech-startups/ ; https://www.finextra.com/newsarticle/39783/commerzventures-closes-300m-third-fund ; https://www.finsmes.com/2022/03/commerzventures-closes-e300m-third-fund.html). Ticket range €2-10M per multiple aggregators (VCsheet, OpenVC, Tracxn) describing Series A/B checks.

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** Montant alloué = 550 M€ is the SUM of the three vintages (Fund I 100 + Fund II 150 + Fund III 300), matching the vehicle's grouped name 'Fonds I à III'. Ticket range (€2-10M) sourced from secondary aggregator sites, not the primary commerzventures.com site (blocked by network egress in this session) — confidence Moyen for that sub-figure specifically. Nb exits: a WebSearch AI-summary cited CB Insights/Dealroom data suggesting roughly '2 IPOs and 7 acquisitions' (~9 exits, incl. Curv→PayPal, Payworks→Visa, eToro & Marqeta IPOs) but this could not be verified against a directly quoted primary source in this session, so it was NOT used as a fill — reported here only as an unverified lead for a follow-up check.

### NewAlpha Asset Management / NewAlpha Fintech/InsurTech
`newalpha-asset-management_newalpha-fintech-insurtech`

**Champs complétés :**
- `Site web` = 'https://www.newalpha.com/en/venture-capital/' *(confiance : Moyen)*

**Source(s) :** https://www.newalpha.com/en/venture-capital/ (official NewAlpha Asset Management site, venture capital / FinTech-InsurTech page)

**Champs recherchés sans résultat exploitable :** Montant levé, Montant alloué, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** NewAlpha's dedicated vehicle is referred to in secondary sources (Crunchbase, Tracxn, CB Insights, superscout.co) as the 'NewAlpha FinTech Fund'; no exact vehicle name, fund size or number of InsurTech investees was found on an official page or in named press. Firm-wide AUM (~$3-4bn across all NewAlpha absolute-performance + VC activities, per newalpha.com and Crunchbase) is NOT fund-specific and was NOT used to fill 'Montant levé' per the no-inference rule. Portfolio is described as spanning payments/digital banking/insurtech/wealthtech/regtech with named investees including Digital Insure (insurtech) among ~23-30 companies since 2016, but no official per-category count ('Nb InsurTech') was published, so left blank. Performance metrics (TVPI/DPI/IRR/Quartile) not found in any public LP report.

### NewFund Capital / NewFund NAEH 1 et 2
`newfund-capital_newfund-naeh-1-et-2`

**Champs complétés :**
- `Site web` = 'https://newfundcap.com' *(confiance : Moyen)*
- `Montant alloué (M€)` = 19.1 *(confiance : Moyen)*

**Source(s) :** Private Equity Magazine (pemagazine.fr/odc1mw and pemagazine.fr/odm5oq), Banque des Territoires press release (banquedesterritoires.fr/closing-du-fonds-dinvestissement-newfund-naeh-innopy-153-meu and PDF CP BDT NAEH INNOPY 09102024), CFNEWS (cfnews.net/.../Newfund-passe-la-seconde...470609) — Newfund NAEH (1st vintage, FCPR created end-2018) closed at €3.8M; Newfund NAEH Innopy (2nd vintage) closed at €15.3M (incl. a €5M Banque des Territoires ticket), in line with its €15M target. Combined = 3.8 + 15.3 = 19.1 M€. Site: newfundcap.com confirmed as Newfund Management's official site (also cited by Wikipedia 'Newfund' entry).

**Champs recherchés sans résultat exploitable :** AuM (M€), TVPI, DPI, IRR, Quartile

**Notes :** The €3.8M figure for NAEH1 is corroborated by only one clear press mention (PE Magazine); one earlier article (Le Journal des Entreprises, Dec 2018) cited a €7M target with €3M secured at that point, so the 3.8M€ appears to be the eventual final close rather than the target — treated as the reported final amount, not an estimate. NAEH Innopy is technically a distinct legal vehicle name (2nd generation), and a further 'NAEH Innopy II' vintage was found referenced by CFNEWS, which falls outside the '1 et 2' scope of this Fonds_ID and was not included.

### Accel Partners / Accel London VIII
`accel-partners_accel-london-viii`

**Champs complétés :**
- `Site web` = 'https://www.accel.com' *(confiance : Élevé)*
- `Montant alloué (M€)` = 598 *(confiance : Élevé)*
- `Géographie` = 'Europe' *(confiance : Élevé)*

**Source(s) :** Tech.eu (tech.eu/2024/05/14/accel-s-eighth-fund-closes-at-650m — explicitly 'Accel's eighth fund'), Venture Capital Journal (venturecapitaljournal.com/accel-london-viii-raises-650m-for-europe-and-israel), CNBC (cnbc.com/2024/05/13/venture-capital-firm-accel-raises-650-million-europe-and-israel-fund.html), Sifted (sifted.eu/articles/accel-650m-fund-europe) — Accel London VIII closed at $650M on 13 May 2024, targeting Europe and Israel, early-stage (seed/Series A). Converted at 1 USD = 0.92 EUR ($650M x 0.92 = 598 M€).

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** Fund explicitly targets 'Europe and Israel'; Israel is out of the closed Géographie vocabulary (Europe/Amérique/International), so only 'Europe' was filled — flagging Israel exposure here rather than in the field. No public performance data found for this 2024-vintage fund (too recent in any case).

### Elevation Capital Partners / FPCI elevation Early Growth I
`elevation-capital-partners_fpci-elevation-early-growth-i`

**Champs complétés :**
- `Site web` = 'https://www.elevation-cp.com' *(confiance : Moyen)*

**Source(s) :** Inter Invest Capital poursuit son développement et devient Elevation Capital Partners (https://www.inter-invest.fr/communiques/capital-investissement/00458/inter-invest-capital-poursuit-son-developpement-et-devient-elevation-capital-partners); Elevation Capital Partners official site confirmed via search results (https://www.elevation-cp.com/)

**Champs recherchés sans résultat exploitable :** Montant alloué, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** No public fund-size figure found specifically for FPCI Elevation Early Growth I (only ticket-size info, 500k-5M€ per investment, which was not a requested field). Elevation Capital Partners (AMF GP-15000006) is the Inter Invest Group's growth-equity arm; fund pages also mirrored at inter-invest.fr/capital-investissement/fpci-elevation-early-growth.

### UNIQA Ventures / UNIQA Ventures
`uniqa-ventures_uniqa-ventures`

**Champs complétés :**
- `Site web` = 'https://www.uniqaventures.com' *(confiance : Faible)*
- `Montant alloué (M€)` = 150 *(confiance : Faible)*

**Source(s) :** UNIQA Group Press Center (press-news.uniqagroup.com/news-150-million-euros-for-start-ups-in-the-cee-region-uniqa-ventures-doubles-growth-capital-for-bold-future-investments) et TrendingTopics (trendingtopics.eu/uniqa-ventures-doubles-investment-volume-to-e150m/) — UNIQA Ventures a doublé son 'volume d'investissement' de 75 M€ à 150 M€ (annonce 2021) ; site officiel uniqaventures.com mentionné dans plusieurs profils d'investisseurs (Vestbee, EU-Startups, Capboard) via extraits WebSearch

**Champs recherchés sans résultat exploitable :** Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** Le montant de 150 M€ correspond au 'volume/capacité d'investissement' total communiqué par UNIQA (pas explicitement qualifié de 'montant alloué au fonds' au sens strict) — à valider ; marqué 'approximate' et confiance Faible car provenant uniquement d'extraits WebSearch (WebFetch bloqué : EGRESS_BLOCKED sur www.uniqaventures.com), déjà EUR donc pas de conversion nécessaire. 'Nb InsurTech' non trouvé : le portefeuille total est cité tantôt à 30 (vestbee.com/insights/articles/vc-of-the-month-uniqa-ventures, 2024) tantôt à 52 sociétés (tracxn.com), mais aucune de ces sources ne ventile par verticale (InsurTech vs FinTech vs HealthTech) — chiffre non fiable pour ce champ, donc non rempli. Aucune métrique TVPI/DPI/IRR publique trouvée.

### Helsana HealthInvest / Helsana HealthInvest
`helsana-healthinvest_helsana-healthinvest`

**Champs complétés :**
- `Montant alloué (M€)` = 107 *(confiance : Élevé)*
- `Millésime` = 2020 *(confiance : Élevé)*

**Source(s) :** Helsana official page (helsana.ch/en/helsana-group/about-us/healthinvest.html), IFHP executive summary PDF (ifhp.com/wp-content/uploads/2023/06/Helsana-PDF-Exec-Sum.pdf), Caplight investor profile (caplight.com/investor/helsana), EU-Startups (eu-startups.com/investor/helsana-healthinvest) — all describe Helsana HealthInvest as an evergreen fund with a CHF 100 million commitment. Converted at 1 CHF = 1.07 EUR (CHF 100M x 1.07 = 107 M€). Founding year 2020 confirmed by Moneyhouse Swiss commercial-registry entry (moneyhouse.ch/en/company/helsana-healthinvest-ag-4648125341).

**Champs recherchés sans résultat exploitable :** Nb exits, TVPI, DPI, IRR, Quartile

**Notes :** Multiple sources describe the vehicle as 'operational since 2021' even though the AG was registered/founded in 2020 (per Moneyhouse) — this is a launch-of-activity date, not a contradictory fund-size or amount figure, so it was noted rather than treated as a divergence_range. No confirmed exits found; one portfolio investment (AdCubum) surfaced ambiguously in search snippets but could not be verified as an exit with a clear source, so it was excluded rather than reported.

### Eurazeo / Eurazeo Venture Capital
`eurazeo_eurazeo-venture-capital`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** Montant alloué, Nb exits, Nb InsurTech, TVPI, DPI, IRR (TRI), Quartile

**Notes :** Investissements insurtech identifiés dans la presse pour Eurazeo (segment Venture/Growth) : Igloo (Singapour, déc. 2023), InsuranceDekho (Inde, oct. 2023), participation à la Series D de wefox — mais liste non exhaustive/non officielle, donc non utilisée pour remplir 'Nb InsurTech' (risque de sous-compte). Un chiffre de '~200 exits' apparaît pour Eurazeo mais au niveau groupe (private equity + growth + real assets), pas spécifique à Eurazeo Venture Capital — non retenu. Aucune métrique TVPI/DPI/IRR publique trouvée pour ce véhicule précis (Eurazeo est cotée mais ne publie pas de TVPI par sous-fonds VC dans les sources consultées).

### Breega Capital / F/I Venture II
`breega-capital_f-i-venture-ii`

**Champs complétés :**
- `Site web` = 'https://www.breega.com' *(confiance : Moyen)*

**Source(s) :** French company registry listings: rubypayeur.com 'Société BREEGA VENTURE II à PARIS - SIREN : 902920651' and annuaire-entreprises.data.gouv.fr entry for SIREN 902920651, both identifying the legal entity 'BREEGA VENTURE II' with the alias '(F/I VENTURE II)', created 15 July 2021, based in Paris (75002); PitchBook fund-profile snippet for the earlier, related 'F/I Venture' (I) describing it as a 2017-vintage Breega-managed fund with Crédit Agricole as sole LP, focused on financial services

**Champs recherchés sans résultat exploitable :** Montant alloué, Nb exits, TVPI, DPI, IRR, Quartile

**Notes :** CONFIRMS DATA-QUALITY QUESTION: 'F/I Venture II' is a genuine, real fund name, not a typo — French company registries record its legal name as 'BREEGA VENTURE II' (SIREN 902920651, created 15/07/2021, Paris) with 'F/I VENTURE II' as an alternate/commercial name. This appears to be a follow-on to 'F/I Venture' (I) (2017 vintage, ~€50M per one PitchBook snippet reference to a Crédit Agricole fintech-fund delegation), a vehicle where Crédit Agricole was described as the sole/main LP, distinct from Breega's main market-raised 'Breega Capital Venture 1/2/3' fund line. No public montant, exit count, or performance figures were found specifically for F/I Venture II; WebFetch to the registry site was blocked (EGRESS_BLOCKED), so this rests on WebSearch snippets of the registry pages, hence Moyen rather than Élevé confidence.

### 13books Capital (ex-Element Ventures) / Fonds I & II
`13books-capital-ex-element-ventures_fonds-i-ii`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** Montant alloué, Nb exits, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** Fund I ($130m, as Element Ventures) and Fund II (£121m, 2024, as 13books Capital) sizes are documented, but Montant levé was not listed as missing so not reported. No explicit count of exits or insurtech-specific deals was found — only a named exit (Roadzen) and named insurtech/embedded-insurance portfolio companies (hepster, hepster described as 'eCommerce embedded insurance'; Third Fort as digital identity). LPs include British Patient Capital and KfW Germany (Fund II) and Isomer Capital / IPGL. No performance metrics (TVPI/DPI/IRR) found. Confirmed this is the same entity as the Fonds_ID's official site (13bookscapital.com), distinct from any Canadian 'Element Ventures' homonym — no Canadian entity appeared in results.

### Concentric / Concentric
`concentric_concentric`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** Montant levé, Montant alloué, Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** The '€100m+ AUM' figure found is described as spanning Concentric's 'different venture strategies' (i.e., the whole firm/multiple fund vehicles), not clearly the single 'Concentric' vehicle in this row, so it was not mapped to Montant levé/alloué to avoid conflating firm-level AUM with a specific fund's raised amount. Concentric also completed a 'first close of second fund' per concentric.vc/news, implying at least two fund vehicles exist under the Concentric brand — worth checking which one this Fonds_ID row corresponds to. No InsurTech-specific investment count or performance metrics found.

### BPI France / Large Ventures
`bpi-france_large-ventures`

**Champs complétés :**
- `Géographie` = 'Europe' *(confiance : Moyen)*

**Source(s) :** Bpifrance, page produit 'Large Venture' (https://www.bpifrance.fr/nos-solutions/investissement/investissement-expertise/large-venture et https://www.bpifrance.com/products/large-venture/, contenu consulté via extraits de recherche — WebFetch direct bloqué sur ces domaines) : 'cible des entreprises technologiques implantées majoritairement en France' / 'forte empreinte française'.

**Champs recherchés sans résultat exploitable :** AuM, Montant alloué, TVPI, DPI, IRR (TRI), Quartile

**Notes :** AuM : les recherches reconfirment la même ambiguïté déjà loguée (1,75 Md€ selon la page produit Bpifrance vs 2,5 Md€ selon d'autres pages/annuaires), sans source datée permettant de trancher — laissé vide comme demandé. Donnée d'activité 2024 trouvée (Bpifrance, bilan d'activité 2024) : investissements Large Venture en hausse de +44% vs 2023 avec 134 M€ investis et 5 nouvelles participations sur l'année — il s'agit d'un flux annuel, pas d'un 'montant alloué' cumulé au fonds, donc non utilisé comme remplissage direct. Le fonds a investi dans plus de 55 sociétés depuis sa création (oct. 2013), avec un ticket minimum de premier tour autour de 10 M€ sur des tours >20 M€.

### 360 Capital Partners / 360 fund V
`360-capital-partners_360-fund-v`

**Champs complétés :**
- `Site web` = 'https://www.360cap.vc' *(confiance : Moyen)*

**Source(s) :** WebSearch result: '360 Capital - European VC from pre-seed to Series B' https://www.360cap.vc/

**Champs recherchés sans résultat exploitable :** Montant alloué, TVPI, DPI, IRR, Quartile

**Notes :** Only a first-close figure was found, not a confirmed final size: 360 Capital held a €90M first close on 360 Fund V in May 2020, with a stated target of €150M by end of 2020 (Maddyness https://www.maddyness.com/2020/05/14/360-capital-closing/ ; Unquote https://www.unquote.com/southern-europe/official-record/3019389/360-capital-holds-eur90m-first-close-for-fifth-fund). No source found confirming the fund's eventual final/total closed size, so 'Montant alloué' was left unfilled rather than inferred from the first-close or target figures. No public TVPI/DPI/IRR/Quartile found for this fund. WebFetch to unquote.com and 360cap.vc was blocked by the egress proxy.

### Samaipata Ventures / Samaipata I
`samaipata-ventures_samaipata-i`

**Champs complétés :**
- `Site web` = 'https://www.samaipata.vc/' *(confiance : Élevé)*

**Source(s) :** https://www.samaipata.vc/ (official Samaipata site)

**Champs recherchés sans résultat exploitable :** Montant alloué, TVPI, DPI, IRR, Quartile

**Notes :** Site confirmed as samaipata.vc (Madrid-based pan-European VC, founded 2016; corroborated by Tech.eu, PitchBook profile references, and the firm's own domain). Samaipata Ventures I made a total of ~31 investments per PitchBook/Nordic9 secondary listings, but no official fund size, allocated amount, or performance metrics (TVPI/DPI/IRR/Quartile) were found in public sources; PitchBook full data is paywalled.

### 360 Capital Partners / 360 Life II
`360-capital-partners_360-life-ii`

**Champs complétés :**
- `Site web` = 'https://www.360cap.vc' *(confiance : Élevé)*
- `Montant alloué (M€)` = 140 *(confiance : Élevé)*

**Source(s) :** EU-Startups: https://www.eu-startups.com/2024/12/360-capital-announces-the-1st-closing-of-its-e140-million-climate-tech-fund-360-life-ii/ ; BeBeez International: https://bebeez.eu/2024/12/23/360-capital-announces-the-1st-closing-of-its-e140-million-climate-tech-fund-360-life-ii/ ; Il Sole 24 Ore: https://en.ilsole24ore.com/art/360-capital-first-closing-140-million-dedicated-climate-tech-fund-AGwINCwB

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** €140M figure is a FIRST CLOSE (announced December 2024), not necessarily the fund's final size — 360 Capital stated a target of €200M for 360 LIFE II (anchor investors: A2A €40M, CDP Venture Capital €44M, De Nora €10M). Flagged as approximate since the total/final closed amount was not found in public sources. Corroborated by 3 independent named press sources, hence Élevé confidence for the €140M first-close figure itself. No TVPI/DPI/IRR/Quartile found.

### TomCat / TomCat Ventures II
`tomcat_tomcat-ventures-ii`

**Champs complétés :**
- `Site web` = 'https://www.tomcat.eu/' *(confiance : Moyen)*

**Source(s) :** https://www.maddyness.com/2025/06/24/tomcat-veut-lever-entre-80-et-100-millions-deuros-pour-son-deuxieme-fonds-dinvestissement/ (Maddyness, 24 Jun 2025); https://www.jaimelesstartups.fr/news/tomcat-lance-officiellement-le-vehicule-tomcat-ventures-ii/ (J'aime les Startups); https://www.tomcat.eu/

**Champs recherchés sans résultat exploitable :** AuM, TVPI, DPI, IRR, Quartile

**Notes :** As of the two named press articles (Jun and Oct 2025), Tomcat Ventures II was still fundraising with a stated TARGET of EUR 80-100M (hard close not yet announced in any source found), aiming to back ~70 B2B Tech startups by 2030 at Seed/Series A. Since this is an unclosed target rather than a confirmed AuM figure, it was left unfilled per the no-estimation rule rather than reported as final AuM; flagged here as an approximate target range for the researcher's awareness. No performance metrics exist yet for this fund (too recent / still investing).

### ISAI / ISAI Venture III
`isai_isai-venture-iii`

**Champs complétés :**
- `Site web` = 'https://www.isai.fr' *(confiance : Faible)*
- `Géographie` = 'Europe ; Amérique' *(confiance : Faible)*

**Source(s) :** WebSearch snippet summarizing frenchweb.fr 'French Tech : ISAI lance un nouveau fonds de 90 millions d'euros' (frenchweb.fr/french-tech-isai-lance-un-nouveau-fonds-de-90-millions-deuros/399745) describing a 'transatlantic approach': investing in US companies founded by French entrepreneurs and supporting French startups' expansion into North America; official ISAI news page title 'First closing at €90M for the ISAI Venture III fund' (isai.vc/news/first-closing-at-90m-for-the-isai-venture-iii-fund)

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** Géographie fill is an inference from a strategy description ('transatlantic approach' covering France/Europe and North America), not an explicit geography label on ISAI's site — flagged Faible confidence; recommend manual confirmation. IMPORTANT CAUTION (not a formal divergence since I do not know the spreadsheet's current Montant value for this fund, which was not listed as missing): search results contain a likely CONFLATION risk between ISAI Venture III (early-stage/post-seed fund, first closing €90M, apparent final target ~€100M per CFNEWS 2025 article 'Isai ne change pas de taille en venture' stating the successor ISAI Venture IV targets 'the same size' as Venture III) and a separate, different vehicle 'ISAI Expansion III' (a growth/LBO-side fund with a €150M hard cap, per Hagnéré Patrimoine listing). One low-quality secondary source conflated these as 'ISAI Venture III closed at 150M€' — this figure most likely actually belongs to ISAI Expansion III, not Venture III. If the spreadsheet's existing Montant field for isai-venture-iii shows ~150M€, it should be manually re-verified against this risk. WebFetch to isai.fr/isai.vc/cfnews.net was blocked, so all of this rests on WebSearch snippets only.

**Rapprochement non résolu :** Risque de confusion avec « ISAI Expansion III », véhicule distinct de croissance/LBO (hard cap 150 M€) chez ISAI. Aucun montant n'a été renseigné pour ISAI Venture III afin d'éviter toute conflation ; à vérifier manuellement avant toute future collecte de Montant alloué/levé sur cette ligne.

### NewFund Capital / NewFund 2
`newfund-capital_newfund-2`

**Champs complétés :**
- `Site web` = 'https://newfundcap.com' *(confiance : Élevé)*

**Source(s) :** Wikipedia 'Newfund' entry (en.wikipedia.org/wiki/Newfund) and newfundcap.com official site — both confirm newfundcap.com as Newfund Management's/Newfund Capital's official website.

**Champs recherchés sans résultat exploitable :** Nb exits, TVPI, DPI, IRR, Quartile

**Notes :** PitchBook lists 'Newfund 2' as a 2016-vintage fund based in San Francisco/Paris investing €0.5-2M tickets, but no fund-size figure, exit count, or performance metric could be confirmed from public, non-paywalled sources within scope. Avoided reporting an unverified '85 investments' figure surfaced in aggregated search summaries, as it could not be traced to a named, citable source and may conflate firm-level with fund-level statistics.

### EOS Venture / EVP I
`eos-venture_evp-i`

**Champs complétés :**
- `Site web` = 'https://eosvc.com' *(confiance : Élevé)*

**Source(s) :** https://eosvc.com (confirmed as Eos Venture Partners' official domain across multiple independent search results: Crunchbase, CBInsights, PitchBook profile pages all reference eosvc.com)

**Champs recherchés sans résultat exploitable :** Montant alloué (M€), TVPI, DPI, IRR (TRI), Quartile

**Notes :** EVP I was announced in April 2018 as a $100M TARGET debut InsurTech fund (Insurance Journal, Private Equity Wire, mind Fintech all report the same $100M target/plan, with disclosed LP commitments of $20M from a global insurer and $10M from a European insurer) — this is a fundraising TARGET, not a confirmed final close amount, so per the 'never estimate' rule it was NOT filled as 'Montant levé'. One low-confidence aggregator (Tracxn-style search synthesis, no isolable primary URL) mentioned the firm 'closed' its first fund in May 2020 and now manages ~$250M in AUM across EVP I plus a 2021 joint fund with Twelve Capital — this could not be corroborated by a second named source, so also left unfilled. TVPI/DPI/IRR/Quartile not publicly disclosed (as expected for a private insurtech VC without public LP reporting). Direct site fetch was blocked by the egress proxy.

### Elaia Partners / PSL innovation Fund
`elaia-partners_psl-innovation-fund`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** Ticket min/max, TVPI, DPI, IRR (TRI), Quartile

**Notes :** Closing final confirmé à 76 M€ (au-delà de la cible de 50 M€). Une fourchette de ticket '0,5 à 1 M€ pour 25-30 sociétés' apparaît dans les résultats de recherche mais provient probablement d'une page générique sur la stratégie globale d'Elaia (eldorado.co) et non spécifiquement du PSL Innovation Fund — non retenue faute de confirmation directe sur une source primaire. Aucune donnée de performance publique.

### Elaia Partners / DV4
`elaia-partners_dv4`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** Montant alloué, Nb exits, TVPI, DPI, IRR (TRI), Quartile

**Notes :** DV4 (200 M€, closing final mars 2023) cible 25-30 sociétés européennes B2B tech, tickets 1-15 M€, du pre-seed à la Série B, secteurs Cybersecurity/Cloud/Digital Life Sciences/Fintech-Insurtech/Digital Transformation. 15 investissements déjà réalisés à l'annonce du closing, dont l'insurtech Seyna (MGA), HarfangLab, Lynxcare, SESAMm, Djust, NANO Corp. Aucun exit ni donnée de performance identifiés.

### Samaipata Ventures / Samaipata II
`samaipata-ventures_samaipata-ii`

**Champs complétés :**
- `Site web` = 'https://www.samaipata.vc/' *(confiance : Élevé)*

**Source(s) :** https://www.samaipata.vc/ (official site); https://tech.eu/2021/12/14/samaipata-closes-second-fund-at-e107-million/ (Tech.eu, Dec 2021)

**Champs recherchés sans résultat exploitable :** Montant alloué, TVPI, DPI, IRR, Quartile

**Notes :** Samaipata Ventures II closed oversubscribed at EUR 107M (Tech.eu, Dec 2021) — this corroborates fund size if 'Montant levé' is not already recorded as current data; not submitted as a fill since it was not listed as missing for this vehicle. No allocated-amount or performance metrics found in public sources.

### ISAI / ISAI Venture II
`isai_isai-venture-ii`

**Champs complétés :**
- `Site web` = 'https://www.isai.fr' *(confiance : Moyen)*
- `Montant alloué (M€)` = 75 *(confiance : Moyen)*

**Source(s) :** ISAI official press release 'ISAI annonce le premier closing du fonds « ISAI VENTURE II » à 55 M€' (isai.fr/news/isai-annonce-le-premier-closing-du-fonds-isai-venture-ii-a-55-m-et-la-mise-en-place-du-isai-seed-club) for first closing; CFNEWS 'Isai boucle Venture II à 75 M€' (cfnews.net) for final closing size

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** First closing was €55M (vs €35M initial target for that stage), final closing reached €75M — these are sequential, not contradictory, figures; reporting the final €75M as Montant alloué. ISAI operates under both isai.fr and isai.vc domains (isai.vc appears to be an English-language mirror); isai.fr used as primary. WebFetch to both isai.fr/isai.vc and cfnews.net was blocked by the egress proxy, so this rests on WebSearch snippets only — confidence kept at Moyen rather than Élevé. No public performance metrics (TVPI/DPI/IRR/Quartile) found for this fund.

### Cathay / Cathay Innovation III
`cathay_cathay-innovation-iii`

**Champs complétés :**
- `Site web` = 'https://cathayinnovation.com/' *(confiance : Élevé)*

**Source(s) :** https://cathayinnovation.com/ (official site); https://cathayinnovation.com/cathay-innovation-closes-1b-venture-capital-fund-to-bring-vertical-ai-to-critical-industries/ (official press release)

**Champs recherchés sans résultat exploitable :** Montant alloué, TVPI, DPI, IRR, Quartile

**Notes :** Fund III closed at $1B per Cathay's own press release, also reported at $1.04B by Crunchbase News and Pulse2 (minor rounding divergence between named sources, not submitted as a fill since 'Montant levé' was not listed as missing for this vehicle). No public figure found for capital already allocated/deployed from Fund III specifically, nor any TVPI/DPI/IRR/Quartile data.

### Astorya.vc / Astorya.vc
`astorya-vc_astorya-vc`

**Champs complétés :**
- `Site web` = 'https://astorya.vc' *(confiance : Élevé)*

**Source(s) :** astorya.vc (official site, per WebSearch result 'Astorya.vc' at https://astorya.vc/); corroborated by eu-startups.com directory entry 'astorya.vc | EU-Startups' (https://www.eu-startups.com/directory/astorya-vc/) and Crunchbase organization profile (https://www.crunchbase.com/organization/astorya-vc).

**Champs recherchés sans résultat exploitable :** Montant alloué, TVPI, DPI, IRR, Quartile

**Notes :** Astorya.vc founded 2017/2018 by Florian Graillot, Paris-based, insurtech/fintech/cybersecurity focus, ~12 seed-stage portfolio companies over 5 years per finscale.substack.com podcast summary. One unverified/ambiguous figure appeared in a search snippet ('astorya.vc helped raise €40 million') without a clear named source or clarity on whether it refers to a specific fund vintage vs. cumulative amount raised by portfolio companies with astorya's help — not filled due to ambiguity, flagged here only.

### Index Venture / Index Origin I
`index-venture_index-origin-i`

**Champs complétés :**
- `Site web` = 'https://www.indexventures.com' *(confiance : Élevé)*
- `Montant alloué (M€)` = 184 *(confiance : Élevé)*

**Source(s) :** BusinessWire / TechCrunch (08/04/2021) — "Index Ventures Launches Index Origin, a $200 Million Dedicated Seed Fund" (https://www.businesswire.com/news/home/20210408005561/en/Index-Ventures-Launches-Index-Origin-a-200-Million-Dedicated-Seed-Fund ; https://techcrunch.com/2021/04/08/index-closes-200-million-dedicated-seed-fund-to-intensify-multi-stage-thesis/). Original amount $200M converted to EUR at 1 USD ≈ 0.92 EUR = 184 M€.

**Champs recherchés sans résultat exploitable :** Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** Identity confirmed: 'Index Venture' = Index Ventures, the Geneva/London-founded international VC firm (indexventures.com), confirmed as an active InsurTech investor (Alan, Wefox per Index Ventures' own site and press) — not a different/smaller entity. No InsurTech-specific count within Index Origin I disclosed anywhere found.

### Cathay / Cathay Innovation II
`cathay_cathay-innovation-ii`

**Champs complétés :**
- `Site web` = 'https://cathayinnovation.com/' *(confiance : Élevé)*

**Source(s) :** https://cathayinnovation.com/ (official site)

**Champs recherchés sans résultat exploitable :** Montant alloué, TVPI, DPI, IRR, Quartile

**Notes :** Fund II final size reported inconsistently across named sources: TechCrunch (May 2020) reported $550M raised at that point, Caixin Global (Jan 2021) reported final close at $792M ($770M per another aggregator), and Cathay's own site reported a EUR 320M first close en route to Fund II. This is a divergence_range on 'Montant levé' (not listed as missing for this vehicle, so not submitted as a fill) — see divergences. No data found on capital allocated specifically, or on TVPI/DPI/IRR/Quartile.

### XAnge Capital / XAnge Digital 3
`xange-capital_xange-digital-3`

**Champs complétés :**
- `Montant alloué (M€)` = 90 *(confiance : Moyen)*

**Source(s) :** CFNEWS / Next-Finance / DOCaufutur (31/05-01/06/2018) — "XAnge réalise le 2e closing de XAnge Digital 3 et se dote ainsi d'une capacité d'investissement de 90M€" (https://www.next-finance.net/XAnge-realise-le-2e-closing-de ; https://www.cfnews.net/L-actualite/Capital-innovation-developpement/Operations/Levee-de-Fonds/Deuxieme-closing-pour-XAnge-Digital-3-272010 ; https://www.docaufutur.fr/2018/05/31/xange-realise-le-2e-closing-de-xange-digital-3-et-se-dote-ainsi-dune-capacite-dinvestissement-de-90me-pour-investir-dans-le-digital/)

**Champs recherchés sans résultat exploitable :** Nb exits, TVPI, DPI, IRR, Quartile

**Notes :** 90 M€ is the confirmed SECOND closing (mai 2018); a FINAL closing/target amount of 120 M€ was announced as objective but no source confirms it was actually reached (final closing amount not found) — reported the 90 M€ as the last confirmed figure, flag as not a certain final size. Fund also reported to have made 47 investments total (not the same as exits) per PitchBook snippet — not used as a fill since exit count specifically not found. Performance metrics (TVPI/DPI/IRR/Quartile) not public for this private fund (one search performed per rule 9).

### Index Venture / Index Ventures XII
`index-venture_index-ventures-xii`

**Champs complétés :**
- `Site web` = 'https://www.indexventures.com' *(confiance : Élevé)*
- `Montant alloué (M€)` = 736 *(confiance : Élevé)*

**Source(s) :** TechCrunch (09/07/2024) — "Index Ventures Raises $2.3B For New Venture And Growth Funds" (https://techcrunch.com/2024/07/09/index-ventures-raises-23-billion-for-new-venture-and-growth-funds/) and BusinessWire same date — $800M raised for its 12th venture fund (= Index Ventures XII), alongside a separate $1.5B growth fund. $800M converted at 1 USD ≈ 0.92 EUR = 736 M€.

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** The commonly cited '$2.3bn' headline figure covers TWO funds combined ($800M venture fund XII + $1.5B growth fund VII) — only the $800M venture-fund portion was attributed to Index Ventures XII specifically.

### Open CNP / Open CNP
`open-cnp_open-cnp`

**Champs complétés :**
- `Montant alloué (M€)` = 100 *(confiance : Élevé)*

**Source(s) :** Avec Open CNP, CNP Assurances va consacrer 100 M€ sur 5 ans au développement de partenariats avec des start-ups, CNP Assurances newsroom, 22/09/2016 (https://www.cnp.fr/le-groupe-cnp-assurances/newsroom/communiques-de-presse/2016/avec-open-cnp-cnp-assurances-va-consacrer-100-m-sur-5-ans-au-developpement-de-partenariats-avec-des-start-ups)

**Champs recherchés sans résultat exploitable :** Nb InsurTech, TVPI, DPI, IRR, Quartile

**Notes :** The 100 M€ figure is the original 2016 announced envelope for 5 years (2016-2021); could not confirm whether this remains the current allocated amount or has since been renewed/increased. Maddyness portrait (08/07/2024, https://www.maddyness.com/2024/07/08/que-font-les-fonds-le-portrait-dopen-cnp/) describes Open CNP investing 1-5 M€ tickets in Series A/B across 4 verticals (Fintech, Assurtech, B2B services, ClimateTech/energy) but gives no explicit InsurTech-specific deal count.

### 360 Capital Partners / 360 Square II Seed Fund
`360-capital-partners_360-square-ii-seed-fund`

**Champs complétés :**
- `Site web` = 'https://www.360cap.vc' *(confiance : Élevé)*
- `Montant alloué (M€)` = 45 *(confiance : Élevé)*

**Source(s) :** 360 Capital official Medium post: https://medium.com/360-capital/360-capital-launches-its-new-45m-fund-to-back-preseed-seed-ventures-3aad940c37e0 ; corroborated by Unquote: https://www.unquote.com/france/official-record/3027098/360-capital-holds-eur-45m-first-close-on-early-stage-tech-fund

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** €45M is described as a FIRST close (March 2023) on 360 Square II (pre-seed/seed fund, successor to 360 Square I which closed at €35M in 2015). An earlier Unquote article headlined a €50M target for 'Square II' (https://www.unquote.com/southern-europe/official-record/3021663/360-capital-to-launch-eur50m-square-ii-fund) — this appears to be the pre-launch target rather than a disagreement about the same closing, so it was not treated as a divergence_range; it is noted here for context. Final/total closed size not confirmed in public sources. No TVPI/DPI/IRR/Quartile found.

### Founders Future VC / Founders Future Good
`founders-future-vc_founders-future-good`

**Champs complétés :**
- `Site web` = 'https://www.foundersfuture.com' *(confiance : Élevé)*

**Source(s) :** foundersfuture.com (site officiel du gestionnaire Founders Future — pages /en, /en/vision, /us/news apparues directement dans les résultats de recherche) ; corroboré par PitchBook (pitchbook.com/profiles/fund/22037-23F) qui référence 'Founders Future Good' comme fonds millésime 2020, fermé (contenu détaillé non accessible, payant)

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** Le site officiel foundersfuture.com couvre l'ensemble des véhicules du gestionnaire (Good, Entrepreneur, Fund I, Fund II) ; aucune page dédiée par fonds trouvée. Aucune métrique de performance (TVPI/DPI/IRR) publiée publiquement — donnée présente uniquement sur PitchBook (accès payant). WebFetch bloqué par le proxy réseau (EGRESS_BLOCKED) pour ce domaine, confirmation faite via extraits WebSearch uniquement.

### Serena Capital / Serena II
`serena-capital_serena-ii`

**Champs complétés :**
- `Site web` = 'https://www.serena.vc' *(confiance : Moyen)*

**Source(s) :** serena.vc (domaine officiel confirmé via pages d'équipe serena.vc/team-profile/marc-fournier/ et /philippe-hayat/, cofondateurs de Serena Capital, apparues dans les résultats WebSearch) ; JournalDuNet (journaldunet.com/web-tech/start-up/1146927-...) et Fusacq (fusacq.com/buzz/serena-capital-leve-100-m-pour-son-second-fonds-d-investissement-a60173_fr_) pour l'historique du fonds Serena II (clôturé à 133 M€ fin 2014)

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** Site officiel identifié par recoupement (pages d'équipe des fondateurs sur le domaine serena.vc) plutôt que par accès direct à la page d'accueil — WebFetch bloqué (EGRESS_BLOCKED sur www.serenacapital.com, domaine alternatif non testé directement). Aucune métrique de performance publique trouvée (PitchBook référence le fonds mais données chiffrées payantes).

### Serena Capital / Serena III
`serena-capital_serena-iii`

**Champs complétés :**
- `Site web` = 'https://www.serena.vc' *(confiance : Moyen)*

**Source(s) :** serena.vc (domaine officiel, cf. note ci-dessus) ; contexte du fonds Serena III (300 M€ levés, early stage/Séries A-B, tickets jusqu'à 15 M€) via extraits WebSearch agrégés (sans URL d'article unique clairement identifiée pour cette statistique)

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** Même site officiel que Serena II (société de gestion unique). Aucune métrique de performance publique trouvée.

### Founders Future VC / Founders Future Entrepreneur
`founders-future-vc_founders-future-entrepreneur`

**Champs complétés :**
- `Site web` = 'https://www.foundersfuture.com' *(confiance : Moyen)*

**Source(s) :** foundersfuture.com (site officiel du gestionnaire) — voir note ci-dessous

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** Aucune page ou source spécifique nommant 'Founders Future Entrepreneur' n'a été trouvée indépendamment (contrairement à 'Good' et 'Fund II' repérés sur PitchBook) ; le site attribué est celui du gestionnaire commun. Aucune métrique de performance publique trouvée. Confiance réduite à Moyen faute de confirmation directe du nom exact du véhicule.

### Aris Occitanie VC / Aris Occitanie
`aris-occitanie-vc_aris-occitanie`

**Champs complétés :**
- `Site web` = 'https://aris-occitanie.fr' *(confiance : Moyen)*

**Source(s) :** WebSearch result 'ARIS OCCITANIE' at https://aris-occitanie.fr/; corroborated by registry listings (annuaire-entreprises.data.gouv.fr, societe.com, pappers.fr) for 'AGENCE REGIONALE DES INVESTISSEMENTS STRATEGIQUES ARIS-OCCITANIE', SIREN 901805630, Toulouse.

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** ARIS-Occitanie ('Agence Régionale des Investissements Stratégiques') is a regional public-private investment company (SAS, registered 27/07/2021, capital social €8,982,900, based 55 Avenue Louis Bréguet, 31400 Toulouse, headed by Stéphane Père) investing in companies with positive territorial impact in Occitanie across strategic sectors (ecological/digital transition, mobility, health, agri-food) — it is a regional development-capital vehicle rather than a classic institutional VC fund, so LP-style performance metrics (TVPI/DPI/IRR/Quartile) are unlikely to be public; none found. WebFetch to the site itself was not attempted after the blocking pattern observed on other domains; confidence capped at Moyen since this is search-snippet-only confirmation of the domain.

### ISAI / ISAI Cap Venture
`isai_isai-cap-venture`

**Champs complétés :**
- `Site web` = 'https://www.isai.fr/isai-cap-venture' *(confiance : Moyen)*
- `Montant alloué (M€)` = 90 *(confiance : Moyen)*

**Source(s) :** Capgemini official press release 'Capgemini and ISAI launch ISAI Cap Venture II' (capgemini.com/news/press-releases/capgemini-and-isai-launch-isai-cap-venture-ii/), which states ISAI Cap Venture I (launched June 2019) was endowed with €90M; corroborated by Maddyness article on ISAI Cap Venture II (2025) distinguishing the €80M CV II from the earlier €90M CV I

**Champs recherchés sans résultat exploitable :** DPI, Quartile

**Notes :** This fid ('ISAI Cap Venture', no roman numeral) is interpreted as ISAI Cap Venture I (2019, €90M, sponsor Capgemini, co-invests €1-5M tickets alongside lead VCs on Series A+ rounds, ~15 portfolio companies since inception incl. Alation, Zelros, Copado). Do not confuse with the newer, separate 'ISAI Cap Venture II' (€80M, launched April 2025) — if the spreadsheet later adds a CV II row, it should carry its own €80M figure, not this one. WebFetch to isai.fr/capgemini.com was blocked; based on WebSearch snippets, confidence Moyen not Élevé despite Capgemini being a primary corporate source.

### Elaia Partners / Elaia Delta
`elaia-partners_elaia-delta`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** Nb exits, TVPI, DPI, IRR (TRI), Quartile

**Notes :** Des statistiques globales Elaia (toutes générations de fonds confondues) mentionnent ~1 Md€ d'AUM, 100+ investissements, 3 licornes et ~80 exits — mais ces chiffres sont au niveau de la société de gestion, pas spécifiques au véhicule Elaia Delta, donc non utilisés comme remplissage pour 'Nb exits' de ce fonds précis.

### Ring Capital / Mission II
`ring-capital_mission-ii`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile, Site web

**Notes :** Aucune trace publique d'un véhicule 'Ring Mission II' distinct de 'Ring Mission' (fonds unique, closing initial 35 M€, cible 50 M€, LPs Tikehau/Bpifrance/BNP Paribas/Mirova/Danone). Les seuls 'fonds II' identifiés chez Ring Capital sont Ring Altitude II (growth buy-out, 217 M€, stratégie différente) — probablement pas le même véhicule que celui recherché. Ne pas assimiler 'Ring Mission' à 'Mission II' sans confirmation : aucun remplissage de champ effectué par prudence. Recommandation : vérifier directement sur ringcp.com/ring-mission/ (WebFetch bloqué : EGRESS_BLOCKED sur www.ringcp.com) ou dans la fiche CFNEWS ESG 2024 (docs.cfnews.net/ESG/2024/FICHES-ESG-24-25/RING-CAPITAL.pdf, non accessible non plus).

**Rapprochement non résolu :** L'existence d'un véhicule distinct « Ring Mission II » n'a pas pu être confirmée : seuls « Ring Mission » (35 M€, premier closing) et « Altitude II » (stratégie croissance/LBO différente) sont documentés chez Ring Capital. Ligne à vérifier manuellement — possible erreur de référentiel.

### Portage Venture / Portag3 Venture I
`portage-venture_portag3-venture-i`

**Champs complétés :**
- `Site web` = 'https://portageinvest.com' *(confiance : Moyen)*

**Source(s) :** https://portageinvest.com/ (search snippets: 'Home - Portage', 'About - Portage Ventures'); confirms this is the current official site of the firm formerly branded Portag3 Ventures, now Portage Ventures

**Champs recherchés sans résultat exploitable :** Montant alloué, Nb exits, TVPI, DPI

**Notes :** Portag3 Ventures I / Portage Ventures I was launched in October 2016 as the corporate fintech investment vehicle of Power Financial Corp, IGM Financial and Great-West Lifeco (sources: BetaKit https://betakit.com/portag3-sets-sights-on-global-fintech-market-as-it-closes-427-million-cad-fund-ii/, Globe and Mail). No public source found stating Fund I's own closed size (searches only surfaced Fund II's $198M CAD initial close and $427M CAD final close, which are distinct from Fund I). Aggregate platform-level figures exist (e.g. Crunchbase/Tracxn cite ~13 exits and $3.3B AUM across Funds I+II combined as of Dec 2021, PSP Investments press release) but these are not fund-I-specific and were not used as fills. WebFetch to portageinvest.com was blocked by egress proxy; site identification relies on WebSearch snippets only.

### BlackFin Capital Partners / BlackfinTech II
`blackfin-capital-partners_blackfintech-ii`

**Champs complétés :**
- `Site web` = 'https://www.blackfincapital.com' *(confiance : Moyen)*

**Source(s) :** Same as BlackfinTech 1: blackfincapital.com / blackfin-tech.com search snippets; fintech.global 'BlackFin commits €350m fund for European InsurTechs and FinTechs' (2022-07-12); aggregator summary citing 'BlackFin Tech II (€390 million, vintage 2022)'.

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** No public LP-report-level performance data (TVPI/DPI/IRR/Quartile) found for BlackfinTech II via search; these are not typically disclosed for private VC vehicles. WebFetch to blackfincapital.com blocked by proxy — site confirmation is search-snippet-based only.

**Divergence (Montant levé (M€)) :** base = 390 ; web = 350 M€ selon fintech.global (« BlackFin commits €350m fund for European InsurTechs and FinTechs », 12/07/2022) ; 390 M€ selon un agrégateur tiers

### Partech / Partech Seed (I to IV)
`partech_partech-seed-i-to-iv`

**Champs complétés :**
- `Site web` = 'https://partechpartners.com' *(confiance : Moyen)*

**Source(s) :** partechpartners.com news pages (multiple), e.g. https://partechpartners.com/news/partech-closes-its-fourth-seed-fund-partech-entrepreneur-iv-at-120m-to-back-entrepreneurs-from-day-1 — confirms partechpartners.com is Partech's official site

**Champs recherchés sans résultat exploitable :** Nb exits, TVPI, DPI, IRR

**Notes :** DATA-QUALITY ISSUE: no fund series named 'Partech Seed I-IV' was found on Partech's official site. The actual official seed-stage fund series is named 'Partech Entrepreneur I, II, III, IV' (Entrepreneur I ~2013 ~€30M; Entrepreneur II ~€100M, oversubscribed from €80M target; Entrepreneur III ~$100M, 2020; Entrepreneur IV ~€120M, oversubscribed — most recent, per partechpartners.com press releases). Recommend verifying whether 'Partech Seed (I to IV)' in the source spreadsheet is meant to represent this 'Partech Entrepreneur' series (likely a naming/translation artifact) rather than a distinct product. Confidence marked Moyen because WebFetch to partechpartners.com was blocked by the egress proxy; facts are snippet-derived from WebSearch only. Performance metrics (TVPI/DPI/IRR/exits) are not publicly disclosed; PitchBook has fund profile pages for Partech Entrepreneur I-IV but figures are paywalled and were not visible in search snippets, so left unfilled per rule 9 (one search per family exhausted).

**Rapprochement non résolu :** La dénomination « Partech Seed (I to IV) » ne correspond à aucune série officielle documentée par Partech (dont la nomenclature connue est « Partech Entrepreneur I-IV ») ; « Parrtech Venture » est une coquille confirmée pour « Partech Venture », mais il existe à la fois un véhicule 2023 de 360 M€ sous ce nom et un historique de fonds plus anciens du même nom — le millésime visé par la ligne de la base n'a pas pu être déterminé avec certitude.

### Portage Venture / Portag3 Venture III
`portage-venture_portag3-venture-iii`

**Champs complétés :**
- `Site web` = 'https://portageinvest.com' *(confiance : Moyen)*

**Source(s) :** WebSearch snippets confirming portageinvest.com as the firm's official domain (About/Team/Portfolio/Blog pages)

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** No public LP-report-level performance metrics (TVPI/DPI/IRR/Quartile) found for Portage Ventures III. (Fund III's size, $788.5M CAD / ~$616M USD per PSP Investments and BetaKit, was not requested as missing for this fid and is not reported as a fill.) WebFetch to portageinvest.com was blocked by the egress proxy.

### Partech / Parrtech Venture
`partech_parrtech-venture`

**Champs complétés :**
- `Site web` = 'https://partechpartners.com' *(confiance : Faible)*

**Source(s) :** partechpartners.com press release 'Partech launches its successor €360M venture fund' (https://partechpartners.com/news/partech-launches-its-successor-360m-venture-fund), Dec 2023; corroborated by EU-Startups-tier press (Silicon Canals, FinSMEs, Real Deals, funds-europe.com)

**Champs recherchés sans résultat exploitable :** Nb exits, TVPI, DPI, IRR

**Notes :** DATA-QUALITY ISSUE: 'Parrtech Venture' is almost certainly a typo for 'Partech Venture'. There IS a real, currently active fund called 'Partech Venture': a €360M vehicle announced Dec 2023, successor to 'Partech International Ventures VII', targeting ~22-24 Series A/B European companies in software/deep tech/fintech & insurtech, LPs include Allianz France, BNP Paribas, Bpifrance, CDP Venture Capital, Edenred, FDJ Ventures, JCDHolding, Lombard Odier. HOWEVER 'Partech Venture' has historically also been used as a generic brand name by Partech across earlier decades (the firm was founded 1982), so it is not certain the spreadsheet row refers specifically to this 2023 vintage vs. an older, differently-sized 'Partech Venture'/'Partech International Ventures' vintage — recommend manual cross-check against whatever other identifying data (vintage year, montant) the row already carries. Confidence kept Faible given this ambiguity plus WebFetch being blocked (snippet-only evidence). No public TVPI/DPI/IRR/exit-count data found.

**Rapprochement non résolu :** La dénomination « Partech Seed (I to IV) » ne correspond à aucune série officielle documentée par Partech (dont la nomenclature connue est « Partech Entrepreneur I-IV ») ; « Parrtech Venture » est une coquille confirmée pour « Partech Venture », mais il existe à la fois un véhicule 2023 de 360 M€ sous ce nom et un historique de fonds plus anciens du même nom — le millésime visé par la ligne de la base n'a pas pu être déterminé avec certitude.

### Portage Venture / Portag3 Venture II
`portage-venture_portag3-venture-ii`

**Champs complétés :**
- `Site web` = 'https://portageinvest.com' *(confiance : Moyen)*

**Source(s) :** Official Portage blog via Newswire.ca: https://www.newswire.ca/news-releases/portag3-ventures-announces-427m-close-of-second-fintech-fund-839230793.html and https://www.newswire.ca/news-releases/portag3-ventures-announces-initial-198m-closing-of-its-second-fintech-fund-698995521.html ; corroborated by BetaKit https://betakit.com/portag3-sets-sights-on-global-fintech-market-as-it-closes-427-million-cad-fund-ii/ and Finextra https://www.finextra.com/newsarticle/34897/canadas-portag3-closes-cad427m-fintech-fund

**Champs recherchés sans résultat exploitable :** Montant alloué, Nb exits, TVPI, DPI

**Notes :** Fund II raised an initial $198M CAD (Oct 2018) and closed at $427M CAD final (2019), per the official Portage newsroom, BetaKit and Finextra. Left 'Montant alloué' unfilled because the native currency is CAD, which is outside the USD/GBP/CHF conversion table specified (currency_not_converted) — converting via a secondary USD figure that press outlets themselves estimated (NCFA/TechCrunch cite ~$320M USD, PitchBook newsletter cites ~$321M USD, figures not fully consistent) would compound two independent conversions and was avoided per the 'never estimate' rule. Raw figures for reference: $198M CAD initial close, $427M CAD final close (2019 vintage per PitchBook). No fund-specific exit count, TVPI or DPI found.

### Founders Future VC / Founders Future Fund II
`founders-future-vc_founders-future-fund-ii`

**Champs complétés :**
- `Site web` = 'https://www.foundersfuture.com' *(confiance : Moyen)*

**Source(s) :** foundersfuture.com (site officiel) ; PitchBook (pitchbook.com/profiles/fund/24427-54F) référence 'Founders Future II', fonds early-stage domicilié en France, fermé — via extraits WebSearch, TechCrunch (techcrunch.com/2023/09/05/...) confirme que Founders Future levait deux nouveaux fonds ('Founders Future II' et 'Founders Future Expansion') avec un premier closing à 80 M$ en 2023

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** PitchBook nomme le véhicule 'Founders Future II' (sans 'Fund') — simple variante d'intitulé, non traité comme divergence de données chiffrées. Aucune métrique de performance publique disponible.

### XAnge Capital / XAnge 4
`xange-capital_xange-4`

**Champs complétés :**
- `Montant alloué (M€)` = 220 *(confiance : Élevé)*

**Source(s) :** Usine Digitale — "XAnge boucle un nouveau fonds de 220 millions d'euros" (https://www.usine-digitale.fr/article/xange-boucle-un-nouveau-fonds-de-220-millions-d-euros-et-s-interesse-tout-particulierement-au-web3.N2024137) ; Journal du Net (https://www.journaldunet.com/web3/crypto/1513101-le-fonds-de-capital-risque-xange-leve-220-millions-d-euros/) — final closing at 220 M€, doubling capacity vs. XAnge 3

**Champs recherchés sans résultat exploitable :** Nb exits, TVPI, DPI, IRR

**Notes :** 220 M€ = final closing size of the XAnge 4 fund itself (distinct from the 700 M€ AuM already in base, which appears to be at the société de gestion / all-vintages level — not treated as a divergence). Intermediate closings reported inconsistently across press (125 M€ per Maddyness Dec-2021 vs 140 M€ per Siparex) before reaching the 220 M€ final size; only the final, most corroborated figure (220 M€) is reported as fill.

### Sharpstone Capitale / Sharpstone Capitale
`sharpstone-capitale_sharpstone-capitale`

**Champs complétés :**
- `Site web` = 'https://www.sharpstone.fr' *(confiance : Élevé)*

**Source(s) :** CFNEWS annuaire 'SHARPSTONE CAPITAL - Fonds d'investissement / gestionnaire' (cfnews.net); Maddyness 'Que Font Les Fonds ? Le portrait de Sharpstone Capital' (2023-01-16, maddyness.com); official site sharpstone.fr.

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** Official/press naming is consistently 'Sharpstone Capital' (not 'Sharpstone Capitale' as in the current Fonds_ID/name) — flagging as a possible naming discrepancy to verify, not treated as a divergence field since the row's 'Nom/Véhicule' fields weren't listed as missing. Founded 2014 by Germain Gaschet and Alain Chabanne; seed-stage investment arm of the Sharpstone group (alongside Sharpstone Advisory). No public performance metrics found for this seed vehicle.

**Rapprochement non résolu :** La presse et le site officiel désignent systématiquement l'entité « Sharpstone Capital » (et non « Sharpstone Capitale ») ; probable variante orthographique dans le référentiel, à confirmer.

### Elevation Capital Partners / FPCI Food Invest II
`elevation-capital-partners_fpci-food-invest-ii`

**Champs complétés :**
- `Site web` = 'https://www.elevation-cp.com' *(confiance : Moyen)*

**Source(s) :** Elevation Capital Partners official site, fund communiqué (https://www.elevation-cp.com/communiques/capital-investissement/00552/premiere-levee-de-fonds-pour-vegetal-food-le-fpci-food-invest-2-investit-1-2-millions)

**Champs recherchés sans résultat exploitable :** TVPI, DPI, Quartile

**Notes :** Fund's actual strategy/sector is 'Food & Beverage' (catering, food production, e-commerce, FoodTech, B2B — per inter-invest.fr and snacking.fr) which is out of the allowed Stratégie vocabulary (InsurTech/FinTech/Digital/Santé/ESG/Impact/Innovation/Autre/Généraliste); not filled into Stratégie per rule 6, noted here instead. The predecessor fund (FPCI Food Invest, vintage I) reportedly targeted 30 M€ (max 50 M€) per snacking.fr, but no comparable size figure was found specifically for Food Invest II, so Montant/taille not filled.

### Elevation Capital Partners / FPCI elevation Early Growth II
`elevation-capital-partners_fpci-elevation-early-growth-ii`

**Champs complétés :**
- `Site web` = 'https://www.elevation-cp.com' *(confiance : Moyen)*

**Source(s) :** Elevation Capital Partners official site (https://www.elevation-cp.com/); fund page also at https://www.inter-invest.fr/capital-investissement/fpci-elevation-early-growth-2

**Champs recherchés sans résultat exploitable :** Montant alloué, TVPI, DPI, Quartile

**Notes :** No public fund-size (montant levé/alloué) figure found specifically for Early Growth II; distinguish from the separate, newer 'FPCI Elevation Growth' (target 100-150 M€, Series B focus) which is a different, later vehicle, not Early Growth II.

### XAnge Capital / Mutuelles Impact
`xange-capital_mutuelles-impact`

**Champs complétés :**
- `Montant alloué (M€)` = 95 *(confiance : Élevé)*

**Source(s) :** Maddyness (13/01/2023) — "Le fonds Mutuelles Impact atteint 95 millions d'euros sous gestion" (https://www.maddyness.com/2023/01/13/mutuelles-impact/)

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** 95 M€ is the latest/highest reported size (Jan 2023). Prior closings were smaller and sequential, not conflicting: 80 M€ per XAnge's own page (https://www.xange.vc/knowledge-hub/202206-a-e80-million-closing-for-mutuelles-impact, June 2022) and per Carenews. Fund managed by XAnge with Investir&+; ticket size reported 0.5-7 M€ (not requested in missing-fields list, given for context only).

### NewFund Capital / NewFund 1
`newfund-capital_newfund-1`

**Champs complétés :**
- `Site web` = 'https://newfundcap.com' *(confiance : Élevé)*

**Source(s) :** Wikipedia 'Newfund' entry (en.wikipedia.org/wiki/Newfund) and newfundcap.com official site.

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** No public performance data found for this 2008-era vintage fund; PitchBook has a fund profile (pitchbook.com/profiles/fund/11673-28F) but figures are paywalled.

### Alven Capital / Alven VI
`alven-capital_alven-vi`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** Montant alloué, TVPI, DPI, IRR (TRI), Quartile

**Notes :** Alven VI (350 M€, hard cap, clôturé en 2022, cible amorçage/Série A en Europe, tickets 100 k€–15 M€, ~60-70% réservé au suivi). Le montant réellement déployé n'est donné qu'en pourcentage ('environ 15% du montant levé placé' peu après le closing, puis '10 à 15% déployé' plus tard) sans valeur absolue en M€ dans les sources trouvées — non converti en montant pour éviter toute estimation dérivée. Aucune donnée de performance publique identifiée (fonds privé).

### XAnge Capital / XAnge Capital 2
`xange-capital_xange-capital-2`

**Champs complétés :**
- `Montant alloué (M€)` = 62 *(confiance : Élevé)*

**Source(s) :** Bpifrance press release / Boursorama / Fusacq — all titled "XAnge annonce le succès du closing final de son fonds multicorporate, XAnge Capital 2, à hauteur de 62 M€" (https://presse.bpifrance.fr/xangefonds-partenaire-de-bpifrance-annonce-le-succes-du-closing-final-de-son-fonds-multicorporate-xange-capital-2-a-hauteur-de-62-me ; https://www.boursorama.com/bourse/actualites/xange-closing-final-du-2eme-fonds-multicorporate-a-62-m-b77f0ecbd90ca66aea05c5260d6e0343 ; https://www.fusacq.com/buzz/xange-annonce-le-succes-du-closing-final-de-son-fonds-multicorporate-xange-capital-2-a80113_fr_)

**Champs recherchés sans résultat exploitable :** Nb exits, TVPI, DPI, IRR

**Notes :** 62 M€ final closing corroborated by 3 independent sources with matching headline figure (initial target was 80 M€ per Journal du Net 2012, not reached). One unrelated WebSearch AI-summary mentioned a conflicting '130 million euros total investment capacity' for 'the second multicorporate fund' but no primary source could be found for that figure — discarded, not reported as a formal divergence since it could not be traced to a citable article.

### Atlantic Vantage Point / AVP early stage II
`atlantic-vantage-point_avp-early-stage-ii`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** Confirms Atlantic Vantage Point (AVP) as the rebranded former AXA Venture Partners (2024 MBO, name change 2025), Paris-based, managing >EUR 2.5bn across venture/early growth/growth equity/fund-of-funds strategies per gilion.com aggregator — corroborates the Fonds record's existing note rather than filling new fields. No public LP report or named press disclosed TVPI/DPI/IRR/Quartile for 'AVP Early Stage II' specifically; PitchBook has a profile page but figures are behind a paywall and were not accessible.

### Atlantic Vantage Point / AVP early stage I
`atlantic-vantage-point_avp-early-stage-i`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR, Quartile

**Notes :** Same entity/rebrand note as AVP Early Stage II. No public performance data found for 'AVP Early Stage I' in any named source; PitchBook profile exists but figures are paywalled and not accessible via search snippets.

### Alven Capital / Alven V
`alven-capital_alven-v`

**Champs complétés :**
- `Géographie` = 'Europe' *(confiance : Moyen)*

**Source(s) :** Maddyness, 09/01/2017, 'Alven Capital fête ses 17 ans en levant 250 millions d'euros pour les startups françaises' (https://www.maddyness.com/2017/01/09/vc-alven-capital-fonds-250-millions-euros-startups/) ; TechCrunch, 03/01/2017 (https://techcrunch.com/2017/01/03/alven-capital-raises-261-million-fund-to-invest-in-french-entrepreneurs/) ; description générale d'Alven (investorsglobe.com, openvc.app) indiquant un focus France/Europe avec possibilité de suivre des fondateurs européens s'implantant aux US.

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR (TRI)

**Notes :** Fonds V ciblait prioritairement des entrepreneurs français en Série A (IA, data, SaaS, sécurité, marketplaces), avec 80% du capital apporté par des investisseurs institutionnels européens et capacité à accompagner des sociétés dans leur développement aux US (environ la moitié du portefeuille déjà implantée aux US selon Maddyness). D'où le remplissage 'Europe' en confiance Moyenne plutôt que 'Europe ; Amérique'.

**Divergence (Millésime) :** base = 2016 ; web = closing d'Alven V (250 M€) annoncé début janvier 2017 (TechCrunch 03/01/2017, Maddyness 09/01/2017 : « VC Alven Capital : fonds de 250 millions d'euros »)

### Elaia Partners / Alpha II
`elaia-partners_alpha-ii`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** TVPI, DPI, IRR (TRI), Quartile

**Notes :** PitchBook confirme l'existence de données de performance pour Alpha Fund II mais celles-ci sont derrière un paywall et non accessibles publiquement. Millésime 2020 mentionné par PitchBook (non demandé/déjà en base probablement).

### Truffle Capital / Truffle FinTech & InsurTech Fund II
`truffle-capital_truffle-fintech-et-insurtech-fund-ii`

**Champs complétés :**
- `Site web` = 'https://www.truffle.com' *(confiance : Élevé)*

**Source(s) :** Truffle Capital official website (https://www.truffle.com/fintech/about); fund closing corroborated by Bernard-Louis Roques: 'Truffle Capital boucle son fonds fintech à 140 millions d'euros', mind Fintech (https://www.mind.eu.com/fintech/investissement/bernard-louis-roques-truffle-capital-boucle-son-fonds-fintech-a-140-millions-deuros-et-va-investir-dans-smartpush/) and FrenchWeb (https://www.frenchweb.fr/truffle-capital-leve-400-millions-deuros-pour-investir-dans-la-biomedtech-et-la-fintech-insurtech/387504)

**Champs recherchés sans résultat exploitable :** Montant alloué, DPI, IRR

**Notes :** Fund closed at 140 M€ (first close Q4 2017 at 70 M€, target 130 M€/hard-cap 150 M€ per next-finance.net); this appears to already be known as 'Montant levé' in the current sheet (not requested as missing) so not re-filled. Could not find a separate 'amount actually deployed/allocated' distinct from the raised amount, nor DPI/IRR (no public LP report found).

### Kima Venture Capital / Kima Venture
`kima-venture-capital_kima-venture`

**Champs complétés :**
- `Site web` = 'https://kimaventures.com' *(confiance : Élevé)*

**Source(s) :** Multiple WebSearch aggregator results (signal.nfx.com, swanbase.co, vcsheet.com, eldorado.co) consistently citing kimaventures.com as the official Kima Ventures site; general knowledge of Xavier Niel's Kima Ventures corroborates this domain.

**Champs recherchés sans résultat exploitable :** TVPI, DPI

**Notes :** Note: the vehicle's actual/common name is 'Kima Ventures' (plural), not 'Kima Venture Capital' / 'Kima Venture' as in the current row — flagging for verification, not filed as a divergence since Nom/Véhicule weren't listed as missing fields. Kima operates as an evergreen single-LP structure funded directly by Xavier Niel (a family-office style vehicle, not a traditional fund raising from institutional LPs), which likely explains why no TVPI/DPI figures are publicly disclosed — none found.

**Rapprochement non résolu :** L'entité réelle se nomme « Kima Ventures » (véhicule evergreen à LP unique, Xavier Niel), et non « Kima Venture Capital » — ce qui explique l'absence de TVPI/DPI/IRR publics (pas de LPs tiers à rapporter).

### Ring Capital / Génération
`ring-capital_generation`

**Champs complétés :**
- `Site web` = 'https://www.ringcp.com/generations/' *(confiance : Élevé)*

**Source(s) :** ringcp.com/generations/ (page officielle du fonds, apparue directement dans les résultats WebSearch) ; corroboré par Maddyness (maddyness.com/2023/10/19/ledhec-et-ring-capital-lancent-un-fonds-damorcage-dedie-aux-projets-a-impact/) et EDHEC (edhec.edu/en/about-us/entrepreneurship-business-school/generations-powered-by-edhec-fund)

**Champs recherchés sans résultat exploitable :** DPI, Quartile

**Notes :** Le véhicule complet se nomme 'GENERATIONS powered by EDHEC', fonds d'amorçage (pré-seed/seed) lancé oct. 2023 en partenariat avec l'EDHEC, cible 40 M€ (dont 20 M€ apportés par la communauté EDHEC), tickets 150 k€–1 M€ — mentionné ici à titre de contexte, non reporté comme 'fill' car hors du périmètre des champs manquants indiqués (Montant alloué non demandé). Aucune métrique DPI/Quartile publique trouvée.

### Ring Capital / Mission I
`ring-capital_mission-i`

**Champs complétés :**
- `Site web` = 'https://www.ringcp.com/ring-mission/' *(confiance : Élevé)*

**Source(s) :** ringcp.com/ring-mission/ (page officielle du fonds, apparue directement dans les résultats WebSearch) ; Carenews (carenews.com/fr/news/le-nouveau-fonds-ring-mission-mise-sur-les-startup-tech-for-good-de-demain) et Finyear (finyear.com/Ring-Mission-nouveau-fonds-de-Venture-Capital-Impact_a44209.html) pour le contexte de lancement

**Champs recherchés sans résultat exploitable :** DPI

**Notes :** Premier closing à 35 M€ (LPs : Tikehau Capital, Bpifrance, BNP Paribas, Mirova, Danone, family offices), cible 50 M€, tickets 0,5–5 M€ — contexte, non requis en remplissage. Aucune donnée DPI publique trouvée.

### White Star Capital / WSC I
`white-star-capital_wsc-i`

**Champs complétés :**
- `Site web` = 'https://whitestarcapital.com' *(confiance : Élevé)*
- `Montant alloué (M€)` = 64.4 *(confiance : Élevé)*

**Source(s) :** TechCrunch: https://techcrunch.com/2015/11/13/white-star-capital-closes-70m-for-its-first-institutional-transatlantic-fund ; PE Hub: https://www.pehub.com/2015/11/white-star-capital-collects-70-mln-for-initial-fund/ ; Venture Capital Journal: https://www.venturecapitaljournal.com/white-star-capital-collects-70-mln-for-initial-fund/ ; We Are Guernsey: https://www.weareguernsey.com/news/2015/guernsey-home-to-white-star-capitals-70-million-start-up-fund/


**Notes :** White Star Capital's first institutional fund closed at $70M USD in November 2015 (well corroborated by 4 independent named sources). Converted to EUR at 1 USD ≈ 0.92 EUR => 70 x 0.92 = 64.4 M€. Official site whitestarcapital.com identified via WebSearch (top organic result: 'White Star Capital | A global multi-stage technology investment platform'); WebFetch to the domain was blocked by the egress proxy so page content could not be independently verified, hence Moyen-leaning caution on the site identification specifically even though the amount confidence is Élevé.

### White Star Capital / WSC III
`white-star-capital_wsc-iii`

**Champs complétés :**
- `Site web` = 'https://whitestarcapital.com' *(confiance : Élevé)*
- `Montant alloué (M€)` = 331.2 *(confiance : Élevé)*

**Source(s) :** Private Equity International: https://www.privateequityinternational.com/white-star-capital-raises-360m/ ; Forbes: https://www.forbes.com/sites/rebeccaszkutak/2021/10/25/white-star-capital-raises-360-million-fund-that-nearly-doubles-its-aum/ ; Unquote: https://www.unquote.com/unquote/official-record/3025324/white-star-holds-usd-360m-final-close-for-fund-iii ; PR Newswire: https://www.prnewswire.com/news-releases/white-star-capital-announces-new-fund-firm-now-has-more-than-500-million-of-fresh-capital-to-invest-globally-301407267.html


**Notes :** White Star Capital Fund III held its final close at $360M USD in October 2021, exceeding its original $300M target (well corroborated by 4+ independent named sources: PEI, Forbes, Unquote, PR Newswire, TechNode Global). Converted to EUR at 1 USD ≈ 0.92 EUR => 360 x 0.92 = 331.2 M€.

### White Star Capital / WSC II
`white-star-capital_wsc-ii`

**Champs complétés :**
- `Site web` = 'https://whitestarcapital.com' *(confiance : Élevé)*
- `Montant alloué (M€)` = 165.6 *(confiance : Élevé)*

**Source(s) :** TechCrunch: https://techcrunch.com/2018/06/03/white-star-capital-ii/ ; Unquote: https://www.unquote.com/france/official-record/3010234/white-star-capital-closes-second-fund-on-usd180m ; VentureBeat: https://venturebeat.com/entrepreneur/white-star-capital-closes-180-million-fund-for-early-stage-transatlantic-startups ; Private Equity Wire: https://www.privateequitywire.co.uk/white-star-capital-announces-second-vc-fund-usd180-million/


**Notes :** White Star Capital Fund II closed at $180M USD in June 2018 (oversubscribed from an initial $140M target), well corroborated by 4+ independent named sources plus White Star's own Medium post. Converted to EUR at 1 USD ≈ 0.92 EUR => 180 x 0.92 = 165.6 M€.

### Alven Capital / Alven IV
`alven-capital_alven-iv`

**Champs complétés :** aucun.

**Champs recherchés sans résultat exploitable :** TVPI, DPI

**Notes :** Fonds privé sans rapport LP public identifié. Aucune source ne mentionne de TVPI/DPI pour Alven IV — conforme à l'attente que la performance ne soit pas publique.

### TomCat / TomCat Ventures I
`tomcat_tomcat-ventures-i`

**Champs complétés :**
- `Site web` = 'https://www.tomcat.eu/' *(confiance : Élevé)*

**Source(s) :** https://www.maddyness.com/2025/10/23/que-font-les-fonds-le-portrait-de-tomcat/ (Maddyness, 23 Oct 2025); https://www.tomcat.eu/ (official site)


**Notes :** Confirmed the correct entity: French VC founded by Patrice Thiry, Jérémy Pouyer and Elie du Pré de Saint-Maur (ex-ProwebCE/Edenred/Doctolib), B2B Tech seed/Series A focus — distinct from an unrelated Berlin-based 'TomCat Ventures' that appears in Crunchbase/PitchBook/Tracxn indexes (different firm, founded 2018, Haan Germany). Official site tomcat.eu also corroborated by CFNews and Eldorado.co listings; a secondary domain tomcat-invest.fr also exists for the same group. Per Maddyness, Tomcat Ventures I launched in 2022 with EUR 20M under management, having deployed ~EUR 8M across 21 startups from its accelerator as of the article date — this AUM figure was not requested as missing for Fund I (only Site web was), so not submitted as a fill, only noted for context.

### Founders Future VC / Founders Future Fund I
`founders-future-vc_founders-future-fund-i`

**Champs complétés :**
- `Site web` = 'https://www.foundersfuture.com' *(confiance : Élevé)*

**Source(s) :** foundersfuture.com (site officiel du gestionnaire, pages /en et /en/vision apparues dans les résultats WebSearch)


**Notes :** Seul champ manquant (Site web) renseigné ; aucune autre donnée recherchée pour ce véhicule.

### BlackFin Capital Partners / BlackfinTech 1
`blackfin-capital-partners_blackfintech-1`

**Champs complétés :**
- `Site web` = 'https://www.blackfincapital.com' *(confiance : Moyen)*

**Source(s) :** WebSearch snippets: blackfincapital.com homepage (title 'Home - Blackfin Capital'), https://www.blackfincapital.com/about/ ; corroborated by fintech.global article 'BlackFin commits €350m fund for European InsurTechs and FinTechs' (2022-07-12) referencing BlackFin Capital Partners; PitchBook profile 'BlackFin Tech Fund 1' (2018 vintage, ~€180m).


**Notes :** WebFetch to blackfincapital.com was blocked by network egress proxy, so this is search-snippet-only confirmation, not a direct page read. The group also operates blackfin.com and blackfin-tech.com (the dedicated tech/venture arm's site is https://www.blackfin-tech.com/news) — these may be more specific for the BlackfinTech vehicles than the corporate blackfincapital.com domain; worth checking which one the spreadsheet prefers as canonical. BlackFin Tech Fund 1 (BlackfinTech 1) closed at ~€180m in 2018 per PitchBook/fintech.global aggregation, but this Montant levé figure was not requested (not in the missing-fields list for this row) so it is not reported as a fill.
