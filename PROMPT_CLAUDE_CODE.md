# Prompt à coller dans Claude Code

> Place d'abord dans un même dossier : `VC_Database_Standardisee_v3.xlsx`, `todo_collecte_web.csv`,
> `restructuration_vc.py`, `maj_web_vc.py`. Puis lance `claude` dans ce dossier et colle le texte
> ci-dessous. Autorise les outils WebSearch / WebFetch et l'écriture de fichiers.

---

## Contexte

Tu travailles sur une base de données standardisée de fonds de Venture Capital (périmètre
insurtech Europe/UK). Le dossier contient :

- `VC_Database_Standardisee_v3.xlsx` — la base à enrichir, 5 onglets : `Fonds` (111 véhicules,
  49 colonnes), `Participations`, `Tours_de_table`, `Dictionnaire`, `Anomalies` (178 lignes,
  numérotation `ANO-0001` … `ANO-0178`).
- `todo_collecte_web.csv` — plan de collecte : une ligne par véhicule, avec la liste exacte des
  champs vides à chercher (806 champs manquants au total), le site officiel quand il est déjà
  connu, une colonne `Statut collecte` et une colonne `Notes`.
- `maj_web_vc.py` — script de la passe web précédente : il montre le format attendu des mises à
  jour (`MAJ_WEB`), des divergences (`DIVERGENCES`) et du journal d'anomalies. **Réutilise sa
  structure**, ne repars pas de zéro.
- `restructuration_vc.py` — contient `write_excel()`, qui applique la mise en forme du classeur
  (tableaux structurés, filtres, volets figés, formats de nombres). Importe-le, ne réécris pas
  la mise en forme.

## Objectif

Compléter un maximum de champs vides de l'onglet `Fonds` à partir des **sites officiels des
fonds** et de sources publiques fiables, en produisant `VC_Database_Standardisee_v4.xlsx`.

## Méthode de travail — par lots, reprenable

Traite les véhicules **par lots de 10**, dans l'ordre du fichier `todo_collecte_web.csv`
(déjà trié par nombre de champs manquants décroissant, mais commence plutôt par les entités
les plus documentées : sociétés de gestion européennes établies, CVC d'assureurs, puis les
petites structures).

Pour chaque lot :

1. pour chaque véhicule, cherche d'abord le **site officiel** de la société de gestion, puis la
   page « funds / about / portfolio » et les communiqués de closing ;
2. complète ensuite avec presse spécialisée nommée (CFNEWS, Maddyness, Sifted, TechCrunch,
   EU-Startups, Global Corporate Venturing, Insurance Insider, Tech.eu) ou registres publics
   (AMF, Companies House, Climate Transparency Hub / ADEME, Bundesanzeiger) ;
3. écris les résultats dans le classeur, mets à jour `Statut collecte` et `Notes` dans
   `todo_collecte_web.csv`, **et sauvegarde les deux fichiers avant de passer au lot suivant** ;
4. affiche un récapitulatif du lot : champs remplis, champs restés vides, divergences détectées.

Si tu es interrompu, tu dois pouvoir reprendre uniquement à partir de `todo_collecte_web.csv`.

## Règles de données — non négociables

Ces règles viennent du cahier des charges initial de la base, respecte-les strictement.

1. **Aucune estimation.** Si l'information n'est pas publiée, la cellule reste vide. Ne déduis
   jamais un AuM d'une somme de tickets ni un millésime d'une date d'article.
2. **On ne remplit que des cellules vides.** Une valeur déjà présente n'est jamais écrasée. Si la
   source web contredit la base, tu conserves la base et tu ajoutes une ligne dans `Anomalies`
   de type `Différence entre description et structure réelle`, avec les deux valeurs.
   Seule exception : un booléen de stade à `FAUX` sur une ligne par ailleurs vide peut être
   corrigé en `VRAI` si le fonds documente explicitement ce stade — avec anomalie
   `Correction de nom` à l'appui.
3. **Fourchettes et valeurs divergentes selon les bases** (« 44 (Tracxn) à 97 (PitchBook) ») :
   cellule laissée vide, anomalie `Valeur en fourchette non retenue`, les deux chiffres dans le
   commentaire. Une valeur approximative unique (`~10`, `>150M`) peut être retenue avec une
   anomalie `Valeur approximative conservée`.
4. **Devises.** Les colonnes montants sont en M€ (capital social en k€). Conversions autorisées,
   avec le taux écrit dans le commentaire : `1 USD ≈ 0,92 EUR`, `1 GBP ≈ 1,17 EUR`,
   `1 CHF ≈ 1,07 EUR`. Toute autre devise : cellule vide + anomalie `Conversion non effectuée`.
5. **Unités et formats.** TVPI, DPI, RVPI, MoC, MoEP = multiples décimaux (jamais en %).
   IRR = décimal (0,30 s'affiche 30 %). Millésime = entier. `RVPI = TVPI − DPI` uniquement si
   les deux sont présents.
6. **Vocabulaires fermés.** `Statut` ∈ {Ouvert, Fermé} · `Phase` ∈ {Investissement, Gestion, Clos}
   · `Géographie` ∈ {Europe, Amérique, International} concaténés par « ; » ·
   `Stratégie` ∈ {InsurTech, FinTech, Digital, Santé, ESG, Impact, Innovation, Autre, Généraliste}
   concaténés par « ; » · stades ∈ {Pré-Seed, Seed, Pré-Série A, Série A, Série B, Série C, Série D},
   booléens VRAI/FAUX + concaténation ordonnée dans `Stages pratiqués`.
   Un libellé source qui ne rentre pas dans ces vocabulaires va dans le commentaire d'anomalie,
   pas dans la cellule.
7. **Jamais de marqueur textuel** (`n.a`, `NA`, `-`, `Non disponible`) : une valeur manquante est
   une cellule vide. Un zéro n'est conservé que s'il est manifestement réel.
8. **Traçabilité obligatoire.** Chaque cellule remplie entraîne :
   - `Source_AuM` complété au format `[MAJ web JJ/MM/AAAA] <source précise>` (en concaténant à
     l'existant avec « | », sans l'écraser) ;
   - `Date_MAJ` = date du jour ;
   - une ligne dans `Anomalies` de type `Mise à jour web`, avec la liste des champs écrits, les
     valeurs retenues, l'URL ou le nom de la source, et le niveau de confiance.
   La numérotation des anomalies **reprend à ANO-0179** et reste séquentielle.
9. **Colonne `Site web`** : URL du site officiel de la société de gestion, à compléter pour les
   72 lignes qui en sont dépourvues.
10. **Cohérence intra-société.** L'AuM est une donnée de société de gestion : s'il diffère d'un
    véhicule à l'autre pour une même société, ne l'harmonise pas en silence — signale-le et
    indique la date de référence de chaque valeur dans le commentaire.

## Priorités de recherche

Par ordre d'intérêt, en sachant que les performances ne sont presque jamais publiques :

1. `Site web`, `AuM (M€)`, `Montant levé (M€)`, `Millésime` — les plus atteignables ;
2. `Ticket min (M€)` / `Ticket max (M€)`, `Stages pratiqués` + booléens, `Géographie`,
   `Stratégie` — souvent sur la page « what we look for » du fonds ;
3. `Nb participations (déclaré)`, `Nb exits`, `Nb InsurTech` — via la page portefeuille du site
   officiel, en privilégiant le décompte du fonds lui-même ;
4. `Statut`, `Phase` — dernier closing annoncé et rythme d'investissement observé ;
5. `TVPI`, `DPI`, `IRR (TRI)`, `Quartile` — ne les cherche que pour les fonds ayant des LPs
   publics (fonds de fonds cotés, rapports d'investisseurs publics type British Business Bank,
   EIF, Bpifrance, caisses de retraite américaines). Pour tous les autres, note une fois pour
   toutes qu'il n'y a pas de source publique et passe.

N'invente jamais une performance. Un fonds sans reporting public reste vide sur ces colonnes.

## Cas particuliers déjà identifiés

- Quatre entités sont des **business angels solos ou des structures sans portefeuille vérifiable**
  (David Semmens, Patrice Fleurquin, M. H. Tavangar, The49) : ne cherche pas plus de 5 minutes
  chacune, marque `Statut collecte = non exploitable` et passe.
- **Kickstart Innovation** et **NCA** sont des accélérateurs zero-equity, **Seraphim Space** est
  un fonds SpaceTech hors périmètre insurtech, **Blast.Club** et **South East Angels** sont des
  clubs/réseaux : ni AuM ni ticket au sens VC. Complète seulement site web, millésime, portefeuille.
- **Atlantic Vantage Point** = ex-AXA Venture Partners (MBO 2024, rebranding avril 2025) :
  même entité, ne crée pas de doublon.
- **Element Ventures** est devenu **13books Capital** ; attention à l'homonyme canadien sur
  PitchBook. **Mash VC** a un homonyme américain.

## Livrables

1. `VC_Database_Standardisee_v4.xlsx` — mêmes 5 onglets, mêmes colonnes dans le même ordre,
   `Dictionnaire` avec taux de remplissage recalculés, `Anomalies` complété.
2. `collecte_web_vc.py` — le script rejouable qui applique la table des valeurs collectées au
   classeur v3 (données de collecte en dur dans une structure lisible, comme `MAJ_WEB` dans
   `maj_web_vc.py`). La transformation doit être reproductible sans refaire les recherches.
3. `todo_collecte_web.csv` mis à jour, avec pour chaque véhicule le statut et les notes.
4. `RAPPORT_COLLECTE.md` — pour chaque véhicule : champs remplis, sources retenues (URL),
   champs restés introuvables et pourquoi. Plus une synthèse : taux de remplissage avant/après
   par colonne, nombre d'anomalies ajoutées, liste des divergences base/web à arbitrer
   manuellement.

## Contrôles avant de rendre

- Les 111 lignes et l'ordre des colonnes sont inchangés ; aucune valeur préexistante modifiée
  sans anomalie associée.
- `Fonds_ID` toujours unique ; chaque `Fonds_ID` de `Participations` existe dans `Fonds`.
- Aucun marqueur textuel d'absence de donnée dans le classeur.
- Multiples non formatés en pourcentage, IRR formaté en pourcentage, années en entiers.
- Le classeur se rouvre sans erreur avec `openpyxl` et les 5 onglets sont conformes.
- Affiche à la fin : lignes par onglet, cellules remplies par cette passe, taux de remplissage
  par colonne clé, divergences non arbitrées, véhicules restés sans aucune donnée.

Commence par lire les trois fichiers, affiche-moi ton plan de lots, puis enchaîne les lots sans
me redemander confirmation entre chacun.
