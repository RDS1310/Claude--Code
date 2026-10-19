# Radar Insurtech — project guide for Claude

Suite of 5 web pages built for Rui (Groupama Vol'terre Investissement, GVI — Groupama's CVC) to track European insurtech VC/CVC funds, their portfolio companies, GVI's start-up pipeline, news, and insurance M&A.

## Working with this user
- Chat in **English**. Everything on the pages stays in **French**. The user writes in both.
- The user is **cost-sensitive** (Claude plan credits). Be lean: no repo exploration beyond what the task needs, no re-reading history, answer simple questions without tools, and don't burn cloud-session credits on simple writing tasks.
- Prefer one validated publish over several speculative ones.

## The 5 pages (Claude Artifacts, shared by link)

| Page | Artifact URL | Data | Generator | Template |
|---|---|---|---|---|
| Radar Insurtech VC | https://claude.ai/artifact/X6o5LvDnYe6Qo5fPYFkX7g | `VC_Database_Standardisee_vN.xlsx` (latest N auto-detected, currently v26) | `generer_site_vc.py` → `radar_insurtech_vc.html` | `_site_template.html` |
| Stratégie & Pipe GVI | https://claude.ai/artifact/58n6yRVTLgDuhavqhxihGn | `gvi_data/GVI_Status_Report_Startups.xlsx` (tab `Start-up`) | `generer_page_gvi.py` → `gvi_strategie_pipe.html` | `_gvi_template.html` |
| Veille Insurtech | https://claude.ai/artifact/5FXTpkCo6KvGu6YyeYAHiw | `veille_data/entries.json` | `generer_page_veille.py` → `veille_insurtech.html` | `_veille_template.html` |
| Étude M&A Assurance | https://claude.ai/artifact/BPSeTnKnWY2PkAGcH5s2ai | `ma_data/deals.json` | `generer_page_ma.py` → `etude_marche_ma.html` | `_ma_template.html` |
| Analyses GVI (private, internal) | https://claude.ai/artifact/Mor6DTNTSbekT4irHSLyJe | `analyses_data/analyses.json` | `generer_page_analyses.py` → `analyses_gvi.html` | `_analyses_template.html` |

Each page embeds its data as JSON (`window.__DATA__`) in a static HTML file. A cross-nav bar links the 5 pages.

### Radar Insurtech VC — tabs
Fonds (fund sheets: key figures, stages, contacts, participations, news, sources) · Start-ups · Comparateur · Tendances (stacked bars by year × 8 sectors + country breakdown) · Fiabilité (only the exceptions: participations with confidence below Élevé, with their reason, + ~42 open points from the Anomalies sheet — divergences, homonymy risks, approximate/unconfirmed values; types listed in `TYPES_POINTS_OUVERTS` in `generer_site_vc.py`). The full Anomalies log stays in the Excel only (replaced the old Anomalies tab on 02/10, user found it unhelpful). Rail filters include a Récence filter (news-based for funds, year-based for start-ups).

### Analyses GVI — quick screening sheets
On request, the user sends a start-up name/site/docs in chat; Claude researches (WebSearch, ≤ a few searches), adds one entry to `analyses_data/analyses.json`, regenerates, publishes. Entry: `verdict` (`etudier` / `surveiller` / `ecarter`) + `verdictTexte`, `confiance`, `identite`, `notes` on the 6 GVI criteria of the `grille` (sponsorship, equipe, marche, tech, traction, finance — note 1–3 or null), `sections` (secteur, financements, management, produit…), `questions` for the team, `documentsFournis`, `sources`. Same data rules: no fabrication, "non trouvé". Klaimee is the first example (À surveiller).

## Data model (Excel, 5 sheets)
`Fonds` (125 vehicles, 67 cols) · `Participations` (263) · `Tours_de_table` (58) · `Dictionnaire` (106 fields) · `Anomalies` (341, IDs `ANO-0001`…).

Last full web refresh: 02/10/2026 → v24 (`maj_radar_v24.py` applies the verified findings stored in `radar_refresh/maj_2026_10.json`: 6 new vehicles, Munich Re Ventures closed to new investments, 2 participations, Alan valuation, 5 news items). Next refresh: same pattern (findings JSON + saved script → vN+1).

## How to update a page
1. Change the data. A new Excel version = **a saved script** producing `vN+1` (never an inline heredoc — v21/v22 were made that way and aren't reproducible).
2. Run the generator (`python3 generer_*.py`).
3. Extract the logic `<script>` and run `node --check` on it.
4. Publish to the **existing** artifact URL (if the publish guard refuses, read the live version first, merge, republish).
5. Commit and push.

## Data rules (non-negotiable)
- **Never fabricate.** Every data point has a source; unknown = "non trouvé". Confidence levels: Élevé / Moyen / Faible.
- **Log decisions** (ambiguity, conflict, exclusion, reclassification) as a new row in the `Anomalies` sheet instead of silently changing or dropping data.
- **Scope: Europe.** A CVC is in scope only if its parent company is headquartered in Europe (13 non-European CVCs were purged, ANO-0314).
- Investment dates are mostly **year-only** — never invent a month.
- Currency → M€: `$ × 0.92`, `£ × 1.17`.
- Sector taxonomy (Tendances chart): InsurTech (cœur), FinTech liée assurance, Data/IA & distribution, Télématique/mobilité & prévention IoT, Santé & prévoyance, Cyber, Emploi & avantages sociaux, Climat & risques catastrophes. Sub-classification of the "adjacent" bucket lives in `ADJACENT_SOUS_SECTEUR` in `generer_site_vc.py`.
- Pipe GVI page: no contact name/phone/email/deck columns. Veille page: factual press news only, never pipe statuses.

## Design
Site tokens: petrol `#0C5C5E`, ochre `#A8681F`, Newsreader (display) + IBM Plex Sans/Mono. Light + dark themes. Chart colors use the dataviz skill's validated categorical palette (the brand colors fail its chroma check).

## Automations
- **Daily collection (free, no AI)**: GitHub Actions `.github/workflows/collecte_news.yml` runs `collecte_news_quotidienne.py` at 05:47 UTC (GitHub actually starts it ~5 h late) (also manual via workflow_dispatch). 50 sources: direct RSS feeds + Google News `site:` searches for French press without usable feeds (IT/business sites filtered on insurance keywords), 2-day window, retry on Google News 503/429 (all 39 GN feeds 503'd on 03/10) + 2 s spacing, insurance keyword filter + tracked-name matching (funds + short aliases, portfolio start-ups, active GVI pipe), 30-day dedupe on link, per-feed health. Output: `veille_data/raw/AAAA-MM-JJ.json`, committed by github-actions[bot] to `claude/new-session-esly1p` (the default branch — scheduled workflows only run there). Replaces n8n `fetch_news`.
- **Claude routines** (fresh session per run):
  - weekly veille `trig_01C8tsihdPN4XackPZwGAsq7`, Mondays 06:00 UTC — since 02/10 it reads the raw files since `derniereMaj` (≤3 WebSearch), verifies, updates `entries.json`, republishes, and reports feed health + M&A leads; falls back to roundup WebSearch if no raw files. (The 28/09 run found nothing reliable and correctly pushed nothing, ~$1.90 / 38 searches.) 05/10 run: added 2 entries (~$0.41) and published, but its push got 403 (add_repo without push access) — entries merged back by hand on 05/10; prompt now requires `access: "push"` and push-before-publish.
  - monthly M&A `trig_013VoXmnanak2NpoDXPuPUU1`, 1st of the month 06:00 UTC.
  - **Known issue (unresolved)**: routine sessions can't push — 403, the repo isn't among the routine's sources and `add_repo` isn't available there (19/10 run). The 01/10 M&A and 05/10–19/10 Veille runs published live pages without pushing; their data was merged back into the repo by hand on 19/10. Fix: add `RDS1310/Claude--Code` as a source of both routines (with push access) from the routine settings. Until then, before regenerating Veille/M&A, read the live artifact and merge its data.
- **n8n** (user's trial, being retired): workflow `fetch_news` and an agent with a knowledge base — superseded by the GitHub Actions collection.

## Open items
- Verify TheFamily acquisition (17/02/2026, CB Insights only — ANO-0327). Beazley/Zurich deal completed 01/10/2026: update `ma_data` status at next M&A run.
- Radar template has `VALO_DATES` (dated valuations, e.g. Alan 06/2026) — it was published from an uncommitted session on 01/10 and merged back on 02/10.
- Review the 4 start-ups tagged "adjacent" but out of insurance scope (ANO-0315: MuchBetter.ai, Gretel, hypt., Value Factory).
- Collection: watch feed health in the raw files (`flux[].ok`); feeding fund news into the Excel (vN+1) not automated yet.
- Missing LinkedIn: Dietrich Aumann (Helsana, 2nd contact). Contact emails: v25 (15/10, `maj_radar_v25.py` + `radar_refresh/emails_apollo_2026_10.json`) — user revealed them manually in Apollo: 27 verified, 2 guessed (ANO-0331/0332, in Fiabilité), 56 unavailable (red on the page). Apollo Free plan blocks API enrichment. v26 (`maj_radar_v26.py` + `radar_refresh/contacts_societe_2026_10.json`): `Contact_Societe_*` (generic address/form published on the fund's own site, 25 funds, 20 with email — found via WebSearch restricted to the fund domain; WebFetch can't reach fund sites), 3 partner emails published by the fund (Menasé, Molitor; Thorne = old doc, grey), `Contact_0x_A_Remplacer` for 4 outdated contacts (Cara Morton, Munich Re ×2, Christian Huber). Hampus (TrueSight) has a 1880.vc email — to check (ANO-0341). Rendering: `STATUTS_EMAIL` in `generer_site_vc.py`, `emailChip` in the template.
- Deferred: case-based session on how GVI assesses a start-up (to improve "Pertinence pour GVI"); login gate for Radar (undecided).

## Git
Repo `RDS1310/Claude--Code`, working branch so far: `claude/new-session-esly1p`.
