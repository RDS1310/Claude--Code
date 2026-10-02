# Radar Insurtech — project guide for Claude

Suite of 4 web pages built for Rui (Groupama Vol'terre Investissement, GVI — Groupama's CVC) to track European insurtech VC/CVC funds, their portfolio companies, GVI's start-up pipeline, news, and insurance M&A.

## Working with this user
- Chat in **English**. Everything on the pages stays in **French**. The user writes in both.
- The user is **cost-sensitive** (Claude plan credits). Be lean: no repo exploration beyond what the task needs, no re-reading history, answer simple questions without tools, and don't burn cloud-session credits on simple writing tasks.
- Prefer one validated publish over several speculative ones.

## The 4 pages (Claude Artifacts, shared by link)

| Page | Artifact URL | Data | Generator | Template |
|---|---|---|---|---|
| Radar Insurtech VC | https://claude.ai/artifact/X6o5LvDnYe6Qo5fPYFkX7g | `VC_Database_Standardisee_vN.xlsx` (latest N auto-detected, currently v23) | `generer_site_vc.py` → `radar_insurtech_vc.html` | `_site_template.html` |
| Stratégie & Pipe GVI | https://claude.ai/artifact/58n6yRVTLgDuhavqhxihGn | `gvi_data/GVI_Status_Report_Startups.xlsx` (tab `Start-up`) | `generer_page_gvi.py` → `gvi_strategie_pipe.html` | `_gvi_template.html` |
| Veille Insurtech | https://claude.ai/artifact/5FXTpkCo6KvGu6YyeYAHiw | `veille_data/entries.json` | `generer_page_veille.py` → `veille_insurtech.html` | `_veille_template.html` |
| Étude M&A Assurance | https://claude.ai/artifact/BPSeTnKnWY2PkAGcH5s2ai | `ma_data/deals.json` | `generer_page_ma.py` → `etude_marche_ma.html` | `_ma_template.html` |

Each page embeds its data as JSON (`window.__DATA__`) in a static HTML file. A cross-nav bar links the 4 pages.

### Radar Insurtech VC — tabs
Fonds (fund sheets: key figures, stages, contacts, participations, news, sources) · Start-ups · Comparateur · Tendances (stacked bars by year × 8 sectors + country breakdown) · Anomalies (searchable data-quality log). Rail filters include a Récence filter (news-based for funds, year-based for start-ups).

## Data model (Excel, 5 sheets)
`Fonds` (119 vehicles, 59 cols: 100 VC indépendants, 19 European CVC) · `Participations` (261) · `Tours_de_table` (58) · `Dictionnaire` (106 fields) · `Anomalies` (315, IDs `ANO-0001`…).

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
- **Claude routines** (fresh session per run):
  - weekly veille `trig_01C8tsihdPN4XackPZwGAsq7`, Mondays 06:00 UTC — the 28/09 run "succeeded" but pushed nothing (no alert);
  - monthly M&A `trig_013VoXmnanak2NpoDXPuPUU1`, 1st of the month 06:00 UTC.
- **n8n** (user's own instance, trial):
  - workflow `fetch_news` (published): 26 RSS feeds incl. 4 Google News searches, keyword filter, `days` = 2, snippets capped at 300 chars;
  - an n8n agent with a knowledge base (guide + CSV exports of funds, participations, active pipe) — the user had trouble running it.
- **Target setup**:
  - n8n `daily_news_to_sheet` (Schedule 08:00 Paris → `fetch_news` → Google Sheet `Veille Insurtech`, tab `Brut`, dedupe on `link`);
  - then a **weekly** Claude/Cowork task that reads the sheet, validates, and updates `entries.json` / the Excel and republishes;
  - if that works, the Claude weekly routine becomes redundant.

## Open items
- Review the 4 start-ups tagged "adjacent" but out of insurance scope (ANO-0315: MuchBetter.ai, Gretel, hypt., Value Factory).
- n8n: error-alert workflow, `add_entries` / sheet output, broken feeds (403: FinSMEs, Les Echos Start, La Tribune de l'Assurance…).
- Missing LinkedIn: Dietrich Aumann (Helsana, 2nd contact). Contact emails empty (planned Apollo enrichment).
- Deferred: case-based session on how GVI assesses a start-up (to improve "Pertinence pour GVI"); login gate for Radar (undecided).

## Git
Repo `RDS1310/Claude--Code`, working branch so far: `claude/new-session-esly1p`.
