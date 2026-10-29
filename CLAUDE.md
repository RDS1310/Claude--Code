# Radar Insurtech — project guide for Claude

Suite of 5 web pages built for Rui (Groupama Vol'terre Investissement, GVI — Groupama's CVC) to track European insurtech VC/CVC funds, their portfolio companies, GVI's start-up pipeline, news, and insurance M&A.

## Working with this user
- Chat in **English**. Everything on the pages stays in **French**. The user writes in both.
- The user is **cost-sensitive** (Claude plan credits). Be lean: no repo exploration beyond what the task needs, no re-reading history, answer simple questions without tools, and don't burn cloud-session credits on simple writing tasks.
- Prefer one validated publish over several speculative ones.

## The 5 pages (Claude Artifacts, shared by link)

| Page | Artifact URL | Data | Generator | Template |
|---|---|---|---|---|
| Radar Insurtech VC | https://claude.ai/artifact/X6o5LvDnYe6Qo5fPYFkX7g | `VC_Database_Standardisee_vN.xlsx` (latest N auto-detected, currently v29) | `generer_site_vc.py` → `radar_insurtech_vc.html` | `_site_template.html` |
| Stratégie & Pipe GVI | https://claude.ai/artifact/58n6yRVTLgDuhavqhxihGn | `gvi_data/GVI_Status_Report_Startups.xlsx` (tab `Start-up`) | `generer_page_gvi.py` → `gvi_strategie_pipe.html` | `_gvi_template.html` |
| Veille Insurtech | https://claude.ai/artifact/5FXTpkCo6KvGu6YyeYAHiw | `veille_data/entries.json` | `generer_page_veille.py` → `veille_insurtech.html` | `_veille_template.html` |
| Étude M&A Assurance | https://claude.ai/artifact/BPSeTnKnWY2PkAGcH5s2ai | `ma_data/deals.json` | `generer_page_ma.py` → `etude_marche_ma.html` | `_ma_template.html` |
| Analyses GVI (private, internal) | https://claude.ai/artifact/Mor6DTNTSbekT4irHSLyJe | `analyses_data/analyses.json` | `generer_page_analyses.py` → `analyses_gvi.html` | `_analyses_template.html` |

Each page embeds its data as JSON (`window.__DATA__`) in a static HTML file. A cross-nav bar links the 5 pages.

**Cross-links (29/10, `liens_croises.py`, imported by all 5 generators)**: one index by normalized company/fund name (latest Radar Excel + Pipe xlsx + `analyses.json` + Veille entries + M&A deals). Radar start-up sheet ← news (Veille/M&A), "Suivie dans le pipe GVI", "Analyse GVI", other Radar funds on the cap table; Pipe detail ← "Fiche Radar", Radar funds, news, analysis (+ "actu" badge in the table); Veille/M&A → "Radar · <fonds/start-up>" chips (**never** a Pipe link on Veille); Analyses ← `RAW.liens` index matched client-side (`normNom` mirrors `norm`). Deep links: Radar `#s-<slug>` / `#f-<Fonds_ID>`, Pipe `#p-<slug>`, Veille `#v-<md5(url)[:8]>`, Analyses `#<id>`. Ambiguous names in `DENY`. Regenerate in order Radar → Pipe (reads radar html) → Veille → M&A → Analyses. Both routines now rebuild all 5 pages when something new was added (read pages with `Artifact read` + `path: "index.html"` so the 650 KB Radar isn't loaded into context).

### Radar Insurtech VC — tabs
Fonds (fund sheets: key figures, stages, contacts, participations, news, sources) · Start-ups · Comparateur · Tendances (stacked bars by year × 8 sectors + country breakdown) · Fiabilité (only the exceptions: participations with confidence below Élevé, with their reason, + ~42 open points from the Anomalies sheet — divergences, homonymy risks, approximate/unconfirmed values; types listed in `TYPES_POINTS_OUVERTS` in `generer_site_vc.py`). The full Anomalies log stays in the Excel only (replaced the old Anomalies tab on 02/10, user found it unhelpful). Rail filters include a Récence filter (news-based for funds, year-based for start-ups).

### Analyses GVI — quick screening sheets
Two ways to create a sheet (6 GVI criteria of the `grille`: sponsorship, equipe, marche, tech, traction, finance — note 1–3 or null; verdict `etudier` / `surveiller` / `ecarter`):
1. **On the page (main path, since 19/10)**: "Nouvelle analyse" form — name, site, context, file drop (PDF via pdf.js 3.11.174, .docx/.pptx via JSZip, Excel via SheetJS 0.18.5, all lazy-loaded from cdnjs; text-less PDF pages rendered as images, ≤8). Text extracted in the browser, sent to Claude through the `sample` capability (`modelTier: "complex"`, `sample.json`, images allowed) with a French prompt embedding the grille + data rules ("(doc.)" vs "(à vérifier)" tags, no web, no URLs). Result saved to the artifact **db**, collection `analyses` (`origine: "page"`, `identite` stored as `[{k,v}]` — no nested arrays), shown with a "documents seuls, non vérifiée" banner; deletable in-page. Documents themselves are not stored. Capabilities: `{sample:{images:true}, db:{}, user:{}}` — restate them on every full-set publish.
2. **Web verification = button "Collecter les données web" / "Actualiser avec le web"** at the bottom of a db sheet (user authorized page-launched runs on 19/10): sets `web: {etat: "demandée", demande}` then calls `fire_trigger` (mcp capability, connector "Claude Code Remote", tool `fire_trigger` only) on the fixed, unscheduled task **`trig_01SB5bBresd7Jx4QU5KGcJcV` "Analyses GVI · vérification web"** (fresh session; reads `analyses` via ArtifactData, 4–8 WebSearch, rewrites the sheet with "(web)" tags + `sources`, `origine: "web"`, `web.etat: "terminée"`; its prompt holds the GVI doctrine and the data rules). The page updates live via onSnapshot; a request older than 30 min can be relaunched. First real run on Seedext (19/10) worked (7 → 9 sources); the task writes made-up times in `web.debut/fin` (it has no clock) — harmless, only the date is shown. Can also be done in chat ("vérifie X") the same way. Static sheets in `analyses_data/analyses.json` (Klaimee) remain for chat-made analyses.
3. **PDF export** ("Exporter en PDF", `downloads` capability, jsPDF 2.5.1 from cdnjs): clean A4 for e-mail — origin tags "(doc.)", "(web)", "(à vérifier)" stripped (`sansMarqueurs`), no internal banners; verdict box, identity, 6-criteria grid with dots, sections, questions, sources; confidential footer. Text sanitized to WinAnsi (`pdfTexte`).
Capabilities (full set, restate on every publish that passes `capabilities`): `{sample:{images:true}, db:{}, user:{}, downloads:true, mcp:{servers:[{server:"Claude Code Remote", tools:["fire_trigger"]}]}}`.
Traffic-light colours (19/10): note 3 = vert « Fort », 2 = orange « Moyen », 1 or null = rouge « Faible » / « Pas de données » (`niveau()`, tokens good/ochre/bad) — on criterion cards, a 6-cell summary strip under the verdict, mini squares in the rail list, and in the PDF.
Grille + doctrine follow GVI's external deck (Sept. 2026): criteria titles/questions in `grille`, thesis/domains/terms (ticket 200 k€–2 M€, 5–20 %, follower or co-lead, proven Groupama partnership) in `these`, injected into the page prompt.
Local test harness for the upload flow (mocked `window.claude`, libs from npm): scratchpad only, not committed.

## Data model (Excel, 5 sheets)
`Fonds` (128 vehicles, 67 cols) · `Participations` (448) · `Tours_de_table` (58) · `Dictionnaire` (106 fields) · `Anomalies` (368, IDs `ANO-0001`…).

v29 (29/10, `maj_radar_v29.py` + `radar_refresh/cvc_2026_10.json`): CVC enrichment from official portfolio pages (fetched via `recuperer_pages.yml`; axavp.com only answers over **http**, ternel.vc is behind a Vercel bot check): **Allianz X** (+28, 40 listed / 14 exits, AuM 1,7 Md€, 1 Md€ envelope, Ledgebrook 2026 co-lead), **MAIF Avenir / Ternel** (+35, entry stages from maif-avenir.fr, tickets 1–8 M€, 42 cos / 15 exits per Ternel), new firm **AVP (ex-AXA Venture Partners)** — vehicles « AVP Venture & Growth » (+60 participations) and « AVP Late Growth », classed VC indépendant since the 2024 MBO, AuM 1 360 M€ (ADEME FY2024; >2 Md€ claimed), contacts F. Robinet / B. Fosseprez. No investment years published → these rows don't feed Tendances (ANO-0351…0368). v28 (29/10, `maj_radar_v28.py`): exact-equivalent labels merged (Series→Série, Actif→Active, US/UK→États-Unis/Royaume-Uni; ANO-0348…0350).

v27 (20/10/2026, `maj_radar_v27.py` + `radar_refresh/portefeuille_sagard_2026_10.json`): full Portage (Sagard) portfolio from sagard.com — 62 rows added to Portag3 Venture I/II/III, Portage Ventures IV and the new vehicle **Portage Capital Solutions I** (PCS I). Sagard publishes no date/amount/stage (left empty); "previous" = Exit O. Non-insurance companies are kept but labelled `FinTech (hors assurance)` / `Hors périmètre assurance`; `secteur_bucket` excludes any label containing "hors" from Tendances (ANO-0342…0347).

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
  - **Push blockage worked around (19/10)**: routine sessions can't push (403 — the repo isn't among their sources and `add_repo` isn't available there; the trigger API can't add sources). Both routine prompts were rewritten: clone read-only → `Artifact read` of the live page → `python3 fusion_depuis_live.py veille|ma <live.html>` (newest `derniereMaj` wins; veille also unions entries by url) → update → regenerate → publish, **no commit/push**. The live page is the reference; **in any session, before regenerating Veille or M&A, read the live artifact and run `fusion_depuis_live.py`, then commit** — that's how the repo catches up. (Adding `RDS1310/Claude--Code` as a push source in the routines' settings on claude.ai would remove the need, but isn't required anymore.)
- **Fetching blocked websites**: `.github/workflows/recuperer_pages.yml` — put URLs (one per line) in `outils/pages_a_recuperer.txt` and push; GitHub Actions downloads them into `outils/pages/` (+ `_journal.txt` with HTTP codes) and commits (~1 min). Used for sagard.com (WordPress REST: `/wp-json/wp/v2/projects?per_page=100&page=N`, fund tags in `class_list` `cat_source-pv1…`). Wait with a background `git fetch` loop, not sleep.
- **n8n** (user's trial, being retired): workflow `fetch_news` and an agent with a knowledge base — superseded by the GitHub Actions collection.

## Open items
- Verify TheFamily acquisition (17/02/2026, CB Insights only — ANO-0327). Beazley/Zurich marked « Finalisée le 01/10/2026 » in `ma_data` on 29/10.
- Radar template has `VALO_DATES` (dated valuations, e.g. Alan 06/2026) — it was published from an uncommitted session on 01/10 and merged back on 02/10.
- Review the 4 start-ups tagged "adjacent" but out of insurance scope (ANO-0315: MuchBetter.ai, Gretel, hypt., Value Factory).
- Collection: watch feed health in the raw files (`flux[].ok`); feeding fund news into the Excel (vN+1) not automated yet.
- Missing LinkedIn: Dietrich Aumann (Helsana, 2nd contact). Contact emails: v25 (15/10, `maj_radar_v25.py` + `radar_refresh/emails_apollo_2026_10.json`) — user revealed them manually in Apollo: 27 verified, 2 guessed (ANO-0331/0332, in Fiabilité), 56 unavailable (red on the page). Apollo Free plan blocks API enrichment. v26 (`maj_radar_v26.py` + `radar_refresh/contacts_societe_2026_10.json`): `Contact_Societe_*` (generic address/form published on the fund's own site, 25 funds, 20 with email — found via WebSearch restricted to the fund domain; WebFetch can't reach fund sites), 3 partner emails published by the fund (Menasé, Molitor; Thorne = old doc, grey), `Contact_0x_A_Remplacer` for 4 outdated contacts (Cara Morton, Munich Re ×2, Christian Huber). Hampus (TrueSight) has a 1880.vc email — to check (ANO-0341). Rendering: `STATUTS_EMAIL` in `generer_site_vc.py`, `emailChip` in the template.
- User preference (29/10): enrich **CVCs** rather than independent VCs (Sagard/Portage enrichment declined). Duplicate Analyses artifact 4jYs3DtMATnczgcmemsh8P: user declined deletion — leave it.
- Deferred: case-based session on how GVI assesses a start-up (to improve "Pertinence pour GVI"); login gate for Radar (undecided).

## Git
Repo `RDS1310/Claude--Code`, working branch so far: `claude/new-session-esly1p`.
