#!/usr/bin/env python3
"""
scraper_portefeuilles.py — Relevé des pages portefeuille des sites de fonds (GitHub Actions)
============================================================================================

Les sites des fonds sont inaccessibles depuis les sessions Claude ; ce script tourne dans GitHub Actions
(.github/workflows/scraper_portefeuilles.yml) avec un vrai navigateur (Playwright/Chromium), car beaucoup
de pages portefeuille sont rendues en JavaScript.

Pour chaque société de gestion du dernier classeur Radar (colonne « Site web ») :
1. page d'accueil rendue → liens du même domaine dont le chemin évoque un portefeuille ;
2. chemins usuels (/portfolio, /portefeuille, /companies…) et sitemap.xml ;
3. jusqu'à MAX_PAGES pages candidates rendues, défilées jusqu'en bas (chargements différés), dont on garde :
   liens (texte + adresse), textes alternatifs des logos, intitulés, et le texte visible (tronqué).

Aucune interprétation ici : le rapprochement avec les participations est fait ensuite en session, et
chaque ajout au Radar passe par un script maj_radar_vN avec la page comme source.
Sortie : outils/portefeuilles/releve.json (+ _journal.txt).
"""
from __future__ import annotations

import json
import re
import sys
import urllib.request
from datetime import date
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
from playwright.sync_api import sync_playwright

BASE_DIR = Path(__file__).resolve().parent.parent
SORTIE = BASE_DIR / "outils" / "portefeuilles"
MOTIF = re.compile(r"portfolio|portefeuille|compan|investment|participation|startups?\b|backed|entreprises|our-?team-?s?|"
                   r"invest(i|e)?ssements?|beteiligung|portfolio-?companies", re.I)
CHEMINS = ["/portfolio", "/portfolio/", "/en/portfolio", "/portefeuille", "/fr/portefeuille", "/companies",
           "/our-companies", "/investments", "/participations"]
# Pages portefeuille connues quand la page d'accueil ne les expose pas (ou que le site est celui du groupe)
CONNUES = {
    "Allianz X": ["https://www.allianzx.com/ourcompanies"],
    "MAIF Avenir": ["https://maif-avenir.fr/portefeuille/"],
    "AVP (ex-AXA Venture Partners)": ["http://www.axavp.com/companies/"],
    "Munich Re Ventures": ["https://www.munichreventures.com/portfolio", "https://www.munichre.com/mrv/en/portfolio.html"],
    "SCOR Ventures": ["https://www.scor.com/en/scor-ventures"],
    "Aviva Ventures": ["https://www.aviva.com/about-us/aviva-ventures/"],
    "Generali Ventures": ["https://www.generali.com/our-responsibilities/innovation"],
    "Zurich Insurance Group": ["https://www.zurich.com/en/about-us/zurich-ventures"],
    "Helsana HealthInvest": ["https://www.helsana.ch/en/helsana-group/about-us/healthinvest.html"],
    "Howden Ventures": ["https://www.howdengroupholdings.com/howden-ventures"],
    "Portage Venture": ["https://www.sagard.com/portfolio/"],
    "Atlantic Vantage Point": ["https://www.avpcap.com/portfolio"],
}
MAX_PAGES = 6
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"


def dernier_classeur() -> Path:
    return sorted(BASE_DIR.glob("VC_Database_Standardisee_v*.xlsx"),
                  key=lambda p: int(re.search(r"_v(\d+)\.xlsx$", p.name).group(1)))[-1]


def sitemap(site: str) -> list[str]:
    urls = []
    for chemin in ("/sitemap.xml", "/sitemap_index.xml", "/wp-sitemap.xml"):
        try:
            with urllib.request.urlopen(urllib.request.Request(urljoin(site, chemin), headers={"User-Agent": UA}), timeout=20) as r:
                xml = r.read(3_000_000).decode("utf-8", "ignore")
        except Exception:
            continue
        locs = re.findall(r"<loc>\s*([^<\s]+)\s*</loc>", xml)
        # Index de sitemaps : on suit les sous-sitemaps qui évoquent un portefeuille
        for sous in [u for u in locs if u.endswith(".xml") and MOTIF.search(u)][:4]:
            try:
                with urllib.request.urlopen(urllib.request.Request(sous, headers={"User-Agent": UA}), timeout=20) as r:
                    locs += re.findall(r"<loc>\s*([^<\s]+)\s*</loc>", r.read(3_000_000).decode("utf-8", "ignore"))
            except Exception:
                pass
        urls += [u for u in locs if not u.endswith(".xml")]
        if urls:
            break
    return sorted(set(u for u in urls if MOTIF.search(urlparse(u).path)))[:400]


JS_RELEVE = """() => {
  const vis = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
  const liens = [...document.querySelectorAll('a[href]')].map(a => ({t: (a.innerText || a.getAttribute('aria-label') || a.title || '').trim().slice(0, 120), h: a.href})).filter(l => l.h.startsWith('http'));
  const alts = [...document.querySelectorAll('img[alt], svg[aria-label], [role=img][aria-label]')].map(i => (i.getAttribute('alt') || i.getAttribute('aria-label') || '').trim()).filter(Boolean);
  const titres = [...document.querySelectorAll('h1,h2,h3,h4,h5,h6')].filter(vis).map(h => h.innerText.trim()).filter(Boolean);
  return {titre: document.title, liens, alts: [...new Set(alts)], titres: [...new Set(titres)], texte: (document.body ? document.body.innerText : '').slice(0, 40000)};
}"""


def rendre(page, url: str) -> dict:
    try:
        rep = page.goto(url, wait_until="domcontentloaded", timeout=45000)
        page.wait_for_timeout(2500)
        for _ in range(12):  # défilement : listes chargées au fil de l'eau
            page.mouse.wheel(0, 4000)
            page.wait_for_timeout(400)
        for libelle in ("Load more", "Show more", "Voir plus", "Afficher plus", "View all", "Tout voir", "See all"):
            try:
                b = page.get_by_text(libelle, exact=False).first
                for _ in range(8):
                    if b.is_visible(timeout=500):
                        b.click(timeout=2000)
                        page.wait_for_timeout(1200)
            except Exception:
                pass
        r = page.evaluate(JS_RELEVE)
        r.update(url=url, final=page.url, statut=rep.status if rep else None)
        return r
    except Exception as e:
        return {"url": url, "erreur": str(e)[:200]}


def main() -> int:
    fonds = pd.read_excel(dernier_classeur(), sheet_name="Fonds")
    sites = fonds.groupby("Nom du fonds")["Site web"].first().dropna().to_dict()
    seulement = set(sys.argv[1:])
    SORTIE.mkdir(parents=True, exist_ok=True)
    releve, journal = {}, []
    with sync_playwright() as pw:
        nav = pw.chromium.launch()
        ctx = nav.new_context(user_agent=UA, viewport={"width": 1400, "height": 2000}, locale="fr-FR")
        page = ctx.new_page()
        for firme, site in sorted(sites.items()):
            if seulement and firme not in seulement:
                continue
            accueil = rendre(page, site)
            dom = urlparse(accueil.get("final") or site).netloc.replace("www.", "")
            candidats = list(CONNUES.get(firme, []))
            candidats += [l["h"].split("#")[0] for l in accueil.get("liens", [])
                          if dom in urlparse(l["h"]).netloc and MOTIF.search(urlparse(l["h"]).path + " " + l["t"])
                          and urlparse(l["h"]).path.count("/") <= 3]
            candidats += [urljoin(accueil.get("final") or site, c) for c in CHEMINS]
            vus, pages = {(accueil.get("final") or site).rstrip("/")}, []
            for u in candidats:
                u = u.rstrip("/")
                if u in vus or len(pages) >= MAX_PAGES:
                    continue
                vus.add(u)
                r = rendre(page, u)
                if r.get("erreur") or (r.get("statut") or 200) >= 400 or r.get("final", "").rstrip("/") in vus - {u}:
                    journal.append(f"{firme}\t{u}\t{r.get('statut') or r.get('erreur', '')[:60]}")
                    continue
                vus.add(r.get("final", "").rstrip("/"))
                pages.append(r)
                journal.append(f"{firme}\t{u}\tOK {len(r['texte'])} car. {len(r['alts'])} logos")
            releve[firme] = {"site": site, "accueil": accueil,  # sites d'une seule page : le portefeuille y est
                             "sitemap": sitemap(accueil.get("final") or site), "pages": pages}
            print(firme, len(pages), "pages", flush=True)
        nav.close()
    (SORTIE / "releve.json").write_text(json.dumps({"date": date.today().isoformat(), "fonds": releve},
                                                   ensure_ascii=False), encoding="utf-8")
    (SORTIE / "_journal.txt").write_text("\n".join(journal), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
