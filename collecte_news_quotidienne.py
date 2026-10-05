#!/usr/bin/env python3
"""Collecte quotidienne des actualités (remplace le workflow n8n `fetch_news`).

Lancé chaque matin par GitHub Actions (.github/workflows/collecte_news.yml),
sans IA : lit les flux RSS de FLUX, garde les articles des JOURS derniers jours
qui parlent d'assurance / insurtech OU citent un nom suivi (fonds du Radar,
start-ups en portefeuille, pipe GVI actif), et écrit
veille_data/raw/AAAA-MM-JJ.json.

La routine Claude hebdomadaire lit ces fichiers, vérifie et sélectionne, puis
met à jour veille_data/entries.json. Ce script ne publie rien.

Usage : python3 collecte_news_quotidienne.py [--jours 2]
"""
import argparse
import datetime as dt
import email.utils
import html
import json
import re
import unicodedata
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path

RACINE = Path(__file__).resolve().parent
SORTIE = RACINE / "veille_data" / "raw"
RADAR = RACINE / "radar_insurtech_vc.html"
GVI = RACINE / "gvi_data" / "GVI_Status_Report_Startups.xlsx"
JOURS_DEDOUBLONNAGE = 30
SNIPPET_MAX = 300

GN = "https://news.google.com/rss/search?q={q}&hl={hl}&gl={gl}&ceid={gl}:{hl}"
ASSURANCE_FR = "(assurance OR assureur OR insurtech OR assurtech OR mutuelle OR courtier OR réassurance)"


def gn_site(domaine: str, filtre: str = "") -> str:
    """Recherche Google News limitée à un site (7 jours), en français.
    `filtre` restreint les sites généralistes aux sujets assurance."""
    q = urllib.parse.quote_plus(f"site:{domaine} {filtre} when:7d".replace("  ", " "))
    return GN.format(q=q, hl="fr", gl="FR")


FLUX = {
    # Écosystème start-up / VC Europe
    "EU-Startups": "https://www.eu-startups.com/feed/",
    "Tech.eu": "https://tech.eu/feed/",
    "Sifted": "https://sifted.eu/feed",
    "Maddyness": "https://www.maddyness.com/feed/",
    "FinSMEs": GN.format(q="site:finsmes.com+insurance+when:7d", hl="en-GB", gl="GB"),
    "Finextra": "https://www.finextra.com/rss/headlines.aspx",
    "FinTech Global": "https://fintech.global/feed/",
    # Assurance / insurtech
    "Coverager": "https://coverager.com/feed/",
    "Insurtech Insights": GN.format(q="site:insurtechinsights.com+when:7d", hl="en-GB", gl="GB"),
    "Reinsurance News": "https://www.reinsurancene.ws/feed/",
    "Artemis": "https://www.artemis.bm/feed/",
    "Insurance Journal": "https://www.insurancejournal.com/rss/news/",
    "Claims Journal": "https://www.claimsjournal.com/rss/",
    "Insurance Times": GN.format(q="site:insurancetimes.co.uk+when:7d", hl="en-GB", gl="GB"),
    "News Assurances Pro": GN.format(q="site:newsassurancespro.com+when:7d", hl="fr", gl="FR"),
    "L'Argus de l'assurance": GN.format(q="site:argusdelassurance.com+when:7d", hl="fr", gl="FR"),
    # (FinSMEs, Insurtech Insights, Insurance Times, News Assurances Pro, L'Argus : flux directs
    #  en 403/404 ou XML invalide,
    #  lus via Google News « site: ».)
    # Presse française start-up / VC / deals (sans filtre : les noms suivis sont détectés ensuite)
    "FrenchWeb": gn_site("frenchweb.fr"),
    "CFNEWS": gn_site("cfnews.net"),
    "Décideurs Corporate Finance": gn_site("magazine-decideurs.com"),
    "Les Echos Start": gn_site("lesechos.fr", "(start-up OR levée OR insurtech OR assurtech)"),
    "Journal du Net": gn_site("journaldunet.com", "(levée OR " + ASSURANCE_FR[1:]),
    "Siècle Digital": gn_site("siecledigital.fr", "(levée OR " + ASSURANCE_FR[1:]),
    "Dynamique Mag": gn_site("dynamique-mag.com", "(levée OR " + ASSURANCE_FR[1:]),
    "Entreprendre": gn_site("entreprendre.fr", "(levée OR " + ASSURANCE_FR[1:]),
    # Presse économique, financière et assurance
    "BFM Business": gn_site("bfmtv.com", ASSURANCE_FR),
    "L'Agefi": gn_site("agefi.fr", ASSURANCE_FR),
    "Revue Banque": gn_site("revue-banque.fr", ASSURANCE_FR),
    "La Tribune de l'Assurance": gn_site("tribune-assurance.optionfinance.fr"),
    # Presse tech / IT / data (filtre assurance : volume élevé hors sujet)
    "L'Usine Digitale": gn_site("usine-digitale.fr", ASSURANCE_FR),
    "LeMagIT": gn_site("lemagit.fr", ASSURANCE_FR),
    "Le Monde Informatique": gn_site("lemondeinformatique.fr", ASSURANCE_FR),
    "CIO Online": gn_site("cio-online.com", ASSURANCE_FR),
    "ZDNet.fr": gn_site("zdnet.fr", ASSURANCE_FR),
    "Silicon.fr": gn_site("silicon.fr", ASSURANCE_FR),
    "ActuIA": gn_site("actuia.com", ASSURANCE_FR),
    "IT for Business": gn_site("itforbusiness.fr", ASSURANCE_FR),
    "Alliancy": gn_site("alliancy.fr", ASSURANCE_FR),
    "Solutions Numériques": gn_site("solutions-numeriques.com", ASSURANCE_FR),
    "InformatiqueNews": gn_site("informatiquenews.fr", ASSURANCE_FR),
    "L'Informaticien": gn_site("linformaticien.com", ASSURANCE_FR),
    "Distributique": gn_site("distributique.com", ASSURANCE_FR),
    "ChannelNews": gn_site("channelnews.fr", ASSURANCE_FR),
    "Decideo": gn_site("decideo.fr", ASSURANCE_FR),
    "IT Social": gn_site("itsocial.fr", ASSURANCE_FR),
    "LeBigData.fr": gn_site("lebigdata.fr", ASSURANCE_FR),
    # Recherches Google News (7 derniers jours ; le filtre JOURS s'applique ensuite)
    "GN insurtech funding": GN.format(q="insurtech+funding+when:7d", hl="en-GB", gl="GB"),
    "GN insurtech raises": GN.format(q="insurtech+raises+Europe+when:7d", hl="en-GB", gl="GB"),
    "GN assurtech levée": GN.format(q="assurtech+lev%C3%A9e+when:7d", hl="fr", gl="FR"),
    "GN CVC assureur": GN.format(q="%22corporate+venture%22+insurer+when:7d", hl="en-GB", gl="GB"),
    "GN M&A assurance": GN.format(q="acquisition+courtier+OR+assureur+when:7d", hl="fr", gl="FR"),
}

MOTS_CLES = [
    "insurtech", "assurtech", "insurance", "insurer", "reinsur", "réassur", "assurance",
    "assureur", "mutuelle", "courtier", "broker", "mga", "underwrit", "sinistre", "claims",
    "prévoyance", "embedded insurance", "parametric", "paramétrique", "bancassur",
]


def normaliser(txt: str) -> str:
    txt = unicodedata.normalize("NFKD", txt or "").encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", txt.lower()).strip()


# Suffixes retirés pour obtenir le nom court d'un fonds (« BlackFin Capital Partners » → « BlackFin »),
# sauf si le nom court est un mot courant (STOP).
SUFFIXES = re.compile(r"(\s*\(.*?\))|\s+(venture capital|capital partners|asset management|partners|"
                      r"capital|ventures?|vc)$", re.I)
STOP = {"index", "start", "step", "ring", "committed", "elevation", "insurtech", "concentric", "cadence growth"}


def noms_suivis() -> dict:
    """{nom normalisé: (nom affiché, catégorie)} — fonds, participations, pipe GVI actif."""
    noms = {}

    def ajouter(nom, cat):
        n = normaliser(str(nom or ""))
        if len(n) >= 4 and n not in noms:  # trop court = faux positifs
            noms[n] = (str(nom).strip(), cat)

    h = RADAR.read_text(encoding="utf-8")
    m = re.search(r"window\.__DATA__\s*=\s*(\{.*?\});\s*</script>", h, re.S)
    data = json.loads(m.group(1))
    for f in data["fonds"]:
        ajouter(f["nom"], "fonds")
        court = f["nom"]
        while (c := SUFFIXES.sub("", court).strip()) != court:
            court = c
        if normaliser(court) not in STOP:
            ajouter(court, "fonds")
    for s in data["startups"]:
        ajouter(s["nom"], "participation")

    try:
        import openpyxl
        ws = openpyxl.load_workbook(GVI, read_only=True, data_only=True)["Start-up"]
        lignes = ws.iter_rows(values_only=True)
        entetes = [str(c or "") for c in next(lignes)]
        i_nom, i_int, i_vg = (entetes.index(c) for c in ("Start-up", "Intérêt GVI", "Volt'terre \nGVI"))
        for r in lignes:
            interet = str(r[i_int] or "").strip().removesuffix(".0")
            if interet in {"1", "2", "3"} or str(r[i_vg] or "").strip().lower() == "oui":
                ajouter(r[i_nom], "pipe GVI")
    except Exception as e:  # le pipe est un bonus : la collecte continue sans
        print(f"Pipe GVI non chargé : {e}")
    return noms


def texte(el, *tags):
    for t in tags:
        x = el.find(t)
        if x is not None and (x.text or "").strip():
            return x.text.strip()
    return ""


def parse_date(s: str):
    if not s:
        return None
    try:
        d = email.utils.parsedate_to_datetime(s)
    except (TypeError, ValueError):
        try:
            d = dt.datetime.fromisoformat(s.replace("Z", "+00:00"))
        except ValueError:
            return None
    return d if d.tzinfo else d.replace(tzinfo=dt.timezone.utc)


def lire_flux(url: str):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (veille-insurtech RSS reader)"})
    for essai in range(3):  # Google News renvoie parfois 503/429 en rafale (ex. 03/10/2026)
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                brut = r.read()
            break
        except urllib.error.HTTPError as e:
            if e.code not in (429, 503) or essai == 2:
                raise
            time.sleep(10 * (essai + 1))
    fin = max(brut.rfind(b"</rss>") + 6, brut.rfind(b"</feed>") + 7)
    if fin > 6:
        brut = brut[:fin]  # contenu parasite après la balise de fin (Maddyness)
    try:
        racine = ET.fromstring(brut)
    except ET.ParseError:  # « & » non échappés (Insurance Times)
        racine = ET.fromstring(re.sub(rb"&(?!#?\w+;)", b"&amp;", brut))
    atom = "{http://www.w3.org/2005/Atom}"
    items = racine.findall(".//item") or racine.findall(f".//{atom}entry")
    for it in items:
        lien = texte(it, "link")
        if not lien:
            l = it.find(f"{atom}link")
            lien = l.get("href", "") if l is not None else ""
        resume = texte(it, "description", f"{atom}summary", f"{atom}content")
        resume = html.unescape(re.sub(r"<[^>]+>", " ", resume))
        yield {
            "titre": html.unescape(texte(it, "title", f"{atom}title")),
            "lien": lien.strip(),
            "date": parse_date(texte(it, "pubDate", f"{atom}published", f"{atom}updated")),
            "extrait": re.sub(r"\s+", " ", resume).strip()[:SNIPPET_MAX],
        }


def liens_deja_vus(aujourdhui: dt.date) -> set:
    vus = set()
    for p in SORTIE.glob("*.json"):
        try:
            if (aujourdhui - dt.date.fromisoformat(p.stem)).days <= JOURS_DEDOUBLONNAGE:
                vus.update(a["lien"] for a in json.loads(p.read_text(encoding="utf-8"))["articles"])
        except (ValueError, KeyError, json.JSONDecodeError):
            continue
    return vus


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--jours", type=int, default=2)
    args = ap.parse_args()

    maintenant = dt.datetime.now(dt.timezone.utc)
    limite = maintenant - dt.timedelta(days=args.jours)
    SORTIE.mkdir(parents=True, exist_ok=True)
    fichier = SORTIE / f"{maintenant.date().isoformat()}.json"
    vus = liens_deja_vus(maintenant.date())
    if fichier.exists():  # relance le même jour : on complète le fichier existant
        vus -= {a["lien"] for a in json.loads(fichier.read_text(encoding="utf-8"))["articles"]}
    noms = noms_suivis()
    motifs = [(re.compile(rf"\b{re.escape(n)}\b"), nom, cat) for n, (nom, cat) in noms.items()]

    articles, etat_flux = {}, {}
    for source, url in FLUX.items():
        if "news.google.com" in url:
            time.sleep(2)  # espacer les requêtes Google News pour éviter le 503
        try:
            items = list(lire_flux(url))
        except Exception as e:
            etat_flux[source] = {"ok": False, "erreur": str(e)[:200]}
            continue
        retenus = 0
        for a in items:
            if not a["lien"] or a["lien"] in vus or a["lien"] in articles:
                continue
            if a["date"] and a["date"] < limite:
                continue
            corps = normaliser(f"{a['titre']} {a['extrait']}")
            mots = sorted({k for k in MOTS_CLES if normaliser(k) in corps})
            cites = sorted({f"{nom} ({cat})" for rx, nom, cat in motifs if rx.search(corps)})
            if not mots and not cites:
                continue
            articles[a["lien"]] = {
                "source": source,
                "titre": a["titre"],
                "lien": a["lien"],
                "date": a["date"].date().isoformat() if a["date"] else None,
                "extrait": a["extrait"],
                "motsCles": mots,
                "nomsSuivis": cites,
            }
            retenus += 1
        etat_flux[source] = {"ok": True, "items": len(items), "retenus": retenus}

    tri = sorted(articles.values(), key=lambda a: a["date"] or "", reverse=True)
    tri.sort(key=lambda a: not a["nomsSuivis"])  # noms suivis en tête, puis du plus récent
    sortie = {
        "collecte": maintenant.isoformat(timespec="seconds"),
        "fenetreJours": args.jours,
        "nbArticles": len(tri),
        "nbAvecNomSuivi": sum(1 for a in tri if a["nomsSuivis"]),
        "flux": etat_flux,
        "articles": tri,
    }
    fichier.write_text(json.dumps(sortie, ensure_ascii=False, indent=1), encoding="utf-8")
    ko = [s for s, e in etat_flux.items() if not e["ok"]]
    print(f"{fichier.name} : {len(tri)} articles ({sortie['nbAvecNomSuivi']} citant un nom suivi), "
          f"{len(FLUX) - len(ko)}/{len(FLUX)} flux OK" + (f" — en échec : {', '.join(ko)}" if ko else ""))


if __name__ == "__main__":
    main()
