#!/usr/bin/env python3
"""
generer_page_accueil.py — Page « Accueil » du site Radar Insurtech (tableau de bord d'entrée)
==========================================================================================

Première page pour un lecteur qui découvre le site : à quoi sert chaque page, chiffres clés,
levées récentes par priorité GVI, parcours conseillé, rythme de mise à jour, lexique.
Les chiffres sont lus dans les pages déjà générées (radar_insurtech_vc.html, gvi_strategie_pipe.html)
et dans veille_data/entries.json et ma_data/deals.json : à régénérer en dernier.
Aucune donnée confidentielle (pas de valorisation chiffrée).

Usage : python3 generer_page_accueil.py
"""
from __future__ import annotations

import html
import json
import re
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE = BASE_DIR / "_accueil_template.html"
SORTIE = BASE_DIR / "accueil.html"
URL = {
    "accueil": "https://claude.ai/artifact/3XXDgSBc7G557ghyD4ysqL",
    "radar": "https://claude.ai/artifact/X6o5LvDnYe6Qo5fPYFkX7g",
    "pipe": "https://claude.ai/artifact/58n6yRVTLgDuhavqhxihGn",
    "analyses": "https://claude.ai/artifact/Mor6DTNTSbekT4irHSLyJe",
    "veille": "https://claude.ai/artifact/5FXTpkCo6KvGu6YyeYAHiw",
    "ma": "https://claude.ai/artifact/BPSeTnKnWY2PkAGcH5s2ai",
}
e = html.escape


def donnees(fichier: str) -> dict:
    h = (BASE_DIR / fichier).read_text(encoding="utf-8")
    return json.loads(re.search(r"window\.__DATA__ = (\{.*?\});\s*</script>", h, re.S).group(1))


def nb(x, d=0) -> str:
    s = f"{x:,.{d}f}".replace(",", " ").replace(".", ",")
    return s


def nav() -> str:
    items = [("Accueil", None), ("Radar Insurtech VC", URL["radar"]), ("Stratégie &amp; Pipe GVI", URL["pipe"]),
             ("Analyses GVI", URL["analyses"]), ("Veille Insurtech", URL["veille"]), ("Étude M&amp;A Assurance", URL["ma"])]
    corps = '\n  <span class="sep">·</span>\n  '.join(
        f'<span class="current">{n}</span>' if u is None else f'<a href="{u}">{n}</a>' for n, u in items)
    js = ("<script>(() => { const e = document.querySelector('.entete-fixe'); if (!e) return; const f = () => "
          "document.documentElement.style.setProperty('--entete-h', e.offsetHeight + 'px'); f(); "
          "if (window.ResizeObserver) new ResizeObserver(f).observe(e); })();</script>")
    return f'<div class="entete-fixe pad">\n<nav class="crossnav">\n  {corps}\n</nav>\n</div>\n{js}'


def main() -> int:
    r, g = donnees("radar_insurtech_vc.html"), donnees("gvi_strategie_pipe.html")
    v = json.loads((BASE_DIR / "veille_data" / "entries.json").read_text(encoding="utf-8"))
    m = json.loads((BASE_DIR / "ma_data" / "deals.json").read_text(encoding="utf-8"))
    mk = r["marche"]
    gestion = len({f.get("nom") for f in r["fonds"]})
    cibles_neuves = sum(1 for c in r["cibles"] if c["score"] >= 50 and not c.get("pipe") and not c.get("portefeuille"))
    an = mk["anneeRef"]

    chiffres = [
        (nb(len(r["fonds"])), "véhicules d'investissement suivis", f"{gestion} sociétés de gestion, VC et CVC européens", URL["radar"]),
        (nb(len(r["startups"])), "participations recensées", f"{nb(mk['nbSocietes'])} start-ups distinctes", URL["radar"]),
        (nb(mk["nbLevees"]), f"levées en {an - 1}-{an}", f"{nb(mk['montantLevees'] / 1000, 2)} Md€ levés au total", URL["radar"] + "#marche"),
        (nb(cibles_neuves), "cibles nouvelles pour GVI", f"notées 50/100 ou plus, sur {len(r['cibles'])} évaluées", URL["radar"] + "#cibles"),
        (nb(g["total"]), "start-ups dans le pipe GVI", "suivi des dossiers de l'équipe", URL["pipe"]),
        (nb(len(v["entries"])), "actualités vérifiées", f"et {len(m['confirmees'])} opérations M&A assurance", URL["veille"]),
    ]
    tuiles = "\n".join(f'<a class="chiffre" href="{u}"><b>{c}</b><span>{e(l)}</span><small>{e(s)}</small></a>' for c, l, s, u in chiffres)

    pages = [
        ("Radar Insurtech VC", "Qui investit dans l'insurtech en Europe, et dans quoi ?",
         "Les fonds de capital-risque et les fonds des assureurs (CVC), leurs participations, les levées récentes, les tendances par secteur et par pays.",
         "Pour comprendre le marché et repérer des co-investisseurs.", URL["radar"],
         [("Marché", "#marche"), ("Cibles GVI", "#cibles"), ("Tendances", "#trends")]),
        ("Stratégie & Pipe GVI", "Quelles start-ups GVI suit‑elle, et où en est chaque dossier ?",
         "Le pipe de l'équipe : statut, secteur, maturité, fonds déjà au capital, données publiques des sociétés françaises.",
         "Pour l'équipe d'investissement.", URL["pipe"], []),
        ("Analyses GVI", "Cette start-up mérite‑t‑elle une étude approfondie ?",
         "Une pré-analyse en quelques minutes à partir du deck : six critères notés de 1 à 3, une recommandation, les questions à poser.",
         "Pour trier les dossiers entrants.", URL["analyses"], []),
        ("Valorisation du portefeuille", "Que valent nos participations aujourd'hui ?",
         "Valeur estimée de chaque participation selon la part du business plan réalisée, recoupée par les multiples du marché, et ce qu'elle peut rapporter à la sortie.",
         "Pour le suivi du portefeuille et le reporting.", URL["analyses"] + "#valo", []),
        ("Veille Insurtech", "Que s'est-il passé dans l'insurtech ces dernières semaines ?",
         "Les actualités vérifiées : levées, partenariats, nouveaux fonds. Collecte automatique chaque jour, sélection chaque lundi.",
         "Pour rester à jour en cinq minutes.", URL["veille"], []),
        ("Étude M&A Assurance", "Qui rachète qui dans l'assurance ?",
         "Les fusions et acquisitions du secteur, confirmées ou en rumeur, avec le contexte de marché.",
         "Pour anticiper les mouvements des concurrents et des partenaires.", URL["ma"], []),
    ]
    cartes = "\n".join(
        f'''<article class="page"><span class="nom">{e(n)}</span><h3 class="q">{e(q)}</h3><p>{e(t)}</p><span class="pour">{e(pour)}</span>
        {('<div class="voir">' + "".join(f'<a href="{u}{h}">{e(lib)}</a>' for lib, h in voir) + "</div>") if voir else ""}
        <a class="ouvrir" href="{u}">Ouvrir la page →</a></article>''' for n, q, t, pour, u, voir in pages)

    doms = mk.get("domaines") or []
    mx = max([d["n"] for d in doms] or [1])
    barres = "\n".join(f'<div class="bar"><span>{e(d["nom"])}</span><span class="t"><i style="width:{max(2, round(d["n"] / mx * 100))}%"></i></span><b>{d["n"]}</b></div>' for d in doms)

    contenu = f"""
<section class="hero">
  <h1>Radar Insurtech</h1>
  <p>L'outil de veille et d'aide à la décision de Groupama Vol'terre Investissement (GVI), le fonds de capital-risque de Groupama. Il rassemble en six pages ce qu'il faut savoir pour repérer, analyser et suivre les start-ups de l'assurance en Europe.</p>
  <p class="sous">Données au {e(r["genere"])} · chaque chiffre renvoie à une source, les points incertains sont signalés.</p>
</section>

<section>
  <h2>En chiffres</h2>
  <p class="intro">Cliquez sur un chiffre pour ouvrir la page correspondante.</p>
  <div class="chiffres">{tuiles}</div>
</section>

<section>
  <h2>Six pages, six questions</h2>
  <p class="intro">Chaque page répond à une question. La barre en haut de l'écran permet de passer de l'une à l'autre à tout moment.</p>
  <div class="pages">{cartes}</div>
</section>

<section>
  <h2>Où vont les levées récentes</h2>
  <p class="intro">Levées de {an - 1}-{an} dans les quatre priorités actuelles de GVI (une levée peut relever de plusieurs priorités). Le détail est dans l'onglet <a href="{URL["radar"]}#marche">Marché</a> du Radar.</p>
  <div class="bars">{barres}</div>
</section>

<section>
  <h2>Par où commencer</h2>
  <ol class="etapes">
    <li><b>Lire le marché</b>Ouvrez l'onglet <a href="{URL["radar"]}#marche">Marché</a> du Radar : levées récentes, fonds les plus actifs, valorisations, sorties.</li>
    <li><b>Repérer des cibles</b>L'onglet <a href="{URL["radar"]}#cibles">Cibles GVI</a> classe les start-ups européennes selon la thèse de GVI ; filtrez par priorité.</li>
    <li><b>Analyser un dossier</b>Dans <a href="{URL["analyses"]}">Analyses GVI</a>, déposez le deck d'une société pour obtenir une pré-analyse en six critères.</li>
    <li><b>Suivre le portefeuille</b>La <a href="{URL["analyses"]}#valo">Valorisation du portefeuille</a> compare chaque participation à son business plan.</li>
  </ol>
</section>

<section>
  <h2>Comment les données sont tenues à jour</h2>
  <div class="table-w"><table class="maj">
    <thead><tr><th>Page</th><th>Mise à jour</th><th>Dernière</th></tr></thead>
    <tbody>
      <tr><td>Veille Insurtech</td><td>Collecte automatique chaque jour dans une cinquantaine de sources ; sélection et vérification chaque lundi</td><td>{e(v.get("derniereMaj", ""))}</td></tr>
      <tr><td>Étude M&amp;A Assurance</td><td>Chaque mois, le 1er</td><td>{e(m.get("derniereMaj", ""))}</td></tr>
      <tr><td>Radar Insurtech VC</td><td>Par versions successives de la base (sites des fonds, presse, registres publics) ; l'onglet Fiabilité liste les points à confirmer</td><td>{e(r["genere"])} (version {e(str(r["version"]))})</td></tr>
      <tr><td>Stratégie &amp; Pipe GVI</td><td>À partir du fichier de suivi de l'équipe GVI</td><td>{e(g.get("genere", ""))}</td></tr>
      <tr><td>Analyses et valorisations</td><td>À la demande, à partir des documents transmis par les sociétés</td><td>—</td></tr>
    </tbody></table></div>
</section>

<section>
  <h2>Lexique</h2>
  <dl class="lexique">
    <div><dt>VC (capital-risque)</dt><dd>Fonds qui investit dans des start-ups en échange d'une part de leur capital.</dd></div>
    <div><dt>CVC (corporate venture capital)</dt><dd>Fonds de capital-risque d'un groupe industriel ou d'un assureur, comme GVI pour Groupama.</dd></div>
    <div><dt>Insurtech</dt><dd>Start-up qui utilise la technologie pour transformer l'assurance : distribution, tarification, sinistres, prévention.</dd></div>
    <div><dt>Levée, tour, Seed, Série A</dt><dd>Augmentation de capital d'une start-up ; Seed et Série A désignent les premiers tours.</dd></div>
    <div><dt>Ticket</dt><dd>Montant investi par un fonds dans un tour.</dd></div>
    <div><dt>Valorisation pré-money / post-money</dt><dd>Valeur de la société avant / après l'argent levé lors d'un tour.</dd></div>
    <div><dt>ARR</dt><dd>Revenu récurrent annualisé (abonnements) : l'indicateur clé d'une start-up logicielle.</dd></div>
    <div><dt>Business plan (BP)</dt><dd>Trajectoire de chiffre d'affaires présentée lors d'une levée ; la valorisation compare le réalisé à ce plan.</dd></div>
    <div><dt>Exit (sortie)</dt><dd>Cession de la société ou introduction en bourse, qui permet au fonds de récupérer son investissement.</dd></div>
    <div><dt>TVPI</dt><dd>(Montants distribués + valeur actuelle) / montants investis : la performance d'un portefeuille à date.</dd></div>
    <div><dt>DPI</dt><dd>Montants effectivement distribués / montants investis : ce qui est déjà revenu en trésorerie.</dd></div>
    <div><dt>Préférence de liquidation</dt><dd>Droit d'un investisseur à récupérer sa mise en priorité en cas de vente.</dd></div>
    <div><dt>Pipe</dt><dd>Liste des start-ups suivies par l'équipe, de la première rencontre à la décision d'investir.</dd></div>
  </dl>
</section>

<p class="foot">Radar Insurtech — outil interne de GVI. Données publiques (presse, sites des fonds, registres) et données de l'équipe ; chaque information est sourcée sur sa page.</p>
"""
    page = TEMPLATE.read_text(encoding="utf-8").replace("__NAV__", nav()).replace("__CONTENU__", contenu)
    SORTIE.write_text(page, encoding="utf-8")
    print(f"Page produite : {SORTIE.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
