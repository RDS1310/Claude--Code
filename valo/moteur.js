/* =====================================================================================
   Moteur « Valorisation du portefeuille » — partagé par la page Analyses GVI (onglet privé,
   nominatif) et la page de démonstration anonymisée. Injecté par les générateurs.
   Aucune donnée ici : les participations sont lues dans la base de la page (collection
   « valorisations »), écrites par Claude (ArtifactData) ou depuis la page.

   Méthode (lignes directrices IPEV, pratique des fonds de VC) :
   1. point de départ = prix du dernier tour (post-money) ;
   2. ajusté du % du business plan atteint à date (ARR réalisé / ARR attendu au BP de référence,
      interpolé au mois), avec une élasticité (k = 1 : proportionnel) et un crédit partiel de la
      surperformance tant qu'aucun nouveau tour ne la confirme ;
   3. recoupé par les multiples d'ARR des comparables (+ trésorerie nette) ;
   4. valeur de la ligne GVI après droits (préférence de liquidation 1x non participative) ;
   5. revue trimestrielle : valeur figée par date, alertes, scénarios de sortie (multiple, TRI, DPI).
   ===================================================================================== */
window.ValoApp = function (root, opts) {
  opts = opts || {};
  const COLL = opts.collection || 'valorisations';
  const VS = {docs: [], id: null, brouillon: null, sale: false, msg: '', msgErr: false, nouvelle: false, maj: false, charge: false};
  let VDB = null, PEUT = null;

  /* ---------- Formats ---------- */
  const esc = s => (s == null ? '' : String(s)).replace(/[&<>"]/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]));
  const ok = x => x != null && x !== '' && isFinite(x);
  const nf = (x, d) => ok(x) ? (+x).toLocaleString('fr-FR', {minimumFractionDigits: d, maximumFractionDigits: d}) : '—';
  const eur = x => !ok(x) ? '—' : +x === 0 ? '0 €' : Math.abs(x) < 1 ? nf(x * 1000, 0) + ' K€' : nf(x, Math.abs(x) < 10 ? 2 : 1) + ' M€';
  const fois = x => ok(x) ? '×' + nf(x, x < 10 ? 2 : 1) : '—';
  const pct = (x, d = 0) => ok(x) ? nf(x * 100, d) + ' %' : '—';
  const clone = o => JSON.parse(JSON.stringify(o));
  const BORNES = ['bas', 'central', 'haut'];
  const MOIS = ['janv.', 'févr.', 'mars', 'avr.', 'mai', 'juin', 'juil.', 'août', 'sept.', 'oct.', 'nov.', 'déc.'];
  const mi = s => { const m = String(s || '').match(/^(\d{4})(?:-(\d{2}))?/); return m ? +m[1] * 12 + ((+m[2] || 12) - 1) : null; };
  const libMois = i => ok(i) ? `${MOIS[i % 12]} ${Math.floor(i / 12)}` : '—';
  const peutEcrire = () => !!VDB && PEUT !== false;

  /* ---------- Moteur ---------- */
  const postMoney = d => { const e = d.entree || {}; return ok(e.preMoney) && ok(e.leve) ? e.preMoney + e.leve : null; };
  const detention = d => { const e = d.entree || {}; if (ok(e.detention)) return e.detention; const p = postMoney(d); return p && ok(e.ticket) ? e.ticket / p : null; };
  const anneeEntree = d => { const i = mi((d.entree || {}).date); return ok(i) ? i / 12 : null; };
  const serie = pts => (pts || []).filter(p => ok(mi(p.date)) && ok(p.arr)).map(p => ({m: mi(p.date), v: +p.arr, p})).sort((a, b) => a.m - b.m);
  function interp(pts, m){
    if (!pts.length || m < pts[0].m || m > pts[pts.length - 1].m) return null;
    for (let i = 0; i < pts.length; i++) {
      if (pts[i].m === m) return pts[i].v;
      if (pts[i].m > m) { const a = pts[i - 1], b = pts[i]; return a.v + (b.v - a.v) * (m - a.m) / (b.m - a.m); }
    }
    return null;
  }
  const decote = d => ((d.bp || {}).decote) || 0;
  const bpRetenu = d => serie((d.bp || {}).points).map(x => ({...x, v: x.v * (1 - decote(d))}));
  function atteinte(d){
    const r = serie(d.realise), dern = r[r.length - 1];
    if (!dern) return {};
    const bp = interp(bpRetenu(d), dern.m), bpM = interp(serie((d.bp || {}).points), dern.m);
    return {arr: dern.v, m: dern.m, bp, bpM, r: ok(bp) && bp > 0 ? dern.v / bp : null};
  }
  function facteurBP(r, k, part){
    if (!ok(r)) return null;
    return r < 1 ? Math.pow(Math.max(r, 0), ok(k) ? k : 1) : 1 + (ok(part) ? part : 0.5) * (r - 1);
  }
  function dcf(flux, wacc, g, an){
    const f = (flux || []).filter(x => x.annee > an && ok(x.valeur)).sort((a, b) => a.annee - b.annee);
    if (!f.length || !ok(wacc) || !ok(g) || wacc <= g) return null;
    let v = 0;
    f.forEach((x, i) => { v += x.valeur / Math.pow(1 + wacc, i + 1); });
    const der = f[f.length - 1].valeur;
    if (der > 0) v += der * (1 + g) / (wacc - g) / Math.pow(1 + wacc, f.length);
    return Math.max(0, v);
  }
  const tresoNette = d => { const t = d.tresorerie || {}; return ok(t.nette) ? +t.nette : null; };
  function valeurMethode(d, m){
    const r = {};
    if (m.type === 'bp') {
      const a = atteinte(d), v0 = ok(m.valeurRef) ? m.valeurRef : postMoney(d), aj = ok(m.ajustMarche) ? m.ajustMarche : 0;
      r.atteinte = a; r.v0 = v0;
      BORNES.forEach(b => { const f = facteurBP(a.r, (m.elasticite || {})[b], (m.partHausse || {})[b]); r[b] = ok(f) && ok(v0) ? v0 * f * (1 + aj) : null; });
    } else if (m.type === 'multiples') {
      const a = atteinte(d), tn = m.avecTresorerie !== false ? tresoNette(d) : null;
      r.arr = a.arr; r.tn = tn;
      BORNES.forEach(b => { r[b] = ok(m[b]) && ok(a.arr) ? m[b] * a.arr + (ok(tn) ? tn : 0) : null; });
    } else if (m.type === 'tour') {
      const base = ok(m.base) ? m.base : postMoney(d);
      BORNES.forEach(b => { const v = (m.variation || {})[b]; r[b] = ok(base) && ok(v) ? base * (1 + v) : null; });
    } else if (m.type === 'dcf') {
      BORNES.forEach(b => { const sc = (m.scenarios || []).find(s => s.cle === (m.choix || {})[b]); r[b] = sc ? dcf(sc.flux, m.wacc, m.g, m.anneeValo) : null; });
    }
    return r;
  }
  // Ce que reçoit GVI pour une valeur V des titres : préférence 1x non participative, pari passu avec
  // les autres actions de préférence (total P) ; GVI convertit si sa quote-part dépasse sa préférence.
  function produitGVI(d, V, det, P){
    const e = d.entree || {}, t = e.ticket;
    if (!ok(V) || !ok(det)) return null;
    if (!e.preference || !ok(P) || !ok(t)) return det * V;
    const mult = ok(e.prefMultiple) ? e.prefMultiple : 1, pref = t * mult, Ptot = P * mult;
    if (V <= Ptot) return V * pref / Ptot;
    return Math.max(pref, det * V);
  }
  function synthese(d){
    const methodes = (d.methodes || []).map((m, i) => ({m, i, r: valeurMethode(d, m)}));
    const s = {methodes, det: detention(d)};
    BORNES.forEach(b => {
      let n = 0, w = 0;
      methodes.forEach(({m, r}) => { if (ok(r[b]) && m.poids > 0) { n += m.poids * r[b]; w += m.poids; } });
      s[b] = w ? n / w : null;
    });
    const e = d.entree || {}, t = e.ticket;
    s.ligne = {}; s.converti = {};
    BORNES.forEach(b => { s.ligne[b] = produitGVI(d, s[b], s.det, e.prefTotal); s.converti[b] = ok(s[b]) && ok(s.det) ? s[b] * s.det : null; });
    s.tvpi = ok(s.ligne.central) && t ? (s.ligne.central + (d.distributions || 0)) / t : null;
    s.atteinte = atteinte(d);
    return s;
  }
  function sortie(d){
    const so = d.sortie || {}, det0 = detention(d), t = (d.entree || {}).ticket, a0 = anneeEntree(d), dil = so.dilution || 0;
    const det = ok(det0) ? det0 * (1 - dil) : null, P = ok(so.prefTotal) ? so.prefTotal : (d.entree || {}).prefTotal;
    const lignes = (so.scenarios || []).map(x => {
      const ve = ok(x.valeur) ? +x.valeur : ok(x.arr) && ok(x.multiple) ? x.arr * x.multiple : null;
      const prod = produitGVI(d, ve, det, P);
      const moic = ok(prod) && t ? prod / t : null;
      const an = x.annee || so.annee, n = ok(an) && ok(a0) ? an + 0.5 - a0 : null;
      const tri = ok(moic) && n > 0 ? (moic > 0 ? Math.pow(moic, 1 / n) - 1 : -1) : null;
      return {...x, ve, prod, moic, tri, n, an};
    });
    const sp = lignes.reduce((a, x) => a + (x.proba || 0), 0);
    const esp = lignes.reduce((a, x) => a + (x.proba || 0) * (x.prod || 0), 0);
    const nEsp = sp ? lignes.reduce((a, x) => a + (x.proba || 0) * (x.n || 0), 0) / sp : null;
    const moicEsp = t ? esp / t : null;
    const triEsp = ok(moicEsp) && nEsp > 0 && moicEsp > 0 ? Math.pow(moicEsp, 1 / nEsp) - 1 : null;
    return {lignes, sp, esp, moicEsp, triEsp, dil, det, P};
  }
  // Alertes calculées + jalons saisis : {niveau: 'rouge'|'orange'|'vert'|'gris', libelle, detail}
  function alertes(d){
    const out = [], a = atteinte(d), t = d.tresorerie || {};
    if (ok(a.r)) out.push({niveau: a.r < 0.7 ? 'rouge' : a.r < 0.9 ? 'orange' : 'vert', libelle: `Atteinte du BP : ${pct(a.r)}`, detail: `ARR ${eur(a.arr)} en ${libMois(a.m)}, contre ${eur(a.bp)} attendus${decote(d) ? ` (BP de référence, ${pct(decote(d))} sous celui du management)` : ''}.`});
    else out.push({niveau: 'gris', libelle: 'Atteinte du BP non calculable', detail: 'Renseignez l\'ARR réalisé et la trajectoire du BP.'});
    if (ok(t.montant)) {
      const mois = ok(t.burn) && t.burn > 0 ? t.montant / t.burn : null;
      out.push(ok(mois) ? {niveau: mois < 6 ? 'rouge' : mois < 12 ? 'orange' : 'vert', libelle: `Runway : ${mois > 36 ? 'plus de 36' : nf(mois, 0)} mois`, detail: `Trésorerie ${eur(t.montant)} (${esc(t.date || '')}), consommation ${eur(t.burn)} par mois.`}
        : {niveau: 'vert', libelle: 'Autofinancée', detail: `Trésorerie ${eur(t.montant)} (${esc(t.date || '')}), sans consommation nette.`});
    } else out.push({niveau: 'gris', libelle: 'Trésorerie non communiquée', detail: 'Runway inconnu : à demander à la société.'});
    if (ok(a.m)) {
      const age = (new Date().getFullYear() * 12 + new Date().getMonth()) - a.m;
      if (age > 6) out.push({niveau: age > 9 ? 'rouge' : 'orange', libelle: `Données de ${libMois(a.m)}`, detail: 'Dernier chiffre réalisé connu : à actualiser.'});
    }
    return out.concat((d.jalons || []).map(j => ({niveau: j.statut || 'gris', libelle: j.libelle, detail: j.commentaire || ''})));
  }
  const NIV = {rouge: 'Alerte', orange: 'À surveiller', vert: 'OK', gris: 'n.d.'};

  /* ---------- Champs modifiables ---------- */
  const getP = (o, p) => p.split('.').reduce((a, k) => a == null ? a : a[k], o);
  function setP(o, p, v){ const ks = p.split('.'); let a = o; ks.slice(0, -1).forEach(k => { if (a[k] == null) a[k] = /^\d+$/.test(k) ? [] : {}; a = a[k]; }); a[ks[ks.length - 1]] = v; }
  function champ(d, p, fmt = 'eur', titre = ''){
    const v = getP(d, p);
    if (!peutEcrire()) return fmt === 'pct' ? pct(v, 1) : fmt === 'x' ? fois(v) : fmt === 'an' ? (v ?? '—') : fmt === 'n' ? nf(v, 2) : eur(v);
    const aff = !ok(v) ? '' : fmt === 'pct' ? +(v * 100).toFixed(2) : +(+v).toFixed(4);
    const unite = {eur: 'M€', pct: '%', x: '×', n: '', an: ''}[fmt];
    return `<span class="vin-w"><input class="vin" type="number" step="any" inputmode="decimal" data-p="${esc(p)}" data-f="${fmt}" value="${aff}" aria-label="${esc(titre || p)}">${unite ? `<i>${unite}</i>` : ''}</span>`;
  }

  /* ---------- Graphique : ARR réalisé contre BP ---------- */
  function graphe(d){
    const R = serie(d.realise), B = bpRetenu(d), BM = decote(d) ? serie((d.bp || {}).points) : [], F = serie(d.prevision);
    if (!R.length || !B.length) return '';
    const m0 = Math.min(R[0].m, B[0].m), mFin = Math.max(R[R.length - 1].m, F.length ? F[F.length - 1].m : 0) + 3;
    const m1 = Math.min(Math.max(mFin, R[R.length - 1].m + 6), B[B.length - 1].m);
    const coupe = s => { if (!s.length) return []; const pts = s.filter(x => x.m >= m0 && x.m <= m1); [m0, m1].forEach(m => { const v = interp(s, m); if (ok(v) && !pts.some(x => x.m === m)) pts.push({m, v}); }); return pts.sort((a, b) => a.m - b.m); };
    const series = [
      {cle: 'r', nom: 'ARR réalisé', court: 'Réalisé', pts: R, cls: 's1', marque: true},
      {cle: 'b', nom: decote(d) ? 'BP de référence' : 'Business plan', court: 'BP', pts: coupe(B), cls: 's2', marque: false},
      ...(F.length ? [{cle: 'f', nom: 'Prévision de la société', court: 'Prévision', pts: F, cls: 's3 tir', marque: true}] : []),
      ...(BM.length ? [{cle: 'm', nom: 'BP du management', pts: coupe(BM), cls: 'ref', marque: false}] : [])];
    const vmax = Math.max(...series.flatMap(s => s.pts.map(p => p.v))) * 1.12 || 1;
    const W = 700, H = 250, g = 46, dr = 140, ht = 14, bs = 30;
    const X = m => g + (m - m0) / Math.max(1, m1 - m0) * (W - g - dr), Y = v => H - bs - v / vmax * (H - bs - ht);
    const pas = [0.1, 0.2, 0.25, 0.5, 1, 2, 2.5, 5, 10, 20, 25, 50].find(p => vmax / p <= 5) || 100;
    const ticksY = []; for (let v = 0; v <= vmax; v += pas) ticksY.push(v);
    const ticksX = []; for (let m = m0 - (m0 % 3); m <= m1; m += 3) if (m >= m0) ticksX.push(m);
    const ligne = s => s.pts.map((p, i) => `${i ? 'L' : 'M'}${X(p.m).toFixed(1)},${Y(p.v).toFixed(1)}`).join('');
    const etiq = []; series.filter(s => s.cle !== 'm').forEach(s => { const p = s.pts[s.pts.length - 1]; if (p) etiq.push({s, p, y: Y(p.v)}); });
    etiq.sort((a, b) => a.y - b.y); for (let i = 1; i < etiq.length; i++) if (etiq[i].y - etiq[i - 1].y < 15) etiq[i].y = etiq[i - 1].y + 15;
    const a = atteinte(d);
    const json = esc(JSON.stringify({m0, m1, W, g, dr, series: series.map(s => ({nom: s.nom, cls: s.cls, pts: s.pts.map(p => [p.m, p.v])}))}));
    return `<figure class="vg" data-g="${json}">
      <svg viewBox="0 0 ${W} ${H}" role="img" aria-label="ARR réalisé comparé au business plan">
        ${ticksY.map(v => `<line class="gr" x1="${g}" x2="${W - dr}" y1="${Y(v)}" y2="${Y(v)}"/><text class="ax" x="${g - 6}" y="${Y(v) + 4}" text-anchor="end">${nf(v, pas < 1 ? 1 : 0)}</text>`).join('')}
        ${ticksX.map(m => `<text class="ax" x="${X(m)}" y="${H - 10}" text-anchor="middle">${m % 12 === 0 ? 'janv. ' + Math.floor(m / 12) : MOIS[m % 12]}</text>`).join('')}
        ${ok(a.m) && ok(a.bp) ? `<line class="ecart" x1="${X(a.m)}" x2="${X(a.m)}" y1="${Y(a.bp)}" y2="${Y(a.arr)}"/><text class="ecart-t" x="${X(a.m) - 6}" y="${(Y(a.bp) + Y(a.arr)) / 2 + 4}" text-anchor="end">${pct(a.r)} du BP</text>` : ''}
        ${series.map(s => `<path class="ln ${s.cls}" d="${ligne(s)}"/>${s.marque ? s.pts.map(p => `<circle class="pt ${s.cls}" cx="${X(p.m)}" cy="${Y(p.v)}" r="4"/>`).join('') : ''}`).join('')}
        ${etiq.map(e => `<text class="dl" x="${W - dr + 8}" y="${e.y + 4}">${esc(e.s.court)} · ${eur(e.p.v)}</text>`).join('')}
        <line class="xh" x1="0" x2="0" y1="${ht}" y2="${H - bs}" hidden/>
        <rect class="hit" x="${g}" y="0" width="${W - g - dr}" height="${H - bs}"/>
      </svg>
      <div class="vg-tip" hidden></div>
      <figcaption class="vg-leg">${series.map(s => `<span><i class="k ${s.cls}"></i>${esc(s.nom)}</span>`).join('')}<span class="vg-u">ARR en M€</span></figcaption>
    </figure>`;
  }
  function brancherGraphe(){
    root.querySelectorAll('figure.vg').forEach(fig => {
      const G = JSON.parse(fig.dataset.g), svg = fig.querySelector('svg'), hit = fig.querySelector('.hit'), xh = fig.querySelector('.xh'), tip = fig.querySelector('.vg-tip');
      const toM = px => G.m0 + (px - G.g) / (G.W - G.g - G.dr) * (G.m1 - G.m0);
      const xs = [...new Set(G.series.flatMap(s => s.pts.map(p => p[0])))].sort((a, b) => a - b);
      const val = (s, m) => { const p = s.pts.find(q => q[0] === m); if (p) return p[1]; if (s.cls.includes('s1')) return null; return interp(s.pts.map(q => ({m: q[0], v: q[1]})), m); };
      const montrer = ev => {
        const r = svg.getBoundingClientRect(), px = (ev.clientX - r.left) / r.width * G.W, m = toM(px);
        const mm = xs.reduce((a, b) => Math.abs(b - m) < Math.abs(a - m) ? b : a, xs[0]);
        const x = G.g + (mm - G.m0) / (G.m1 - G.m0) * (G.W - G.g - G.dr);
        xh.setAttribute('x1', x); xh.setAttribute('x2', x); xh.hidden = false;
        tip.textContent = '';
        const h = document.createElement('b'); h.textContent = libMois(mm); tip.appendChild(h);
        G.series.forEach(s => { const v = val(s, mm); if (!ok(v)) return; const l = document.createElement('div'); const k = document.createElement('i'); k.className = 'k ' + s.cls; const st = document.createElement('strong'); st.textContent = eur(v); const n = document.createElement('span'); n.textContent = ' ' + s.nom; l.append(k, st, n); tip.appendChild(l); });
        tip.hidden = false;
        const fx = x / G.W * r.width; tip.style.left = Math.min(Math.max(0, fx + 12), r.width - tip.offsetWidth) + 'px';
      };
      hit.addEventListener('pointermove', montrer);
      hit.addEventListener('pointerleave', () => { xh.hidden = true; tip.hidden = true; });
    });
  }

  /* ---------- Rendu ---------- */
  const doc = () => VS.brouillon || VS.docs.find(x => x.id === VS.id);
  const pastille = n => `<span class="vniv ${n}">${NIV[n] || n}</span>`;
  function enBref(d, s){
    const a = s.atteinte, t = (d.entree || {}).ticket, pm = postMoney(d);
    const parts = [];
    if (ok(a.r)) parts.push(`À ${libMois(a.m)}, la société a réalisé <b>${pct(a.r)} de son business plan</b> (ARR ${eur(a.arr)} contre ${eur(a.bp)} attendus).`);
    if (ok(s.central)) parts.push(`Sa valeur est estimée à <b>${eur(s.central)}</b> pour 100 % des titres (${eur(s.bas)} – ${eur(s.haut)}), ${ok(pm) ? `soit ${s.central >= pm ? '+' : '−'}${pct(Math.abs(s.central / pm - 1))} par rapport au prix d'entrée de GVI` : ''}.`);
    if (ok(s.ligne.central) && t) parts.push(`La ligne de GVI vaut <b>${eur(s.ligne.central)}</b> pour ${eur(t)} investis (${fois(s.tvpi)})${ok(s.converti.central) && s.ligne.central > s.converti.central + 1e-9 ? `, dont ${eur(s.ligne.central - s.converti.central)} apportés par la préférence de liquidation` : ''}.`);
    return parts.join(' ');
  }
  function portefeuilleHtml(){
    const ds = VS.docs; if (!ds.length) return '';
    let inv = 0, dist = 0, fv = {bas: 0, central: 0, haut: 0}, esp = 0, nbAl = 0, complet = true;
    const parAn = {};
    const cartes = ds.map(d => {
      const s = synthese(d), x = sortie(d), t = (d.entree || {}).ticket || 0, al = alertes(d), rouges = al.filter(z => z.niveau === 'rouge').length;
      inv += t; dist += d.distributions || 0; esp += x.esp || 0; nbAl += rouges;
      BORNES.forEach(b => { if (ok(s.ligne[b])) fv[b] += s.ligne[b]; else complet = false; });
      x.lignes.forEach(l => { if (l.an && l.prod) parAn[l.an] = (parAn[l.an] || 0) + (l.proba || 0) * l.prod; });
      const r = s.atteinte.r, w = ok(r) ? Math.min(r, 1.3) / 1.3 * 100 : 0;
      return `<button type="button" class="vcarte ${d.id === VS.id ? 'sel' : ''}" data-vid="${esc(d.id)}">
        <span class="vc-h"><b>${esc(d.nom)}</b><small>${esc(d.secteur || '')}</small></span>
        <span class="vc-bp"><span>Atteinte du BP</span><span class="jauge"><i class="${ok(r) ? (r < 0.7 ? 'rouge' : r < 0.9 ? 'orange' : 'vert') : ''}" style="width:${w.toFixed(0)}%"></i><em style="left:${(100 / 1.3).toFixed(1)}%"></em></span><b>${pct(r)}</b></span>
        <span class="vc-v"><span>Valeur de la ligne</span><b>${eur(s.ligne.central)}</b><small>${fois(s.tvpi)} le ticket de ${eur(t)}</small></span>
        <span class="vc-al">${al.slice(0, 4).map(z => `<span class="vdot ${z.niveau}" title="${esc(z.libelle)}"></span>`).join('')}${rouges ? `<small>${rouges} alerte${rouges > 1 ? 's' : ''}</small>` : '<small>aucune alerte</small>'}</span>
      </button>`;
    }).join('');
    let cumul = 0;
    const dpi = Object.keys(parAn).sort().map(a => { cumul += parAn[a]; return `<span><b>${a}</b> ${fois(inv ? (dist + cumul) / inv : null)}</span>`; }).join('');
    return `
    <div class="vkpis">
      <div><span>Investi</span><b>${eur(inv)}</b><small>${ds.length} participation${ds.length > 1 ? 's' : ''}</small></div>
      <div><span>Valeur des lignes GVI</span><b>${eur(fv.central)}</b><small>${eur(fv.bas)} – ${eur(fv.haut)}${complet ? '' : ' · incomplète'}</small></div>
      <div><span>TVPI</span><b>${fois(inv ? (dist + fv.central) / inv : null)}</b><small>(distribué + valeur) / investi</small></div>
      <div><span>DPI réalisé</span><b>${fois(inv ? dist / inv : null)}</b><small>${dist ? eur(dist) + ' distribués' : 'aucune cession à date'}</small></div>
      <div><span>Multiple espéré à la sortie</span><b>${fois(inv ? esp / inv : null)}</b><small>${dpi ? 'DPI projeté : ' : ''}${dpi.replace(/<\/?b>/g, '').replace(/<\/?span>/g, ' ')}</small></div>
      <div class="${nbAl ? 'al' : ''}"><span>Alertes</span><b>${nbAl}</b><small>${nbAl ? 'à traiter en revue trimestrielle' : 'aucun signal rouge'}</small></div>
    </div>
    <div class="vcartes">${cartes}</div>`;
  }
  function footballHtml(d, s){
    const lignes = s.methodes.filter(x => BORNES.some(b => ok(x.r[b])));
    const pm = postMoney(d);
    const vals = [pm, s.haut, ...lignes.map(x => x.r.haut)].filter(ok);
    if (!vals.length) return '<p class="vnote">Renseignez au moins une méthode pour obtenir une fourchette.</p>';
    const max = Math.max(...vals) * 1.08, pos = v => (Math.max(0, v) / max * 100).toFixed(2) + '%';
    const barre = (r, cls, lib, sous) => `<div class="ff-row ${cls}">
        <div class="ff-l"><b>${lib}</b>${sous ? `<small>${sous}</small>` : ''}</div>
        <div class="ff-t">${pm ? `<i class="ff-entree" style="left:${pos(pm)}"></i>` : ''}
          ${ok(r.bas) && ok(r.haut) ? `<span class="ff-b" style="left:${pos(Math.min(r.bas, r.haut))};width:calc(${pos(Math.max(r.bas, r.haut))} - ${pos(Math.min(r.bas, r.haut))})"></span>` : ''}
          ${ok(r.central) ? `<span class="ff-c" style="left:${pos(r.central)}"></span>` : ''}</div>
        <div class="ff-v">${eur(r.bas)} · <b>${eur(r.central)}</b> · ${eur(r.haut)}</div></div>`;
    return `<div class="ff">
      ${lignes.map(({m, r}) => barre(r, m.poids > 0 ? '' : 'off', esc(m.nom), m.poids > 0 ? `poids ${pct(m.poids)}` : 'pour information (poids 0)')).join('')}
      ${barre(s, 'syn', 'Valeur retenue', 'synthèse pondérée, 100 % des titres')}
      <div class="ff-leg"><span><i class="lb"></i>fourchette basse – haute</span><span><i class="lc"></i>valeur centrale</span>${pm ? `<span><i class="le"></i>prix d'entrée de GVI (post-money ${eur(pm)})</span>` : ''}</div>
    </div>`;
  }
  function methodesHtml(d, s){
    return s.methodes.map(({m, i, r}) => {
      let corps = '';
      const tete = cols => `<thead><tr><th></th>${cols.map(c => `<th class="num">${c}</th>`).join('')}</tr></thead>`;
      if (m.type === 'bp') {
        const a = r.atteinte || {};
        corps = `<ol class="vetapes">
          <li><span>Prix d'entrée (post-money)</span><b>${ok(m.valeurRef) ? champ(d, `methodes.${i}.valeurRef`, 'eur', 'Valeur de départ') : eur(r.v0)}</b></li>
          <li><span>ARR attendu au BP de référence en ${libMois(a.m)}</span><b>${eur(a.bp)}</b></li>
          <li><span>ARR réalisé</span><b>${eur(a.arr)}</b></li>
          <li class="hl"><span>Atteinte du BP</span><b>${pct(a.r)}</b></li></ol>
          <p class="vnote">Sous le plan, la valeur suit l'atteinte à la puissance k (k = 1 : proportionnel ; k &gt; 1 : le marché sanctionne plus que proportionnellement une croissance moindre). Au-dessus du plan, seule une part de la surperformance est créditée tant qu'un nouveau tour ne l'a pas confirmée.</p>
          <div class="vtable-w"><table class="vtable">${tete(['Bas', 'Central', 'Haut'])}<tbody>
          <tr><td>Élasticité k (sous le plan)</td>${BORNES.map(b => `<td class="num">${champ(d, `methodes.${i}.elasticite.${b}`, 'n', `Élasticité ${b}`)}</td>`).join('')}</tr>
          <tr><td>Surperformance créditée</td>${BORNES.map(b => `<td class="num">${champ(d, `methodes.${i}.partHausse.${b}`, 'pct', `Part créditée ${b}`)}</td>`).join('')}</tr>
          <tr><td>Facteur appliqué au prix</td>${BORNES.map(b => `<td class="num">${fois(facteurBP(a.r, (m.elasticite || {})[b], (m.partHausse || {})[b]))}</td>`).join('')}</tr>
          <tr class="tot"><td>Valeur à 100 %</td>${BORNES.map(b => `<td class="num">${eur(r[b])}</td>`).join('')}</tr></tbody></table></div>
          <p class="vnote">Ajustement de marché (évolution des multiples du secteur depuis l'entrée) : ${champ(d, `methodes.${i}.ajustMarche`, 'pct', 'Ajustement de marché')}. Décote du BP de référence par rapport à celui du management : ${champ(d, 'bp.decote', 'pct', 'Décote du BP')}${(d.bp || {}).decoteNote ? ` — ${esc(d.bp.decoteNote)}` : ''}.</p>`;
      } else if (m.type === 'multiples') corps = `<div class="vtable-w"><table class="vtable">${tete(['Bas', 'Central', 'Haut'])}<tbody>
        <tr><td>Multiple d'ARR</td>${BORNES.map(b => `<td class="num">${champ(d, `methodes.${i}.${b}`, 'x', `Multiple ${b}`)}</td>`).join('')}</tr>
        <tr><td>× ARR réalisé</td>${BORNES.map(() => `<td class="num">${eur(r.arr)}</td>`).join('')}</tr>
        <tr><td>+ trésorerie nette</td>${BORNES.map(() => `<td class="num">${ok(r.tn) ? eur(r.tn) : 'n.d.'}</td>`).join('')}</tr>
        <tr class="tot"><td>Valeur à 100 %</td>${BORNES.map(b => `<td class="num">${eur(r[b])}</td>`).join('')}</tr></tbody></table></div>`;
      else if (m.type === 'tour') corps = `<div class="vtable-w"><table class="vtable">${tete(['Down', 'Flat', 'Up'])}<tbody>
        <tr><td>Variation au prochain tour</td>${BORNES.map(b => `<td class="num">${champ(d, `methodes.${i}.variation.${b}`, 'pct', `Variation ${b}`)}</td>`).join('')}</tr>
        <tr class="tot"><td>Valeur à 100 %</td>${BORNES.map(b => `<td class="num">${eur(r[b])}</td>`).join('')}</tr></tbody></table></div>`;
      else if (m.type === 'dcf') {
        const scs = m.scenarios || [], ans = [...new Set(scs.flatMap(x => (x.flux || []).map(f => f.annee)))].sort();
        corps = `<p class="vnote">Flux ${esc(m.fluxLibelle || '')} postérieurs à ${esc(m.anneeValo)}, actualisés à ${champ(d, `methodes.${i}.wacc`, 'pct', 'Taux d\'actualisation')}, croissance à l'infini ${champ(d, `methodes.${i}.g`, 'pct', 'Croissance')}.</p>
        <div class="vtable-w"><table class="vtable"><thead><tr><th>Scénario</th>${ans.map(a => `<th class="num">${a}</th>`).join('')}<th class="num">Valeur</th></tr></thead><tbody>
        ${scs.map(sc => `<tr><td>${esc(sc.nom)}</td>${ans.map(a => { const f = (sc.flux || []).find(x => x.annee === a); return `<td class="num">${f ? eur(f.valeur) : ''}</td>`; }).join('')}<td class="num"><b>${eur(dcf(sc.flux, m.wacc, m.g, m.anneeValo))}</b></td></tr>`).join('')}
        </tbody></table></div>`;
      }
      return `<div class="vmeth ${m.poids > 0 ? '' : 'off'}">
        <div class="vmeth-h"><b>${esc(m.nom)}</b><span>poids ${champ(d, `methodes.${i}.poids`, 'pct', 'Poids')}</span></div>
        ${m.source ? `<p class="vsrc">${esc(m.source)}</p>` : ''}${corps}</div>`;
    }).join('');
  }
  function droitsHtml(d, s){
    const e = d.entree || {};
    const lignes = BORNES.map(b => `<tr><td>${{bas: 'Bas', central: 'Central', haut: 'Haut'}[b]}</td><td class="num">${eur(s[b])}</td><td class="num">${eur(s.converti[b])}</td><td class="num"><b>${eur(s.ligne[b])}</b></td><td class="num">${fois(e.ticket ? s.ligne[b] / e.ticket : null)}</td></tr>`).join('');
    return `<p class="vnote">${e.preference ? `Hypothèse : actions de préférence 1x non participatives, au même rang que les autres préférences (total ${champ(d, 'entree.prefTotal', 'eur', 'Total des préférences')}). En cas de cession, GVI reçoit le plus élevé de sa préférence (${eur(e.ticket)}) et de sa quote-part (${pct(s.det, 2)}). ${esc(e.prefNote || '')}` : 'Aucune préférence de liquidation saisie : la ligne vaut la quote-part de GVI.'}</p>
      <div class="vtable-w"><table class="vtable"><thead><tr><th></th><th class="num">Valeur 100 %</th><th class="num">Quote-part GVI</th><th class="num">Ligne GVI (droits inclus)</th><th class="num">Multiple</th></tr></thead><tbody>${lignes}</tbody></table></div>`;
  }
  function sensiHtml(d){
    const mb = (d.methodes || []).find(m => m.type === 'bp'); if (!mb) return '';
    const v0 = ok(mb.valeurRef) ? mb.valeurRef : postMoney(d), det = detention(d), t = (d.entree || {}).ticket, a = atteinte(d);
    if (!ok(v0)) return '';
    const niveaux = [0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1, 1.1, 1.25, 1.5];
    const col = r => { const f = facteurBP(r, (mb.elasticite || {}).central, (mb.partHausse || {}).central), V = v0 * f * (1 + (mb.ajustMarche || 0)), L = produitGVI(d, V, det, (d.entree || {}).prefTotal); return {V, L, x: t ? L / t : null}; };
    const proche = ok(a.r) ? niveaux.reduce((p, n) => Math.abs(n - a.r) < Math.abs(p - a.r) ? n : p, niveaux[0]) : null;
    return `<div class="vtable-w"><table class="vtable vsensi"><thead><tr><th>Atteinte du BP</th>${niveaux.map(n => `<th class="num ${n === proche ? 'ici' : ''}">${pct(n)}</th>`).join('')}</tr></thead><tbody>
      <tr><td>Valeur 100 %</td>${niveaux.map(n => `<td class="num ${n === proche ? 'ici' : ''}">${eur(col(n).V)}</td>`).join('')}</tr>
      <tr><td>Ligne GVI</td>${niveaux.map(n => { const c = col(n); return `<td class="num ${c.x >= 1 ? 'up' : c.x >= 0.7 ? 'mid' : 'down'} ${n === proche ? 'ici' : ''}">${eur(c.L)}<small>${fois(c.x)}</small></td>`; }).join('')}</tr>
    </tbody></table></div>`;
  }
  function sortieHtml(d){
    const x = sortie(d), so = d.sortie || {};
    return `<p class="vnote">Dilution d'ici la sortie : ${champ(d, 'sortie.dilution', 'pct', 'Dilution')} (détention à la sortie ${pct(x.det, 2)}) ; préférences totales à la sortie ${champ(d, 'sortie.prefTotal', 'eur', 'Préférences à la sortie')}. ${esc(so.note || '')}</p>
    <div class="vtable-w"><table class="vtable"><thead><tr><th>Scénario</th><th class="num">Proba.</th><th class="num">Année</th><th class="num">ARR de sortie</th><th class="num">Multiple</th><th class="num">Valeur 100 %</th><th class="num">Produit GVI</th><th class="num">Multiple</th><th class="num">TRI</th></tr></thead><tbody>
    ${x.lignes.map((l, i) => `<tr><td>${esc(l.nom)}${l.hypothese ? `<small>${esc(l.hypothese)}</small>` : ''}</td>
      <td class="num">${champ(d, `sortie.scenarios.${i}.proba`, 'pct', 'Probabilité')}</td>
      <td class="num">${champ(d, `sortie.scenarios.${i}.annee`, 'an', 'Année')}</td>
      <td class="num">${ok(l.valeur) ? '—' : champ(d, `sortie.scenarios.${i}.arr`, 'eur', 'ARR de sortie')}</td>
      <td class="num">${ok(l.valeur) ? '—' : champ(d, `sortie.scenarios.${i}.multiple`, 'x', 'Multiple de sortie')}</td>
      <td class="num">${eur(l.ve)}</td><td class="num">${eur(l.prod)}</td><td class="num">${fois(l.moic)}</td><td class="num">${pct(l.tri)}</td></tr>`).join('')}
    <tr class="tot"><td>Espérance pondérée${Math.abs(x.sp - 1) > 0.001 ? ` <span class="vniv rouge">probabilités = ${pct(x.sp)}</span>` : ''}</td><td class="num">${pct(x.sp)}</td><td colspan="4"></td><td class="num">${eur(x.esp)}</td><td class="num">${fois(x.moicEsp)}</td><td class="num">${pct(x.triEsp)}</td></tr>
    </tbody></table></div>`;
  }
  function historiqueHtml(d, s){
    const h = (d.historique || []).slice().sort((a, b) => String(a.date).localeCompare(String(b.date)));
    const rows = h.map((x, i) => { const p = h[i - 1]; return `<tr><td class="mono">${esc(x.date)}</td><td class="num">${pct(x.atteinte)}</td><td class="num">${eur(x.valeur)}</td><td class="num">${eur(x.ligne)}</td><td class="num">${p && ok(p.ligne) && ok(x.ligne) ? (x.ligne >= p.ligne ? '+' : '−') + eur(Math.abs(x.ligne - p.ligne)) : '—'}</td><td>${esc(x.note || '')}</td></tr>`; }).join('');
    return `<div class="vtable-w"><table class="vtable"><thead><tr><th>Date</th><th class="num">Atteinte BP</th><th class="num">Valeur 100 %</th><th class="num">Ligne GVI</th><th class="num">Variation</th><th>Commentaire</th></tr></thead><tbody>
      ${rows || '<tr><td colspan="6">Aucune valorisation figée.</td></tr>'}
      <tr class="cour"><td class="mono">${esc(d.dateValo || 'en cours')}</td><td class="num">${pct(s.atteinte.r)}</td><td class="num">${eur(s.central)}</td><td class="num">${eur(s.ligne.central)}</td><td class="num">${h.length && ok(h[h.length - 1].ligne) && ok(s.ligne.central) ? (s.ligne.central >= h[h.length - 1].ligne ? '+' : '−') + eur(Math.abs(s.ligne.central - h[h.length - 1].ligne)) : '—'}</td><td>valeur en cours (non figée)</td></tr>
      </tbody></table></div>
      ${peutEcrire() ? `<div class="vactions"><button class="btn ghost" type="button" data-va="figer" ${VS.sale ? 'disabled title="Enregistrez d\'abord les hypothèses"' : ''}>Figer la valorisation au ${esc(d.dateValo || '…')}</button></div>` : ''}`;
  }
  function majHtml(){
    return `<form class="vnew" id="v-maj"><h3>Mettre à jour les chiffres</h3>
      <p class="vnote">Ajoute un point réalisé (ARR récurrent annualisé) et, si vous l'avez, la trésorerie. La valeur se recalcule ; enregistrez ensuite.</p>
      <div class="vnew-g">
        <label>Mois (AAAA-MM)<input name="date" required pattern="\\d{4}-\\d{2}" placeholder="2026-09"></label>
        <label>ARR réalisé (M€)<input name="arr" required type="number" step="any"></label>
        <label>Trésorerie (M€)<input name="cash" type="number" step="any"></label>
        <label>Trésorerie nette de dettes (M€)<input name="nette" type="number" step="any"></label>
        <label>Consommation mensuelle (M€, 0 si autofinancée)<input name="burn" type="number" step="any"></label>
      </div>
      <div class="vactions"><button class="btn" type="submit">Ajouter</button><button class="btn ghost" type="button" data-va="maj-non">Annuler</button></div></form>`;
  }
  const liste = (arr, cls = '') => (arr || []).length ? `<ul class="vlist ${cls}">${arr.map(p => `<li>${esc(p)}</li>`).join('')}</ul>` : '';
  function ficheHtml(d){
    const s = synthese(d), e = d.entree || {}, al = alertes(d);
    const actions = peutEcrire() ? `<div class="vactions">
        ${VS.sale ? '<span class="status">Modifications non enregistrées.</span>' : '<span class="status">Les cases grisées sont modifiables : tout se recalcule.</span>'}
        <button class="btn" type="button" data-va="save" ${VS.sale ? '' : 'disabled'}>Enregistrer</button>
        <button class="btn ghost" type="button" data-va="reset" ${VS.sale ? '' : 'disabled'}>Annuler les modifications</button>
        ${VS.msg ? `<span class="status ${VS.msgErr ? 'err' : ''}">${esc(VS.msg)}</span>` : ''}</div>` : '';
    return `<article class="vfiche">
      <div class="vf-top">
        <div class="mono">${esc(d.secteur || '')} · entrée ${esc(e.tour || '')} ${esc(e.date || '')}</div>
        <h2>${esc(d.nom)}</h2>
        <div class="vbadges"><span class="vbadge">valorisation au ${esc(d.dateValo || '—')}</span><span class="vbadge">${esc(d.statut || 'provisoire')}</span>${d.maj ? `<span class="vbadge">mise à jour ${esc(d.maj)}</span>` : ''}</div>
      </div>
      <div class="vbref">${enBref(d, s)}</div>
      <div class="vkpis small">
        <div><span>Atteinte du BP</span><b>${pct(s.atteinte.r)}</b><small>${libMois(s.atteinte.m)}</small></div>
        <div><span>Valeur à 100 %</span><b>${eur(s.central)}</b><small>${eur(s.bas)} – ${eur(s.haut)}</small></div>
        <div><span>Ticket GVI</span><b>${champ(d, 'entree.ticket', 'eur', 'Ticket GVI')}</b><small>détention ${pct(s.det, 2)}${e.ticketNote ? ' · ' + esc(e.ticketNote) : ''}</small></div>
        <div class="hl"><span>Valeur de la ligne GVI</span><b>${eur(s.ligne.central)}</b><small>${eur(s.ligne.bas)} – ${eur(s.ligne.haut)} · ${fois(s.tvpi)} le ticket</small></div>
      </div>
      <section class="vsec"><h3>Signaux</h3><ul class="valertes">${al.map(z => `<li>${pastille(z.niveau)}<b>${esc(z.libelle)}</b><span>${esc(z.detail)}</span></li>`).join('')}</ul></section>
      <section class="vsec"><h3>ARR réalisé contre business plan</h3>${graphe(d)}
        ${peutEcrire() && !VS.maj ? '<div><button class="btn ghost" type="button" data-va="maj">+ Mettre à jour les chiffres</button></div>' : ''}${VS.maj ? majHtml() : ''}</section>
      ${(d.lecture || []).length ? `<section class="vsec"><h3>Lecture</h3>${liste(d.lecture, 'lecture')}</section>` : ''}
      <section class="vsec"><h3>De la performance à la valeur</h3>${footballHtml(d, s)}</section>
      <section class="vsec"><h3>Ce que vaut la ligne de GVI</h3>${droitsHtml(d, s)}</section>
      <section class="vsec"><h3>Si la société atteint x % de son BP</h3><p class="vnote">Méthode du BP seule, paramètres centraux. Colonne encadrée : la situation actuelle.</p>${sensiHtml(d)}</section>
      <section class="vsec"><h3>Projection de sortie</h3>${sortieHtml(d)}</section>
      <section class="vsec"><h3>Historique des valorisations</h3>${historiqueHtml(d, s)}</section>
      <details class="vsec vdet"><summary>Détail des méthodes et hypothèses</summary>${methodesHtml(d, s)}</details>
      <details class="vsec vdet"><summary>Données à obtenir, limites, sources</summary>
        ${(d.aFournir || []).length ? `<h4>Données à obtenir</h4>${liste(d.aFournir)}` : ''}
        ${(d.limites || []).length ? `<h4>Limites</h4>${liste(d.limites)}` : ''}
        ${(d.sources || []).length ? `<h4>Sources</h4>${liste(d.sources, 'vsrcs')}` : ''}</details>
      ${actions}
    </article>`;
  }
  function nouvelleHtml(){
    return `<form class="vnew" id="v-new"><h3>Nouvelle participation</h3>
      <p class="vnote">Créez la fiche, puis ajoutez la trajectoire du BP et les chiffres réalisés (ou transmettez les documents à Claude).</p>
      <div class="vnew-g">
        <label>Nom affiché<input name="nom" required></label>
        <label>Secteur<input name="secteur"></label>
        <label>Date d'entrée (AAAA-MM)<input name="date" pattern="\\d{4}-\\d{2}" placeholder="2026-01"></label>
        <label>Tour<input name="tour" placeholder="Seed, Série A…"></label>
        <label>Pré-money (M€)<input name="preMoney" type="number" step="any"></label>
        <label>Montant levé (M€)<input name="leve" type="number" step="any"></label>
        <label>Ticket GVI (M€)<input name="ticket" type="number" step="any"></label>
      </div>
      <div class="vactions"><button class="btn" type="submit">Créer la fiche</button><button class="btn ghost" type="button" data-va="new-non">Annuler</button></div></form>`;
  }
  const METHODE = `<details class="vmethode"><summary>Notre méthode, en cinq points</summary><ol>
    <li><b>Le prix du dernier tour est le point de départ.</b> C'est la pratique des fonds de capital-risque (lignes directrices IPEV, suivies par France Invest).</li>
    <li><b>Il est ajusté du pourcentage du business plan atteint.</b> Le prix payé reposait sur une trajectoire de croissance ; si la société n'en réalise que 70 %, la valeur baisse en proportion (ou davantage si le marché sanctionne le ralentissement). Une surperformance n'est créditée qu'en partie tant qu'un nouveau tour ne la confirme pas.</li>
    <li><b>Le résultat est recoupé par les multiples du marché</b> (valeur / ARR de sociétés comparables), trésorerie nette incluse.</li>
    <li><b>On applique les droits de GVI.</b> Avec une préférence de liquidation, GVI récupère au moins sa mise tant que la société se vend au-dessus du total des préférences.</li>
    <li><b>Revue trimestrielle.</b> La valeur est figée à chaque date de reporting et les signaux (atteinte du BP, runway, fraîcheur des données) déclenchent une revue.</li></ol></details>`;
  function render(){
    if (!VS.charge) { root.innerHTML = `<p class="vnote vwait">${VDB ? 'Chargement des valorisations…' : 'Les valorisations sont conservées dans la base privée de la page : elles s\'affichent à l\'ouverture sur claude.ai.'}</p>`; return; }
    if (VS.id && !VS.docs.some(x => x.id === VS.id) && !VS.brouillon) VS.id = null;
    if (!VS.id && VS.docs.length) VS.id = VS.docs[0].id;
    const d = doc();
    root.innerHTML = `
      <div class="vintro"><p>${esc(opts.intro || 'Juste valeur indicative des participations et projection de sortie.')} <b>Aide à la décision, non auditée.</b></p>
        ${peutEcrire() && !VS.nouvelle ? '<button class="btn ghost" type="button" data-va="new">+ Ajouter une participation</button>' : ''}</div>
      ${METHODE}
      ${VS.nouvelle ? nouvelleHtml() : ''}
      ${VS.docs.length ? portefeuilleHtml() : '<p class="vnote">Aucune valorisation enregistrée.</p>'}
      ${d ? ficheHtml(d) : ''}
      <p class="foot">Montants en M€ ; « 100 % » = valeur de l'ensemble des titres, « ligne GVI » = ce qui revient à GVI. ARR = revenu récurrent annualisé. Les hypothèses proposées par Claude sont signalées et restent modifiables ; validation par l'équipe d'investissement avant tout usage.</p>`;
    brancherGraphe();
  }

  /* ---------- Interactions ---------- */
  async function enregistrer(d){
    d.maj = new Date().toISOString().slice(0, 10);
    await VDB.collection(COLL).doc(d.id).set(d);
  }
  root.addEventListener('click', async e => {
    const c = e.target.closest('[data-vid]');
    if (c) { choisir(c.dataset.vid); return; }
    const b = e.target.closest('[data-va]'); if (!b) return;
    const a = b.dataset.va;
    if (a === 'new') { VS.nouvelle = true; render(); }
    else if (a === 'new-non') { VS.nouvelle = false; render(); }
    else if (a === 'maj') { VS.maj = true; render(); }
    else if (a === 'maj-non') { VS.maj = false; render(); }
    else if (a === 'reset') { VS.brouillon = null; VS.sale = false; VS.msg = ''; render(); }
    else if (a === 'save' || a === 'figer') {
      const d = clone(VS.brouillon || VS.docs.find(x => x.id === VS.id));
      if (a === 'figer') {
        const s = synthese(d), h = (d.historique || []).filter(x => x.date !== d.dateValo);
        h.push({date: d.dateValo || new Date().toISOString().slice(0, 10), atteinte: s.atteinte.r ?? null, valeur: s.central, bas: s.bas, haut: s.haut, ligne: s.ligne.central, note: 'Valorisation figée'});
        d.historique = h;
      }
      b.disabled = true;
      try { await enregistrer(d); VS.brouillon = null; VS.sale = false; VS.msg = a === 'figer' ? 'Valorisation figée.' : 'Enregistré.'; VS.msgErr = false; }
      catch (_) { VS.msg = 'Enregistrement impossible : vos droits sur cette page ne le permettent pas.'; VS.msgErr = true; }
      render();
    }
  });
  root.addEventListener('change', e => {
    const i = e.target.closest('input.vin'); if (!i) return;
    if (!VS.brouillon) VS.brouillon = clone(VS.docs.find(x => x.id === VS.id));
    const raw = i.value.trim().replace(',', '.');
    let v = raw === '' ? null : Number(raw);
    if (v != null && !isFinite(v)) return;
    if (v != null && i.dataset.f === 'pct') v = v / 100;
    setP(VS.brouillon, i.dataset.p, v);
    VS.sale = true; VS.msg = '';
    const y = window.scrollY; render(); window.scrollTo(0, y);
  });
  root.addEventListener('submit', async e => {
    e.preventDefault();
    const f = Object.fromEntries(new FormData(e.target)), n = k => f[k] === '' || f[k] == null ? null : Number(f[k]);
    if (e.target.id === 'v-maj') {
      if (!VS.brouillon) VS.brouillon = clone(VS.docs.find(x => x.id === VS.id));
      const d = VS.brouillon;
      d.realise = (d.realise || []).filter(p => p.date !== f.date).concat([{date: f.date, arr: n('arr'), source: 'Saisie sur la page'}]);
      if (ok(n('cash'))) d.tresorerie = {...(d.tresorerie || {}), montant: n('cash'), date: f.date, ...(ok(n('nette')) ? {nette: n('nette')} : {}), ...(ok(n('burn')) ? {burn: n('burn')} : {})};
      const [y, m] = f.date.split('-').map(Number), fin = new Date(Date.UTC(y, m, 0)).toISOString().slice(0, 10);
      if (!d.dateValo || fin > d.dateValo) d.dateValo = fin;
      VS.maj = false; VS.sale = true; VS.msg = ''; render(); return;
    }
    if (e.target.id !== 'v-new') return;
    const id = String(f.nom).normalize('NFKD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'participation';
    const pre = n('preMoney'), lev = n('leve');
    const d = {id, nom: f.nom, secteur: f.secteur || '', statut: 'à compléter', creeLe: new Date().toISOString(), dateValo: '',
      entree: {date: f.date || '', tour: f.tour || '', preMoney: pre, leve: lev, ticket: n('ticket'), preference: true, prefMultiple: 1, prefTotal: lev},
      bp: {libelle: 'Business plan présenté au tour', decote: 0, points: []}, realise: [], prevision: [], tresorerie: {},
      methodes: [
        {type: 'bp', nom: 'Prix du dernier tour ajusté de l\'atteinte du BP', poids: 0.7, elasticite: {bas: 1.5, central: 1, haut: 0.75}, partHausse: {bas: 0, central: 0.5, haut: 1}, ajustMarche: 0, source: 'Méthode principale (IPEV : prix récent recalibré sur les jalons).'},
        {type: 'multiples', nom: 'Multiples d\'ARR des comparables', poids: 0.3, bas: null, central: null, haut: null, source: 'Multiples à renseigner (source à citer).'}],
      jalons: [], sortie: {annee: (parseInt(f.date) || new Date().getFullYear()) + 5, dilution: 0.25, prefTotal: lev, scenarios: [
        {nom: 'Échec', proba: 0.25, annee: (parseInt(f.date) || new Date().getFullYear()) + 3, valeur: 0}, {nom: 'Défavorable', proba: 0.3, arr: null, multiple: 3},
        {nom: 'Central', proba: 0.3, arr: null, multiple: 5}, {nom: 'Favorable', proba: 0.15, arr: null, multiple: 8}]},
      historique: [], lecture: [], aFournir: ['Trajectoire du BP (ARR par date)', 'ARR réalisé, trésorerie et consommation à date', 'Termes du pacte : préférences de liquidation'], sources: []};
    try { await enregistrer(d); VS.nouvelle = false; VS.id = id; VS.brouillon = null; VS.sale = false; }
    catch (_) { VS.msg = 'Création impossible : vos droits sur cette page ne le permettent pas.'; VS.msgErr = true; }
    render();
  });
  function choisir(id){
    if (VS.sale && id !== VS.id) { VS.msg = 'Enregistrez ou annulez les modifications avant de changer de participation.'; VS.msgErr = true; render(); return; }
    VS.id = id; VS.brouillon = null; VS.msg = ''; VS.maj = false; render();
    if (opts.onSelect) opts.onSelect(id);
    const f = root.querySelector('.vfiche'); if (f) f.scrollIntoView({behavior: 'smooth', block: 'start'});
  }

  /* ---------- Démarrage ---------- */
  if (opts.id) VS.id = opts.id;
  // Données publiées avec la page (démonstration) : affichées sans compte ni base ; la base, si elle
  // répond avec des fiches, prend le relais (modifications des éditeurs).
  const STATIQUES = Array.isArray(opts.donnees) ? opts.donnees.slice().sort((x, y) => String(x.nom).localeCompare(String(y.nom))) : [];
  if (STATIQUES.length) { VS.docs = STATIQUES; VS.charge = true; }
  render();
  if (window.claude && claude.use) {
    claude.use('user').then(u => { if (u && u.canEdit) Promise.resolve(u.canEdit()).then(v => { PEUT = !!v; if (VS.charge) render(); }).catch(() => {}); });
    claude.use('db').then(db => {
      if (!db) return;
      VDB = db; render();
      db.collection(COLL).orderBy('nom').onSnapshot(snap => {
        const docs = snap.docs.map(x => x.data()).filter(Boolean);
        VS.docs = docs.length ? docs : STATIQUES;
        VS.charge = true;
        if (!VS.sale) render();
      }, () => { VS.charge = true; render(); });
    });
  }
  return {render, choisir};
};
