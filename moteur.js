/* Moteur de l'appli YOUSDAM, qui tourne entièrement dans le téléphone.
   Même logique que data.py / whatsapp.py / server.py (version PC), traduite en JavaScript.
   Lecture seule de l'Excel. Les livraisons « Passer à Excel » vont dans une file d'attente,
   écrite dans l'Excel par l'agent du PC dès qu'il est allumé. */
"use strict";

const PAR_PALETTE = {B7: 256, B10: 180, B12: 128, B50: 56};
const CASA = new Set(["casa", "casablanca", "mohamadia", "mohammedia", "mediouna", "had soualem", "dar bouazza", "bouskoura", "berrechid", "bouznika", "nouaceur", "tit mellil"]);

/* ---------- petits outils ---------- */
const sStr = v => v == null ? "" : String(v).trim();
const sNum = v => { const n = parseFloat(v); return isFinite(n) ? n : 0; };
const cle = n => sStr(n).toLowerCase().split(/\s+/).filter(Boolean).join(" ");
const isoDate = d => {                                   // Date (SheetJS) -> "AAAA-MM-JJ" ; sinon null
  if (!(d instanceof Date) || isNaN(d)) return null;
  const z = new Date(d.getTime() + 12 * 3600e3);          // évite les décalages de fuseau autour de minuit
  return `${z.getUTCFullYear()}-${String(z.getUTCMonth() + 1).padStart(2, "0")}-${String(z.getUTCDate()).padStart(2, "0")}`;
};
const jourUTC = s => Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10));
const ecartJours = (a, b) => Math.round((jourUTC(a) - jourUTC(b)) / 864e5);     // a - b en jours
const ajouteJours = (s, n) => { const d = new Date(jourUTC(s) + n * 864e5); return d.toISOString().slice(0, 10); };
const moisDe = s => s.slice(0, 7);
const aujourdhui = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`; };
const f0 = x => Math.round(x).toLocaleString("fr-FR").replace(/[  ]/g, " ");
const pc = x => Math.round(x * 100) + " %";
const dec3 = x => x.toFixed(3).replace(".", ",");

/* ---------- lecture de l'Excel (SheetJS) ---------- */
function cellule(ws, ref) { const c = ws[ref]; return c ? (c.t === "d" ? c.v : c.v) : null; }
function lireClasseur(buf) {
  const wb = XLSX.read(buf, {type: "array", cellDates: true, cellFormula: false, cellHTML: false, cellText: false});
  const p = wb.Sheets["PRINCIPAL"];
  if (!p) throw new Error("Onglet PRINCIPAL introuvable dans l'Excel.");
  const dern = XLSX.utils.decode_range(p["!ref"]).e.r + 1;
  const g = (c, r) => cellule(p, c + r);
  const ventes = [];
  for (let r = 4; r <= dern; r++) {
    const client = g("B", r), date = isoDate(g("A", r));
    if (!client || !date) continue;
    ventes.push({row: r, date, client: sStr(client), brut: String(client), ck: cle(client), palettes: sNum(g("C", r)), type: sStr(g("D", r)).toUpperCase(),
      qte: sNum(g("E", r)), pu: sNum(g("F", r)), tu: sNum(g("G", r)), pv: sNum(g("H", r)), ca: sNum(g("I", r)), ville: sStr(g("J", r)),
      chantier: sStr(g("K", r)), client_final: sStr(g("L", r)), decharge: sStr(g("M", r)), mdech: sNum(g("N", r)), benef: sNum(g("O", r)), paiement: sStr(g("P", r))});
  }
  const villes = [];
  for (let r = 4; r < 60; r++) { const v = g("AB", r); if (v == null) break; villes.push({ville: sStr(v), decharge: sNum(g("AC", r))}); }
  const solde = {initial: sNum(g("B", 1)), consomme: sNum(g("F", 1)), courant: sNum(g("Q", 4)), encaisse: sNum(g("R", 4))};
  // espèces tirées : le tableau « Tiré | Source » de l'onglet PRINCIPAL (trouvé par son titre, où qu'il soit)
  let sommeO = 0; for (let r = 4; r <= dern; r++) { const o = g("O", r); if (typeof o === "number") sommeO += o; }
  const especes = [];
  const R = XLSX.utils.decode_range(p["!ref"]);
  trouve: for (let r = 1; r <= dern; r++) for (let c = 0; c <= Math.min(R.e.c, 40); c++) {
    const ref = XLSX.utils.encode_col(c);
    if (sStr(g(ref, r)).toLowerCase() === "tiré" || sStr(g(ref, r)).toLowerCase() === "tire") {
      const src = XLSX.utils.encode_col(c + 1);
      for (let k = r + 1; k <= Math.min(dern, r + 60); k++) { const m = g(ref, k); if (typeof m === "number" && m > 0) especes.push({row: k, montant: m, source: sStr(g(src, k))}); }
      break trouve;
    }
  }
  const rg = wb.Sheets["REGLEMENT"], regs = [];
  if (rg) {
    const fin = XLSX.utils.decode_range(rg["!ref"]).e.r + 1;
    for (let r = 2; r <= fin; r++) {
      const client = cellule(rg, "A" + r);
      if (!client || cle(client) === "client") continue;
      const e = rg["E" + r];
      let photo = "";
      if (e && e.l && e.l.Target && !/^https?:/i.test(e.l.Target)) photo = decodeURIComponent(e.l.Target).replace(/\\/g, "/").split("/").pop();
      const rec = cellule(rg, "B" + r);
      regs.push({row: r, client: sStr(client), ck: cle(client), recu: isoDate(rec) || sStr(rec), periode: sStr(cellule(rg, "C" + r)), depot: isoDate(cellule(rg, "D" + r)),
        ref: sStr(cellule(rg, "E" + r)), echeance: isoDate(cellule(rg, "F" + r)), montant: sNum(cellule(rg, "G" + r)), remarque: sStr(cellule(rg, "H" + r)), photo});
    }
  }
  return {ventes, villes, solde, regs, especes, sommeO};
}

function lireProspects(buf) {
  if (!buf) return [];
  const wb = XLSX.read(buf, {type: "array", cellDates: true});
  const top = {};
  const t = wb.Sheets["Casa - à rencontrer"];
  if (t) { const fin = XLSX.utils.decode_range(t["!ref"]).e.r + 1; for (let r = 2; r <= fin; r++) { const n = cellule(t, "B" + r); if (n) top[n] = [cellule(t, "A" + r), sStr(cellule(t, "C" + r))]; } }
  const out = [];
  for (const nom of wb.SheetNames) {
    if (nom === "Mode d'emploi" || nom === "Casa - à rencontrer") continue;
    const ws = wb.Sheets[nom], fin = XLSX.utils.decode_range(ws["!ref"]).e.r + 1;
    for (let r = 2; r <= fin; r++) {
      const n = cellule(ws, "B" + r); if (!n) continue;
      const tp = top[n];
      out.push({feuille: nom, row: r, prio: sStr(cellule(ws, "A" + r)), nom: sStr(n), zone: sStr(cellule(ws, "C" + r)), tel: sStr(cellule(ws, "D" + r)),
        remarque: sStr(cellule(ws, "H" + r)), statut: sStr(cellule(ws, "I" + r)) || "À appeler", date: isoDate(cellule(ws, "J" + r)), notes: sStr(cellule(ws, "K" + r)),
        rang: tp ? tp[0] : null, pourquoi: tp ? tp[1] : ""});
    }
  }
  return out;
}

/* ---------- prix, doublons ---------- */
function zoneTarif(ville) {
  const vk = cle(ville).replace(/é/g, "e");
  if (CASA.has(vk)) return "Casa";
  for (const [k, z] of [["jadida", "El Jadida"], ["marrakech", "Marrakech"], ["benguerir", "Benguerir"], ["ben guerir", "Benguerir"], ["beni mellal", "Beni Mellal"],
    ["bouznika", "Bouznika"], ["rabat", "Rabat"], ["kenitra", "Kenitra"], ["tetouan", "Tetouan"]]) if (vk.includes(k)) return z;
  return null;
}
function prixDefaut(ventes, client, typ, ville, tarifs) {
  const ck = cle(client), vk = cle(ville), T = tarifs || {};
  const v = ventes.slice().sort((a, b) => a.date.localeCompare(b.date) || a.row - b.row).reverse();
  const histo = (v.find(x => x.type === typ && cle(x.ville) === vk && x.pu) || {}).pu;
  let pu = (T[zoneTarif(ville)] || {})[typ] || histo || (v.find(x => x.type === typ && x.pu) || {}).pu || 0;
  if (!pu) return {pu: "", pv: ""};
  const z = zoneTarif(ville), meme = x => cle(x.ville) === vk || (z && zoneTarif(x.ville) === z);
  let pv = (v.find(x => x.ck === ck && x.type === typ && meme(x) && x.pv) || v.find(x => x.ck === ck && x.type === typ && x.pv) || {}).pv || 0;
  pv = Math.max(pv, Math.round((pu + 0.10) * 100) / 100);
  return {pu, pv};
}
function doublon(ventes, client, palettes, client_final, jour) {
  const ck = cle(client), cf = cle(client_final);
  const x = ventes.find(x => x.ck === ck && x.palettes === +palettes && (!cf || cle(x.client_final) === cf) && Math.abs(ecartJours(x.date, jour)) <= 10);
  return x ? {date: x.date, row: x.row, chantier: x.chantier, client_final: x.client_final} : null;
}
function dansExcel(ventes, prises, client, palettes, client_final, chantier, jour, avant, apres) {
  if (!jour) return null;
  const ck = cle(client), cf = cle(client_final || ""), ch = cle(chantier || "");
  if (!cf && !ch) return null;
  for (const v of ventes.slice().sort((a, b) => a.date.localeCompare(b.date))) {
    if (prises.has(v.row) || v.ck !== ck || v.palettes !== +palettes) continue;
    const e = ecartJours(v.date, String(jour).slice(0, 10));
    if (e < avant || e > apres) continue;
    if ((cf && cle(v.client_final) === cf) || (ch && cle(v.chantier) === ch)) return {row: v.row, date: v.date, pv: v.pv, client_final: v.client_final, chantier: v.chantier};
  }
  return null;
}
function situation(ventes, client) {
  const ck = cle(client);
  return ventes.filter(v => v.ck === ck).sort((a, b) => a.date.localeCompare(b.date) || (a.row || 1e9) - (b.row || 1e9));
}

/* ---------- analyse (bilan, points forts / faibles) ---------- */
function analyse(ventes, attente, regs, solde, objectif, today) {
  today = today || aujourdhui();
  const vide = () => ({benef: 0, ca: 0, palettes: 0, livraisons: 0, briques: 0});
  const mois = new Map();
  for (const v of ventes.slice().sort((a, b) => a.date.localeCompare(b.date))) {
    const k = moisDe(v.date); if (!mois.has(k)) mois.set(k, vide());
    const m = mois.get(k); m.benef += v.benef; m.ca += v.ca; m.palettes += v.palettes; m.livraisons += 1; m.briques += v.qte;
  }
  const cur = moisDe(today);
  const d0 = new Date(jourUTC(today)); const prem = Date.UTC(d0.getUTCFullYear(), d0.getUTCMonth(), 1);
  const prev = new Date(prem - 864e5).toISOString().slice(0, 7);
  const an = `${d0.getUTCFullYear() - 1}-${String(d0.getUTCMonth() + 1).padStart(2, "0")}`;
  const mc = mois.get(cur) || vide();
  const nbJours = new Date(Date.UTC(d0.getUTCFullYear(), d0.getUTCMonth() + 1, 0)).getUTCDate();
  const jour = d0.getUTCDate();
  const projection = jour ? mc.benef / jour * nbJours : 0;

  const lundi = ajouteJours(today, -((d0.getUTCDay() + 6) % 7));
  const somme = (a, b) => ventes.filter(v => v.date >= a && v.date < b).reduce((s, v) => s + v.benef, 0);
  const sem = ventes.filter(v => v.date >= lundi && v.date <= today).reduce((s, v) => s + v.benef, 0);
  const semPrev = somme(ajouteJours(lundi, -7), lundi);
  const semaines = []; for (let k = 9; k >= 1; k--) semaines.push(somme(ajouteJours(lundi, -7 * k), ajouteJours(lundi, -7 * (k - 1)))); semaines.push(sem);

  const clients = new Map();
  for (const v of ventes) {
    if (!clients.has(v.ck)) clients.set(v.ck, {nom: v.client, benef: 0, ca: 0, palettes: 0, livraisons: 0, briques: 0, derniere: v.date, en_cours: 0, retard: 0, mois: {}});
    const c = clients.get(v.ck);
    c.benef += v.benef; c.ca += v.ca; c.palettes += v.palettes; c.livraisons += 1; c.briques += v.qte;
    if (v.date > c.derniere) c.derniere = v.date;
    c.mois[moisDe(v.date)] = (c.mois[moisDe(v.date)] || 0) + v.benef;
    const p = v.paiement.toLowerCase();
    if (p.includes("retard")) c.retard += v.ca; else if (p.includes("cours")) c.en_cours += v.ca;
  }
  const tot = [...clients.values()].reduce((s, c) => s + c.benef, 0) || 1;
  for (const c of clients.values()) { c.part = c.benef / tot; c.marge = c.briques ? c.benef / c.briques : 0; c.jours = ecartJours(today, c.derniere); }
  const cl = [...clients.values()].sort((a, b) => b.benef - a.benef);

  const marge = sel => { const b = sel.reduce((s, v) => s + v.qte, 0); return b ? sel.reduce((s, v) => s + v.benef, 0) / b : 0; };
  const mMois = marge(ventes.filter(v => moisDe(v.date) === cur));
  const trois = [1, 2, 3].map(i => new Date(prem - 30 * i * 864e5).toISOString().slice(0, 7));
  const mAvant = marge(ventes.filter(v => trois.includes(moisDe(v.date))));
  const an12 = ventes.filter(v => ecartJours(today, v.date) <= 365), b12 = an12.reduce((s, v) => s + v.benef, 0) || 1;
  const partMk = an12.filter(v => cle(v.ville) === "marrakech").reduce((s, v) => s + v.benef, 0) / b12;
  const partB10 = an12.filter(v => v.type === "B10").reduce((s, v) => s + v.benef, 0) / b12;
  const impayes = regs.filter(r => /impay/i.test(r.remarque));
  const old = attente.filter(c => c.date && ecartJours(today, c.date.slice(0, 10)) > 14);

  const diag = [];
  const D = (niveau, titre, chiffre, detail, action, cible) => diag.push({niveau, titre, chiffre: chiffre.replace("%", " %").replace("  ", " "), detail, action: action || null, cible: cible || null});
  if (mc.benef >= objectif) D("fort", "Objectif du mois atteint", `${f0(mc.benef)} DH`, `Tu as dépassé ton objectif de ${f0(objectif)} DH.`);
  else if (projection >= objectif) D("fort", "En route vers l'objectif", `${f0(projection)} DH`, "À ce rythme, tu finis le mois au-dessus de ton objectif.");
  else D("grave", "Objectif du mois en retard", `−${f0(objectif - projection)} DH`,
    `Au rythme actuel, tu finis le mois à ${f0(projection)} DH sur ${f0(objectif)}. Chaque camion de B7 (52 palettes) te rapporte environ ${f0(13312 * (mMois || mAvant))} DH à ta marge actuelle.`, "Appeler mes prospects", "prospects");
  if (cl.length && cl[0].part > 0.6) D("grave", `Trop dépendant de ${cl[0].nom.trim()}`, pc(cl[0].part),
    "de ton bénéfice vient d'un seul client. S'il s'arrête, tu perds presque tout. Vise moins de 60 % avec 1 ou 2 nouveaux clients moyens.", "Trouver un nouveau client", "prospects");
  else if (cl.length) D("fort", "Clients bien répartis", pc(cl[0].part), "pour ton plus gros client : ton risque est réparti.");
  if (mAvant && mMois && mMois < mAvant * 0.95) D("faible", "Ta marge par brique baisse", `${dec3(mMois)} DH`,
    `contre ${dec3(mAvant)} DH les 3 mois d'avant. L'usine a monté ses prix sans que ta marge suive. Négocie le prix avance et vérifie tes tarifs.`, "Vérifier mes tarifs", "reglages");
  else if (mAvant && mMois > mAvant * 1.05) D("fort", "Ta marge par brique monte", `${dec3(mMois)} DH`, `contre ${dec3(mAvant)} DH les 3 mois d'avant.`);
  if (semPrev && sem < semPrev * 0.7) D("faible", "Semaine plus calme", `−${pc(1 - sem / semPrev)}`, `${f0(sem)} DH cette semaine contre ${f0(semPrev)} DH la semaine dernière.`, "Voir les commandes", "commandes");
  else if (sem && sem >= semPrev) D("fort", "Bonne semaine", `${f0(sem)} DH`, `Mieux que la semaine dernière (${f0(semPrev)} DH).`);
  const inactifs = cl.filter(c => c.livraisons >= 3 && c.jours > 30);
  if (inactifs.length) D("faible", `${inactifs.length} client${inactifs.length > 1 ? "s" : ""} ne commande${inactifs.length > 1 ? "nt" : ""} plus`, String(inactifs.length),
    "Sans commande depuis plus d'un mois : " + inactifs.slice(0, 4).map(c => `${c.nom.trim()} (${c.jours} j)`).join(", ") + ". Appelle-les pour savoir s'ils ont un chantier.", "Voir mes clients", "clients");
  if (impayes.length) D("grave", `${impayes.length} règlement${impayes.length > 1 ? "s" : ""} impayé${impayes.length > 1 ? "s" : ""}`, `${f0(impayes.reduce((s, r) => s + r.montant, 0))} DH`,
    "C'est toi qui portes le risque auprès de l'usine. Relance le client et demande un règlement certifié.", "Voir les règlements", "reglements");
  if (partB10 < 0.1) D("faible", "Le B10 est sous-exploité", pc(partB10), "de ton bénéfice vient du B10, alors que c'était ta meilleure marge par brique. Propose-le quand le chantier le permet.");
  const ma = mois.get(an);
  if (ma && ma.benef && mc.benef >= ma.benef) D("fort", "Mieux que l'an dernier", `+${pc((mc.benef - ma.benef) / ma.benef)}`, `par rapport au même mois l'an dernier (${f0(ma.benef)} DH).`);
  if (old.length) D("faible", "Commandes qui attendent", String(old.length), "en attente depuis plus de 2 semaines : livre-les ou annule-les.", "Voir les commandes", "commandes");
  diag.sort((a, b) => ({grave: 0, faible: 1, fort: 2}[a.niveau] - {grave: 0, faible: 1, fort: 2}[b.niveau]));

  return {diag, today, objectif, mois_cle: cur, mois_courant: mc, mois_prec: mois.get(prev) || null, mois_an: ma || null, projection,
    semaine: sem, semaine_prec: semPrev, semaines, mois: [...mois.entries()].map(([k, v]) => ({mois: k, ...v})).slice(-13),
    marge_mois: mMois, marge_avant: mAvant, part_marrakech: partMk, part_b10: partB10, clients: cl, forts: [], faibles: [], conseils: [], solde,
    nb_attente: attente.length, impayes: impayes.length};
}

/* ---------- lecture d'un message WhatsApp ---------- */
const UNITES = {solo: 24, remorque: 48, semi: 48, camion: 48, plateau: 48};
// plateau : 44, 48 ou 52 palettes ; avec grue : 44 ou 48 seulement
const CHOIX_PLATEAU = [44, 48, 52], CHOIX_GRUE = [44, 48];
const VILLES_WA = [["jadida", "Jadida"], ["azemmour", "Azemour"], ["azemour", "Azemour"], ["mohammedia", "Mohamadia"], ["mohamadia", "Mohamadia"], ["mohammadia", "Mohamadia"],
  ["mansouria", "Mohamadia"], ["marrakech", "Marrakech"], ["kech", "Marrakech"], ["benguerir", "Benguerir"], ["bouznika", "Bouznika"], ["settat", "Settat"],
  ["khouribga", "Khouribga"], ["berrechid", "Berrechid"], ["mediouna", "Mediouna"], ["soualem", "Had Soualem"], ["temara", "Temara"], ["rabat", "Rabat"],
  ["sale", "Salé"], ["kenitra", "Kenitra"], ["beni mellal", "Beni Mellal"]];
const sansAccent = s => s.normalize("NFKD").replace(/[̀-ͯ]/g, "").toLowerCase();
function villeDe(chantier) { const p = sansAccent(chantier || ""); for (const [k, v] of VILLES_WA) if (new RegExp(`\\b${k}\\b`).test(p)) return v; return "Casa"; }
function lireBloc(txt) {
  const p = sansAccent(txt), o = {palettes: null, unite: "", a_confirmer: false, type: "B7", chantier: "", client_final: "", contact: "", brut: txt.trim()};
  const p1 = p.replace(/\b(?:un|une)\s+(?=(?:solo|remorque|semi|camion|plateau))/g, "1 ");
  let m = p1.match(/(\d+)\s*(solos?|remorques?|semis?|camions?|plateaux|plateaus?)\b/);
  if (m) { const u = m[2].replace(/(x|s)$/, ""); o.unite = u; o.palettes = +m[1] * UNITES[u]; o.a_confirmer = u !== "solo"; }
  else if (/\bplateau/.test(p)) { o.unite = "plateau"; o.palettes = 48; o.a_confirmer = true; }
  o.grue = /\bgrue\b/.test(p);
  if (o.unite !== "solo" && o.a_confirmer) o.choix = o.grue ? CHOIX_GRUE : CHOIX_PLATEAU;
  else if ((m = p.match(/(\d+)\s*pal/))) o.palettes = +m[1];
  if ((m = p.match(/\b(?:brs?|briques?|b)\s*\.?\s*(7|10|12|15|50)\b/))) o.type = {7: "B7", 10: "B10", 12: "B12", 15: "B12", 50: "B50"}[m[1]];
  if ((m = txt.match(/chantier\s*:?\s*([^\n(]*)(?:\(([^)]*)\))?/i))) { o.chantier = m[1].replace(/^[\s:-]+|[\s:-]+$/g, ""); o.client_final = (m[2] || "").trim(); }
  if ((m = txt.match(/(?:contact|t[eé]l)\s*:?\s*([+\d][\d .\-]{5,})/i))) o.contact = m[1].replace(/\s+/g, " ").trim();
  const maj1 = t => t ? t.charAt(0).toUpperCase() + t.slice(1) : t; o.chantier = maj1(o.chantier); o.client_final = maj1(o.client_final);  // « tit melil » -> « Tit melil », comme dans ton Excel
  o.ville = villeDe(o.chantier); o.quantite = (o.palettes || 0) * PAR_PALETTE[o.type];
  return o;
}
function lireMessageWA(message) {
  const txt = message.replace(/[‎‏‪-‮⁦-⁩\xa0]/g, " ");
  return txt.split(/(?=merci de (?:me |nous )?livrer)/i).filter(b => /livrer|chantier|solo|remorque|plateau/i.test(b)).map(lireBloc).filter(c => c.palettes || c.chantier);
}

/* ---------- vérification rapide (window.testMoteur()) ---------- */
function testMoteur() {
  const [a, b] = lireMessageWA("Bonjour\nMerci de me livrer\n1SOLO de  brs 7\nChantier quartier nord  (Y )\nContact :0600000000\n\nMerci de me livrer\n1 remorque brs 7\nChantier :  el jadida  (X )\nContact ‪+212 600 00 00 00‬");
  console.assert(a.palettes === 24 && a.type === "B7" && a.chantier === "Quartier nord" && a.client_final === "Y" && a.ville === "Casa" && a.quantite === 6144 && !a.a_confirmer, "wa 1", a);
  console.assert(b.palettes === 48 && b.ville === "Jadida" && b.a_confirmer && b.contact.startsWith("+212"), "wa 2", b);
  const [c] = lireMessageWA("Merci de me livrer un plateau avec grue brs 7\nChantier : bouskoura (Z)");
  console.assert(c.palettes === 48 && c.grue && c.choix.join() === "44,48", "wa plateau grue", c);
  const [e] = lireMessageWA("Merci de me livrer 1 plateau brs 10\nChantier : had soualem (Z)");
  console.assert(e.palettes === 48 && !e.grue && e.choix.join() === "44,48,52" && e.type === "B10", "wa plateau", e);
  const V = Array.from({length: 40}, (_, i) => ({row: i, date: ajouteJours("2026-11-24", -3 * i), client: "A", ck: "a", palettes: 24, type: "B7", qte: 6144, pu: 2.00, tu: 0, pv: 2.10,
    ca: 12902.4, ville: "Casa", chantier: "x", client_final: "y", decharge: "Non", mdech: 0, benef: 614.4, paiement: "Payé"}));
  const A = analyse(V, [], [], {}, 30000, "2026-11-24");
  console.assert(A.mois_courant.livraisons === 8 && Math.abs(A.mois_courant.benef - 8 * 614.4) < 1e-6, "analyse", A.mois_courant);
  console.assert(Math.abs(A.marge_mois - 0.1) < 1e-9, "marge");
  const p = prixDefaut(V, "A", "B7", "Casa"); console.assert(p.pu === 2.00 && p.pv === 2.10, "prix", p);
  console.assert(doublon(V, "A", 24, "y", "2026-11-24").row === 0 && doublon(V, "A", 48, "y", "2026-11-24") === null, "doublon");
  return "moteur ok";
}
