/* Version SANS inscription Microsoft : tout est gardé sur le téléphone.
   - L'Excel (et le fichier de prospection) sont importés depuis l'app OneDrive / Fichiers, puis gardés ici (IndexedDB).
   - Les commandes et réglages de l'appli sont gardés sur ce téléphone (localStorage).
   - « Envoyer au PC » crée un fichier à enregistrer dans OneDrive > YOUSDAM-appli : l'agent du PC l'écrit dans l'Excel. */
"use strict";

const BASE = (() => {
  let p = null;
  const ouvrir = () => p || (p = new Promise((ok, ko) => {
    const r = indexedDB.open("yousdam", 1);
    r.onupgradeneeded = () => r.result.createObjectStore("fichiers");
    r.onsuccess = () => ok(r.result); r.onerror = () => ko(r.error);
  }));
  const op = (mode, fn) => ouvrir().then(db => new Promise((ok, ko) => { const t = db.transaction("fichiers", mode); const q = fn(t.objectStore("fichiers")); t.oncomplete = () => ok(q && q.result); t.onerror = () => ko(t.error); }));
  return {lire: k => op("readonly", s => s.get(k)), ecrire: (k, v) => op("readwrite", s => s.put(v, k)), effacer: k => op("readwrite", s => s.delete(k))};
})();

class StockLocal {
  async init() {}
  connecte() { return true; }
  compte() { return "Sur ce téléphone"; }
  prenom() { return localStorage.getItem("prenom") || ""; }
  connexion() {} deconnexion() {}
  cleDe(chemin) { return chemin === CHEMINS.excel ? "excel" : chemin === F_PROSP ? "prosp" : null; }
  async infos(chemin) {
    if (chemin === F_APPLI) return {etag: localStorage.getItem("appli_v") || "0"};
    if (chemin === F_FAIT) return {etag: "0"};
    const k = this.cleDe(chemin); if (!k) return null;
    const m = await BASE.lire(k + "_meta"); return m ? {etag: m.importe, modifie: m.importe, nom: m.nom} : null;
  }
  async binaire(chemin) { const k = this.cleDe(chemin); return k ? (await BASE.lire(k)) || null : null; }
  async json(chemin) {
    if (chemin === F_APPLI) return {data: JSON.parse(localStorage.getItem("appli") || "null"), etag: localStorage.getItem("appli_v") || "0"};
    return {data: {faits: {}, erreurs: {}}, etag: "0"};
  }
  async ecrireJSON(chemin, data) {
    if (chemin !== F_APPLI) return "0";
    localStorage.setItem("appli", JSON.stringify(data)); const v = String(Date.now()); localStorage.setItem("appli_v", v); return v;
  }
  async urlPhoto(nom) {
    const b = await BASE.lire("photo:" + nom.toLowerCase());
    if (!b) throw new Error(`Photo « ${nom} » pas encore importée. Réglages > Photos des règlements > Importer.`);
    return URL.createObjectURL(b);
  }
  async importerPhotos(fichiers) {          // photos choisies dans OneDrive > REGLEMENTS (plusieurs d'un coup)
    let n = 0;
    for (const f of fichiers) { await BASE.ecrire("photo:" + f.name.toLowerCase(), f); n++; }
    const l = new Set(JSON.parse(localStorage.getItem("photos") || "[]")); [...fichiers].forEach(f => l.add(f.name.toLowerCase()));
    localStorage.setItem("photos", JSON.stringify([...l]));
    return n;
  }
  /* « donnees-yousdam.json », préparé par l'agent du PC : l'Excel + les photos de ses liens hypertexte + la prospection */
  async importerDonnees(fichier) {
    const d = JSON.parse(await fichier.text());
    if (!d || d.type !== "yousdam-donnees" || !d.excel) throw new Error("Ce n'est pas le fichier « donnees-yousdam.json » de l'appli.");
    const versBuf = b64 => Uint8Array.from(atob(b64), c => c.charCodeAt(0)).buffer;
    const xl = versBuf(d.excel); lireClasseur(xl);
    await BASE.ecrire("excel", xl);
    await BASE.ecrire("excel_meta", {nom: d.excel_nom || "Excel", importe: d.cree || new Date().toISOString()});
    if (d.prospection) { const p = versBuf(d.prospection); await BASE.ecrire("prosp", p); await BASE.ecrire("prosp_meta", {nom: "Prospection", importe: d.cree}); }
    localStorage.setItem("couleurs", JSON.stringify(d.couleurs || {}));   // couleurs des clients dans l'Excel (aperçu)
    const noms = [];
    for (const [nom, ph] of Object.entries(d.photos || {})) { await BASE.ecrire("photo:" + nom, new Blob([versBuf(ph.b64)], {type: ph.type})); noms.push(nom); }
    const l = new Set(JSON.parse(localStorage.getItem("photos") || "[]")); noms.forEach(n => l.add(n));
    localStorage.setItem("photos", JSON.stringify([...l]));
    return {photos: noms.length};
  }
  aPhoto(nom) { return JSON.parse(localStorage.getItem("photos") || "[]").includes(String(nom).toLowerCase()); }
  async importer(type, fichier) {
    if (/\.json$/i.test(fichier.name)) return this.importerDonnees(fichier);
    const buf = await fichier.arrayBuffer();
    if (type === "excel") lireClasseur(buf);            // vérifie que c'est bien le bon fichier avant de le garder
    else lireProspects(buf);
    await BASE.ecrire(type, buf);
    await BASE.ecrire(type + "_meta", {nom: fichier.name, importe: new Date().toISOString()});
  }
}

/* l'appli locale n'a pas d'agent qui lui répond : une livraison envoyée au PC est considérée comme écrite
   dès qu'on la retrouve dans l'Excel réimporté (même client, palettes, client final ou chantier, ± 3 jours) */
async function relierEnvoisRetrouves() {
  const v0 = CACHE.classeur.ventes, cmds = CACHE.appli.commandes;
  const prises = new Set(cmds.map(c => c.excel_row).filter(Boolean));
  const liens = [];
  for (const c of cmds) {
    if (c.statut !== "livree" || !c.excel_demande || c.excel_row) continue;
    const L = c.livraison, m = dansExcel(v0, prises, c.client, L.palettes, L.client_final, L.chantier, L.date, -3, 3);
    if (m) { liens.push([c.id, m.row]); prises.add(m.row); }
  }
  if (liens.length) await majAppli(d => liens.forEach(([id, row]) => { const c = d.commandes.find(x => x.id === id); if (c) { c.excel_row = row; delete c.excel_demande; } }));
  return liens.length;
}

/* fichier d'envoi vers le PC (à enregistrer dans OneDrive > YOUSDAM-appli) */
function fichierEnvoi() {
  const a_envoyer = CACHE.appli.commandes.filter(c => c.statut === "livree" && c.excel_demande && !c.excel_row && !c.excel_envoye);
  if (!a_envoyer.length) return null;
  const d = new Date(), t = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, "0")}${String(d.getDate()).padStart(2, "0")}-${String(d.getHours()).padStart(2, "0")}${String(d.getMinutes()).padStart(2, "0")}${String(d.getSeconds()).padStart(2, "0")}`;
  const contenu = JSON.stringify({type: "yousdam-envoi", cree: d.toISOString(), commandes: a_envoyer}, null, 1);
  return {ids: a_envoyer.map(c => c.id), nb: a_envoyer.length, fichier: new File([contenu], `envoi_${t}.json`, {type: "application/json"})};
}
