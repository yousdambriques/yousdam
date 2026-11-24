/* Lien avec OneDrive + « serveur » local : remplace server.py, tout tourne dans le téléphone.
   Les données ne passent que par ton compte Microsoft (OneDrive). */
"use strict";

/* ---------- réglages ---------- */
const CHEMINS = {
  excel: localStorage.getItem("fichier_excel") || "",      // choisi au premier lancement, gardé sur le téléphone
  reglements: "REGLEMENTS",                                 // dossier des photos
  dossier: "YOUSDAM-appli",                                 // fichiers de l'appli
};
const F_APPLI = CHEMINS.dossier + "/appli.json";            // écrit par les téléphones
const F_FAIT = CHEMINS.dossier + "/excel_fait.json";        // écrit par l'agent du PC
const F_PROSP = CHEMINS.dossier + "/Prospection_Promoteurs.xlsx";
const TARIFS_DEFAUT = {};                                   // tes tarifs sont dans ton OneDrive (appli.json), jamais dans le code

/* ---------- stockage : OneDrive (Microsoft Graph) ---------- */
const GRAPH = "https://graph.microsoft.com/v1.0/me/drive/root:/";
const enc = p => p.split("/").map(encodeURIComponent).join("/");
class StockOneDrive {
  constructor(clientId) {
    this.msal = new msal.PublicClientApplication({auth: {clientId, authority: "https://login.microsoftonline.com/consumers", redirectUri: location.origin + location.pathname},
      cache: {cacheLocation: "localStorage"}});
    this.scopes = ["Files.ReadWrite", "User.Read"];
  }
  async init() {
    await this.msal.initialize();
    const r = await this.msal.handleRedirectPromise();
    if (r && r.account) this.msal.setActiveAccount(r.account);
    if (!this.msal.getActiveAccount() && this.msal.getAllAccounts().length) this.msal.setActiveAccount(this.msal.getAllAccounts()[0]);
  }
  connecte() { return !!this.msal.getActiveAccount(); }
  compte() { const a = this.msal.getActiveAccount(); return a ? a.username : ""; }
  connexion() { return this.msal.loginRedirect({scopes: this.scopes, prompt: "select_account"}); }
  deconnexion() { return this.msal.logoutRedirect(); }
  async jeton() {
    try { return (await this.msal.acquireTokenSilent({scopes: this.scopes})).accessToken; }
    catch (e) { await this.msal.acquireTokenRedirect({scopes: this.scopes}); throw new Error("Reconnexion…"); }
  }
  async req(url, opt = {}) {
    const t = await this.jeton();
    const r = await fetch(url, {...opt, headers: {...(opt.headers || {}), Authorization: "Bearer " + t}});
    return r;
  }
  async infos(chemin) {
    const r = await this.req(GRAPH + enc(chemin) + "?$select=eTag,lastModifiedDateTime,size");
    if (r.status === 404) return null;
    if (!r.ok) throw new Error("OneDrive : " + r.status);
    const j = await r.json(); return {etag: j.eTag, modifie: j.lastModifiedDateTime};
  }
  async binaire(chemin) {
    const r = await this.req(GRAPH + enc(chemin) + ":/content");
    if (r.status === 404) return null;
    if (!r.ok) throw new Error(`Impossible de lire « ${chemin} » dans OneDrive (${r.status}).`);
    return r.arrayBuffer();
  }
  async json(chemin) {
    const i = await this.infos(chemin); if (!i) return {data: null, etag: null};
    const b = await this.binaire(chemin);
    return {data: JSON.parse(new TextDecoder().decode(b) || "null"), etag: i.etag};
  }
  async ecrireJSON(chemin, data, etag) {
    const h = {"Content-Type": "application/json"}; if (etag) h["If-Match"] = etag;
    const r = await this.req(GRAPH + enc(chemin) + ":/content", {method: "PUT", headers: h, body: JSON.stringify(data, null, 1)});
    if (r.status === 412) { const e = new Error("conflit"); e.conflit = true; throw e; }
    if (!r.ok) throw new Error("Enregistrement OneDrive impossible (" + r.status + ").");
    return (await r.json()).eTag;
  }
  async listeExcel() {
    const r = await this.req("https://graph.microsoft.com/v1.0/me/drive/root/children?$select=name,file&$top=200");
    if (!r.ok) throw new Error("OneDrive : " + r.status);
    return (await r.json()).value.filter(x => x.file && /\.xlsx$/i.test(x.name)).map(x => x.name);
  }
  prenom() { const a = this.msal.getActiveAccount(); return a && a.name ? a.name.split(" ")[0] : ""; }
  async urlPhoto(nom) {
    const b = await this.binaire(CHEMINS.reglements + "/" + nom);
    if (!b) throw new Error("Photo introuvable dans OneDrive.");
    return URL.createObjectURL(new Blob([b]));
  }
}

/* ---------- stockage de test (fichiers servis par dev_server.py, jamais tes vrais fichiers) ---------- */
class StockDev {
  async init() {} connecte() { return true; } compte() { return "mode test"; } connexion() {} deconnexion() {}
  async infos(chemin) { const r = await fetch("/dev/" + enc(chemin), {method: "HEAD"}); return r.ok ? {etag: r.headers.get("ETag"), modifie: r.headers.get("Last-Modified")} : null; }
  async binaire(chemin) { const r = await fetch("/dev/" + enc(chemin)); return r.ok ? r.arrayBuffer() : null; }
  async json(chemin) { const r = await fetch("/dev/" + enc(chemin)); if (!r.ok) return {data: null, etag: null}; return {data: await r.json(), etag: r.headers.get("ETag")}; }
  async ecrireJSON(chemin, data, etag) {
    const r = await fetch("/dev/" + enc(chemin), {method: "PUT", headers: {"Content-Type": "application/json", ...(etag ? {"If-Match": etag} : {})}, body: JSON.stringify(data, null, 1)});
    if (r.status === 412) { const e = new Error("conflit"); e.conflit = true; throw e; }
    return r.headers.get("ETag");
  }
  async urlPhoto(nom) { return "/dev/" + enc(CHEMINS.reglements + "/" + nom); }
  async listeExcel() { return (await (await fetch("/dev/__liste")).json()); }
  prenom() { return "Test"; }
}

let STOCK = null;
const CACHE = {excelEtag: null, classeur: null, prospEtag: null, prospects: [], appli: null, appliEtag: null, fait: null, maj: null};

/* ---------- chargement (avec copie locale pour le hors-ligne) ---------- */
async function chargerDonnees() {
  const [ie, ip] = await Promise.all([STOCK.infos(CHEMINS.excel), STOCK.infos(F_PROSP)]);
  if (!ie) throw new Error(`Fichier « ${CHEMINS.excel} » introuvable à la racine de ton OneDrive.`);
  if (ie.etag !== CACHE.excelEtag) { CACHE.classeur = lireClasseur(await STOCK.binaire(CHEMINS.excel)); CACHE.excelEtag = ie.etag; }
  CACHE.maj = ie.modifie;
  if (ip && ip.etag !== CACHE.prospEtag) { CACHE.prospects = lireProspects(await STOCK.binaire(F_PROSP)); CACHE.prospEtag = ip.etag; }
  const [a, f] = await Promise.all([STOCK.json(F_APPLI), STOCK.json(F_FAIT)]);
  CACHE.appli = a.data || {commandes: [], prospects: {}, tarifs: TARIFS_DEFAUT, objectif: 30000};
  CACHE.appliEtag = a.etag; CACHE.fait = f.data || {faits: {}, erreurs: {}};
  try { localStorage.setItem("yousdam_cache", JSON.stringify({classeur: CACHE.classeur, prospects: CACHE.prospects, appli: CACHE.appli, fait: CACHE.fait, maj: CACHE.maj})); } catch (e) {}
}
function chargerHorsLigne() {
  const c = JSON.parse(localStorage.getItem("yousdam_cache") || "null");
  if (!c) return false;
  Object.assign(CACHE, c); CACHE.horsLigne = true; return true;
}

/* modification du fichier de l'appli : relit, applique, enregistre (et recommence si un autre appareil a écrit entre-temps) */
async function majAppli(fn) {
  for (let essai = 0; essai < 4; essai++) {
    const a = await STOCK.json(F_APPLI);
    const d = a.data || {commandes: [], prospects: {}, tarifs: TARIFS_DEFAUT, objectif: 30000};
    const r = fn(d);
    try { CACHE.appliEtag = await STOCK.ecrireJSON(F_APPLI, d, a.etag); CACHE.appli = d; return r; }
    catch (e) { if (!e.conflit) throw e; }
  }
  throw new Error("Un autre appareil modifie l'appli en même temps. Réessaie.");
}
const idNouveau = () => Math.random().toString(16).slice(2, 12);
const trouver = (d, id) => { const c = d.commandes.find(x => x.id === id); if (!c) throw new Error("Commande introuvable."); return c; };

/* ---------- état complet pour l'interface (même forme que /api/etat sur PC) ---------- */
function ligneExcelDe(c) { return c.excel_row || (CACHE.fait.faits[c.id] || {}).row || null; }
/* les livraisons envoyées à Excel mais pas encore écrites par le PC comptent déjà dans les chiffres */
function ventesCompletes() {
  const {ventes: v0, villes} = CACHE.classeur, cmds = CACHE.appli.commandes;
  const brut = {}; v0.forEach(v => { if (!(v.ck in brut)) brut[v.ck] = v.brut; });
  const virtuelles = cmds.filter(c => c.statut === "livree" && c.excel_demande && !ligneExcelDe(c)).map(c => {
    const L = c.livraison, md = L.decharge ? ((villes.find(x => x.ville === L.ville) || {}).decharge || 0) : 0;
    return {row: null, virtuelle: true, date: L.date, client: c.client.trim(), brut: brut[cle(c.client)] || c.client, ck: cle(c.client), palettes: L.palettes, type: L.type, qte: L.qte,
      pu: L.pu, tu: L.qte * L.pu + md, pv: L.pv, ca: L.qte * L.pv + md, ville: L.ville, chantier: L.chantier, client_final: L.client_final,
      decharge: L.decharge ? "Oui" : "Non", mdech: md, benef: L.qte * (L.pv - L.pu), paiement: "En cours"}; });
  return {ventes: v0.concat(virtuelles), brut};
}
function etatComplet() {
  const {ventes: v0, villes, solde, regs} = CACHE.classeur;
  const A = CACHE.appli, cmds = A.commandes;
  const {ventes, brut} = ventesCompletes();
  const attente = cmds.filter(c => c.statut === "attente").map(c => ({...c}));
  const a = analyse(ventes, attente, regs, solde, A.objectif || 30000);
  const prises = new Set(cmds.map(ligneExcelDe).filter(Boolean));
  attente.forEach(c => c.dans_excel = dansExcel(v0, prises, c.client, c.palettes, c.client_final, c.chantier, c.date, 0, 14));
  const aExcel = cmds.filter(c => c.statut === "livree" && !ligneExcelDe(c)).map(c => ({...c, en_file: !!c.excel_demande,
    erreur: (CACHE.fait.erreurs || {})[c.id] || null,
    dans_excel: c.excel_demande ? null : dansExcel(v0, prises, c.client, c.livraison.palettes, c.livraison.client_final, c.livraison.chantier, c.livraison.date, -3, 3)}));
  const prospects = CACHE.prospects.map(p => ({...p, ...((A.prospects || {})[p.feuille + "|" + p.nom] || {})}));
  const livrees = ventes.slice().sort((x, y) => y.date.localeCompare(x.date) || (y.row || 1e9) - (x.row || 1e9)).slice(0, 60);
  const noms = [...new Set(v0.map(x => x.client))].sort((x, y) => x.toLowerCase().localeCompare(y.toLowerCase()));
  const m = CACHE.maj ? new Date(CACHE.maj) : new Date();
  const maj = `${m.getFullYear()}-${String(m.getMonth() + 1).padStart(2, "0")}-${String(m.getDate()).padStart(2, "0")}T${String(m.getHours()).padStart(2, "0")}:${String(m.getMinutes()).padStart(2, "0")}`;
  return {analyse: a, attente, livrees, reglements: regs.slice().sort((x, y) => (y.echeance || "").localeCompare(x.echeance || "")), villes, clients_noms: noms, clients_brut: brut,
    prospects, maj, a_excel: aExcel, passees: cmds.filter(ligneExcelDe).slice(-20).reverse().map(c => ({...c, excel_row: ligneExcelDe(c)})),
    tarifs: A.tarifs || TARIFS_DEFAUT, hors_ligne: !!CACHE.horsLigne, compte: STOCK.compte(), prenom: STOCK.prenom ? STOCK.prenom() : "", fichier: CHEMINS.excel};
}

/* ---------- code PIN (sur l'appareil) ---------- */
async function hacher(pin, sel) {
  const b = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(sel + ":" + pin));
  return [...new Uint8Array(b)].map(x => x.toString(16).padStart(2, "0")).join("");
}

/* ---------- routeur : mêmes adresses que le serveur PC, pour garder la même interface ---------- */
function livraisonDe(b) {
  const typ = b.type || "B7", pal = parseInt(b.palettes, 10);
  const pu = parseFloat(String(b.pu).replace(",", ".")), pv = parseFloat(String(b.pv).replace(",", "."));
  if (!(pu > 0 && pv > 0)) throw new Error("Prix usine et prix de vente obligatoires.");
  return {date: (b.date || aujourdhui()).slice(0, 10), palettes: pal, type: typ, qte: pal * (PAR_PALETTE[typ] || 256), pu, pv, ville: b.ville || "Casa",
    chantier: (b.chantier || "").trim(), client_final: (b.client_final || "").trim(), decharge: !!b.decharge};
}
async function routeur(path, body) {
  const u = new URL(path, location.href), q = Object.fromEntries(u.searchParams), b = body || {};
  switch (u.pathname.replace(/^.*\/api\//, "/api/")) {
    case "/api/session": return {pin_defini: !!localStorage.getItem("pin"), connecte: sessionStorage.getItem("ok") === "1"};
    case "/api/pin": {
      if (!/^\d{4,8}$/.test(String(b.pin))) throw new Error("Le code doit faire 4 à 8 chiffres.");
      const sel = crypto.getRandomValues(new Uint32Array(2)).join("");
      localStorage.setItem("pin", JSON.stringify({sel, h: await hacher(String(b.pin), sel)})); sessionStorage.setItem("ok", "1"); return {ok: true};
    }
    case "/api/connexion": {
      const p = JSON.parse(localStorage.getItem("pin") || "null");
      const essais = JSON.parse(sessionStorage.getItem("essais") || "[0,0]");
      if (Date.now() < essais[1]) throw new Error("Trop d'essais. Réessaie dans quelques minutes.");
      if (!p || await hacher(String(b.pin), p.sel) !== p.h) {
        essais[0]++; if (essais[0] >= 5) { essais[0] = 0; essais[1] = Date.now() + 300e3; }
        sessionStorage.setItem("essais", JSON.stringify(essais)); throw new Error("Code incorrect.");
      }
      sessionStorage.setItem("ok", "1"); return {ok: true};
    }
    case "/api/etat":
      try { await chargerDonnees(); CACHE.horsLigne = false; }
      catch (e) { if (!navigator.onLine && chargerHorsLigne()) {} else if (!CACHE.classeur) throw e; }
      return etatComplet();
    case "/api/maj": { const i = await STOCK.infos(CHEMINS.excel), a = await STOCK.infos(F_APPLI), f = await STOCK.infos(F_FAIT);
      return {ts: i && i.etag, app: (a && a.etag) + "|" + (f && f.etag)}; }
    case "/api/prix": return prixDefaut(CACHE.classeur.ventes, q.client || "", q.type || "B7", q.ville || "Casa", CACHE.appli.tarifs || TARIFS_DEFAUT);
    case "/api/situation": return {lignes: situation(ventesCompletes().ventes, q.client || "")};
    case "/api/lire_message": return {commandes: lireMessageWA(b.texte || "")};
    case "/api/verifier_livraison": return {doublon: doublon(CACHE.classeur.ventes, b.client, b.palettes, b.client_final || "", (b.date || aujourdhui()).slice(0, 10))};
    case "/api/commande": {
      const enFile = [];
      await majAppli(d => {
        for (const c of b.commandes) {
          const typ = c.type || "B7", pal = parseInt(c.palettes, 10);
          const n = {id: idNouveau(), client: b.client.trim(), statut: "attente", cree: new Date().toISOString().slice(0, 16), date: (c.date || aujourdhui()).slice(0, 10),
            palettes: pal, type: typ, qte: pal * (PAR_PALETTE[typ] || 256), chantier: (c.chantier || "").trim(), ville: c.ville || "Casa",
            client_final: (c.client_final || "").trim(), contact: (c.contact || "").trim(), message: c.brut || ""};
          if (c.livree) { n.statut = "livree"; n.livraison = livraisonDe(c); if (b.excel) { n.excel_demande = true; n.excel_demande_le = new Date().toISOString(); enFile.push("file"); } }
          d.commandes.push(n);
        }
      });
      return {ok: true, lignes: enFile};
    }
    case "/api/livrer": await majAppli(d => { const c = trouver(d, b.id); c.statut = "livree"; c.livraison = livraisonDe(b); if (b.excel) { c.excel_demande = true; c.excel_demande_le = new Date().toISOString(); } });
      return {ok: true, lignes: b.excel ? ["file"] : []};
    case "/api/excel": await majAppli(d => { const c = trouver(d, b.id); if (c.statut !== "livree") throw new Error("Cette commande n'est pas livrée."); c.excel_demande = true; c.excel_demande_le = new Date().toISOString(); });
      return {ok: true, lignes: ["file"]};
    case "/api/annuler": await majAppli(d => { trouver(d, b.id).statut = "annulee"; }); return {ok: true};
    case "/api/remettre": await majAppli(d => { const c = trouver(d, b.id); if (ligneExcelDe(c) || c.excel_demande) throw new Error("Déjà envoyée à Excel."); c.statut = "attente"; delete c.livraison; }); return {ok: true};
    case "/api/classer": {
      const v = CACHE.classeur.ventes.find(x => x.row === +b.row); if (!v) throw new Error("Ligne introuvable dans l'Excel. Actualise et réessaie.");
      await majAppli(d => { const c = trouver(d, b.id); if (cle(c.client) !== v.ck) throw new Error("Cette ligne ne correspond pas à ce client.");
        if (c.statut === "attente") { c.statut = "livree"; c.livraison = {date: v.date, palettes: v.palettes, type: v.type, qte: v.qte, pu: v.pu, pv: v.pv, ville: v.ville, chantier: v.chantier, client_final: v.client_final, decharge: v.decharge === "Oui"}; }
        c.excel_row = v.row; delete c.excel_demande; });
      return {ok: true};
    }
    case "/api/reessayer": await majAppli(d => { const c = trouver(d, b.id); c.excel_demande = true; c.excel_reessai = new Date().toISOString(); }); return {ok: true};
    case "/api/prospect": await majAppli(d => { const k = b.feuille + "|" + b.nom; d.prospects = d.prospects || {}; d.prospects[k] = {...(d.prospects[k] || {}), statut: b.statut, ...(b.date ? {date: b.date.slice(0, 10)} : {}), ...(b.notes != null ? {notes: b.notes} : {})}; }); return {ok: true};
    case "/api/tarifs": await majAppli(d => { d.tarifs = b.tarifs; }); return {ok: true};
    case "/api/objectif": await majAppli(d => { d.objectif = Math.max(0, parseInt(b.objectif, 10) || 0); }); return {ok: true};
  }
  throw new Error("Action inconnue : " + path);
}
