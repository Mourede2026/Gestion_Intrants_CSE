/**
 * API « Gestion des intrants » v2 : utilisateurs, rôles, lots, péremption, stock de sécurité, alertes.
 * Coller dans Extensions > Apps Script, puis Déployer > Gérer les déploiements > Modifier > Nouvelle version.
 * Alertes par e-mail : exécuter une fois installerAlerteQuotidienne().
 */
var P = ['id', 'nom', 'unite', 'stock_depart', 'stock_securite', 'peremption_depart', 'catalogue_id'];
var C = ['id', 'nom', 'forme', 'dosage', 'conditionnement'];
var M = ['bordereau_id', 'type', 'date', 'ref', 'type_destination', 'partenaire', 'produit_id', 'qte', 'lot', 'peremption', 'utilisateur'];
var U = ['id', 'nom', 'login', 'mot_de_passe', 'role', 'actif', 'email'];

function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return json({ ok: true, info: 'API Gestion des intrants' }); }

function sh(name, head) {
  var ss = SpreadsheetApp.getActive(), s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.setFrozenRows(1); }
  s.getRange(1, 1, 1, head.length).setValues([head]);
  return s;
}
function rows(name, head) { return sh(name, head).getDataRange().getValues().slice(1).filter(function (r) { return r[0] !== ''; }); }
function fd(v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v || ''); }
function withLock(fn) { var l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }
function today() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

/* ---------- Utilisateurs ---------- */
function hash(pw, salt) { return salt + ':' + Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + pw)); }
function users() {
  return rows('Utilisateurs', U).map(function (r) { return { id: String(r[0]), name: r[1], login: r[2], hash: r[3], role: r[4], actif: r[5] === true || r[5] === 'TRUE', email: r[6] }; });
}
function createUser(b, role) {
  var login = String(b.login || '').trim().toLowerCase();
  if (!String(b.name || '').trim() || !login) throw new Error('Nom et identifiant requis.');
  if (String(b.password || '').length < 6) throw new Error('Mot de passe : 6 caractères minimum.');
  if (users().some(function (u) { return u.login === login; })) throw new Error('Cet identifiant existe déjà.');
  sh('Utilisateurs', U).appendRow([Utilities.getUuid(), String(b.name).trim(), login, hash(b.password, Utilities.getUuid().slice(0, 8)), role, true, b.email || '']);
}
function login(b) {
  var l = String(b.login || '').trim().toLowerCase();
  var u = users().filter(function (x) { return x.login === l; })[0];
  if (!u || !u.actif || hash(String(b.password || ''), u.hash.split(':')[0]) !== u.hash) throw new Error('Identifiant ou mot de passe incorrect.');
  var t = Utilities.getUuid();
  CacheService.getScriptCache().put(t, u.id, 21600);
  return json({ ok: true, token: t });
}
function auth(t) {
  var id = t && CacheService.getScriptCache().get(t);
  var u = id && users().filter(function (x) { return x.id === id; })[0];
  if (!u || !u.actif) throw new Error('SESSION');
  return u;
}
function updateUser(b) {
  var s = sh('Utilisateurs', U), r = s.getDataRange().getValues();
  for (var i = 1; i < r.length; i++) {
    if (String(r[i][0]) !== b.id) continue;
    var role = b.role || r[i][4], actif = b.actif === undefined ? r[i][5] : b.actif;
    if (r[i][4] === 'gestionnaire' && (role !== 'gestionnaire' || actif === false)) {
      if (users().filter(function (u) { return u.role === 'gestionnaire' && u.actif && u.id !== b.id; }).length === 0)
        throw new Error('Il doit rester au moins un gestionnaire actif.');
    }
    s.getRange(i + 1, 5, 1, 2).setValues([[role, actif]]);
    if (b.name) s.getRange(i + 1, 2).setValue(b.name);
    if (b.email !== undefined) s.getRange(i + 1, 7).setValue(b.email);
    if (b.password) {
      if (b.password.length < 6) throw new Error('Mot de passe : 6 caractères minimum.');
      s.getRange(i + 1, 4).setValue(hash(b.password, Utilities.getUuid().slice(0, 8)));
    }
  }
}

/* ---------- Données ---------- */
function groupMoves() {
  var map = {}, arr = [];
  rows('Mouvements', M).forEach(function (r) {
    var m = map[r[0]];
    if (!m) { m = map[r[0]] = { id: String(r[0]), type: r[1], date: fd(r[2]), ref: r[3], dest: r[4], party: r[5], user: r[10] || '', lines: [] }; arr.push(m); }
    m.lines.push({ pid: String(r[6]), qty: +r[7], lot: r[8] || '', exp: fd(r[9]) });
  });
  return arr;
}
function byExp(a, b) { var x = a.exp || '9999', y = b.exp || '9999'; return x < y ? -1 : x > y ? 1 : 0; }
function lotsOf(p, moves) {
  var m = {};
  function add(l, e, q) { var k = l + '|' + e; m[k] = m[k] || { lot: l || 'Sans lot', exp: e, qty: 0 }; m[k].qty += q; }
  if (p.initial) add('Stock initial', p.expInit, p.initial);
  moves.forEach(function (mv) { mv.lines.forEach(function (l) { if (l.pid === p.id) add(l.lot, l.exp, mv.type === 'in' ? l.qty : -l.qty); }); });
  return Object.keys(m).map(function (k) { return m[k]; }).filter(function (x) { return x.qty !== 0; }).sort(byExp);
}
function getProducts(moves) {
  var t = today();
  return rows('Produits', P).map(function (r) {
    var p = { id: String(r[0]), name: r[1], unit: r[2], initial: +r[3] || 0, secu: +r[4] || 0, expInit: fd(r[5]), cid: String(r[6] || '') };
    p.lots = lotsOf(p, moves);
    p.stock = 0; p.usable = 0;
    p.lots.forEach(function (l) { p.stock += l.qty; if (!l.exp || l.exp >= t) p.usable += l.qty; });
    return p;
  });
}
function getData(u) {
  var moves = groupMoves();
  var d = { me: { id: u.id, name: u.name, role: u.role }, products: getProducts(moves), catalogue: catalogue(), moves: moves.reverse() };
  if (u.role === 'gestionnaire') d.users = users().map(function (x) { return { id: x.id, name: x.name, login: x.login, role: x.role, actif: x.actif, email: x.email }; });
  return d;
}

/* ---------- Écritures ---------- */
function addProduct(p) {
  if (!String(p.name || '').trim()) throw new Error('Nom du produit requis.');
  sh('Produits', P).appendRow([Utilities.getUuid(), String(p.name).trim(), p.unit || '', +p.initial || 0, +p.secu || 0, p.expInit || '']);
}
function updateProduct(p) {
  var s = sh('Produits', P), r = s.getDataRange().getValues();
  for (var i = 1; i < r.length; i++) if (String(r[i][0]) === p.id) {
    s.getRange(i + 1, 2, 1, 5).setValues([[p.name || r[i][1], p.unit === undefined ? r[i][2] : p.unit, +p.initial || 0, +p.secu || 0, p.expInit || '']]);
  }
}
function deleteProduct(id) {
  var used = groupMoves().some(function (m) { return m.lines.some(function (l) { return l.pid === id; }); });
  if (used) throw new Error('Ce produit a des mouvements : suppression impossible.');
  var s = sh('Produits', P), r = s.getDataRange().getValues();
  for (var i = r.length - 1; i >= 1; i--) if (String(r[i][0]) === id) s.deleteRow(i + 1);
}
function saveMove(m, u) {
  if (!m || !m.lines || !m.lines.length) throw new Error('Aucune ligne de produit.');
  var moves = groupMoves(), prods = getProducts(moves), t = today(), out = [];
  m.lines.forEach(function (l) {
    var p = prods.filter(function (x) { return x.id === l.pid; })[0];
    if (!p || !(+l.qty > 0)) throw new Error('Ligne invalide.');
    if (m.type === 'in') { out.push({ pid: l.pid, qty: +l.qty, lot: l.lot || '', exp: l.exp || '' }); return; }
    var need = +l.qty; // sortie : lots qui expirent le plus tôt d'abord (FEFO), jamais les lots périmés
    p.lots.forEach(function (x) {
      if (need <= 0 || x.qty <= 0 || (x.exp && x.exp < t)) return;
      var q = Math.min(need, x.qty); need -= q;
      out.push({ pid: l.pid, qty: q, lot: x.lot === 'Sans lot' ? '' : x.lot, exp: x.exp });
    });
    if (need > 0) throw new Error('Stock utilisable insuffisant pour ' + p.name + ' (disponible : ' + p.usable + ').');
  });
  var id = Utilities.getUuid(), s = sh('Mouvements', M);
  var data = out.map(function (l) { return [id, m.type, m.date, m.ref || '', m.dest || '', m.party || '', l.pid, l.qty, l.lot, l.exp, u.name]; });
  var start = s.getLastRow() + 1;
  s.getRange(start, 3, data.length, 1).setNumberFormat('@');
  s.getRange(start, 10, data.length, 1).setNumberFormat('@');
  s.getRange(start, 1, data.length, M.length).setValues(data);
}
function deleteMove(id) {
  var s = sh('Mouvements', M), r = s.getDataRange().getValues();
  for (var i = r.length - 1; i >= 1; i--) if (String(r[i][0]) === id) s.deleteRow(i + 1);
}

/* ---------- Catalogue des intrants ---------- */
function catalogue() {
  var s = sh('Catalogue', C), pr = PropertiesService.getScriptProperties();
  if (s.getLastRow() < 2 && !pr.getProperty('catSeed')) { appendText(s, [Utilities.getUuid(), 'Alcool', 'Solution antiseptique', '70°', '1 Litre']); pr.setProperty('catSeed', '1'); }
  return rows('Catalogue', C).map(function (r) { return { id: String(r[0]), name: r[1], forme: r[2], dosage: String(r[3]), cond: r[4] }; });
}
function appendText(s, row) {
  var n = s.getLastRow() + 1, g = s.getRange(n, 1, 1, row.length);
  g.setNumberFormat('@'); g.setValues([row]);
}
function catLabel(c) { return c.name + (c.dosage ? ' ' + c.dosage : '') + (c.cond ? ' (' + c.cond + ')' : ''); }
function addCatalogue(c) {
  if (!String(c.name || '').trim()) throw new Error("Nom de l'intrant requis.");
  catalogue();
  appendText(sh('Catalogue', C), [Utilities.getUuid(), String(c.name).trim(), c.forme || '', c.dosage || '', c.cond || '']);
}
function updateCatalogue(c) {
  var s = sh('Catalogue', C), r = s.getDataRange().getValues();
  for (var i = 1; i < r.length; i++) if (String(r[i][0]) === c.id) {
    var g = s.getRange(i + 1, 2, 1, 4); g.setNumberFormat('@'); g.setValues([[c.name || r[i][1], c.forme || '', c.dosage || '', c.cond || '']]);
  }
}
function deleteCatalogue(id) {
  if (rows('Produits', P).some(function (r) { return String(r[6]) === id; })) throw new Error('Cet intrant est déjà en stock : suppression impossible.');
  var s = sh('Catalogue', C), r = s.getDataRange().getValues();
  for (var i = r.length - 1; i >= 1; i--) if (String(r[i][0]) === id) s.deleteRow(i + 1);
}
function importCatalogue(text) { // lignes : Nom ; Forme ; Dosage ; Conditionnement
  var have = {}, s = sh('Catalogue', C), n = 0;
  catalogue().forEach(function (c) { have[(c.name + '|' + c.dosage + '|' + c.cond).toLowerCase()] = 1; });
  String(text || '').split(/\r?\n/).forEach(function (line) {
    var p = line.split(/[;\t]/).map(function (x) { return x.trim(); });
    if (!p[0] || /^(intrant|nom)$/i.test(p[0])) return;
    var k = (p[0] + '|' + (p[2] || '') + '|' + (p[3] || '')).toLowerCase();
    if (have[k]) return; have[k] = 1;
    appendText(s, [Utilities.getUuid(), p[0], p[1] || '', p[2] || '', p[3] || '']); n++;
  });
  if (!n) throw new Error('Aucune ligne importée (format : Nom ; Forme ; Dosage ; Conditionnement).');
}
function addFromCatalogue(items) {
  var cat = catalogue(), have = {}, s = sh('Produits', P), n = 0;
  rows('Produits', P).forEach(function (r) { have[String(r[6])] = 1; });
  (items || []).forEach(function (it) {
    var c = cat.filter(function (x) { return x.id === it.cid; })[0];
    if (!c || have[c.id]) return;
    s.appendRow([Utilities.getUuid(), catLabel(c), c.cond, +it.initial || 0, +it.secu || 0, it.expInit || '', c.id]); n++;
  });
  if (!n) throw new Error('Aucun intrant sélectionné.');
}

function doPost(e) {
  try {
    var b = JSON.parse(e.postData.contents), a = b.action;
    if (a === 'status') return json({ ok: true, setup: users().length === 0 });
    if (a === 'setup') return withLock(function () {
      if (users().length) throw new Error('Déjà configuré.');
      createUser(b, 'gestionnaire'); return login(b);
    });
    if (a === 'login') return login(b);
    var u = auth(b.token);
    if (a !== 'data') withLock(function () {
      var mgr = ['addProduct', 'updateProduct', 'deleteProduct', 'deleteMove', 'addUser', 'updateUser', 'addCatalogue', 'updateCatalogue', 'deleteCatalogue', 'importCatalogue', 'addFromCatalogue'];
      if (mgr.indexOf(a) > -1 && u.role !== 'gestionnaire') throw new Error('Réservé au gestionnaire.');
      if (a === 'saveMove') saveMove(b.move, u);
      else if (a === 'addProduct') addProduct(b.product);
      else if (a === 'updateProduct') updateProduct(b.product);
      else if (a === 'deleteProduct') deleteProduct(b.id);
      else if (a === 'deleteMove') deleteMove(b.id);
      else if (a === 'addUser') createUser(b.user, b.user.role === 'gestionnaire' ? 'gestionnaire' : 'utilisateur');
      else if (a === 'updateUser') updateUser(b.user);
      else if (a === 'addCatalogue') addCatalogue(b.item);
      else if (a === 'updateCatalogue') updateCatalogue(b.item);
      else if (a === 'deleteCatalogue') deleteCatalogue(b.id);
      else if (a === 'importCatalogue') importCatalogue(b.text);
      else if (a === 'addFromCatalogue') addFromCatalogue(b.items);
      else throw new Error('Action inconnue.');
    });
    return json({ ok: true, data: getData(u) });
  } catch (err) { return json({ ok: false, error: String(err.message || err) }); }
}

/* ---------- Alertes par e-mail (déclencheur quotidien) ---------- */
function envoyerAlertes() {
  var lim = Utilities.formatDate(new Date(Date.now() + 90 * 864e5), Session.getScriptTimeZone(), 'yyyy-MM-dd'), t = today(), L = [];
  getProducts(groupMoves()).forEach(function (p) {
    if (p.stock <= 0) L.push('RUPTURE : ' + p.name);
    else if (p.secu > 0 && p.stock <= p.secu) L.push('Sous le stock de sécurité : ' + p.name + ' (' + p.stock + ' / seuil ' + p.secu + ')');
    p.lots.forEach(function (l) {
      if (l.exp && l.exp < t) L.push('PÉRIMÉ : ' + p.name + ', lot ' + l.lot + ' (' + l.qty + ')');
      else if (l.exp && l.exp <= lim) L.push('Péremption proche : ' + p.name + ', lot ' + l.lot + ', ' + l.exp + ' (' + l.qty + ')');
    });
  });
  var to = users().filter(function (u) { return u.actif && u.role === 'gestionnaire' && u.email; }).map(function (u) { return u.email; }).join(',');
  if (L.length && to) MailApp.sendEmail(to, 'Alertes de stock (' + L.length + ')', L.join('\n'));
}
function installerAlerteQuotidienne() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'envoyerAlertes') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('envoyerAlertes').timeBased().everyDays(1).atHour(7).create();
}
