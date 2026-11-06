/*
 * PV COSMOS : récupère les informations d'un ticket SCTASK dans COSMOS
 * et les prépare pour le PV Excel (installation / restitution PC et iPhone).
 *
 * Utilisation : favori (bookmarklet) à cliquer sur la page d'un ticket SCTASK.
 * Le script ne fait que LIRE des informations visibles avec ton compte ;
 * il ne modifie rien dans COSMOS.
 *
 * Sources utilisées :
 *  - le formulaire du ticket (n° DEM, bénéficiaire, article, variables de
 *    l'onglet "Détail de la demande" : Matériel à récupérer / à restituer...)
 *  - les notes de travail du ticket et des autres tâches de la même demande
 *    (ex. "Approvisionnement matériel", note "Déstockage ... S/N / IMEI")
 *  - les actifs (alm_asset) affectés à / réservés pour le bénéficiaire
 *  - la fiche de l'utilisateur (nom, prénom, site...)
 * Pour chaque lecture distante : API REST de COSMOS d'abord, puis la page
 * HTML classique si l'API est refusée.
 */
(async function () {
  'use strict';

  var VERSION = '2.1';
  var DONE_KEY = 'pvcosmos.faits';
  var GCK = '';
  var TBL = 'sc_task';

  /* Si le formulaire n'est pas accessible (nouveau cadre COSMOS, iframe dans un
     shadow DOM fermé...), on relit la page du ticket à partir du sys_id de l'adresse. */
  async function resolveDoc() {
    if (D.getElementById('sys_uniqueValue')) return 'page';
    var href = location.href;
    try { href = decodeURIComponent(decodeURIComponent(href)); } catch (e) { /* ignoré */ }
    var m = href.match(/(sc_task|incident)\.do\?(?:[^#]*?&)?sys_id=([0-9a-f]{32})/i);
    if (!m) return 'aucun ticket SCTASK ou INC dans l\'adresse';
    var r = await fetch('/' + m[1].toLowerCase() + '.do?sys_id=' + m[2], { credentials: 'same-origin' });
    if (!r.ok) return 'lecture du ticket refusée (HTTP ' + r.status + ')';
    var txt = await r.text();
    var ck = txt.match(/g_ck\s*=\s*['"]([0-9a-f]{20,})['"]/i);
    if (ck) GCK = ck[1];
    D = new DOMParser().parseFromString(txt, 'text/html');
    return D.getElementById('sys_uniqueValue') ? 'relu' : 'page relue sans formulaire';
  }
  /* Trouve la fenêtre du formulaire classique : directement, dans l'iframe
     gsft_main (UI16), ou dans l'iframe cachée dans le shadow DOM du nouveau
     cadre COSMOS (/now/nav/ui/classic/...). */
  function findFrame() {
    function ok(w) { try { return !!(w && w.document && (w.document.getElementById('sys_uniqueValue') || w.g_form || w.g_ck)); } catch (e) { return false; } }
    if (document.getElementById('sys_uniqueValue')) return window;
    var f = document.getElementById('gsft_main');
    if (f && ok(f.contentWindow)) return f.contentWindow;
    var found = null;
    (function walk(root, depth) {
      if (found || !root || depth > 25) return;
      var ifr = root.querySelectorAll ? root.querySelectorAll('iframe') : [];
      for (var i = 0; i < ifr.length && !found; i++) if (ok(ifr[i].contentWindow)) found = ifr[i].contentWindow;
      var all = root.querySelectorAll ? root.querySelectorAll('*') : [];
      for (var j = 0; j < all.length && !found; j++) if (all[j].shadowRoot) walk(all[j].shadowRoot, depth + 1);
    })(document, 0);
    if (found) return found;
    for (var k = 0; k < window.frames.length; k++) if (ok(window.frames[k])) return window.frames[k];
    return window;
  }

  var TOP = window.top || window;
  var TD = TOP.document;
  var W = window;
  var D = document;
  W = findFrame();
  D = W.document;

  var old = TD.getElementById('pvcosmos-panel');
  if (old) old.remove();

  var warnings = [];
  var HEX32 = /^[0-9a-f]{32}$/;

  /* ---------- lecture du formulaire ---------- */

  function clean(s) {
    s = (s == null ? '' : String(s)).replace(/\s+/g, ' ').trim();
    return /^(--\s*aucun\s*--|\(vide\)|null|undefined)$/i.test(s) ? '' : s;
  }

  function val(id, doc) {
    var e = (doc || D).getElementById(id);
    if (!e) return '';
    if (e.tagName === 'SELECT') {
      var o = e.options[e.selectedIndex];
      return o ? o.text : '';
    }
    if (e.value != null) return e.value;
    return e.textContent || '';
  }

  function fld(name, doc, table) {
    table = table || TBL;
    var ids = ['sys_display.' + table + '.' + name, 'sys_readonly.' + table + '.' + name,
      table + '.' + name + '_label', table + '.' + name];
    for (var i = 0; i < ids.length; i++) {
      var v = clean(val(ids[i], doc));
      if (v && !HEX32.test(v)) return v;
    }
    return '';
  }

  function sysid(name, doc) {
    var v = clean(val(TBL + '.' + name, doc));
    return HEX32.test(v) ? v : '';
  }

  function varByLabel(re, doc) {
    var ls = (doc || D).querySelectorAll('label');
    for (var i = 0; i < ls.length; i++) {
      var t = ls[i].textContent.replace(/\s+/g, ' ').trim();
      if (!re.test(t)) continue;
      var id = ls[i].getAttribute('for');
      if (!id) continue;
      var base = id.replace(/^sys_display\./, '');
      var v = clean(val('sys_display.' + base, doc)) || clean(val(base + '_read_only', doc)) || clean(val(base, doc));
      if (v && !HEX32.test(v)) return v;
    }
    return '';
  }

  function journals(doc) {
    var out = [];
    (doc || D).querySelectorAll('.sn-widget-textblock-body').forEach(function (e) { out.push(e.textContent || ''); });
    return out;
  }

  /* ---------- analyse de texte ---------- */

  function grab(re, t) {
    var m = (t || '').match(re);
    return m ? clean(m[1]) : '';
  }

  function cleanPhoneModel(s) {
    return clean(String(s || '').replace(/\s+-\s+inscrit.*$/i, '').replace(/^APPLE\s+/i, ''));
  }

  function parseNote(t) {
    return {
      modele: cleanPhoneModel(grab(/Mod[eèé]le[ \t]*:[ \t]*([^\n\r]+)/i, t)),
      sn: grab(/S\/N[ \t]*:[ \t]*([A-Z0-9-]{5,})/i, t),
      imei: grab(/IMEI[ \t]*:?[ \t]*(\d{14,16})/i, t)
    };
  }

  function isPhoneText(s) { return /iphone|ipad|smartphone|IMEI/i.test(s || ''); }

  /* "IMEI 355108982261812 - GV4KQ2YVN735 - APPLE iPhone 11" */
  function parsePhoneStr(s) {
    var r = { imei: grab(/IMEI\s*:?\s*(\d{14,16})/i, s), sn: '', modele: '' };
    String(s || '').split(/\s+-\s+/).forEach(function (x) {
      x = clean(x);
      if (!x || /^IMEI/i.test(x)) return;
      if (!r.sn && /^[A-Z0-9]{8,14}$/i.test(x)) r.sn = x;
      else if (!r.modele && /iphone|ipad|apple/i.test(x)) r.modele = cleanPhoneModel(x);
    });
    return r;
  }

  /* "GT5ZQN1TC - 5CG936BR5X - HEWLETT PACKARD PROBOOK 650 G4" */
  function parsePcStr(s) {
    var p = String(s || '').split(/\s+-\s+/).map(clean);
    return { netbios: p[0] || '', sn: p[1] || '', modele: p.slice(2).join(' - ') };
  }

  /* "FRANCE/SUD-OUEST/BEZIERS/BAT MONTIMARAN/ETG 00/BUR 014"
     -> ville BEZIERS, bâtiment MONTIMARAN, étage 0, bureau 014
     (FRANCE et la région sont ignorés) */
  function parseLoc(loc) {
    var p = String(loc || '').split('/').map(clean).filter(Boolean);
    var r = { site: p.length >= 3 ? p[2] : '', bat: '', etage: '', bureau: '' };
    p.slice(3).forEach(function (x) {
      x.split(/\s+-\s+|,\s*/).forEach(function (y) {
        var m;
        if ((m = y.match(/^B[AÂ]T(?:IMENT)?\.?\s*(.+)$/i))) r.bat = r.bat || clean(m[1]);
        else if ((m = y.match(/^(?:ETG|[ÉE]TAGE|NIV(?:EAU)?)\.?\s*(.+)$/i))) r.etage = r.etage || clean(m[1]);
        else if (/^(RDC|REZ.DE.CHAUSS[ÉE]E)$/i.test(y)) r.etage = r.etage || '0';
        else if ((m = y.match(/^(?:BUR(?:EAU)?|PI[ÈE]CE|SALLE)\.?\s*(.+)$/i))) r.bureau = r.bureau || clean(m[1]);
      });
    });
    // "ETG 00" -> 0 ; "ETG 02" -> 2 ; "RDC" -> 0
    if (/^RDC$/i.test(r.etage)) r.etage = '0';
    else if (/^-?\d+$/.test(r.etage)) r.etage = String(parseInt(r.etage, 10));
    return r;
  }

  /* Date du dossier : la date de prochaine action (souvent le RDV) si elle est
     à venir, sinon aujourd'hui. Format JJ/MM/AAAA. */
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function fmtDate(d) { return pad(d.getDate()) + '/' + pad(d.getMonth() + 1) + '/' + d.getFullYear(); }
  function dateDefaut(s) {
    var now = new Date(), today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var m = String(s || '').match(/^(\d{2})\/(\d{2})\/(\d{4})/);
    if (m) {
      var d = new Date(+m[3], +m[2] - 1, +m[1]);
      if (d >= today) return fmtDate(d);
    }
    return fmtDate(today);
  }

  /* Combine plusieurs emplacements : chaque information vient du premier qui la donne */
  /* Le premier emplacement est la référence ; un autre emplacement ne
     complète bâtiment / étage / bureau que s'il est dans la même ville. */
  function bestLoc(list) {
    list = list.filter(Boolean);
    var r = parseLoc(list[0] || '');
    list.slice(1).forEach(function (l) {
      var x = parseLoc(l);
      if (r.site && x.site && norm(x.site) !== norm(r.site)) return;
      for (var k in r) if (!r[k] && x[k]) r[k] = x[k];
    });
    return r;
  }

  /* Tous les champs "emplacement" du ticket (variables qui pointent sur
     cmn_location), dans l'ordre : arrivée, emplacement / localisation, autres */
  function locationVars(doc) {
    doc = doc || D;
    var out = [];
    doc.querySelectorAll('input[id^="sys_display.ni."]').forEach(function (e) {
      var base = e.id.replace(/^sys_display\./, '');
      var v = clean(e.value || e.getAttribute('value'));
      if (!v || HEX32.test(v)) return;
      var isLoc = clean(val(base + 'TABLE', doc)) === 'cmn_location' || /cmn_location/.test(e.getAttribute('onfocus') || '') || /^FRANCE\//i.test(v);
      if (!isLoc) return;
      var lb = doc.getElementById('label_' + base) || doc.querySelector('label[for="' + e.id + '"]') || doc.querySelector('label[for="' + base + '"]');
      var lab = lb ? clean(lb.textContent) : '';
      var prio = /arriv/i.test(lab) ? 0 : /emplacement|localisation|site|lieu/i.test(lab) ? 1 : 2;
      out.push({ label: lab || 'Emplacement', value: v, prio: prio });
    });
    out.sort(function (a, b) { return a.prio - b.prio; });
    return out;
  }

  /* ---------- lectures distantes ---------- */

  async function api(table, query, fields) {
    try {
      var u = '/api/now/table/' + table + '?sysparm_display_value=true&sysparm_exclude_reference_link=true' +
        '&sysparm_limit=200&sysparm_query=' + encodeURIComponent(query) +
        (fields ? '&sysparm_fields=' + encodeURIComponent(fields) : '');
      var r = await fetch(u, {
        credentials: 'same-origin',
        headers: { 'Accept': 'application/json', 'X-UserToken': GCK || W.g_ck || TOP.g_ck || '' }
      });
      if (!r.ok) throw new Error('HTTP ' + r.status);
      var j = await r.json();
      return j.result || [];
    } catch (e) {
      warnings.push('API ' + table + ' refusée (' + e.message + '), lecture de la page à la place.');
      return null;
    }
  }

  async function getDoc(url) {
    var r = await fetch(url, { credentials: 'same-origin' });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return new DOMParser().parseFromString(await r.text(), 'text/html');
  }

  async function listHtml(table, query) {
    try {
      var doc = await getDoc('/' + table + '_list.do?sysparm_query=' + encodeURIComponent(query));
      var t = doc.getElementById(table + '_table') || doc.querySelector('table.list_table');
      if (!t) return [];
      var hs = [].map.call(t.querySelectorAll('thead tr:first-child th'), function (th) { return th.getAttribute('name') || ''; });
      return [].map.call(t.querySelectorAll('tbody tr[sys_id]'), function (tr) {
        var o = { sys_id: tr.getAttribute('sys_id') };
        [].forEach.call(tr.children, function (td, i) { if (hs[i]) o[hs[i]] = clean(td.textContent); });
        return o;
      });
    } catch (e) {
      warnings.push('Liste ' + table + ' illisible (' + e.message + ').');
      return [];
    }
  }

  async function query(table, q, fields) {
    var r = await api(table, q, fields);
    return r || listHtml(table, q);
  }

  async function getUser(id) {
    if (!id) return {};
    var r = await api('sys_user', 'sys_id=' + id, '');
    if (r && r[0]) return r[0];
    try {
      var doc = await getDoc('/sys_user.do?sys_id=' + id);
      var o = {};
      doc.querySelectorAll('[id^="sys_user."],[id^="sys_display.sys_user."],[id^="sys_readonly.sys_user."]').forEach(function (e) {
        var k = e.id.replace(/^(sys_display|sys_readonly)\./, '').replace(/^sys_user\./, '');
        var v = clean(val(e.id, doc));
        if (v && (!o[k] || HEX32.test(o[k]))) o[k] = v;
      });
      return o;
    } catch (e) {
      warnings.push('Fiche utilisateur illisible (' + e.message + ').');
      return {};
    }
  }

  async function journalsOf(taskId) {
    var r = await api('sys_journal_field', 'element_id=' + taskId + '^ORDERBYDESCsys_created_on', 'value,element');
    if (r && r.length) return r.map(function (x) { return x.value || ''; });
    try { return journals(await getDoc('/sc_task.do?sys_id=' + taskId)); } catch (e) { return []; }
  }

  function norm(x) { return clean(x).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase(); }

  /* Compte COSMOS d'une personne à partir de son nom / prénom (arrivée externe) */
  async function findUsers(nom, prenom) {
    if (!nom) return [];
    var rows = await query('sys_user', 'last_nameLIKE' + nom, 'sys_id,name,first_name,last_name,user_name');
    if (!rows.length && norm(nom) !== nom.toLowerCase()) rows = await query('sys_user', 'last_nameLIKE' + norm(nom).toUpperCase(), 'sys_id,name,first_name,last_name,user_name');
    var n = norm(nom), p = norm(prenom);
    return rows.filter(function (r) {
      var t = norm(Object.keys(r).map(function (k) { return r[k]; }).join(' '));
      return r.sys_id && t.indexOf(n) >= 0 && (!p || t.indexOf(p) >= 0);
    });
  }

  function userField(u, re) {
    for (var k in u) {
      if (!re.test(k)) continue;
      var v = clean(u[k]);
      if (v && !HEX32.test(v) && v !== 'false' && v !== 'true') return v;
    }
    return '';
  }

  /* ---------- tickets déjà traités (mémoire du navigateur) ---------- */

  function loadDone() { try { return JSON.parse(TOP.localStorage.getItem(DONE_KEY)) || {}; } catch (e) { return {}; } }
  function saveDone(o) { try { TOP.localStorage.setItem(DONE_KEY, JSON.stringify(o)); } catch (e) { /* ignoré */ } }

  /* Sur une liste COSMOS : lignes des tickets déjà traités en vert + pastille */
  function markList(doc) {
    var done = loadDone(), n = 0;
    doc.querySelectorAll('a').forEach(function (a) {
      var t = clean(a.textContent);
      if (!/^(SCTASK|INC)\d+$/i.test(t) || !done[t]) return;
      var tr = a.closest('tr');
      if (!tr || tr.getAttribute('data-pvc')) return;
      tr.setAttribute('data-pvc', '1');
      tr.querySelectorAll('td').forEach(function (td) { td.style.background = '#d9f2df'; });
      var b = doc.createElement('span');
      b.textContent = '✔ PV ' + (done[t].d || '');
      b.title = 'PV créés le ' + (done[t].d || '') + ' : ' + (done[t].a || '');
      b.style.cssText = 'background:#1a7f37;color:#fff;border-radius:9px;padding:0 6px;font-size:11px;margin-left:6px;white-space:nowrap';
      a.parentNode.insertBefore(b, a.nextSibling);
      n++;
    });
    return n;
  }

  function hasList(doc) { return !!doc.querySelector('table.list_table, tr[sys_id], tr[record_class]'); }

  /* ---------- panneau ---------- */

  function el(tag, props, kids) {
    var e = TD.createElement(tag);
    for (var k in (props || {})) {
      if (k === 'style') e.style.cssText = props[k];
      else if (k === 'text') e.textContent = props[k];
      else if (k.slice(0, 2) === 'on') e.addEventListener(k.slice(2), props[k]);
      else e.setAttribute(k, props[k]);
    }
    (kids || []).forEach(function (c) { if (c) e.appendChild(typeof c === 'string' ? TD.createTextNode(c) : c); });
    return e;
  }

  var panel = el('div', {
    id: 'pvcosmos-panel',
    style: 'position:fixed;top:10px;right:10px;width:560px;max-width:96vw;max-height:94vh;overflow:auto;' +
      'z-index:2147483647;background:#fff;color:#222;border:2px solid #0b5394;border-radius:8px;' +
      'box-shadow:0 6px 24px rgba(0,0,0,.35);font:13px Arial,sans-serif;padding:10px 12px;'
  });
  var status = el('div', { style: 'margin:6px 0;color:#555', text: 'Lecture du ticket en cours…' });
  panel.appendChild(el('div', { style: 'display:flex;justify-content:space-between;align-items:center' }, [
    el('b', { style: 'font-size:15px;color:#0b5394', text: 'PV COSMOS v' + VERSION }),
    el('button', { text: '✕', title: 'Fermer', style: 'border:0;background:none;font-size:16px;cursor:pointer', onclick: function () { panel.remove(); } })
  ]));
  panel.appendChild(status);
  TD.body.appendChild(panel);

  try {
    /* ----- 1. le ticket affiché ----- */
    var how = await resolveDoc();
    TBL = /^incident$/i.test(clean(val('sys_target'))) ? 'incident' : 'sc_task';
    var isInc = TBL === 'incident';
    var taskId = clean(val('sys_uniqueValue')) || sysid('sys_id');
    var number = fld('number') || clean(val('sys_displayValue'));
    if (!/^(SCTASK|INC)/i.test(number) && hasList(D)) {
      var nbMarq = markList(D);
      status.textContent = nbMarq ? nbMarq + ' ticket(s) déjà traités marqués en vert dans la liste.' :
        'Aucun ticket de cette liste n\'a encore de PV créé avec le favori.';
      setTimeout(function () { panel.remove(); }, 5000);
      return;
    }
    if (!/^(SCTASK|INC)/i.test(number)) {
      status.textContent = 'Ouvre la page d\'un ticket SCTASK ou d\'un incident INC, puis reclique sur le favori. ' +
        '[diagnostic : ' + how + ' ; cadres : ' + window.frames.length + ' ; ' + location.pathname.slice(0, 40) + ']';
      return;
    }
    // Incident : pas de demande DEM, le n° de dossier est le n° INC ; bénéficiaire = caller_id
    var dem = isInc ? number : fld('request_item');
    var ritm = isInc ? '' : sysid('request_item');
    var article = isInc ? [fld('category'), fld('subcategory')].filter(Boolean).join(' / ') : fld('request_item.cat_item');
    var shortDesc = fld('short_description');
    var benef = (isInc ? fld('caller_id') : fld('request_item.requested_for')) || varByLabel(/^B[ée]n[ée]ficiaire/i);
    var benefId = isInc ? sysid('caller_id') : sysid('request_item.requested_for');
    var descr = val(TBL + '.description');
    var benefOrig = benef;
    var nextAction = fld('u_next_action_datetime');
    var operateur = varByLabel(/Op[ée]rateur/i);
    var ligne = varByLabel(/Num[ée]ro de ligne/i) || varByLabel(/T[ée]l[ée]phone professionnel/i);
    var locs = locationVars(D);
    // "Localisation : FRANCE/..." écrite dans la description (incidents e-maintenance...)
    var locDesc = grab(/Localisation\s*:\s*(FRANCE\/[^\n\r]+)/i, descr);
    if (locDesc) locs.unshift({ label: 'Description', value: locDesc, prio: 0 });
    var locInc = isInc ? fld('location') : '';
    if (locInc) locs.push({ label: 'Emplacement du ticket', value: locInc, prio: 1 });
    // Arrivée externe : le bénéficiaire est un compte générique ("A CREER_...") ;
    // la vraie personne est dans les variables Nom / Prénom (Détail de la demande)
    var varNom = varByLabel(/^\*?\s*Nom\s*$/i);
    var varPrenom = varByLabel(/^\*?\s*Pr[ée]nom\s*$/i);
    var benefGenerique = /A\s*CREER|^\W*-\s/i.test(benef);
    var aRecuperer = varByLabel(/Mat[ée]riel.{0,4}(r[ée]cup[ée]rer|restituer|d[ée]sinstaller|reprendre)/i);
    var snSuppr = varByLabel(/Num[ée]ro de s[ée]rie.{0,4}supprimer/i);
    var texts = journals(D).concat([descr]);

    var isPhone = /iphone|smartphone|t[ée]l[ée]phon/i.test(article + (isInc ? ' ' + shortDesc : ''));
    var type = isPhone ? 'IPHONE' : (/d[ée]sinstall|restitu|r[ée]cup[ée]r|retour/i.test(shortDesc) ? 'REST_PC' : 'INST_PC');

    /* ----- 2. autres tâches de la même demande (Approvisionnement matériel...) ----- */
    status.textContent = 'Lecture des autres tâches de la demande ' + dem + '…';
    var sibs = ritm ? await query('sc_task', 'request_item=' + ritm, 'sys_id,number,short_description,state') : [];
    for (var i = 0; i < sibs.length; i++) {
      var sid = sibs[i].sys_id;
      if (!sid || sid === taskId) continue;
      texts = texts.concat(await journalsOf(sid));
    }

    /* ----- 3. utilisateur ----- */
    status.textContent = 'Lecture de la fiche utilisateur…';
    var user = benefGenerique ? {} : await getUser(benefId);
    var nomAff = clean(benef.replace(/\(.*?\)/g, ''));
    var prenom = varPrenom || clean(user.first_name) || (benefGenerique ? '' : nomAff.split(' ')[0] || '');
    var nom = varNom || clean(user.last_name) || (benefGenerique ? '' : nomAff.split(' ').slice(1).join(' '));
    // l'emplacement indiqué dans le ticket passe avant celui de la fiche utilisateur
    if (clean(user.location)) locs.push({ label: 'Fiche utilisateur', value: clean(user.location), prio: 3 });
    var seenLoc = {};
    locs = locs.filter(function (x) { var k = norm(x.value); if (seenLoc[k]) return false; seenLoc[k] = 1; return true; });
    var L = bestLoc(locs.map(function (x) { return x.value; }));
    // bâtiment / étage / bureau de la fiche utilisateur seulement si c'est la même ville
    var sameSite = !clean(user.location) || !L.site || norm(parseLoc(clean(user.location)).site) === norm(L.site);
    var uf = function (re) { return sameSite ? userField(user, re) : ''; };

    /* ----- 4. actifs affectés / réservés ----- */
    var ids = [];
    var qui = benef;
    if (benefGenerique || (varNom && !benefId)) {
      status.textContent = 'Recherche du compte de ' + prenom + ' ' + nom + '…';
      var found = await findUsers(nom, prenom);
      ids = found.map(function (u) { return u.sys_id; });
      if (found.length) qui = found.map(function (u) { return clean(u.name) || (clean(u.first_name) + ' ' + clean(u.last_name)); }).join(', ');
      else warnings.push('Aucun compte COSMOS trouvé pour ' + prenom + ' ' + nom + ' : le matériel réservé ne peut pas être cherché.');
    } else if (benefId) ids = [benefId];
    benef = qui;
    status.textContent = 'Recherche du matériel affecté à / réservé pour ' + qui + '…';
    var rawAssets = ids.length ? await query('alm_asset', 'assigned_toIN' + ids.join(',') + '^ORreserved_forIN' + ids.join(','),
      'sys_id,asset_tag,serial_number,u_code,model_category,model.display_name,model,install_status,substatus,reserved_for,assigned_to,location') : [];
    var assets = rawAssets.map(function (a) {
      var o = {
        tag: clean(a.asset_tag), sn: clean(a.serial_number), code: clean(a.u_code),
        cat: clean(a.model_category), model: clean(a['model.display_name'] || a.model || a.display_name),
        status: clean([a.install_status, a.substatus].map(clean).filter(Boolean).join(' / ')),
        reserved: clean(a.reserved_for), loc: clean(a.location)
      };
      var s = (o.cat + ' ' + o.model).toLowerCase();
      o.kind = /smartphone|iphone|t[ée]l[ée]phone/.test(s) ? 'PHONE'
        : /station|dock|r[ée]plicat/.test(s) ? 'DOCK'
        : /[ée]cran|moniteur|monitor/.test(s) ? 'SCREEN'
        : /ordinateur|portable|laptop|notebook|\bpc\b/.test(s) ? 'PC' : 'OTHER';
      o.imei = grab(/(\d{14,16})/, o.code);
      return o;
    });

    // Localisation de secours : emplacement du matériel EN SERVICE chez la personne
    // (pas celui du matériel réservé, qui est l'emplacement du stock)
    assets.forEach(function (a) {
      if (a.reserved || !/^FRANCE\//i.test(a.loc) || !/utilisation|install/i.test(a.status)) return;
      if (seenLoc[norm(a.loc)]) return;
      seenLoc[norm(a.loc)] = 1;
      locs.push({ label: 'Matériel en service (' + (a.cat || a.model) + ')', value: a.loc, prio: 4 });
    });
    L = bestLoc(locs.map(function (x) { return x.value; }));

    /* ----- 5. pré-remplissage ----- */
    var F = {
      DEM: dem, SCTASK: number, NOM: nom.toUpperCase(), PRENOM: prenom.toUpperCase(),
      NNI: grab(/\(([A-Z0-9]+)\)\s*$/i, benefOrig) || clean(user.user_name),
      LIGNE: clean(ligne).replace(/^\+33\s*/, '0'), OPERATEUR: operateur,
      DATE: dateDefaut(nextAction),
      SITE: L.site,
      BATIMENT: L.bat || uf( /^(building|u_.*(batiment|building))$/i),
      ETAGE: L.etage || uf( /(floor|etage)/i),
      BUREAU: L.bureau || uf( /(room|bureau|office_num)/i)
    };

    // iPhone à installer : note de travail "Modèle / S/N / IMEI" non vide
    texts.forEach(function (t) {
      var n = parseNote(t);
      if (n.sn && !F.IPI_SN) { F.IPI_SN = n.sn; F.IPI_IMEI = n.imei; F.IPI_MODELE = n.modele; }
    });

    // Matériel à récupérer (onglet Détail de la demande)
    var restSN = '';
    if (aRecuperer) {
      if (isPhoneText(aRecuperer)) {
        var pr = parsePhoneStr(aRecuperer);
        F.IPR_SN = pr.sn || snSuppr; F.IPR_IMEI = pr.imei; F.IPR_MODELE = pr.modele;
        restSN = F.IPR_SN;
      } else {
        var pc = parsePcStr(aRecuperer);
        F.PCR_NETBIOS = pc.netbios; F.PCR_SN = pc.sn; F.PCR_MODELE = pc.modele;
        restSN = pc.sn;
      }
    } else if (snSuppr && isPhone) {
      F.IPR_SN = snSuppr;
    }

    // Actifs : PC / station / écrans / iPhone à installer = ceux qui ne sont pas à restituer
    function pick(kind) {
      var c = assets.filter(function (a) { return a.kind === kind && a.sn && a.sn !== restSN; });
      c.sort(function (a, b) { return (b.reserved ? 1 : 0) - (a.reserved ? 1 : 0); });
      return c;
    }
    var pcs = pick('PC'), docks = pick('DOCK'), screens = pick('SCREEN'), phones = pick('PHONE');
    if (pcs[0]) { F.PCI_MODELE = pcs[0].model; F.PCI_SN = pcs[0].sn; F.PCI_NETBIOS = pcs[0].code; }
    if (docks[0]) F.PCI_STATION = docks[0].sn;
    if (screens[0]) F.PCI_ECRAN1 = screens[0].sn;
    if (screens[1]) F.PCI_ECRAN2 = screens[1].sn;
    if (!F.IPI_SN && phones[0] && phones[0].reserved) {
      F.IPI_SN = phones[0].sn; F.IPI_IMEI = phones[0].imei; F.IPI_MODELE = cleanPhoneModel(phones[0].model);
    }
    // compléter le modèle du matériel restitué à partir des actifs
    assets.forEach(function (a) {
      if (!restSN || a.sn !== restSN) return;
      if (a.kind === 'PHONE') { F.IPR_MODELE = F.IPR_MODELE || cleanPhoneModel(a.model); F.IPR_IMEI = F.IPR_IMEI || a.imei; }
      else { F.PCR_MODELE = F.PCR_MODELE || a.model; F.PCR_NETBIOS = F.PCR_NETBIOS || a.code; }
    });

    /* ----- 6. opérations à faire (cochées par défaut) ----- */
    var txtT = (article + ' ' + shortDesc).toLowerCase();
    var act = {};
    if (isPhone) {
      act.IPHONE = 1;
      if (F.IPI_SN || F.IPI_IMEI) act.CHARTE = 1;
    } else {
      if (F.PCI_SN || /install|appro|arriv|nouveau poste|mise en service/.test(txtT)) act.INST_PC = 1;
      if (F.PCR_SN || /d[ée]sinstall|restitu|r[ée]cup[ée]r|d[ée]part|retour/.test(txtT)) act.REST_PC = 1;
      if (act.REST_PC && (F.IPR_SN || F.IPR_IMEI)) { delete act.REST_PC; act.REST_PC_IP = 1; }
      if (!act.INST_PC && !act.REST_PC && !act.REST_PC_IP) act.INST_PC = 1;
    }
    var ACTIONS = [['INST_PC', 'Installation PC'], ['REST_PC', 'Désinstallation PC'], ['IPHONE', 'iPhone (changement / installation / restitution)'],
      ['REST_PC_IP', 'Désinstallation PC + iPhone'], ['CHARTE', 'Charte iPhone (+ PDF)']];

    /* ----- 7. affichage ----- */
    var GROUPS = [
      { title: 'Utilisateur et localisation', types: 'ALL', fields: [['DEM', 'N° dossier (DEM/INC)'], ['NOM', 'Nom'], ['PRENOM', 'Prénom'], ['SITE', 'Site (ville) *'], ['BATIMENT', 'Bâtiment *'], ['ETAGE', 'Étage *'], ['BUREAU', 'Bureau *'], ['DATE', 'Date du dossier']] },
      { title: 'PC installé', types: 'INST_PC', fields: [['PCI_MODELE', 'Modèle PC'], ['PCI_SN', 'N° série PC'], ['PCI_NETBIOS', 'NetBIOS'], ['PCI_STATION', 'N° série station'], ['PCI_ECRAN1', 'N° série écran 1'], ['PCI_ECRAN2', 'N° série écran 2']] },
      { title: 'PC restitué', types: 'REST_PC REST_PC_IP', fields: [['PCR_MODELE', 'Modèle PC'], ['PCR_SN', 'N° série PC'], ['PCR_NETBIOS', 'NetBIOS'], ['PCR_STATION', 'N° série station'], ['PCR_ECRAN1', 'N° série écran 1'], ['PCR_ECRAN2', 'N° série écran 2']] },
      { title: 'iPhone installé', types: 'IPHONE CHARTE', fields: [['IPI_MODELE', 'Modèle'], ['IPI_SN', 'N° série'], ['IPI_IMEI', 'IMEI']] },
      { title: 'iPhone restitué', types: 'IPHONE REST_PC_IP', fields: [['IPR_MODELE', 'Modèle'], ['IPR_SN', 'N° série'], ['IPR_IMEI', 'IMEI']] },
      { title: 'Charte iPhone', types: 'CHARTE', fields: [['NNI', 'NNI'], ['LIGNE', 'N° de ligne'], ['OPERATEUR', 'Opérateur']] }
    ];
    var REQUIRED = ['SITE', 'BATIMENT', 'ETAGE', 'BUREAU'];
    var inputs = {};
    var groupEls = [];
    var checks = {};

    status.textContent = '';
    panel.appendChild(el('div', { style: 'margin-bottom:6px' }, [
      el('div', { text: number + ' · ' + dem + ' · ' + benef }),
      el('div', { style: 'color:#666;font-size:12px', text: article + (shortDesc ? ' — ' + shortDesc : '') })
    ]));

    // déjà traité ?
    var done = loadDone();
    if (done[number]) {
      var ban = el('div', { style: 'background:#d9f2df;border:1px solid #1a7f37;border-radius:5px;padding:5px 8px;margin:6px 0;color:#145c2a' }, [
        el('b', { text: '✔ PV déjà créés le ' + done[number].d + ' ' }),
        el('span', { text: '(' + (done[number].a || '') + '). Évite les doublons.' }),
        el('button', { text: 'Marquer comme non fait', style: 'margin-left:6px;font-size:11px;cursor:pointer', onclick: function () {
          var o = loadDone(); delete o[number]; saveDone(o); ban.remove();
        } })
      ]);
      panel.appendChild(ban);
    }

    var actBox = el('div', { style: 'margin:6px 0;display:flex;flex-wrap:wrap;gap:4px 12px' }, [el('b', { style: 'width:100%', text: 'Documents à créer :' })]);
    ACTIONS.forEach(function (a) {
      var cb = el('input', { type: 'checkbox', onchange: refresh });
      cb.checked = !!act[a[0]];
      checks[a[0]] = cb;
      actBox.appendChild(el('label', { style: 'display:flex;align-items:center;gap:4px;cursor:pointer' }, [cb, a[1]]));
    });
    panel.appendChild(actBox);

    var locSel = null;
    if (locs.length) {
      locSel = el('select', { style: 'font-size:12px;padding:2px;max-width:100%', onchange: function () {
        var x = bestLoc([locs[+locSel.value].value].concat(locs.map(function (y) { return y.value; })));
        inputs.SITE.value = x.site; inputs.BATIMENT.value = x.bat; inputs.ETAGE.value = x.etage; inputs.BUREAU.value = x.bureau;
        highlight();
      } }, locs.map(function (x, i) { return el('option', { value: String(i), text: x.label + ' : ' + x.value }); }));
      panel.appendChild(el('div', { style: 'margin:6px 0;display:flex;gap:6px;align-items:center;flex-wrap:wrap' }, [el('b', { text: 'Localisation : ' }), locSel]));
    } else {
      panel.appendChild(el('div', { style: 'margin:6px 0;color:#b00;font-weight:bold', text: '⚠ Aucune localisation trouvée dans COSMOS : remplis site, bâtiment, étage et bureau à la main.' }));
    }

    GROUPS.forEach(function (g) {
      var box = el('fieldset', { style: 'border:1px solid #ccd;border-radius:6px;margin:6px 0;padding:4px 8px' },
        [el('legend', { style: 'font-weight:bold;color:#0b5394', text: g.title })]);
      g.fields.forEach(function (f) {
        var inp = el('input', { type: 'text', style: 'flex:1;font-size:13px;padding:2px 4px;border:1px solid #bbb;border-radius:3px', oninput: highlight });
        inp.value = F[f[0]] || '';
        inputs[f[0]] = inp;
        box.appendChild(el('label', { style: 'display:flex;align-items:center;gap:6px;margin:2px 0' },
          [el('span', { style: 'width:130px;color:#444', text: f[1] }), inp]));
      });
      box.dataset.types = g.types;
      groupEls.push(box);
      panel.appendChild(box);
    });

    // tableau des actifs
    if (assets.length) {
      var tb = el('tbody');
      assets.forEach(function (a) {
        var btns = el('td', { style: 'white-space:nowrap' });
        if (a.kind !== 'OTHER') {
          btns.appendChild(el('button', { text: '→ installé', title: 'Utiliser comme matériel installé', style: 'font-size:11px;cursor:pointer', onclick: function () { useAsset(a, 'I'); } }));
          btns.appendChild(el('button', { text: '→ restitué', title: 'Utiliser comme matériel restitué', style: 'font-size:11px;cursor:pointer;margin-left:2px', onclick: function () { useAsset(a, 'R'); } }));
        }
        tb.appendChild(el('tr', { style: 'border-top:1px solid #eee' }, [
          el('td', { text: a.cat }), el('td', { text: a.model }), el('td', { text: a.sn }),
          el('td', { text: a.code }), el('td', { text: a.status + (a.reserved ? ' · réservé' : '') }), btns
        ]));
      });
      panel.appendChild(el('div', { style: 'margin-top:8px' }, [
        el('b', { text: 'Matériel affecté à / réservé pour ' + benef + ' :' }),
        el('table', { style: 'width:100%;border-collapse:collapse;font-size:11px;margin-top:3px' }, [
          el('thead', {}, [el('tr', { style: 'text-align:left;color:#666' },
            ['Catégorie', 'Modèle', 'N° série', 'Code', 'État', ''].map(function (h) { return el('th', { text: h }); }))]),
          tb
        ])
      ]));
    } else {
      warnings.push('Aucun actif trouvé affecté à / réservé pour ' + (benef || 'le bénéficiaire') + '.');
    }

    var msg = el('div', { style: 'margin-top:6px;font-weight:bold' });
    var forceBtn = el('button', { text: 'Créer quand même', style: 'display:none;margin-left:8px;font-size:12px;cursor:pointer', onclick: function () { doCreate(true); } });
    panel.appendChild(el('div', { style: 'margin-top:10px;display:flex;align-items:center;flex-wrap:wrap' }, [
      el('button', { text: '📄 Créer les PV', style: 'font-size:14px;padding:6px 14px;background:#0b5394;color:#fff;border:0;border-radius:5px;cursor:pointer', onclick: function () { doCreate(false); } }),
      forceBtn,
      el('button', { text: '✔ Marquer comme fait (PV à la main)', title: 'Met la pastille verte sans créer de fichier', style: 'margin-left:8px;font-size:12px;padding:5px 10px;background:#fff;color:#1a7f37;border:1px solid #1a7f37;border-radius:5px;cursor:pointer', onclick: markManual })
    ]));
    panel.appendChild(msg);
    panel.appendChild(el('div', { style: 'color:#666;font-size:12px;margin-top:4px', text: 'Excel doit être ouvert : il crée les fichiers dans Documents\\COSMOS\\NOM Prénom\\date, puis ouvre le dossier.' }));
    if (warnings.length) {
      panel.appendChild(el('details', { style: 'margin-top:6px;color:#a15c00;font-size:11px' },
        [el('summary', { text: 'Remarques (' + warnings.length + ')' })].concat(warnings.map(function (w) { return el('div', { text: '• ' + w }); }))));
    }
    refresh();

    function useAsset(a, side) {
      var p = side === 'I' ? 'PCI_' : 'PCR_';
      if (a.kind === 'PC') { inputs[p + 'MODELE'].value = a.model; inputs[p + 'SN'].value = a.sn; inputs[p + 'NETBIOS'].value = a.code; }
      else if (a.kind === 'DOCK') inputs[p + 'STATION'].value = a.sn;
      else if (a.kind === 'SCREEN') {
        var slot = (!inputs[p + 'ECRAN1'].value || inputs[p + 'ECRAN1'].value === a.sn) ? 'ECRAN1' : 'ECRAN2';
        inputs[p + slot].value = a.sn;
      } else if (a.kind === 'PHONE') {
        var q = side === 'I' ? 'IPI_' : 'IPR_';
        inputs[q + 'MODELE'].value = cleanPhoneModel(a.model); inputs[q + 'SN'].value = a.sn; inputs[q + 'IMEI'].value = a.imei;
      }
      highlight();
    }

    function selected() { return ACTIONS.map(function (a) { return a[0]; }).filter(function (k) { return checks[k].checked; }); }

    function refresh() {
      var sel = selected();
      groupEls.forEach(function (g) {
        var show = g.dataset.types === 'ALL' || g.dataset.types.split(' ').some(function (t) { return sel.indexOf(t) >= 0; });
        g.style.display = show ? '' : 'none';
      });
      highlight();
    }

    function highlight() {
      for (var k in inputs) {
        var empty = !inputs[k].value.trim();
        inputs[k].style.background = empty ? (REQUIRED.indexOf(k) >= 0 ? '#ffd9d9' : '#fff6d6') : '#fff';
        inputs[k].style.borderColor = empty && REQUIRED.indexOf(k) >= 0 ? '#c0392b' : '#bbb';
      }
    }

    function payload(sel) {
      var g = function (k) { return clean(inputs[k] ? inputs[k].value : ''); };
      var lines = ['#PVCOSMOS2', 'TICKET=' + number, 'ACTIONS=' + sel.join(',')];
      for (var k in inputs) lines.push(k + '=' + g(k).replace(/[\r\n=]+/g, ' '));
      return lines.join('\r\n');
    }

    function markManual() {
      var o = loadDone();
      o[number] = { d: fmtDate(new Date()), a: 'PV fait à la main', n: clean(inputs.NOM.value) + ' ' + clean(inputs.PRENOM.value) };
      saveDone(o);
      msg.style.color = '#1a7f37';
      msg.textContent = '✔ Ticket marqué « PV fait » (sans création de fichier).';
    }

    function doCreate(force) {
      var sel = selected();
      msg.style.color = '#b00';
      if (!sel.length) { msg.textContent = 'Coche au moins un document à créer.'; return; }
      if (!/^\d{2}\/\d{2}\/\d{4}$/.test(clean(inputs.DATE.value))) { msg.textContent = 'Date du dossier invalide : format JJ/MM/AAAA.'; return; }
      if (!clean(inputs.NOM.value)) { msg.textContent = 'Le nom est vide : il sert à nommer le dossier.'; return; }
      var missing = REQUIRED.filter(function (k) { return !clean(inputs[k].value); });
      if (missing.length && !force) {
        msg.textContent = '⚠ Localisation incomplète (' + missing.map(function (k) { return k.toLowerCase(); }).join(', ') +
          ') : ces champs sont obligatoires dans le PV. Complète-les, ou clique « Créer quand même ».';
        forceBtn.style.display = '';
        highlight();
        return;
      }
      var blob = new Blob([payload(sel)], { type: 'text/plain;charset=utf-8' });
      var url = URL.createObjectURL(blob);
      var a = TD.createElement('a');
      a.href = url; a.download = 'PVCOSMOS_' + number + '.txt'; a.style.display = 'none';
      TD.body.appendChild(a); a.click();
      setTimeout(function () { a.remove(); URL.revokeObjectURL(url); }, 2000);
      var o = loadDone();
      o[number] = { d: fmtDate(new Date()), a: sel.map(function (k) { return ACTIONS.filter(function (x) { return x[0] === k; })[0][1]; }).join(', '), n: clean(inputs.NOM.value) + ' ' + clean(inputs.PRENOM.value) };
      saveDone(o);
      forceBtn.style.display = 'none';
      msg.style.color = '#1a7f37';
      msg.textContent = '✔ Demande envoyée à Excel (' + sel.length + ' document(s)). Le dossier s\'ouvre dans quelques secondes. Ticket marqué « PV fait ».';
    }
  } catch (e) {
    status.style.color = '#b00';
    status.textContent = 'Erreur : ' + (e && e.message ? e.message : e);
  }
}());
