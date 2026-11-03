/*
 * MON SECTEUR : favori COSMOS qui affiche les tâches actives de mes groupes
 * avec leur localisation (ville, bâtiment, étage, bureau), et des filtres :
 * villes, libres / à moi / collègues, type de demande, retard, recherche.
 *
 * Lecture seule : rien n'est modifié dans COSMOS.
 * Localisation : emplacement indiqué dans la demande (variables), sinon
 * emplacement du bénéficiaire, sinon ville dans le nom du bénéficiaire
 * générique ("A CREER_FC SCF BEZIERS").
 */
(async function () {
  'use strict';

  var VERSION = '1.0';
  var GROUP_QUERY = 'active=true^assignment_groupDYNAMICd6435e965f510100a9ad2572f2b47744';
  var STORE = 'monsecteur.v1';
  var TOP = window.top || window;
  var TD = TOP.document;
  var W = window;
  var fr = document.getElementById('gsft_main');
  if (fr && fr.contentWindow) W = fr.contentWindow;
  var HEX32 = /^[0-9a-f]{32}$/;

  var old = TD.getElementById('monsecteur-panel');
  if (old) { old.remove(); return; }

  /* ---------- utilitaires ---------- */

  function clean(s) {
    s = (s == null ? '' : String(s)).replace(/\s+/g, ' ').trim();
    return /^(--\s*aucun\s*--|\(vide\)|null|undefined)$/i.test(s) ? '' : s;
  }
  function norm(x) { return clean(x).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase(); }

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
    if (/^-?\d+$/.test(r.etage)) r.etage = String(parseInt(r.etage, 10));
    return r;
  }

  function load() { try { return JSON.parse(TOP.localStorage.getItem(STORE)) || {}; } catch (e) { return {}; } }
  function save(o) { try { TOP.localStorage.setItem(STORE, JSON.stringify(o)); } catch (e) { /* ignoré */ } }

  /* ---------- lectures COSMOS ---------- */

  var notes = [];

  function flat(r) {
    var o = {};
    for (var k in r) {
      var v = r[k];
      if (v && typeof v === 'object') { o[k] = clean(v.display_value); o[k + '$id'] = clean(v.value); }
      else o[k] = clean(v);
    }
    return o;
  }

  async function api(table, q, fields, limit) {
    var u = '/api/now/table/' + table + '?sysparm_display_value=all&sysparm_exclude_reference_link=true' +
      '&sysparm_limit=' + (limit || 1000) + '&sysparm_query=' + encodeURIComponent(q) +
      (fields ? '&sysparm_fields=' + encodeURIComponent(fields) : '');
    var r = await fetch(u, { credentials: 'same-origin', headers: { 'Accept': 'application/json', 'X-UserToken': W.g_ck || TOP.g_ck || '' } });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return ((await r.json()).result || []).map(flat);
  }

  async function jsonv2(table, q) {
    var r = await fetch('/' + table + '.do?JSONv2&displayvalue=all&sysparm_query=' + encodeURIComponent(q), { credentials: 'same-origin' });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return ((await r.json()).records || []).map(function (x) {
      var o = {};
      for (var k in x) {
        if (k.slice(0, 3) === 'dv_') continue;
        o[k] = clean(x['dv_' + k] != null && x['dv_' + k] !== '' ? x['dv_' + k] : x[k]);
        o[k + '$id'] = clean(x[k]);
      }
      return o;
    });
  }

  function chunks(a, n) { var out = []; for (var i = 0; i < a.length; i += n) out.push(a.slice(i, i + n)); return out; }

  async function byIds(table, ids, fields) {
    var out = [];
    ids = ids.filter(function (x, i) { return x && HEX32.test(x) && ids.indexOf(x) === i; });
    for (var c of chunks(ids, 80)) {
      var q = 'sys_idIN' + c.join(',');
      try { out = out.concat(await api(table, q, fields)); }
      catch (e) { out = out.concat(await jsonv2(table, q)); }
    }
    return out;
  }

  async function loadTasks() {
    var fields = 'sys_id,number,short_description,state,assigned_to,u_next_action_datetime,u_pending_reason,u_expected_action,' +
      'request_item,request_item.cat_item,request_item.requested_for,request_item.requested_for.location';
    var tasks;
    try {
      tasks = await api('sc_task', GROUP_QUERY + '^ORDERBYu_next_action_datetime', fields, 2000);
    } catch (e) {
      notes.push('API refusée (' + e.message + ') : lecture par une autre méthode.');
      tasks = await jsonv2('sc_task', GROUP_QUERY + '^ORDERBYu_next_action_datetime');
      var ritms = await byIds('sc_req_item', tasks.map(function (t) { return t['request_item$id']; }), 'sys_id,cat_item,requested_for');
      var users = await byIds('sys_user', ritms.map(function (r) { return r['requested_for$id']; }), 'sys_id,location');
      var R = {}, U = {};
      ritms.forEach(function (r) { R[r.sys_id$id || r.sys_id] = r; });
      users.forEach(function (u) { U[u.sys_id$id || u.sys_id] = u; });
      tasks.forEach(function (t) {
        var r = R[t['request_item$id']] || {};
        var u = U[r['requested_for$id']] || {};
        t['request_item.cat_item'] = r.cat_item || '';
        t['request_item.requested_for'] = r.requested_for || '';
        t['request_item.requested_for$id'] = r['requested_for$id'] || '';
        t['request_item.requested_for.location'] = u.location || '';
      });
    }
    return tasks;
  }

  /* Emplacements saisis dans les demandes (variables "Emplacement…") */
  async function loadTicketLocations(ritmIds) {
    var res = {};
    try {
      var opts = [];
      for (var c of chunks(ritmIds.filter(Boolean), 60)) {
        opts = opts.concat(await api('sc_item_option_mtom',
          'request_itemIN' + c.join(',') + '^sc_item_option.item_option_new.question_textLIKEmplacement',
          'request_item,sc_item_option.value,sc_item_option.item_option_new.question_text'));
      }
      var locIds = opts.map(function (o) { return o['sc_item_option.value$id'] || o['sc_item_option.value']; });
      var locs = await byIds('cmn_location', locIds, 'sys_id,name,full_name');
      var L = {};
      locs.forEach(function (l) { L[l['sys_id$id'] || l.sys_id] = l.full_name || l.name; });
      opts.forEach(function (o) {
        var rid = o['request_item$id'], lid = o['sc_item_option.value$id'] || o['sc_item_option.value'];
        var q = o['sc_item_option.item_option_new.question_text'] || '';
        if (!rid || !L[lid]) return;
        var prio = /arriv/i.test(q) ? 0 : 1;
        if (!res[rid] || prio < res[rid].prio) res[rid] = { prio: prio, loc: L[lid] };
      });
    } catch (e) {
      notes.push('Emplacements des demandes non lus (' + e.message + ') : localisation du bénéficiaire utilisée.');
    }
    return res;
  }

  async function me() {
    try {
      var r = await api('sys_user', 'sys_id=javascript:gs.getUserID()', 'sys_id,name', 1);
      if (r[0]) return { id: r[0]['sys_id$id'] || r[0].sys_id, name: r[0].name };
    } catch (e) { /* ignoré */ }
    var u = W.g_user || TOP.g_user;
    return { id: u && u.userID || '', name: u && u.fullName || '' };
  }

  function typeOf(t) {
    var s = (t['request_item.cat_item'] + ' ' + t.short_description).toLowerCase();
    if (/iphone|smartphone|t[ée]l[ée]phone/.test(s)) return 'iPhone';
    if (/arriv/.test(s)) return 'Arrivée';
    if (/d[ée]part/.test(s)) return 'Départ';
    if (/mutation/.test(s)) return 'Mutation';
    if (/poste de travail|ordinateur|\bpc\b|portable/.test(s)) return 'PC';
    return 'Autre';
  }

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

  var C = { bleu: '#0b5394', rouge: '#c0392b', orange: '#d68910', vert: '#1a7f37', gris: '#666', fond: '#f4f6f9' };
  var panel = el('div', {
    id: 'monsecteur-panel',
    style: 'position:fixed;inset:2vh 2vw;z-index:2147483647;background:#fff;color:#222;border:2px solid ' + C.bleu +
      ';border-radius:8px;box-shadow:0 8px 30px rgba(0,0,0,.35);font:13px Arial,sans-serif;display:flex;flex-direction:column;overflow:hidden'
  });
  var head = el('div', { style: 'display:flex;align-items:center;gap:10px;padding:8px 12px;background:' + C.bleu + ';color:#fff' });
  var title = el('b', { style: 'font-size:15px', text: 'MON SECTEUR v' + VERSION });
  var count = el('span', { style: 'opacity:.9' });
  head.appendChild(title); head.appendChild(count);
  head.appendChild(el('span', { style: 'flex:1' }));
  head.appendChild(el('button', { text: '⟳ Actualiser', style: 'cursor:pointer;border:0;border-radius:4px;padding:3px 8px', onclick: function () { panel.remove(); start(); } }));
  head.appendChild(el('button', { text: '✕', title: 'Fermer', style: 'cursor:pointer;border:0;background:none;color:#fff;font-size:18px', onclick: function () { panel.remove(); } }));
  var bar = el('div', { style: 'padding:8px 12px;background:' + C.fond + ';display:flex;flex-direction:column;gap:6px;border-bottom:1px solid #dde' });
  var body = el('div', { style: 'flex:1;overflow:auto;padding:0 12px 12px' });
  var status = el('div', { style: 'padding:20px;color:' + C.gris, text: 'Chargement des tâches de mes groupes…' });
  body.appendChild(status);
  panel.appendChild(head); panel.appendChild(bar); panel.appendChild(body);
  TD.body.appendChild(panel);

  start();

  async function start() {
    if (!TD.body.contains(panel)) TD.body.appendChild(panel);
    body.innerHTML = ''; bar.innerHTML = ''; notes = [];
    status.style.color = C.gris; status.textContent = 'Chargement des tâches de mes groupes…';
    body.appendChild(status);
    try {
      var user = await me();
      var tasks = await loadTasks();
      status.textContent = 'Lecture des emplacements (' + tasks.length + ' tâches)…';
      var tl = await loadTicketLocations(tasks.map(function (t) { return t['request_item$id']; }));

      var rows = tasks.map(function (t) {
        var benef = t['request_item.requested_for'];
        var generic = /A\s*CREER|^\W*-\s/i.test(benef);
        var tLoc = (tl[t['request_item$id']] || {}).loc || '';
        var uLoc = generic ? '' : t['request_item.requested_for.location'];
        var L = parseLoc(tLoc || uLoc);
        if (tLoc && uLoc) {
          var L2 = parseLoc(uLoc);
          if (norm(L2.site) === norm(L.site)) ['bat', 'etage', 'bureau'].forEach(function (k) { L[k] = L[k] || L2[k]; });
        }
        var nxtRaw = t['u_next_action_datetime$id'] || '';
        var nxt = nxtRaw ? new Date(nxtRaw.replace(' ', 'T') + 'Z') : null;
        var aId = t['assigned_to$id'] || '';
        return {
          id: t['sys_id$id'] || t.sys_id, number: t.number, benef: benef, generic: generic,
          type: typeOf(t), article: t['request_item.cat_item'], desc: t.short_description,
          state: t.state, pending: t.u_pending_reason, action: t.u_expected_action,
          assigned: t.assigned_to, mine: !!aId && aId === user.id, free: !aId,
          next: nxt, nextTxt: t.u_next_action_datetime, loc: L, locTxt: tLoc || uLoc
        };
      });
      // ville d'un bénéficiaire générique : ville connue présente dans son nom
      var cities = {};
      rows.forEach(function (r) { if (r.loc.site) cities[norm(r.loc.site)] = r.loc.site; });
      rows.forEach(function (r) {
        if (r.loc.site) return;
        var n = norm(r.benef);
        for (var k in cities) if (n.indexOf(k) >= 0) { r.loc.site = cities[k]; break; }
      });
      rows.forEach(function (r) { r.city = r.loc.site || '(ville inconnue)'; });
      render(rows, user);
    } catch (e) {
      status.style.color = C.rouge;
      status.textContent = 'Erreur : ' + (e && e.message ? e.message : e) + '. Envoie une capture de ce message.';
    }
  }

  function render(rows, user) {
    var st = load();
    st.villes = st.villes || [];
    st.qui = st.qui || 'tous';
    st.types = st.types || [];
    var search = '';
    var retardOnly = !!st.retard;

    var now = new Date();
    var today0 = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var today1 = new Date(today0.getTime() + 864e5);

    function chip(label, on, onclick, n) {
      return el('button', {
        style: 'cursor:pointer;border:1px solid ' + (on ? C.bleu : '#bbc') + ';background:' + (on ? C.bleu : '#fff') + ';color:' + (on ? '#fff' : '#222') +
          ';border-radius:14px;padding:2px 10px;font-size:12px', onclick: onclick
      }, [label + (n != null ? ' (' + n + ')' : '')]);
    }

    function filtered(skip) {
      var q = norm(search);
      return rows.filter(function (r) {
        if (skip !== 'villes' && st.villes.length && st.villes.indexOf(r.city) < 0) return false;
        if (skip !== 'qui' && st.qui === 'libres' && !r.free) return false;
        if (skip !== 'qui' && st.qui === 'moi' && !r.mine) return false;
        if (skip !== 'qui' && st.qui === 'libres+moi' && !(r.free || r.mine)) return false;
        if (skip !== 'qui' && st.qui === 'collegues' && (r.free || r.mine)) return false;
        if (skip !== 'types' && st.types.length && st.types.indexOf(r.type) < 0) return false;
        if (retardOnly && !(r.next && r.next < today0)) return false;
        if (q && norm([r.number, r.benef, r.assigned, r.article, r.desc, r.locTxt].join(' ')).indexOf(q) < 0) return false;
        return true;
      });
    }

    function countBy(list, key) {
      var o = {};
      list.forEach(function (r) { o[r[key]] = (o[r[key]] || 0) + 1; });
      return o;
    }

    function draw() {
      save(st);
      bar.innerHTML = '';
      // villes
      var cv = countBy(filtered('villes'), 'city');
      var allCities = Object.keys(countBy(rows, 'city')).sort();
      var lineV = el('div', { style: 'display:flex;flex-wrap:wrap;gap:4px;align-items:center' }, [el('b', { style: 'width:80px', text: 'Villes :' })]);
      lineV.appendChild(chip('Toutes', !st.villes.length, function () { st.villes = []; draw(); }));
      allCities.forEach(function (c) {
        var on = st.villes.indexOf(c) >= 0;
        lineV.appendChild(chip(c, on, function () {
          if (on) st.villes.splice(st.villes.indexOf(c), 1); else st.villes.push(c);
          draw();
        }, cv[c] || 0));
      });
      bar.appendChild(lineV);
      // affectation
      var fq = filtered('qui');
      var nb = { tous: fq.length, libres: 0, moi: 0, 'libres+moi': 0, collegues: 0 };
      fq.forEach(function (r) { if (r.free) nb.libres++; if (r.mine) nb.moi++; if (r.free || r.mine) nb['libres+moi']++; if (!r.free && !r.mine) nb.collegues++; });
      var lineA = el('div', { style: 'display:flex;flex-wrap:wrap;gap:4px;align-items:center' }, [el('b', { style: 'width:80px', text: 'Affectées :' })]);
      [['tous', 'Toutes'], ['libres', 'Libres (à prendre)'], ['moi', 'À moi'], ['libres+moi', 'Libres + à moi'], ['collegues', 'Collègues']].forEach(function (x) {
        lineA.appendChild(chip(x[1], st.qui === x[0], function () { st.qui = x[0]; draw(); }, nb[x[0]]));
      });
      bar.appendChild(lineA);
      // types + retard + recherche
      var ct = countBy(filtered('types'), 'type');
      var lineT = el('div', { style: 'display:flex;flex-wrap:wrap;gap:4px;align-items:center' }, [el('b', { style: 'width:80px', text: 'Type :' })]);
      lineT.appendChild(chip('Tous', !st.types.length, function () { st.types = []; draw(); }));
      ['iPhone', 'PC', 'Arrivée', 'Départ', 'Mutation', 'Autre'].forEach(function (t) {
        if (!ct[t] && st.types.indexOf(t) < 0) return;
        var on = st.types.indexOf(t) >= 0;
        lineT.appendChild(chip(t, on, function () {
          if (on) st.types.splice(st.types.indexOf(t), 1); else st.types.push(t);
          draw();
        }, ct[t] || 0));
      });
      var cb = el('input', { type: 'checkbox', onchange: function () { retardOnly = cb.checked; st.retard = retardOnly; draw(); } });
      cb.checked = retardOnly;
      lineT.appendChild(el('label', { style: 'margin-left:14px;display:flex;align-items:center;gap:4px;color:' + C.rouge }, [cb, 'En retard seulement']));
      var sInput = el('input', { type: 'search', placeholder: 'Rechercher (nom, n°, bâtiment…)', style: 'margin-left:auto;padding:3px 6px;min-width:220px', oninput: function () { search = sInput.value; drawTable(); } });
      sInput.value = search;
      lineT.appendChild(sInput);
      bar.appendChild(lineT);
      if (notes.length) bar.appendChild(el('div', { style: 'font-size:11px;color:' + C.orange, text: notes.join(' · ') }));
      drawTable();
    }

    function drawTable() {
      var list = filtered();
      count.textContent = list.length + ' tâche(s) affichée(s) sur ' + rows.length + ' — ' + (user.name || '');
      body.innerHTML = '';
      if (!list.length) { body.appendChild(el('div', { style: 'padding:20px;color:' + C.gris, text: 'Aucune tâche avec ces filtres.' })); return; }
      var groups = {};
      list.forEach(function (r) { (groups[r.city] = groups[r.city] || []).push(r); });
      Object.keys(groups).sort().forEach(function (city) {
        var g = groups[city];
        body.appendChild(el('div', { style: 'position:sticky;top:0;background:#fff;padding:10px 0 4px;font-weight:bold;font-size:14px;color:' + C.bleu + ';border-bottom:2px solid ' + C.bleu },
          [city + ' — ' + g.length + ' tâche(s)']));
        var tb = el('tbody');
        g.sort(function (a, b) { return (a.next ? a.next.getTime() : 9e15) - (b.next ? b.next.getTime() : 9e15); });
        g.forEach(function (r) {
          var color = !r.next ? C.gris : r.next < today0 ? C.rouge : r.next < today1 ? C.orange : '#222';
          var lieu = [r.loc.bat && 'Bât ' + r.loc.bat, r.loc.etage && 'Étg ' + r.loc.etage, r.loc.bureau && 'Bur ' + r.loc.bureau].filter(Boolean).join(' · ');
          var who = r.free ? el('span', { style: 'color:' + C.vert + ';font-weight:bold', text: 'Libre' })
            : r.mine ? el('span', { style: 'color:' + C.bleu + ';font-weight:bold', text: 'Moi' }) : el('span', { text: r.assigned });
          tb.appendChild(el('tr', { style: 'border-bottom:1px solid #eee;vertical-align:top' + (r.mine ? ';background:#eef4fb' : '') }, [
            el('td', { style: 'white-space:nowrap;color:' + color + ';font-weight:' + (color === C.rouge ? 'bold' : 'normal'), text: r.nextTxt ? r.nextTxt.slice(0, 16) : '—' }),
            el('td', { style: 'white-space:nowrap' }, [el('a', { href: '/sc_task.do?sys_id=' + r.id, target: '_blank', style: 'color:' + C.bleu, text: r.number })]),
            el('td', { text: r.type }),
            el('td', { text: r.benef }),
            el('td', { text: lieu, title: r.locTxt }),
            el('td', {}, [who]),
            el('td', { style: 'color:' + C.gris, text: [r.state, r.pending, r.action].filter(Boolean).join(' · ') }),
            el('td', { style: 'color:' + C.gris, text: r.desc })
          ]));
        });
        var W8 = ['11%', '9%', '6%', '18%', '15%', '12%', '15%', '14%'];
        body.appendChild(el('table', { style: 'width:100%;border-collapse:collapse;margin-top:4px;table-layout:fixed' }, [
          el('thead', {}, [el('tr', { style: 'text-align:left;color:' + C.gris + ';font-size:11px' },
            ['Prochaine action', 'N°', 'Type', 'Bénéficiaire', 'Bât / étage / bureau', 'Affecté à', 'État', 'Description'].map(function (h, i) { return el('th', { style: 'padding:4px 6px 4px 0;width:' + W8[i], text: h }); }))]),
          tb
        ]));
      });
      [].forEach.call(body.querySelectorAll('td'), function (td) { td.style.padding = td.style.padding || '4px 6px 4px 0'; });
    }

    draw();
  }
}());
