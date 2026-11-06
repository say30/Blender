/* PV COSMOS - pastilles : colore en vert le n° et le nom des tickets dont les PV
   ont été créés avec le favori PV COSMOS (bouton "Créer les PV").
   Lecture seule : l'extension lit la liste gardée par le favori dans ce
   navigateur (localStorage "pvcosmos.faits") et ne modifie rien dans COSMOS. */
(function () {
  'use strict';
  var KEY = 'pvcosmos.faits';
  var VERT = '#c8ecd0';

  function faits() {
    try { return JSON.parse(localStorage.getItem(KEY)) || {}; } catch (e) { return {}; }
  }

  function colorer(cellules, oui) {
    cellules.forEach(function (c) {
      if (!c) return;
      if (oui) c.style.setProperty('background-color', VERT, 'important');
      else c.style.removeProperty('background-color');
    });
  }

  function peindre() {
    var d = faits();
    // Listes : la case du n° de ticket et la case suivante (bénéficiaire)
    var liens = document.querySelectorAll('a');
    for (var i = 0; i < liens.length; i++) {
      var a = liens[i];
      var num = (a.textContent || '').trim();
      if (!/^(SCTASK|INC)\d{6,}$/.test(num)) continue;
      var td = a.closest('td');
      if (!td) continue;
      var info = d[num];
      var deja = td.getAttribute('data-pvc') === '1';
      if (info && !deja) {
        td.setAttribute('data-pvc', '1');
        colorer([td, td.nextElementSibling], true);
        var b = document.createElement('span');
        b.className = 'pvc-pastille';
        b.textContent = ' ✔';
        b.title = 'PV créés le ' + (info.d || '') + ' : ' + (info.a || '');
        b.style.cssText = 'color:#1a7f37;font-weight:bold;margin-left:3px';
        a.parentNode.insertBefore(b, a.nextSibling);
      } else if (!info && deja) {
        td.removeAttribute('data-pvc');
        colorer([td, td.nextElementSibling], false);
        var p = td.querySelector('.pvc-pastille');
        if (p) p.remove();
      }
    }
    // Formulaire d'un ticket : petit bandeau vert en haut
    var champ = document.getElementById('sys_displayValue');
    var numero = champ ? (champ.value || '').trim() : '';
    var bandeau = document.getElementById('pvc-bandeau');
    if (numero && d[numero]) {
      if (!bandeau) {
        bandeau = document.createElement('div');
        bandeau.id = 'pvc-bandeau';
        bandeau.style.cssText = 'position:fixed;top:4px;left:50%;transform:translateX(-50%);z-index:99999;' +
          'background:#1a7f37;color:#fff;font:bold 13px Arial,sans-serif;padding:4px 12px;border-radius:12px;' +
          'box-shadow:0 2px 6px rgba(0,0,0,.3);pointer-events:none';
        document.body.appendChild(bandeau);
      }
      bandeau.textContent = '✔ PV faits le ' + (d[numero].d || '') + ' (' + (d[numero].a || '') + ')';
    } else if (bandeau) {
      bandeau.remove();
    }
  }

  peindre();
  var attente = null;
  new MutationObserver(function () {
    if (attente) return;
    attente = setTimeout(function () { attente = null; peindre(); }, 300);
  }).observe(document.documentElement, { childList: true, subtree: true });
  window.addEventListener('storage', function (e) { if (e.key === KEY) peindre(); });
  setInterval(peindre, 3000);
})();
