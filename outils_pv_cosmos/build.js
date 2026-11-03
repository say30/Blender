// Génère installer_favori.html (favoris "PV COSMOS" et "MON SECTEUR")
// Usage : npx terser X.js -c -m --ecma 2017 -o X.min.js (pour pv_cosmos et mon_secteur) puis node build.js
const fs = require('fs');
const esc = s => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
const href = f => 'javascript:' + encodeURIComponent('void ' + fs.readFileSync(__dirname + '/' + f, 'utf8').trim().replace(/;$/, ''));
const tpl = fs.readFileSync(__dirname + '/installer_favori.template.html', 'utf8');
const pv = href('pv_cosmos.min.js'), sec = href('mon_secteur.min.js');
fs.writeFileSync(__dirname + '/installer_favori.html', tpl.split('%%HREF%%').join(esc(pv)).split('%%HREF_SECTEUR%%').join(esc(sec)));
fs.writeFileSync(__dirname + '/favoris.json', JSON.stringify({ pv: pv, secteur: sec }));
console.log('installer_favori.html généré (' + pv.length + ' / ' + sec.length + ' caractères)');
