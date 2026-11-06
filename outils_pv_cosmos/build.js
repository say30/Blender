// Génère installer_favori.html (favori "PV COSMOS") à partir de pv_cosmos.js
// Usage : npx terser pv_cosmos.js -c -m --ecma 2017 -o pv_cosmos.min.js && node build.js
const fs = require('fs');
const esc = s => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
const href = 'javascript:' + encodeURIComponent('void ' + fs.readFileSync(__dirname + '/pv_cosmos.min.js', 'utf8').trim().replace(/;$/, ''));
const tpl = fs.readFileSync(__dirname + '/installer_favori.template.html', 'utf8');
fs.writeFileSync(__dirname + '/installer_favori.html', tpl.split('%%HREF%%').join(esc(href)));
fs.writeFileSync(__dirname + '/favoris.json', JSON.stringify({ pv: href }));
console.log('installer_favori.html généré (' + href.length + ' caractères)');
