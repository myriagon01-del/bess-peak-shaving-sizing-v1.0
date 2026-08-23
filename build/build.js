'use strict';
/* =====================================================================
   build.js — deterministic rebuild of the BESS peak-shaving sizing tool.
   Reads src/ (body.html, app.js, style.css, vendored libs) and the
   battery/PCS library template, assembles a single standalone index.html.
   ===================================================================== */
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const ASSETS = path.join(ROOT, 'src');
const LIB_TEMPLATE = path.join(ROOT, 'library', 'bess_library_template.xlsx');
const OUT = path.join(ROOT, 'index.html');

function read(p) { return fs.readFileSync(p, 'utf8'); }
function countOf(text, needle) { let n = 0, i = 0; while ((i = text.indexOf(needle, i)) !== -1) { n++; i += needle.length; } return n; }
function requireCount(label, text, needle, expected) {
  const got = countOf(text, needle);
  if (got !== expected) throw new Error('Anchor count mismatch [' + label + ']: "' + needle + '" expected ' + expected + ', got ' + got);
  return got;
}

/* ---- 1. read assets ---- */
const styleCss = read(path.join(ASSETS, 'style.css'));
const bodyHtml = read(path.join(ASSETS, 'body.html'));
const appJs = read(path.join(ASSETS, 'app.js'));
const chartJs = read(path.join(ASSETS, 'chart.umd.min.js'));
const sheetJs = read(path.join(ASSETS, 'xlsx.umd.min.js'));

if (!fs.existsSync(LIB_TEMPLATE)) throw new Error('Battery & PCS Excel template is missing: ' + LIB_TEMPLATE);
const libraryTemplateBase64 = fs.readFileSync(LIB_TEMPLATE).toString('base64');

if (!sheetJs || sheetJs.indexOf('xlsx.js') < 0) throw new Error('SheetJS (xlsx.umd.min.js) asset missing');
if (!chartJs || chartJs.length < 1000) throw new Error('Chart.js asset missing or empty');
if (chartJs.indexOf('4.4.1') < 0) throw new Error('Chart.js asset version marker missing');
console.log('[build] vendored SheetJS: ' + sheetJs.length + ' chars, Chart.js: ' + chartJs.length + ' chars');

requireCount('body-open', bodyHtml, '<body>', 1);
requireCount('body-close', bodyHtml, '</body>', 1);

/* ---- 2. compile-check the app JS without executing it ---- */
try {
  new Function(appJs);
  console.log('[build] app.js compiles');
} catch (e) {
  throw new Error('app.js does not compile: ' + e.message);
}

/* ---- 3. assemble ---- */
const head = '<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="UTF-8" />\n<meta name="viewport" content="width=device-width, initial-scale=1.0" />\n<title>BESS Peak-Shaving Sizing</title>\n\n<style>\n' + styleCss + '\n</style>\n</head>\n';
const bodyInner = bodyHtml.replace(/<\/body>\s*$/i, '').replace(/\s+$/, '');
const html = head + bodyInner + '\n'
  + '<script>\n' + sheetJs + '\n</script>\n'
  + '<script>\n' + chartJs + '\n</script>\n'
  + '<script>window.__BESS_LIBRARY_TEMPLATE_BASE64__="' + libraryTemplateBase64 + '";</script>\n'
  + '<script>\n' + appJs + '\n</script>\n'
  + '</body>\n</html>\n';

/* ---- 4. static DOM checks ---- */
const markup = head + bodyInner;
const idSet = new Set();
const idRe = /\bid\s*=\s*"([^"]+)"/g;
let m;
while ((m = idRe.exec(markup)) !== null) {
  if (idSet.has(m[1])) throw new Error('Duplicate DOM id: ' + m[1]);
  idSet.add(m[1]);
}
console.log('[build] unique DOM ids: ' + idSet.size);

const domRefRe = /\$\(\s*'([^']+)'\s*\)/g;
const missing = [];
while ((m = domRefRe.exec(appJs)) !== null) {
  if (!idSet.has(m[1])) missing.push(m[1]);
}
if (missing.length) throw new Error('Missing DOM references in app.js: ' + [...new Set(missing)].join(', '));
console.log('[build] all $() DOM references resolved');

requireCount('output-doctype', html, '<!DOCTYPE html>', 1);
if (!html.trimEnd().endsWith('</html>')) throw new Error('output does not end with </html>');

/* ---- 5. write output ---- */
fs.writeFileSync(OUT, html, 'utf8');
const stat = fs.statSync(OUT);
console.log('[build] wrote output: ' + OUT + ' (' + stat.size + ' bytes)');
console.log('PASS');
