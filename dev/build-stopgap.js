#!/usr/bin/env node
// Builds a copy of the site for a second GitHub Pages host (a stopgap while the
// student web filter mis-classifies the main one). Output: dist/site/ with
// index.html and lesson1/index.html, the deck's app link pointing at that host.
//   node dev/build-stopgap.js https://fis-pe.github.io
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const host = (process.argv[2] || '').replace(/\/$/, '');
if (!/^https:\/\/[a-z0-9-]+\.github\.io$/.test(host)) { console.error('Usage: node dev/build-stopgap.js https://<org>.github.io'); process.exit(1); }
require('child_process').execFileSync(process.execPath, [path.join(__dirname, 'build-site.js')], { stdio: 'inherit' });
const out = path.join(root, 'dist/site');
fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(path.join(out, 'lesson1'), { recursive: true });
fs.copyFileSync(path.join(root, 'index.html'), path.join(out, 'index.html'));
const deck = fs.readFileSync(path.join(root, 'lesson1/index.html'), 'utf8').replace(/https:\/\/scottbain-dot\.github\.io\/8Profile\//g, host + '/');
fs.writeFileSync(path.join(out, 'lesson1/index.html'), deck);
fs.writeFileSync(path.join(out, '.nojekyll'), '');
console.log('Wrote dist/site for', host);
