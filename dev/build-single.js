#!/usr/bin/env node
// Copies the server, src/Code.gs, to dist/Code.gs: the one file the teacher
// pastes into the Sheet's Apps Script editor. (The page no longer lives in
// Apps Script; it is built to index.html by dev/build-site.js.)
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const code = fs.readFileSync(path.join(root, 'src/Code.gs'), 'utf8');
fs.mkdirSync(path.join(root, 'dist'), { recursive: true });
fs.writeFileSync(path.join(root, 'dist/Code.gs'), code);
console.log('Wrote dist/Code.gs', `(${(code.length / 1024).toFixed(0)} KB, ${code.split('\n').length} lines)`);
