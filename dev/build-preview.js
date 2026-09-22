#!/usr/bin/env node
// Builds dev/preview.html: the real Index/Styles/App pages with src/Code.gs
// running in the browser against a fake spreadsheet. Open it in a browser:
//   node dev/build-preview.js && open dev/preview.html?role=student
// See dev/mock-runtime.js for the URL switches. Synthetic data only.
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const read = p => fs.readFileSync(path.join(root, p), 'utf8');

require('child_process').execFileSync(process.execPath, [path.join(__dirname, 'build-rubrics.js')], { stdio: 'inherit' });
const index = read('src/Index.html');
const wrap = js => `<script>\n${js}\n</script>`;
const html = index
  .replace('<?= appTitle ?>', 'My PE Profile (preview)')
  .replace('<?!= fontsLink ?>', '')
  .replace("<?!= include('Styles') ?>", read('src/Styles.html'))
  .replace("<?!= include('Rubrics') ?>", read('src/Rubrics.html'))
  .replace("<?!= include('App') ?>", [wrap(read('dev/fake-sheets.js')), wrap(read('src/Code.gs')), wrap(read('dev/mock-runtime.js')), read('src/App.html')].join('\n'));

const out = path.join(__dirname, 'preview.html');
fs.writeFileSync(out, html);
console.log('Wrote', path.relative(root, out), `(${(html.length / 1024).toFixed(0)} KB)`);
