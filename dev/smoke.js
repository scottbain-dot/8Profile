#!/usr/bin/env node
// Headless Chromium run of dev/preview.html: greeting, style pick + save,
// and the non-student screens. Screenshots land in dev/shots/ (git-ignored).
//   node dev/build-preview.js && NODE_PATH=$(npm root -g) node dev/smoke.js
const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');
const root = path.join(__dirname, '..');
const url = q => 'file://' + path.join(root, 'dev/preview.html') + '?reset=1&latency=20&' + q;
const shots = path.join(root, 'dev/shots'); fs.mkdirSync(shots, { recursive: true });

(async () => {
  let failed = 0;
  const check = (name, ok, extra) => { console.log((ok ? '  ✓ ' : '  ✗ ') + name + (ok ? '' : ' ' + (extra || ''))); if (!ok) failed++; };
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1100, height: 900 } });
  const errors = []; page.on('pageerror', e => errors.push(e.message));

  // First visit, no style yet → lands on the dark "build your card" screen
  await page.goto(url('role=student'));
  await page.waitForSelector('.sbtn');
  check('first visit opens the card builder', await page.evaluate(() => document.body.classList.contains('dark')));
  check('greeted by first name', (await page.textContent('h1')).includes('Hi Sample,'));
  check('card shows the name', (await page.textContent('.cname')).includes('Sample One'));
  check('no style chosen yet', (await page.textContent('#archt')) === 'PICK A STYLE');
  await page.click('.sbtn[data-k="explorer"]');
  await page.waitForFunction(() => document.querySelector('#saved').textContent.includes('saved'));
  check('style saved', (await page.textContent('#archt')) === 'THE EXPLORER');
  await page.screenshot({ path: path.join(shots, 'card.png'), fullPage: true });
  await page.click('#toprofile2');
  await page.waitForSelector('.cardteaser');
  check('back to the light profile', !(await page.evaluate(() => document.body.classList.contains('dark'))));
  check('profile teaser shows the chosen style', (await page.textContent('.ct-arch')).includes('THE EXPLORER'));
  check('nine fundamental skills from the rubric bank', (await page.$$('[data-ov^="fund-"]')).length === 9);
  await page.click('[data-ov="fund-catch"]');
  await page.waitForSelector('.ov.show');
  check('ladder overlay shows four levels', (await page.$$('.ov.show .lad')).length === 4);
  await page.keyboard.press('Escape');
  await page.screenshot({ path: path.join(shots, 'profile.png'), fullPage: true });

  // Return visit with a style → straight to the profile
  await page.goto(url('role=student').replace('reset=1&', ''));
  await page.waitForSelector('.cardteaser');
  check('return visit opens the profile with the saved style', (await page.textContent('.ct-arch')).includes('THE EXPLORER'));
  await page.click('#tocard');
  await page.waitForSelector('.sbtn');
  check('View my card reopens the builder with the choice selected', (await page.getAttribute('.sbtn[data-k="explorer"]', 'aria-pressed')) === 'true');

  await page.goto(url('role=student2').replace('reset=1&', ''));
  await page.waitForSelector('.cardteaser');
  check('"Last, First" roster name → first name only', (await page.textContent('.ct-name')).trim().startsWith('Sample ') && !(await page.textContent('.ct-name')).includes('Two'));
  check('another student has their own choice', (await page.textContent('.ct-arch')).includes('THE IMPROVER'));

  await page.goto(url('role=student&photo=1').replace('reset=1&', ''));
  await page.waitForSelector('.ct-av img');
  check('photo shown when the directory returns one', (await page.$$('.ct-av img')).length === 1);

  await page.goto(url('role=student&fail=1').replace('reset=1&', ''));
  await page.waitForSelector('.cardteaser');
  await page.click('#tocard');
  await page.waitForSelector('.sbtn');
  await page.click('.sbtn[data-k="team"]');
  await page.waitForSelector('.toast.err.show');
  check('failed save reverts to previous style', (await page.textContent('#archt')) === 'THE EXPLORER');

  for (const [role, needle] of [['teacher', 'Teacher view'], ['unknown', 'Not on the Grade 8 list'], ['wrong', 'Use your FIS account'], ['anon', 'confirm who you are']]) {
    await page.goto(url('role=' + role).replace('reset=1&', ''));
    await page.waitForSelector('.msg');
    check(role + ' screen', (await page.textContent('.msg')).includes(needle));
    await page.screenshot({ path: path.join(shots, role + '.png') });
  }
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto(url('role=student').replace('reset=1&', ''));
  await page.waitForSelector('.cardteaser');
  check('no horizontal scroll on phone', await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1));
  await page.screenshot({ path: path.join(shots, 'profile-phone.png'), fullPage: true });
  await page.click('#tocard'); await page.waitForSelector('.sbtn');
  check('no horizontal scroll on phone (card)', await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1));
  await page.screenshot({ path: path.join(shots, 'card-phone.png'), fullPage: true });

  check('no page errors', errors.length === 0, errors.join(' | '));
  await browser.close();
  console.log(failed ? `\n${failed} smoke check(s) FAILED` : '\nsmoke passed');
  process.exit(failed ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
