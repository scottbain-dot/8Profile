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
  // The fake sign-in token carries a photo URL on Google's image host; serve a tiny SVG for it offline.
  await page.route('https://lh3.googleusercontent.com/**', route => route.fulfill({ contentType: 'image/svg+xml', body: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><rect width="10" height="10" fill="#cfe4f2"/></svg>' }));

  // Signed out → sign-in screen; the (fake) Google button signs in
  await page.goto(url('role=anon'));
  await page.waitForSelector('.signin');
  check('signed out shows the sign-in screen', (await page.textContent('.signin')).includes('Sign in with your'));
  await page.click('#fakegsi');
  await page.waitForSelector('.sbtn');
  check('sign-in leads to the profile (first visit → card builder)', await page.evaluate(() => document.body.classList.contains('dark')));

  // A token the server rejects → back to sign-in with a message
  await page.goto(url('role=expired').replace('reset=1&', ''));
  await page.waitForSelector('.signin');
  check('rejected token → sign-in screen with a message', (await page.textContent('.signin')).includes('sign-in expired'));

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

  // Lesson 1 prediction, from the profile panel
  check('profile shows the Lesson 1 call to action', (await page.textContent('.goals.l1')).includes('Predict yourself'));
  await page.click('#topredict');
  await page.waitForSelector('.pitem');
  check('13 items to rate', (await page.$$('.pitem')).length === 13);
  check('save disabled until complete', await page.$eval('#psave', b => b.disabled));
  await page.click('.tl[data-k="cv"][data-rating="strength"]');
  await page.click('.fq[data-k="cv"][data-freq="often"]');
  check('progress counts a completed item', (await page.textContent('#pdone')) === '1');
  check('item marked done', await page.$eval('#pi-cv', el => el.classList.contains('done')));
  const items = ['me', 'power', 'speed', 'throw', 'catch', 'strike', 'dribble', 'kick', 'balance', 'agility', 'jump', 'core'];
  for (let i = 0; i < items.length; i++) {
    const r = ['strength', 'neutral', 'work-on'][i % 3], f = ['often', 'sometimes', 'rarely'][i % 3];
    await page.click('.tl[data-k="' + items[i] + '"][data-rating="' + r + '"]');
    await page.click('.fq[data-k="' + items[i] + '"][data-freq="' + f + '"]');
  }
  check('13 of 13 enables save', !(await page.$eval('#psave', b => b.disabled)));
  await page.screenshot({ path: path.join(shots, 'predict.png'), fullPage: true });
  await page.click('#psave');
  await page.waitForSelector('.sumhero');
  check('summary lists strengths and work-ons', (await page.$$('.sumcard.g .slist li')).length === 5 && (await page.$$('.sumcard.r .slist li')).length === 4);
  await page.screenshot({ path: path.join(shots, 'summary.png'), fullPage: true });
  await page.click('#toprofile');
  await page.waitForSelector('.goals.l1');
  check('profile panel now shows the prediction with chips', (await page.$$('.goals.l1 .chip.g')).length === 5);
  await page.click('#tocompare');
  await page.waitForSelector('.crowi');
  check('compare stub: 13 rows, Combine pending', (await page.$$('.crowi')).length === 13 && (await page.$$('.crowi .pending')).length === 13);
  await page.screenshot({ path: path.join(shots, 'compare.png'), fullPage: true });
  await page.click('#toprofile'); await page.waitForSelector('.goals.l1');
  await page.click('#topredict'); await page.waitForSelector('.pitem');
  check('re-opening pre-fills the saved picks', (await page.$$('.tl.on')).length === 13 && (await page.$$('.fq.on')).length === 13 && !(await page.$eval('#psave', b => b.disabled)));
  await page.click('.tl[data-k="cv"][data-rating="work-on"]');
  await page.click('#psave');
  await page.waitForSelector('.sumhero');
  check('edit re-saves (strength moved to work-on)', (await page.$$('.sumcard.r .slist li')).length === 5);

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
  check('another student sees their own seeded prediction', (await page.$$('.goals.l1 .chip')).length >= 8);

  await page.goto(url('role=student').replace('reset=1&', ''));
  await page.waitForSelector('.ct-av img');
  check('photo from the sign-in token is shown', (await page.$$('.ct-av img')).length === 1);

  await page.goto(url('role=student&fail=1').replace('reset=1&', ''));
  await page.waitForSelector('.cardteaser');
  await page.click('#tocard');
  await page.waitForSelector('.sbtn');
  await page.click('.sbtn[data-k="team"]');
  await page.waitForSelector('.toast.err.show');
  check('failed save reverts to previous style', (await page.textContent('#archt')) === 'THE EXPLORER');

  for (const [role, needle] of [['teacher', 'Teacher view'], ['unknown', 'Not on the Grade 8 list'], ['wrong', 'Use your FIS account']]) {
    await page.goto(url('role=' + role).replace('reset=1&', ''));
    await page.waitForSelector('.msg');
    check(role + ' screen', (await page.textContent('.msg')).includes(needle));
    await page.screenshot({ path: path.join(shots, role + '.png') });
  }
  await page.click('#signout');
  await page.waitForSelector('.signin');
  check('sign out returns to the sign-in screen', true);
  await page.screenshot({ path: path.join(shots, 'signin.png') });
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto(url('role=student').replace('reset=1&', ''));
  await page.waitForSelector('.cardteaser');
  check('no horizontal scroll on phone', await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1));
  await page.screenshot({ path: path.join(shots, 'profile-phone.png'), fullPage: true });
  await page.click('#topredict'); await page.waitForSelector('.pitem');
  check('no horizontal scroll on phone (predict)', await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1));
  await page.screenshot({ path: path.join(shots, 'predict-phone.png'), fullPage: false });
  await page.click('#toprofile'); await page.waitForSelector('.cardteaser');
  await page.click('#tocard'); await page.waitForSelector('.sbtn');
  check('no horizontal scroll on phone (card)', await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1));
  await page.screenshot({ path: path.join(shots, 'card-phone.png'), fullPage: true });

  check('no page errors', errors.length === 0, errors.join(' | '));
  await browser.close();
  console.log(failed ? `\n${failed} smoke check(s) FAILED` : '\nsmoke passed');
  process.exit(failed ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
