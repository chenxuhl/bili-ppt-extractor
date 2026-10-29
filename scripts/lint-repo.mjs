#!/usr/bin/env node
/* Lint: manifest + JS syntax + repo invariants for bili-ppt-extractor.
 * Zero dependencies (Node >= 18). Run: node scripts/lint-repo.mjs
 * Exit code 0 = all pass; 1 = failures. Every error is agent-actionable:
 * WHAT failed, WHY it matters, HOW to fix. */

import { readFileSync, readdirSync, statSync, existsSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const errors = [];
const warnings = [];
const fail = (msg) => errors.push(msg);
const warn = (msg) => warnings.push(msg);
const read = (p) => readFileSync(join(ROOT, p), 'utf8');

// ---------- 1. JS syntax: node --check on every shipped script ----------
const manifest = JSON.parse(read('manifest.json'));
const jsFiles = [
  'background.js',
  'pptx-builder.js',
  ...manifest.content_scripts.flatMap((cs) => cs.js),
];
for (const f of jsFiles) {
  if (!existsSync(join(ROOT, f))) {
    fail(`MISSING FILE: ${f} is referenced by manifest.json but does not exist.
  Fix: restore the file or update manifest.json.`);
    continue;
  }
  try {
    execFileSync(process.execPath, ['--check', join(ROOT, f)], { stdio: 'pipe' });
  } catch (e) {
    fail(`SYNTAX ERROR in ${f}:
  ${String(e.stderr).split('\n').slice(0, 4).join('\n  ')}
  Fix: correct the syntax error before committing.`);
  }
}

// ---------- 2. Manifest invariants ----------
// MV3 permission budget: this extension needs only these three. Anything else
// (especially host_permissions) is a store-review and privacy regression.
const ALLOWED_PERMS = ['scripting', 'activeTab', 'storage'];
for (const p of manifest.permissions || []) {
  if (!ALLOWED_PERMS.includes(p)) {
    fail(`MANIFEST PERMISSION REGRESSION: "${p}" is not in the allowed set [${ALLOWED_PERMS.join(', ')}].
  Why: the extension reads the page's own <video> via content scripts; broad
  host permissions would break the minimal-permission posture documented in
  README (Permissions section) and risk Chrome Web Store review rejection.
  Fix: remove the permission, or add it to ALLOWED_PERMS here WITH a comment
  explaining why the feature cannot work without it.`);
  }
}
if (manifest.host_permissions) {
  fail(`MANIFEST has host_permissions: ${JSON.stringify(manifest.host_permissions)}.
  Why: content_scripts already declare the bilibili match patterns; host_permissions
  would grant silent access to those hosts at all times.
  Fix: delete host_permissions; add page patterns under content_scripts.matches.`);
}
if (manifest.manifest_version !== 3) {
  fail(`MANIFEST_VERSION: expected 3, got ${manifest.manifest_version}.`);
}

// ---------- 2b. Load-order sync: the script list is declared in TWO places ----------
// manifest.json content_scripts[].js and background.js executeScript files[]
// must both be exactly ['pptx-builder.js', 'content.js'] — content.js consumes
// the global BiliPptx that pptx-builder.js defines, so the order is load-bearing.
const EXPECTED_SCRIPTS = ['pptx-builder.js', 'content.js'];
const declaredLists = (manifest.content_scripts || []).map((cs) => cs.js || []);
for (const js of declaredLists) {
  if (JSON.stringify(js) !== JSON.stringify(EXPECTED_SCRIPTS)) {
    fail(`LOAD-ORDER DESYNC: manifest.json content_scripts declares ${JSON.stringify(js)}.
  Expected exactly ${JSON.stringify(EXPECTED_SCRIPTS)}.
  Why: content.js references the global BiliPptx at export time; if pptx-builder.js
  is missing or loads after content.js, every export throws "BiliPptx is not defined".
  Fix: restore the order in manifest.json AND background.js (both lists must match).`);
  }
}
const bg = read('background.js');
const filesMatch = bg.match(/files:\s*\[([^\]]*)\]/);
if (filesMatch) {
  const bgFiles = filesMatch[1].split(',').map((s) => s.trim().replace(/['"]/g, ''));
  if (JSON.stringify(bgFiles) !== JSON.stringify(EXPECTED_SCRIPTS)) {
    fail(`LOAD-ORDER DESYNC: background.js executeScript files = ${JSON.stringify(bgFiles)}.
  Expected exactly ${JSON.stringify(EXPECTED_SCRIPTS)} (must match manifest content_scripts).
  Why: the fallback injection path must load pptx-builder.js before content.js,
  same as the manifest path. Fix both declarations together.`);
  }
}
if (!/^\d+\.\d+\.\d+$/.test(manifest.version || '')) {
  warn(`Version "${manifest.version}" is not plain x.y.z; keep semver-style bumps.`);
}

// ---------- 3. Source invariants (grep-style, each encodes a past bug) ----------
const content = read('content.js');
const builder = read('pptx-builder.js');

// 3a. frame.t: captureAt must attach the timestamp (v1.0-1.1.0 shipped with
// every timecode showing 0:00 because of this exact omission).
if (!/frame\.t\s*=\s*t/.test(content)) {
  fail(`INVARIANT LOST: captureAt no longer attaches the timestamp ("frame.t = t").
  Why: finalizeSegment reads frame.t for thumbnail timecodes and export ZIP
  filenames; without it every page shows 0:00 (bug shipped v1.0.0-1.1.0,
  fixed in 1.1.1). See content.js captureAt / finalizeSegment.
  Fix: restore "if (frame) frame.t = t;" in captureAt.`);
}

// 3b. Trim thresholds must stay RELATIVE. The absolute-threshold version
// cropped 10% off every side of dark-theme slides.
if (!/EDGE_MAX\s*=\s*190/.test(content) || !/GAP\s*=\s*50/.test(content)) {
  fail(`INVARIANT LOST: trimOutputEdges must use the relative edge test
  (EDGE_MAX = 190, GAP = 50, per-edge comparison against 3px inside).
  Why: absolute luminance thresholds were tried and failed: DARK=200 cropped
  10% off each side of dark-background decks (edge rows are ~31 everywhere);
  mean- and median-based inner references failed on text-heavy dark decks.
  Fix: restore the per-edge relative test in trimOutputEdges (content.js).`);
}

// 3c. Storage: settings must live in chrome.storage.local, not page
// localStorage (site cleanup wipes it; isolated world cannot see page vars).
if (/localStorage/.test(content)) {
  fail(`REGRESSION: content.js references localStorage.
  Why: bilibili page JS shares that storage; keys get cleaned and collide.
  Fix: use chrome.storage.local via loadSaved()/persist().`);
}

// 3d. CSP: no inline style attributes in injected UI (bilibili CSP blocks them).
// Allow .style.property assignments used for dynamic values, but raw
// style="..." in HTML strings is a CSP bug.
if (/style="/.test(content)) {
  fail(`CSP RISK: content.js contains a literal style="..." attribute in HTML.
  Why: bilibili page CSP blocks inline style attributes; the element silently
  renders unstyled. Fix: use CSS classes defined in the CSS constant.`);
}

// 3e. pptx-builder esc() must escape the five XML specials including apostrophe.
if (!/replace\([^)]*'\/g,\s*'&apos;'\)/.test(builder)) {
  fail(`INVARIANT LOST: pptx-builder.js esc() no longer escapes single quotes
  to &apos;. Why: apostrophes in titles/attrs break XML well-formedness.
  Fix: restore .replace(/'/g, '&apos;') in esc().`);
}

// 3f. Orphan teardown must exist: stale content scripts after extension
// reload would leave a dead panel otherwise.
if (!/chrome\.runtime\.id/.test(content)) {
  fail(`INVARIANT LOST: orphan-script detection (chrome.runtime.id check) is gone.
  Why: after the extension is reloaded/updated, old content scripts keep
  running with dead chrome APIs; the orphan check tears down the stale UI.
  Fix: keep the orphanTimer block at the bottom of content.js.`);
}

// ---------- 4. Repo hygiene ----------
if (!existsSync(join(ROOT, 'LICENSE'))) warn('LICENSE missing.');
for (const f of readdirSync(ROOT)) {
  // mojibake detector: replacement chars or GBK-as-UTF8 artifacts in text files
  if (statSync(join(ROOT, f)).isFile() && /\.(md|gitignore|mjs|json|js)$/.test(f) || f === '.gitignore') {
    const raw = readFileSync(join(ROOT, f));
    const text = raw.toString('utf8');
    if (/\uFFFD|绯荤粺|鏉傞/.test(text)) {
      warn(`Possible mojibake (encoding damage) in ${f}; rewrite the file as UTF-8.`);
    }
  }
}

// ---------- Report ----------
console.log(`lint-repo: ${jsFiles.length} JS files, manifest, ${'invariants'} checked`);
for (const w of warnings) console.log(`  WARN  ${w.split('\n')[0]}`);
for (const e of errors) console.log(`  FAIL  ${e}`);
if (errors.length) {
  console.log(`\n${errors.length} failure(s).`);
  process.exit(1);
}
console.log('\nAll checks passed.');
