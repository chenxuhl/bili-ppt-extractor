# Development

## Setup

None. The extension is pure vanilla JS with no build step and no
dependencies. Load it via `chrome://extensions` → Developer mode →
"Load unpacked" → select the repository root.

## Repository lint

```powershell
node scripts/lint-repo.mjs
```

Zero dependencies (Node >= 18). Checks JS syntax of every shipped
script, manifest permission budget, and repo invariants that encode past
bugs (timestamp attach, relative trim thresholds, storage API, CSP-safe
styling, XML escaping, orphan teardown). Run it before every commit; it
must exit 0.

## Structure test for the PPTX builder

The `pptx-builder.js` writer is hand-rolled OOXML with no in-tree test
(the chain was stripped in `ac7743d`; see PITFALLS). After any change to
`pptx-builder.js`, restore and run the chain from git history:

```powershell
git show 85fc49a:dev-test/make_slides.py     | Out-File -Encoding utf8 dev-test/make_slides.py
git show 85fc49a:dev-test/build_test_pptx.js | Out-File -Encoding utf8 dev-test/build_test_pptx.js
git show 85fc49a:dev-test/validate_pptx.py   | Out-File -Encoding utf8 dev-test/validate_pptx.py
python dev-test/make_slides.py       # 4 synthetic slides via Pillow
node   dev-test/build_test_pptx.js   # build with the real builder
python dev-test/validate_pptx.py     # zip integrity + python-pptx checks
```

## Real-video acceptance

Automated e2e never went green (see PITFALLS); acceptance is manual:

1. Load the extension in Edge or Chrome.
2. Open a course-style video with static slides, set the player to the
   highest quality.
3. Scan, review thumbnails (timecodes, revisit badges), export PPTX.
4. Open the PPTX and check page count, crop cleanliness, and that
   dark-theme slides keep their full frame.

For automation attempts: use Edge headless (branded Chrome 137+ ignores
`--load-extension`), launch without a startup URL, and create the page
target over CDP.

## Conventions

- Conventional Commits, commit messages in English.
- Version bumps: `fix` → patch, `feat` → minor, in `manifest.json`.
- README.md is the primary (English) document; README.zh-CN.md is the
  Chinese mirror that defers to it.
- Keep permissions at `["scripting", "activeTab", "storage"]`; no
  `host_permissions` (the lint enforces this).
- Any change to thresholds or detection logic needs a real-video sanity
  pass, not just the lint.
