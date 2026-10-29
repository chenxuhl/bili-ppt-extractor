# Pitfalls and environment quirks

Measured, reproduced, or paid-for lessons. Verify before assuming any of
these have changed.

## Browser quirks

### Branded Google Chrome 137+ silently ignores `--load-extension`

Stable-channel Chrome (verified on 156) does not load unpacked
extensions passed via `--load-extension`; it also warns
`--disable-extensions-except is not allowed in Google Chrome, ignoring`
without failing. Use Edge or a Chrome Canary/Dev channel for automated
extension testing. Manual "Load unpacked" in `chrome://extensions` works
everywhere.

### Headless mode cannot combine `--load-extension` with a startup URL

Chromium headless aborts with "Multiple targets are not supported in
headless mode". Launch without a URL and create the page target over CDP
(`Target.createTarget`) instead.

### Bilibili pages render fine headless

No captcha wall observed with a desktop user agent and `zh-CN` locale;
the page reaches the video player and iframes load normally.

### Page CSP blocks inline styles

Inline `style="..."` attributes in injected HTML silently render
unstyled on bilibili pages. Use CSS classes inside the shadow root. The
`lint-repo` script fails on any literal `style="` in content.js.

### Extension reload orphans content scripts

After the extension reloads/updates, old content scripts keep running
with dead `chrome.runtime` APIs. The orphan check
(`chrome.runtime.id` every 5s) tears the stale UI down; if you touch
lifecycle code, keep it. `lint-repo` enforces its presence.

## Algorithm pitfalls

### Absolute thresholds do not generalize across slide themes

Every luminance threshold calibrated on white-background slides failed
on dark decks (see DECISIONS D3). Any future edge/border detection must
be relative (compare a band against its interior), never absolute.

### Sub-pixel residue survives source-coordinate cropping

Even with correct source math, scaling redistributes a cropped 1-2px
edge onto the outermost output pixels (measured: 636px output with a
112-luminance right column, 180-luminance top row, on 249-white
content). That is why `trimOutputEdges` re-measures on the output canvas
per frame.

### MSE seeking is async and lossy

`currentTime` assignment resolves before the frame is painted; the
capture path waits for `seeked` + `readyState >= 2` + a paint tick, with
a 3s timeout. Seeks can also land slightly off-target (±0.3s tolerated).

### Content-surface heuristics are proxies

`isBlank` (signature stddev < 3) can drop a genuinely minimal page
(e.g. a full-screen "thanks" slide); `mergeBuildups` compares signature
stddev as a "fullness" proxy. Both are empirically tuned trade-offs;
changes need real-video validation, not just unit tests.

## Process and tooling

### The test chain lives in git history, not the working tree

`ac7743d` stripped `dev-test/` (structure tests + e2e harness) and
`dev-assets/`. Recover with:

```powershell
git show 85fc49a:dev-test/make_slides.py      > dev-test/make_slides.py
git show 85fc49a:dev-test/build_test_pptx.js  > dev-test/build_test_pptx.js
git show 85fc49a:dev-test/validate_pptx.py    > dev-test/validate_pptx.py
```

Note: two e2e driver spellings existed — `run_e2e.mjs` (added `a2af1d9`,
recoverable at `85fc49a`) and `run-e2e.mjs` (added `07a9f5c`) — both
deleted at `ac7743d`. Check both if you need the harness.

Also: the pre-rewrite Chinese-message commits `96cbe76` and `d5e3038`
are unreachable orphans (verifiable with `git cat-file`) and will
eventually be garbage-collected; cite the rewritten hashes `0ed9b5d` /
`2a04552` in docs.

The structure chain (Pillow slides -> real builder -> python-pptx
validation) is the only automated protection for the hand-rolled OOXML
writer. After touching `pptx-builder.js`, restore and run it, or
re-add it permanently before the next builder change.

### The e2e run never went green

The CDP driver (`85fc49a:dev-test/e2e/run_e2e.mjs`) reached: extension
loaded in headless Edge, bilibili page rendered, panel UI driven over
DevTools. It failed on flaky headless startup races. Real-video
acceptance to date is manual export inspection (which found the
letterbox and timecode bugs). Do not claim automated e2e coverage.

### PowerShell redirect corrupts UTF-8 files

`>` in Windows PowerShell writes UTF-16 (with BOM), breaking Python and
Node scripts extracted from git. Use
`[IO.File]::WriteAllText($path, $content, [Text.UTF8Encoding]::new($false))`
or `git show ... | Out-File -Encoding utf8NoBOM` equivalents.

### Verification toolchain on this machine

- Node: `C:\Users\chenxu.joey\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\node\bin\node.exe`
- Python (with python-pptx, Pillow): same runtime tree, `dependencies\python\python.exe`
- Edge: `C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`
- Repo lint: `node scripts/lint-repo.mjs` (zero dependencies)
