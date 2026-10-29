# AGENTS.md

Entry point for AI coding agents working on this repository. This is a
map, not a manual: it tells you what exists, what is load-bearing, and
where the details live. Read it before changing anything.

## What this project is

A Chrome MV3 extension (zero dependencies, no build step) that extracts
presentation slides from Bilibili videos: it seeks the page's own
`<video>` element, detects slide changes with grayscale signatures, and
assembles a PPTX entirely in the browser. English-primary bilingual
README (README.md / README.zh-CN.md), MIT license.

## Files and their jobs

| File | Job |
|------|-----|
| `manifest.json` | MV3, v1.1.1. Permissions locked to `scripting` + `activeTab` + `storage`. No `host_permissions`, ever. |
| `content.js` | The whole product: video discovery, seek/capture pipeline, change detection, letterbox crop, shadow-DOM panel UI, export. ~1200 lines, one IIFE. |
| `background.js` | Service worker: toolbar click toggles the panel via message, with a `chrome.scripting` injection fallback. Script load order is declared HERE and in the manifest — keep both in sync (`pptx-builder.js` before `content.js`). |
| `pptx-builder.js` | Zero-dependency PPTX writer: CRC32 + store-only ZIP + minimal OOXML. UMD: `BiliPptx` global in the page, CommonJS for tests. |
| `scripts/lint-repo.mjs` | Repository lint. Run it before every commit; it must exit 0. |
| `docs/DECISIONS.md` | Design decisions with the evidence that forced them. |
| `docs/PITFALLS.md` | Browser quirks, algorithm traps, tooling gotchas. |

## The pipeline in one paragraph

`findBestVideo` picks the player video → user or auto crop narrows the
source region (`regionOf` ∩ `detectDarkEdges`, cached per resolution) →
`captureAt` seeks and grabs a frame (`grabFrame`: crop → scale →
`trimOutputEdges` residue trim → 48×27 signature) → blank frames are
skipped (`isBlank`) → a diff above the user threshold cuts a new segment
(eager segmentation) → `finalizeSegment` marks duplicates →
`mergeBuildups` unchecks build-up intermediates → export via
`BiliPptx.buildPptx` / `buildZip`. UI ids used by tests and automation:
`.bp-fab`, `#scanBtn`, `#stopBtn`, `#t0`, `#t1`, `#thumbs`, `#pptxBtn`.

## Non-negotiable invariants

`scripts/lint-repo.mjs` mechanically enforces these; understand them
before you fight the lint:

1. **Timestamp attach**: `captureAt` must set `frame.t` (shipped broken
   as 0:00 everywhere from v1.0.0 to 1.1.0).
2. **Relative edge thresholds**: `trimOutputEdges` compares each edge
   band against its interior; absolute luminance thresholds break
   dark-theme slides (docs/DECISIONS.md D3).
3. **chrome.storage.local**, never page `localStorage`.
4. **CSS classes, never inline `style="..."`** — bilibili CSP blocks it.
5. **Orphan teardown** (`chrome.runtime.id` check) must survive any
   lifecycle refactor.
6. **Permission budget**: `["scripting", "activeTab", "storage"]` only.
7. **pptx-builder esc()** escapes all five XML specials including `'`.

## Before you commit

```powershell
node scripts/lint-repo.mjs        # must exit 0
```

After touching `pptx-builder.js`, also run the structure test chain from
git history (docs/DEVELOPMENT.md has the exact commands — the chain is
not in the working tree by design of commit `ac7743d`).

After touching detection thresholds or crop logic: run a real-video scan
manually and look at the thumbnails. Threshold behavior does not show up
in unit tests (docs/PITFALLS.md, "Content-surface heuristics are
proxies").

## Conventions

- Conventional Commits, messages in English.
- `fix` → patch bump, `feat` → minor bump, in `manifest.json`.
- README.md (English) is authoritative; README.zh-CN.md mirrors it.
- Comments in content.js are Chinese; keep new ones bilingual-friendly
  (Chinese comments are fine).

## Where the bodies are buried

`docs/PITFALLS.md` — read it before: launching browsers for tests
(Chrome 137+ ignores `--load-extension`), extracting files from git
history in PowerShell (redirect corrupts UTF-8), or tuning any
luminance threshold (absolute thresholds never survive contact with
dark slides). `docs/DECISIONS.md` records what was tried and rejected so
you do not re-litigate settled questions.
