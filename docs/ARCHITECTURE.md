# Architecture

## Runtime topology

```
┌────────────────────────────────────────────────────────────┐
│ Browser (bilibili video page)                              │
│                                                            │
│  background.js (service worker, MV3)                       │
│    │ toolbar click ──► chrome.tabs.sendMessage(BILI_PPT_   │
│    │                     PPT_TOGGLE)                        │
│    │ fallback: chrome.scripting.executeScript(content.js)  │
│    ▼                                                       │
│  content.js (isolated world, document_idle)                │
│    ├─ pptx-builder.js loaded first (same isolated world)   │
│    ├─ finds page <video>, seeks/captures via canvas        │
│    ├─ shadow DOM panel (#bili-ppt-host, mode:open)         │
│    └─ exports: BiliPptx.buildPptx / buildZip ──► download  │
│                                                            │
│  chrome.storage.local: { settings, crop, hideFab }         │
└────────────────────────────────────────────────────────────┘
```

Dependency direction: `background.js` and `content.js` reference each
other only through the message contract (`BILI_PPT_TOGGLE`) and the
idempotent injection guard (`window.__BILI_PPT_TOGGLE__`).
`content.js` depends on `BiliPptx` from `pptx-builder.js`. Nothing
depends on `background.js`.

## Extraction pipeline (content.js)

| Step | Function | Notes |
|------|----------|-------|
| Find video | `findBestVideo()` | Scores all `<video>` elements: in-player selector match (+1000) then `videoWidth`. |
| Source region | `regionOf(v)` | User crop (normalized 0..1) → source-pixel rect. |
| Letterbox edges | `detectDarkEdges(v)` | Full-res probe (cap 1920), col/row mean < 60, 3px JPEG-ring tolerance, 45% inner guard, +1px margin. Cached per `videoWidth x videoHeight`; null results are NOT cached (open issue). |
| Seek | `seekTo(v, t)` | `seeked` event + `readyState >= 2` poll + 3s timeout; ±0.3s tolerated. |
| Capture | `captureAt(v, t)` | seekTo → 40ms → paint tick → `grabFrame`; attaches `frame.t` (invariant, see lint). |
| Frame | `grabFrame(v)` | Crop∩edges → scale (≤1920 wide) → `trimOutputEdges` (relative per-edge test: < 190 and 50 darker than 3px inside, cap 4px) → JPEG bytes + 48×27 grayscale signature. |
| Blank skip | `isBlank(sig)` | Signature stddev < 3 → skip entirely (no baseline update). |
| Segmentation | scan loop | Eager: `diff(prevSig, sig) > threshold` cuts immediately; segment frame = last sample before the cut. |
| Duplicate mark | `finalizeSegment(frame)` | `diff <= threshold * 1.4` vs any earlier segment → `dupOf`, unchecked by default, "回翻→N" badge. |
| Buildup merge | `mergeBuildups()` | Post-scan: same family (`diff <= th * 1.4`) and content < 65% of a later segment → uncheck, "中间态" badge. |
| Export | `exportPptx()` / `exportZip()` | Checked segments only; `BiliPptx.buildPptx` (slide size = first slide aspect: landscape cx 12192000 EMU, portrait cy 9144000 EMU) or ZIP of JPEGs named `pNN_MMmSSs.jpg` from `seg.time`. |

## UI structure (shadow DOM, `#bili-ppt-host`)

Panel is a single shadow root created by `buildUI()`. Automation-relevant
ids (stable contract — the e2e driver and any future tests depend on
them): `.bp-fab`, `#scanBtn`, `#stopBtn`, `#reScanBtn`, `#t0`, `#t1`,
`#step`, `#th`, `#cropBtn`, `#cntTotal`, `#cntSel`, `#thumbs`,
`#pptxBtn`, `#zipBtn`, `#fabToggle`, `#vinfo`. Scanning state is
observable via `#scanBtn` carrying the `hidden` class.

## Lifecycle and settings

- Settings (`step`, `threshold`, `startStr`, `endStr`), `crop`, and
  `hideFab` persist to `chrome.storage.local` (`loadSaved`/`persist`).
- Orphan detection: `chrome.runtime.id` checked every 5s; on failure the
  UI tears down and clears `window.__BILI_PPT_TOGGLE__` so a fresh
  injection reinitializes cleanly.
- Video watch: 2s interval re-runs `findBestVideo` and refreshes info;
  `loadstart` clears stale results (unless scanning).

## Known open issues

1. `detectDarkEdges` samples only the first captured frame per
   resolution — intro/logo framing can mis-set the crop for the whole
   scan (docs/DECISIONS.md D4).
2. Null edge-detection results are re-computed every frame (full-res
   probe cost on videos without dark edges).
3. `isBlank` can silently drop minimal-content pages; no UI surfacing
   of skip counts.
4. `.bp-fab.ready` CSS rule is empty (readiness glow lost in the UI
   redesign); the class is still toggled.
5. Mixed aspect ratios (quality switch mid-scan) stretch later slides
   to the first slide's aspect.
