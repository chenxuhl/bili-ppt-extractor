# Design decisions

Each entry states what we chose, what we rejected, and the evidence that
forced the choice. Update this file when a decision is revisited.

## D1. Eager segmentation over anchor-based three-zone detection

Cut a new segment whenever consecutive-sample diff exceeds the user
threshold; identify build-up intermediates after the scan
(`mergeBuildups`).

Rejected: the three-zone anchor approach (hard cuts immediate, moderate
changes compared against a drifting segment anchor). It could never fire
on gradual page turns: each step differed ~8 against a threshold of 15,
and the anchor drifted onto the new page once it settled, so cumulative
distance never crossed the cut line. Real-time and time-sharing slides
were silently swallowed.

Evidence: commit `9544ff9`; the direction of failure matters — an extra
intermediate page costs a checkbox, a lost page costs the content.

## D2. Revisited pages default to unchecked, with a "back to page N" badge

Duplicates (`diff <= threshold * 1.4` against an earlier segment) stay in
the results, unchecked by default, with a badge naming their source page.

History (three states, two flips — do not cite `e6aed53` as current
policy; its commit message says the opposite of the final state):

- v1.0.0 (`0ed9b5d`): unchecked by default.
- `e6aed53`: flipped to checked — revisited pages (an intro diagram
  recurring as the closing summary) were vanishing from exports.
- `fcb1e30`: flipped back to unchecked — revisit detection tested
  accurate and revisits are rarely wanted; the badge keeps the manual
  override one click away.

The current state assumes detection accuracy; if revisit detection ever
regresses, the failure mode is silent page loss, and the flip history
above is the argument map for re-litigating.

## D3. Relative, per-edge residue detection in `trimOutputEdges`

A 1-2px output-canvas edge band is residue only if it is absolutely dim
(< 190) AND clearly darker than the same side 3px inside (gap > 50), cap
4px per side.

Rejected, with evidence from synthetic-frame simulation:

- Absolute threshold 200 (the original): dark-theme decks have edge
  luminance ~31 everywhere; all four sides cropped to the 10% cap.
- Inner reference by mean (inner < 120 skip): white text on dark decks
  pulls the center mean to ~133; misclassified as light-background.
- Inner reference by median: fails on text-heavy dark decks (code
  screenshots) where bright text covers more than half the frame.

The relative per-edge test holds for any background color because it asks
"does the edge look like the interior", not "is the interior white".
True dark letterbox edges are the source-level detector's job (D4).

## D4. Full-resolution, **per-frame** letterbox probing in `detectDarkEdges`

Probe row/column means at capture resolution (capped 1920), `DARK = 60`,
`EDGE_TOLERANCE = 3` for JPEG ringing, `MIN_INNER = 45%`, +1px inward
margin — **on every captured frame, no caching**.

Why full resolution: a real 1-2px video edge diluted to 0.1 of a column
on the original 64px probe canvas and never triggered detection
(commit `e6aed53`, measured luminance 17-97 on exported frames).

Why the 3px tolerance: JPEG edge ringing brightens the outermost column
(measured: x=639 mean 85 while the true 1px edge sits at x=638 mean 17).

Why per-frame (v1.2, replacing the per-resolution cache): letterbox
width drifts *within one video* — measured on the reported video:
L87/R88 in the first half, L82/R82 later, a bottom bar appearing only
in some segments. A cached rect is correct for exactly one moment:
widening leaks black bands of varying width into exports (the reported
bug), narrowing silently crops real content with no correction path.
The cache was also poisonable by a mid-scan dark slide and never
invalidated across scans of the same video. A self-heal on detected
residue was tried and rejected (commit `3f3d6cb`, reverted): it can
only observe the widening direction, and its trigger was blind to the
wide-drift cases it was written for (see the v1.2.0 code review).
Per-frame probing costs ~10ms against hundreds of ms per seek —
accuracy wins (user-set priority: accuracy over speed).

## D5. Programmatic seek of the page's own `<video>` element

The pipeline seeks the page's MSE blob stream, draws frames to canvas,
and builds the PPTX in the page. No network interception, no downloads,
no ffmpeg.

Why: keeps permissions at `scripting`/`activeTab`/`storage`, keeps the
extension dependency-free and store-reviewable, and works with DRM-free
 HLS/DASH streams the browser is already decoding.

## D6. Zero-dependency PPTX writer (`pptx-builder.js`)

CRC32 + store-only ZIP + minimal OOXML, UMD (CommonJS for tests, global
`BiliPptx` in the page). Slide size is inferred per D7; each image is
placed contain-fit, never stretched (mixing aspect ratios mid-scan by
switching quality used to distort later slides — the UI warns against
switching quality during a scan).

Why: no bundler step; the extension loads unpacked sources directly.
The cost is that this hand-rolled OOXML has no test protection in the
working tree (see PITFALLS: deleted test chain).

## D7. Median aspect ratio with preset snapping; per-slide contain placement

The slide size is inferred from the **median** aspect ratio of all selected
slides (not the first slide, which the pre-1.2 code used — an intro logo
hijacked the whole deck). If the median falls within 2% of a mainstream
ratio (16:9, 16:10, 3:2, 4:3, 1:1, 5:4, 9:16, 3:4), the page uses that
ratio exactly: 16:9 and 4:3 at PowerPoint's built-in page sizes
(12192000x6858000 = 13.33x7.5 in, 9144000x6858000 = 10x7.5 in); the rest
derived at a fixed long edge (landscape 12192000 EMU, portrait its flip),
so the geometry is exact even where PowerPoint shows "Custom". Anything
else: a custom page at the median ratio, long edge anchored at 12192000
EMU (landscape) / 9144000 EMU (portrait).

Each slide's image is placed contain-fit and centered — never stretched.
Rejection history:

- First slide's ratio (v1.1 and earlier): one outlier (intro logo, a
  partial-frame crop) fixed the page size for every slide.
- Stretch-to-fill: mixed-ratio decks visibly distorted off-ratio pages;
  also fights the per-frame edge self-healing (D4 follow-up), which
  legitimately produces slightly different per-page ratios after edge
  trimming.

Contain-fit means an off-ratio page shows slide-background margins
instead of distortion; that is the same trade PowerPoint itself makes
when inserting a mismatched image.
