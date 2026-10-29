<div align="center">

# Bilibili slide extractor

**Turn the slides inside a Bilibili video into a real `.pptx` file, entirely in your browser.**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Chrome extension](https://img.shields.io/badge/platform-Chrome%20%2F%20Edge-4285F4.svg)](#installation)
[![Manifest V3](https://img.shields.io/badge/manifest-v3-34A853.svg)](manifest.json)
[![Zero dependencies](https://img.shields.io/badge/dependencies-none-FB7299.svg)](#how-it-works)

**English** | [简体中文](README.zh-CN.md)

</div>

## What it does

Given a Bilibili video page, the extension seeks through the video frame by frame, detects every slide change, and exports one settled frame per slide as a PowerPoint file. It works directly on the page's playing stream: no video download, no server, no third-party library, and the export quality matches whatever quality you are watching.

## Features

- **Automatic slide detection.** Frames are compared as downscaled grayscale signatures; a change above a configurable threshold marks a new slide.
- **Duplicate marking.** When a lecturer flips back to an earlier slide, the revisit is flagged as a suspected duplicate and left unchecked by default.
- **Region cropping.** If a talking head, watermark, or corner overlay covers part of the frame, drag a crop box around the slide area; detection and export then use that region only.
- **Time range.** Scan a segment (for example the lecture part after the intro) instead of the full video.
- **Adjustable sampling.** Sampling interval from 1 to 10 seconds and a change threshold you can tune per video.
- **Two export formats.** A ready-to-open PPTX, or a ZIP of numbered JPEG images with timestamps in the file names.

## Installation

1. Download or clone this repository.
2. Open `chrome://extensions` (or `edge://extensions` in Edge).
3. Enable **Developer mode**.
4. Click **Load unpacked** and select the repository folder.

The toolbar icon opens and closes the panel; on a supported video page the pill button is also available at the bottom right of the page.

## Usage

1. Open a video on `www.bilibili.com` and **switch the player to the highest available quality**, because frames are captured at the quality you are watching (width capped at 1920).
2. Click the pink **提取PPT** pill at the bottom right of the page.
3. Optionally adjust the settings:
   - **Sampling interval** (default 3 s). Lower it for fast slide changes, raise it for long videos.
   - **Change threshold** (default 5). Raise it when small motion triggers false pages, lower it when real pages are missed.
   - **Crop region.** Play the video to a point where the slide is visible, click **选取区域…**, and drag a box.
   - **Time range.** Values such as `90`, `1:30`, or `1:00:30` are accepted; an empty end field means "until the end".
4. Click **开始扫描**. The video pauses and the extension starts seeking; thumbnails appear live as pages are recognized.
5. Click a thumbnail to toggle it, double-click to preview it large. Uncheck anything you do not want.
6. Click **⬇ 导出 PPTX** and the file downloads.

Keep the tab open during the scan and do not change quality or switch parts mid-scan; the extension detects such changes and clears results.

## How it works

The extension runs entirely in the content script of the video page:

1. The page's own `<video>` element is driven programmatically: the player is paused, then `currentTime` is stepped through the requested range.
2. At each sample point a frame is drawn to a canvas. A 48 x 27 grayscale signature is computed for change detection, and a full-resolution JPEG (width capped at 1920) is kept as the candidate image.
3. Consecutive signatures are compared by mean absolute difference. A difference above the threshold starts a new segment; each segment is represented by its **last** frame, the one settled right before the next change.
4. Export builds the PPTX in the browser. `pptx-builder.js` implements CRC32, a store-only ZIP writer, and a minimal but complete OOXML package (presentation, slide master, slide layout, theme, one image per slide) with no dependencies. Danmaku and player UI never appear in the output because frames are read from the video stream itself, not from a screen capture.

## Permissions

| Permission | Why |
| --- | --- |
| `activeTab` | Grants temporary access to the current tab when you click the toolbar icon, used only to inject the script as a fallback. No standing host access. |
| `scripting` | Programmatic injection for the fallback path described above. |
| `storage` | Saves your settings (interval, threshold, crop, time range) in `chrome.storage.local`. The extension reads no site data. |

There are no `host_permissions`, no remote code, and no network calls. Content scripts are declared statically for `www.bilibili.com/video/*`, `/bangumi/play/*`, and `/festival/*`.

## Project structure

```
bili-ppt-extractor/
├── manifest.json           # MV3 manifest
├── background.js           # service worker: toolbar click, fallback injection
├── content.js              # scanning, change detection, dedup, cropping, UI, export
├── pptx-builder.js         # CRC32 + store-only ZIP + minimal OOXML PPTX writer
└── icons/                  # extension icons
```

## FAQ

**The exported images are blurry.**
Set the player to the highest quality before scanning; captures match the playing stream.

**Too many duplicates or false pages.**
Revisited slides are flagged and unchecked by default. If motion (a webcam, an animated watermark) creates false pages, raise the change threshold or crop to the slide area.

**The scan is slow.**
A 1-hour video at a 3-second interval needs roughly 1200 seeks, typically a few minutes. Raise the interval or restrict the time range.

**Do danmaku appear in the export?**
No. Frames are read from the video stream; danmaku and player UI are DOM overlays.

**Live streams and interactive videos?**
Not supported; the scan needs a finite, seekable stream.

## Legal disclaimer

This extension is a technical tool: it captures frames from the video stream that your browser is already playing, on your own machine, for personal study and archival. The slides inside a video are copyrighted by their authors (the uploader or original course provider). Extracting them does not transfer any rights. Whether you may keep, share, or reuse the exported material — for course review, note-taking, teaching, or anything commercial — is your responsibility to determine under the copyright of the original content and Bilibili's terms of service. Do not redistribute extracted slides without the copyright holder's permission.

## Privacy

The extension collects nothing. There are no network requests, no analytics, no telemetry. Your settings (interval, threshold, crop, time range) are stored locally in `chrome.storage.local` and never leave your device. Everything — frame capture, slide detection, PPTX assembly — happens in your browser tab and the exported file goes straight to your downloads folder.

## Contributing

Issues and pull requests are welcome. The project uses Conventional Commits (`feat:`, `fix:`, `docs:`, `chore:`), with commit messages written in English.

## License

[MIT](LICENSE), Copyright (c) 2026 chenxuhl
