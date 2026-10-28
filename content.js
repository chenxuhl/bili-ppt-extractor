/*!
 * B站视频PPT提取器 — content script
 * 原理：直接复用页面正在播放的 <video> 元素做程序化 seek 抽帧，
 * 用小尺寸灰度签名对比相邻帧检测"翻页"，每个稳定片段取最后一帧作为 PPT 页，
 * 纯浏览器内合成 PPTX（依赖 pptx-builder.js 提供的 BiliPptx）。
 */
(() => {
  // 幂等：background 程序化兜底注入时，重复执行 = 切换面板
  if (window.__BILI_PPT_TOGGLE__) {
    try { window.__BILI_PPT_TOGGLE__(); } catch (e) { /* ignore */ }
    return;
  }

  'use strict';

  // ---------------- 工具函数 ----------------
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const nextPaint = () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));

  function formatTime(sec) {
    sec = Math.max(0, Math.floor(sec || 0));
    const h = Math.floor(sec / 3600), m = Math.floor((sec % 3600) / 60), s = sec % 60;
    const mm = String(m).padStart(2, '0'), ss = String(s).padStart(2, '0');
    return h > 0 ? `${h}:${mm}:${ss}` : `${m}:${ss}`;
  }

  function parseTimeStr(str, fallback) {
    if (str == null) return fallback;
    const parts = String(str).trim().split(':').map((x) => x.trim());
    if (!parts.length || parts.some((p) => p === '' || isNaN(+p))) return fallback;
    let t = 0;
    for (const p of parts) t = t * 60 + parseFloat(p);
    return t;
  }

  function videoTitle() {
    let t = (document.title || '').replace(/_哔哩哔哩.*$/, '').trim();
    t = t.replace(/[\\/:*?"<>|]/g, '').slice(0, 80);
    return t || 'bilibili';
  }

  // ---------------- 状态 ----------------
  const LS_KEY = 'bilippt.settings.v1';
  const state = {
    video: null,
    host: null, shadow: null,
    panel: null, fab: null,
    scanning: false,
    abort: false,
    segments: [],   // {bytes,w,h,sig,time,checked,dupOf,url}
    crop: null,     // {x,y,w,h} 归一化 0..1
    settings: { step: 3, threshold: 5, startStr: '', endStr: '' },
  };
  try {
    const saved = JSON.parse(localStorage.getItem(LS_KEY) || '{}');
    if (saved) state.settings = Object.assign(state.settings, saved);
    if (saved.crop) state.crop = saved.crop;
  } catch (e) { /* ignore */ }

  function persist() {
    try {
      localStorage.setItem(LS_KEY, JSON.stringify({ ...state.settings, crop: state.crop }));
    } catch (e) { /* ignore */ }
  }

  // ---------------- 复用画布 ----------------
  const SIG_W = 48, SIG_H = 27;
  const capCanvas = document.createElement('canvas');
  const capCtx = capCanvas.getContext('2d', { willReadFrequently: false });
  const sigCanvas = document.createElement('canvas');
  sigCanvas.width = SIG_W; sigCanvas.height = SIG_H;
  const sigCtx = sigCanvas.getContext('2d', { willReadFrequently: true });

  // ---------------- 视频探测 ----------------
  function findBestVideo() {
    const candidates = [...document.querySelectorAll('video')];
    if (!candidates.length) return null;
    const inPlayer = (v) => !!v.closest('.bpx-player-video-wrap, .bilibili-player-video, #bilibili-player, .squirtle-video-page-player');
    const score = (v) => (inPlayer(v) ? 1000 : 0) + (v.videoWidth || 0);
    return candidates.sort((a, b) => score(b) - score(a))[0];
  }

  function regionOf(v) {
    const vw = v.videoWidth || 2, vh = v.videoHeight || 2;
    if (!state.crop) return { sx: 0, sy: 0, sw: vw, sh: vh };
    const c = state.crop;
    const sx = clamp(Math.round(c.x * vw), 0, vw - 2);
    const sy = clamp(Math.round(c.y * vh), 0, vh - 2);
    const sw = clamp(Math.round(c.w * vw), 2, vw - sx);
    const sh = clamp(Math.round(c.h * vh), 2, vh - sy);
    return { sx, sy, sw, sh };
  }

  function grabFrame(v) {
    const { sx, sy, sw, sh } = regionOf(v);
    const scale = Math.min(1, 1920 / sw);
    const cw = Math.max(2, Math.round(sw * scale));
    const ch = Math.max(2, Math.round(sh * scale));
    capCanvas.width = cw; capCanvas.height = ch;
    try {
      capCtx.drawImage(v, sx, sy, sw, sh, 0, 0, cw, ch);
      sigCtx.drawImage(v, sx, sy, sw, sh, 0, 0, SIG_W, SIG_H);
      var px = sigCtx.getImageData(0, 0, SIG_W, SIG_H).data;
    } catch (e) {
      return Promise.resolve(null); // 画布不可用（理论不会发生：blob 同源）
    }
    const sig = new Uint8Array(SIG_W * SIG_H);
    for (let i = 0; i < sig.length; i++) {
      const o = i * 4;
      sig[i] = (px[o] * 299 + px[o + 1] * 587 + px[o + 2] * 114) / 1000;
    }
    return new Promise((resolve) => {
      capCanvas.toBlob((blob) => {
        if (!blob) { resolve(null); return; }
        blob.arrayBuffer().then((ab) => resolve({ sig, bytes: new Uint8Array(ab), w: cw, h: ch }));
      }, 'image/jpeg', 0.9);
    });
  }

  function diff(a, b) {
    let sum = 0;
    for (let i = 0; i < a.length; i++) sum += Math.abs(a[i] - b[i]);
    return sum / a.length;
  }

  function seekTo(v, t) {
    return new Promise((resolve) => {
      let done = false, poll = null, timer = null;
      const finish = () => {
        if (done) return;
        done = true;
        v.removeEventListener('seeked', finish);
        if (poll) clearInterval(poll);
        if (timer) clearTimeout(timer);
        resolve();
      };
      if (Math.abs(v.currentTime - t) < 0.01 && v.readyState >= 2) { finish(); return; }
      v.addEventListener('seeked', finish);
      try { v.currentTime = t; } catch (e) { finish(); return; }
      poll = setInterval(() => {
        if (v.readyState >= 2 && Math.abs(v.currentTime - t) < 0.3) finish();
      }, 60);
      timer = setTimeout(finish, 3000);
    });
  }

  async function captureAt(v, t) {
    await seekTo(v, t);
    await sleep(40);
    await nextPaint();
    return grabFrame(v);
  }

  // ---------------- 扫描 ----------------
  function clearResults() {
    for (const s of state.segments) if (s.url) URL.revokeObjectURL(s.url);
    state.segments = [];
    renderThumbs();
  }

  function finalizeSegment(frame) {
    if (!frame) return;
    const th = state.settings.threshold;
    let dupOf = -1;
    for (let i = 0; i < state.segments.length; i++) {
      if (diff(state.segments[i].sig, frame.sig) <= th * 1.4) { dupOf = i; break; }
    }
    const seg = {
      bytes: frame.bytes, w: frame.w, h: frame.h, sig: frame.sig,
      time: frame.t, checked: dupOf < 0, dupOf,
      url: URL.createObjectURL(new Blob([frame.bytes], { type: 'image/jpeg' })),
    };
    state.segments.push(seg);
    appendThumb(seg);
  }

  async function startScan() {
    const v = state.video;
    if (!v) { toast('未检测到视频，请先播放'); return; }
    if (state.scanning) return;
    const dur = v.duration;
    if (!isFinite(dur) || dur <= 0) { toast('无法读取视频时长（直播流不支持）'); return; }
    if (v.videoWidth < 200) { toast('当前清晰度太低，请先在播放器切换到 720P 及以上再扫描'); return; }

    const start = clamp(parseTimeStr(state.settings.startStr, 0), 0, dur - 1);
    const end = clamp(parseTimeStr(state.settings.endStr, dur), start + 1, dur);
    const step = clamp(state.settings.step, 1, 10);
    const threshold = clamp(state.settings.threshold, 2, 20);

    clearResults();
    state.scanning = true;
    state.abort = false;
    setScanUI(true);
    try { v.pause(); } catch (e) { /* ignore */ }

    const total = Math.ceil((end - start) / step);
    const t0 = performance.now();
    let count = 0, prevSig = null, lastFrame = null, missed = 0;

    for (let t = start; t < end && !state.abort; t += step) {
      const cap = await captureAt(v, Math.min(t, end - 0.05));
      count++;
      if (cap) {
        if (prevSig) {
          if (diff(prevSig, cap.sig) > threshold) {
            finalizeSegment(lastFrame); // 上一段最后一帧（已稳定）
            lastFrame = cap;
          } else {
            lastFrame = cap; // 同段内持续后移，取靠近翻页前的稳定帧
          }
        } else {
          lastFrame = cap;
        }
        prevSig = cap.sig;
      } else {
        missed++;
      }
      updateProgress(count, total, t0);
    }
    if (state.abort) { if (lastFrame && prevSig) finalizeSegment(lastFrame); }
    else finalizeSegment(lastFrame);

    state.scanning = false;
    setScanUI(false);
    if (missed > 0) toast(`${missed} 次抽帧失败已跳过`);
    toast(state.abort ? `已停止，共识别 ${state.segments.length} 页` : `扫描完成，共识别 ${state.segments.length} 页`);
  }

  // ---------------- 导出 ----------------
  function downloadBlob(bytes, mime, filename) {
    const blob = bytes instanceof Blob ? bytes : new Blob([bytes], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 120000);
  }

  function exportPptx() {
    const chosen = state.segments.filter((s) => s.checked);
    if (!chosen.length) { toast('请至少勾选一页'); return; }
    try {
      const slides = chosen.map((s) => ({ bytes: s.bytes, w: s.w, h: s.h, ext: 'jpg' }));
      const pptx = BiliPptx.buildPptx(slides);
      downloadBlob(pptx, 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        `${videoTitle()}_PPT.pptx`);
      toast(`已导出 ${chosen.length} 页`);
    } catch (e) {
      toast('导出失败：' + (e && e.message ? e.message : e));
    }
  }

  function exportZip() {
    const chosen = state.segments.filter((s) => s.checked);
    if (!chosen.length) { toast('请至少勾选一页'); return; }
    try {
      const files = chosen.map((s, i) => ({
        name: `p${String(i + 1).padStart(2, '0')}_${formatTime(s.time).replace(/:/g, 'm')}s.jpg`,
        data: s.bytes,
      }));
      downloadBlob(BiliPptx.buildZip(files), 'application/zip', `${videoTitle()}_PPT图.zip`);
      toast(`已导出 ${chosen.length} 张图片`);
    } catch (e) {
      toast('导出失败：' + (e && e.message ? e.message : e));
    }
  }

  // ---------------- UI ----------------
  const CSS = `
    :host { all: initial; }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, "Segoe UI Variable", "Segoe UI", "Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC", sans-serif; }
    .num { font-family: ui-monospace, Consolas, monospace; }

    .bp-fab {
      position: fixed; right: 20px; bottom: 100px; z-index: 2147483000;
      display: flex; align-items: center; gap: 7px;
      padding: 11px 16px; border: none; border-radius: 999px; cursor: pointer;
      background: linear-gradient(135deg, #fb7299, #ff5075);
      color: #fff; font-size: 14px; font-weight: 700; letter-spacing: .5px;
      box-shadow: 0 6px 20px rgba(251,114,153,.45), inset 0 1px 0 rgba(255,255,255,.25);
      transition: transform .18s cubic-bezier(.34,1.56,.64,1), box-shadow .18s;
      animation: bp-in .4s cubic-bezier(.34,1.56,.64,1);
    }
    .bp-fab:hover { transform: translateY(-2px) scale(1.04); box-shadow: 0 10px 28px rgba(251,114,153,.55); }
    .bp-fab svg { width: 17px; height: 17px; }
    .bp-fab.ready::after {
      content: ""; position: absolute; top: -3px; right: -3px; width: 10px; height: 10px;
      border-radius: 50%; background: #35c26e; border: 2px solid #17171c;
    }
    @keyframes bp-in { from { opacity: 0; transform: translateY(14px); } to { opacity: 1; transform: none; } }

    .bp-panel {
      position: fixed; right: 14px; top: 66px; bottom: 14px; width: 400px; z-index: 2147483001;
      display: none; flex-direction: column;
      background: rgba(23,23,28,.96);
      border: 1px solid rgba(255,255,255,.09); border-radius: 18px;
      box-shadow: 0 24px 64px rgba(0,0,0,.55);
      color: #e7e7ea; font-size: 13px; overflow: hidden;
      backdrop-filter: blur(18px);
    }
    .bp-panel.open { display: flex; animation: bp-slide .32s cubic-bezier(.34,1.3,.64,1); }
    @keyframes bp-slide { from { opacity: 0; transform: translateX(24px); } to { opacity: 1; transform: none; } }

    .bp-head {
      display: flex; align-items: center; gap: 10px;
      padding: 14px 16px 12px; border-bottom: 1px solid rgba(255,255,255,.08);
    }
    .bp-logo {
      width: 30px; height: 30px; border-radius: 9px; flex: none;
      display: grid; place-items: center;
      background: linear-gradient(135deg, #fb7299, #ff5075);
      box-shadow: 0 4px 12px rgba(251,114,153,.4);
    }
    .bp-logo svg { width: 16px; height: 16px; }
    .bp-title { font-weight: 800; font-size: 14.5px; flex: 1; letter-spacing: .3px; }
    .bp-title small { display: block; font-weight: 400; font-size: 11px; color: #9a9aa5; margin-top: 1px; }
    .bp-x {
      width: 26px; height: 26px; border: none; border-radius: 8px; cursor: pointer;
      background: rgba(255,255,255,.08); color: #cfcfd6; font-size: 15px; line-height: 1;
      transition: background .15s;
    }
    .bp-x:hover { background: rgba(255,255,255,.16); }

    .bp-body { flex: 1; overflow-y: auto; padding: 12px 14px 14px; }
    .bp-body::-webkit-scrollbar { width: 8px; }
    .bp-body::-webkit-scrollbar-thumb { background: rgba(255,255,255,.14); border-radius: 4px; }

    .bp-info {
      display: flex; gap: 8px; align-items: center;
      padding: 9px 11px; margin-bottom: 12px;
      background: rgba(255,255,255,.05); border-radius: 10px;
      font-size: 12px; color: #b9b9c2;
    }
    .bp-info b { color: #fb7299; font-weight: 700; }
    .bp-dot { width: 7px; height: 7px; border-radius: 50%; background: #666; flex: none; }
    .bp-dot.ok { background: #35c26e; box-shadow: 0 0 8px rgba(53,194,110,.7); }

    .bp-sec { margin-bottom: 12px; }
    .bp-sec-title {
      font-size: 11px; font-weight: 700; letter-spacing: 1.5px; color: #84848f;
      text-transform: uppercase; margin: 0 0 8px 2px;
    }
    .bp-row { display: flex; align-items: center; gap: 8px; margin-bottom: 8px; }
    .bp-row label { width: 62px; flex: none; color: #c4c4cc; font-size: 12.5px; }
    .bp-row .val { width: 56px; text-align: right; color: #fb7299; font-size: 12px; font-weight: 700; flex: none; }
    input[type="range"] {
      flex: 1; appearance: none; height: 4px; border-radius: 2px;
      background: linear-gradient(90deg, #fb7299 var(--p, 30%), rgba(255,255,255,.12) var(--p, 30%));
      outline: none; cursor: pointer;
    }
    input[type="range"]::-webkit-slider-thumb {
      appearance: none; width: 14px; height: 14px; border-radius: 50%;
      background: #fff; border: 3px solid #fb7299; cursor: grab;
      box-shadow: 0 2px 6px rgba(0,0,0,.4);
    }
    .bp-tinputs { display: flex; gap: 8px; }
    .bp-tinputs input {
      flex: 1; min-width: 0; padding: 7px 10px; border-radius: 8px;
      border: 1px solid rgba(255,255,255,.12); background: rgba(255,255,255,.06);
      color: #eee; font-size: 12.5px; outline: none;
    }
    .bp-tinputs input:focus { border-color: #fb7299; }
    .bp-hint { font-size: 11px; color: #7c7c87; margin-top: 4px; line-height: 1.6; }

    .bp-btn {
      border: none; border-radius: 10px; cursor: pointer; font-weight: 700;
      transition: transform .15s, box-shadow .15s, background .15s;
    }
    .bp-btn:active { transform: scale(.97); }
    .bp-btn:disabled { opacity: .45; cursor: not-allowed; }
    .bp-btn.primary {
      background: linear-gradient(135deg, #fb7299, #ff5075); color: #fff;
      padding: 11px 18px; font-size: 14px;
      box-shadow: 0 6px 18px rgba(251,114,153,.4);
    }
    .bp-btn.primary:hover:not(:disabled) { box-shadow: 0 8px 24px rgba(251,114,153,.55); }
    .bp-btn.ghost {
      background: rgba(255,255,255,.08); color: #d9d9e0;
      padding: 11px 14px; font-size: 13px; border: 1px solid rgba(255,255,255,.1);
    }
    .bp-btn.ghost:hover:not(:disabled) { background: rgba(255,255,255,.14); }
    .bp-btn.small { padding: 7px 12px; font-size: 12px; background: rgba(255,255,255,.08); color: #d9d9e0; border: 1px solid rgba(255,255,255,.1); font-weight: 600; }
    .bp-btn.stop { background: rgba(255,77,79,.15); color: #ff6f71; border: 1px solid rgba(255,77,79,.35); padding: 11px 14px; font-size: 13px; }

    .bp-progress { display: none; margin-bottom: 12px; }
    .bp-progress.on { display: block; }
    .bp-bar { height: 7px; border-radius: 4px; background: rgba(255,255,255,.1); overflow: hidden; }
    .bp-bar > div {
      height: 100%; width: 0%; border-radius: 4px;
      background: linear-gradient(90deg, #fb7299, #ff9ab0, #fb7299);
      background-size: 200% 100%;
      animation: bp-shimmer 1.2s linear infinite;
      transition: width .3s;
    }
    @keyframes bp-shimmer { from { background-position: 0 0; } to { background-position: -200% 0; } }
    .bp-ptext { margin-top: 6px; font-size: 11.5px; color: #9a9aa5; display: flex; justify-content: space-between; }

    .bp-res-head { display: flex; align-items: center; justify-content: space-between; margin: 4px 0 8px; }
    .bp-res-head b { color: #fff; }
    .bp-res-head b em { color: #fb7299; font-style: normal; }
    .bp-thumbs { display: grid; grid-template-columns: 1fr 1fr; gap: 9px; }
    .bp-thumb {
      position: relative; border-radius: 10px; overflow: hidden; cursor: pointer;
      border: 2px solid rgba(255,255,255,.1);
      background: #000; transition: transform .15s, border-color .15s;
      animation: bp-pop .28s cubic-bezier(.34,1.4,.64,1);
    }
    @keyframes bp-pop { from { opacity: 0; transform: scale(.85); } to { opacity: 1; transform: none; } }
    .bp-thumb:hover { transform: translateY(-2px); }
    .bp-thumb img { display: block; width: 100%; aspect-ratio: 16/9; object-fit: cover; }
    .bp-thumb.off { opacity: .38; filter: grayscale(.6); }
    .bp-check {
      position: absolute; top: 6px; right: 6px; width: 18px; height: 18px;
      border-radius: 6px; border: 2px solid rgba(255,255,255,.85);
      display: grid; place-items: center; pointer-events: none;
      background: rgba(0,0,0,.35);
    }
    .bp-check svg { width: 11px; height: 11px; opacity: 0; }
    .bp-thumb.on .bp-check { background: #fb7299; border-color: #fb7299; }
    .bp-thumb.on .bp-check svg { opacity: 1; }
    .bp-tc {
      position: absolute; left: 6px; bottom: 6px;
      padding: 2px 7px; border-radius: 6px;
      background: rgba(0,0,0,.65); color: #ffd7e2;
      font-size: 10.5px; font-weight: 600; letter-spacing: .5px;
    }
    .bp-dup {
      position: absolute; top: 6px; left: 6px;
      padding: 2px 6px; border-radius: 6px;
      background: rgba(255,170,0,.9); color: #4a3200;
      font-size: 10px; font-weight: 800;
    }
    .bp-empty { grid-column: 1 / -1; text-align: center; color: #6d6d78; padding: 26px 10px; font-size: 12px; line-height: 1.8; border: 1px dashed rgba(255,255,255,.12); border-radius: 12px; }

    .bp-actions { display: flex; gap: 9px; margin-top: 14px; }
    .bp-actions .bp-btn { flex: 1; }

    .bp-foot { padding: 10px 16px; border-top: 1px solid rgba(255,255,255,.08); font-size: 11px; color: #767681; line-height: 1.7; }

    .bp-toast {
      position: fixed; left: 50%; bottom: 40px; transform: translateX(-50%) translateY(20px);
      z-index: 2147483002; background: rgba(30,30,36,.97); color: #fff;
      padding: 10px 18px; border-radius: 12px; font-size: 13px; font-weight: 600;
      border: 1px solid rgba(255,255,255,.14); box-shadow: 0 10px 32px rgba(0,0,0,.5);
      opacity: 0; pointer-events: none; transition: opacity .25s, transform .25s;
      max-width: 340px; text-align: center;
    }
    .bp-toast.show { opacity: 1; transform: translateX(-50%) translateY(0); }

    .bp-modal {
      position: fixed; inset: 0; z-index: 2147483003; display: none;
      align-items: center; justify-content: center;
      background: rgba(0,0,0,.6); backdrop-filter: blur(4px);
    }
    .bp-modal.open { display: flex; }
    .bp-modal-box {
      background: #1d1d23; border: 1px solid rgba(255,255,255,.12); border-radius: 16px;
      padding: 16px; width: 430px; max-width: calc(100vw - 32px); box-shadow: 0 24px 64px rgba(0,0,0,.6);
    }
    .bp-modal-box h3 { font-size: 14px; margin-bottom: 10px; color: #fff; }
    .bp-modal-box .bp-hint { margin-bottom: 10px; }
    .bp-cropwrap { position: relative; border-radius: 10px; overflow: hidden; background: #000; }
    .bp-cropwrap canvas { display: block; width: 100%; cursor: crosshair; touch-action: none; }
    .bp-modal-actions { display: flex; gap: 8px; justify-content: flex-end; margin-top: 12px; }

    .bp-lightbox {
      position: fixed; inset: 0; z-index: 2147483004; display: none;
      align-items: center; justify-content: center;
      background: rgba(0,0,0,.82); cursor: zoom-out;
    }
    .bp-lightbox.open { display: flex; }
    .bp-lightbox img { max-width: 92vw; max-height: 88vh; border-radius: 8px; box-shadow: 0 12px 48px rgba(0,0,0,.7); }
  `;

  const ICON_SLIDES = `<svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="12" rx="2"/><path d="M8 20h8M12 16v4"/><path d="M7 12l3-3 2 2 3-4" stroke-linejoin="round"/></svg>`;
  const ICON_FILM = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><rect x="3" y="4" width="18" height="16" rx="2.5"/><path d="M7 4v16M17 4v16M3 9h4M17 9h4M3 15h4M17 15h4" stroke-width="1.6"/></svg>`;
  const ICON_CHECK = `<svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"><path d="M4.5 12.5l5 5 10-11"/></svg>`;

  function buildUI() {
    const host = document.createElement('div');
    host.id = 'bili-ppt-host';
    document.documentElement.appendChild(host);
    const shadow = host.attachShadow({ mode: 'open' });

    const style = document.createElement('style');
    style.textContent = CSS;
    shadow.appendChild(style);

    const fab = document.createElement('button');
    fab.className = 'bp-fab';
    fab.innerHTML = ICON_FILM + '<span>提取PPT</span>';
    shadow.appendChild(fab);

    const panel = document.createElement('div');
    panel.className = 'bp-panel';
    panel.innerHTML = `
      <div class="bp-head">
        <div class="bp-logo">${ICON_SLIDES}</div>
        <div class="bp-title">PPT 提取器<small>B站视频 · 翻页识别</small></div>
        <button class="bp-x" title="收起">✕</button>
      </div>
      <div class="bp-body">
        <div class="bp-info"><span class="bp-dot" id="dot"></span><span id="vinfo">正在检测播放器…</span></div>

        <div class="bp-sec">
          <div class="bp-sec-title">扫描设置</div>
          <div class="bp-row"><label>采样间隔</label><input type="range" id="step" min="1" max="10" step="1"><span class="val num" id="stepV"></span></div>
          <div class="bp-row"><label>翻页阈值</label><input type="range" id="th" min="2" max="20" step="1"><span class="val num" id="thV"></span></div>
          <div class="bp-tinputs">
            <input id="t0" placeholder="起点 0:00" title="扫描起始时间，支持 90 / 1:30 / 1:00:30">
            <input id="t1" placeholder="终点（默认结尾）" title="扫描结束时间，留空=视频结尾">
          </div>
          <div class="bp-row" style="margin-top:8px;">
            <label>画面裁剪</label>
            <button class="bp-btn small" id="cropBtn">选取区域…</button>
            <span id="cropLabel" style="font-size:11.5px;color:#9a9aa5;">全画面</span>
          </div>
          <div class="bp-hint">阈值越小越容易判定为“新的一页”；摄像头/水印干扰画面时，用裁剪只保留 PPT 区域。</div>
        </div>

        <div class="bp-sec" style="display:flex;gap:9px;">
          <button class="bp-btn primary" id="scanBtn" style="flex:1.4;">开始扫描</button>
          <button class="bp-btn stop" id="stopBtn" style="display:none;">停止</button>
          <button class="bp-btn ghost" id="reScanBtn" title="重新扫描（清空当前结果）">重扫</button>
        </div>

        <div class="bp-progress" id="prog">
          <div class="bp-bar"><div id="bar"></div></div>
          <div class="bp-ptext"><span id="ptxt1">准备中…</span><span class="num" id="ptxt2"></span></div>
        </div>

        <div class="bp-res-head">
          <span>识别结果 <b><em id="cntTotal">0</em></b> 页 · 已选 <b id="cntSel">0</b></span>
          <span style="display:flex;gap:6px;">
            <button class="bp-btn small" id="allBtn">全选</button>
            <button class="bp-btn small" id="noneBtn">全不选</button>
          </span>
        </div>
        <div class="bp-thumbs" id="thumbs">
          <div class="bp-empty">还没有识别结果<br>点「开始扫描」，工具会自动逐段抽帧并找出每一页 PPT</div>
        </div>

        <div class="bp-actions">
          <button class="bp-btn primary" id="pptxBtn">⬇ 导出 PPTX</button>
          <button class="bp-btn ghost" id="zipBtn">图片 ZIP</button>
        </div>
      </div>
      <div class="bp-foot">提示：扫描期间请停留在本页面、不要切换清晰度；建议先把播放器切到最高清晰度再开始。<br>导出的 PPTX 每页即视频中出现的一帧幻灯片。</div>

      <div class="bp-toast" id="toast"></div>

      <div class="bp-modal" id="cropModal">
        <div class="bp-modal-box">
          <h3>选取 PPT 区域</h3>
          <div class="bp-hint">先播放到有 PPT 的画面，再在这里拖拽框选 PPT 所在区域（用于抽帧和变化检测）。</div>
          <div class="bp-cropwrap"><canvas id="cropCanvas"></canvas></div>
          <div class="bp-modal-actions">
            <button class="bp-btn small" id="cropClear">清除裁剪</button>
            <button class="bp-btn small" id="cropCancel">取消</button>
            <button class="bp-btn small" id="cropApply" style="background:#fb7299;color:#fff;border-color:#fb7299;">应用</button>
          </div>
        </div>
      </div>

      <div class="bp-lightbox" id="lightbox"><img id="lightboxImg" alt="预览"></div>
    `;
    shadow.appendChild(panel);

    state.host = host;
    state.shadow = shadow;
    state.panel = panel;
    state.fab = fab;
    wireEvents(shadow);
    syncSettingsToUI();
  }

  function el(id) { return state.shadow.getElementById(id); }

  function wireEvents(shadow) {
    state.fab.addEventListener('click', () => {
      state.panel.classList.toggle('open');
      updateInfo();
    });
    state.shadow.querySelector('.bp-x').addEventListener('click', () => state.panel.classList.remove('open'));

    // 采样间隔
    const step = el('step'), stepV = el('stepV');
    step.addEventListener('input', () => {
      state.settings.step = +step.value;
      stepV.textContent = step.value + 's';
      step.style.setProperty('--p', ((step.value - 1) / 9 * 100) + '%');
      persist();
    });
    // 阈值
    const th = el('th'), thV = el('thV');
    th.addEventListener('input', () => {
      state.settings.threshold = +th.value;
      thV.textContent = th.value;
      th.style.setProperty('--p', ((th.value - 2) / 18 * 100) + '%');
      persist();
    });
    el('t0').addEventListener('change', (e) => { state.settings.startStr = e.target.value.trim(); persist(); });
    el('t1').addEventListener('change', (e) => { state.settings.endStr = e.target.value.trim(); persist(); });

    el('scanBtn').addEventListener('click', startScan);
    el('stopBtn').addEventListener('click', () => { state.abort = true; });
    el('reScanBtn').addEventListener('click', () => { clearResults(); toast('结果已清空'); });

    el('allBtn').addEventListener('click', () => setAllChecked(true));
    el('noneBtn').addEventListener('click', () => setAllChecked(false));

    el('pptxBtn').addEventListener('click', exportPptx);
    el('zipBtn').addEventListener('click', exportZip);

    el('cropBtn').addEventListener('click', openCropEditor);
    el('cropClear').addEventListener('click', () => {
      state.crop = null; persist(); el('cropLabel').textContent = '全画面';
      el('cropModal').classList.remove('open');
      toast('已恢复全画面');
    });
    el('cropCancel').addEventListener('click', () => el('cropModal').classList.remove('open'));
    el('cropApply').addEventListener('click', applyCrop);

    el('lightbox').addEventListener('click', () => el('lightbox').classList.remove('open'));

    chrome.runtime.onMessage.addListener((msg) => {
      if (msg && msg.type === 'BILI_PPT_TOGGLE') {
        state.panel.classList.toggle('open');
        updateInfo();
      }
    });
  }

  function syncSettingsToUI() {
    const step = el('step'), th = el('th');
    step.value = clamp(state.settings.step, 1, 10);
    step.dispatchEvent(new Event('input'));
    th.value = clamp(state.settings.threshold, 2, 20);
    th.dispatchEvent(new Event('input'));
    el('t0').value = state.settings.startStr || '';
    el('t1').value = state.settings.endStr || '';
    el('cropLabel').textContent = state.crop ? '已设置' : '全画面';
  }

  // ---------------- 进度/信息 UI ----------------
  function updateInfo() {
    const v = state.video;
    const dot = el('dot'), info = el('vinfo');
    state.fab.classList.toggle('ready', !!v);
    if (!v) {
      dot.className = 'bp-dot';
      info.innerHTML = '未检测到视频，请先播放';
      el('scanBtn').disabled = true;
      return;
    }
    el('scanBtn').disabled = false;
    if (!isFinite(v.duration) || v.duration <= 0) {
      dot.className = 'bp-dot';
      if (v.readyState < 1) {
        info.innerHTML = '视频加载中，稍候…';
      } else {
        info.innerHTML = '检测到直播流，暂不支持';
      }
      el('scanBtn').disabled = true;
      return;
    }
    dot.className = 'bp-dot ok';
    const res = v.videoWidth ? ` · ${v.videoWidth}×${v.videoHeight}` : '';
    info.innerHTML = `<b>已就绪</b> 时长 ${formatTime(v.duration)}${res}` +
      (v.videoWidth && v.videoWidth < 720 ? ' · <span style="color:#ffb027">清晰度偏低，建议调高</span>' : '');
  }

  function setScanUI(on) {
    el('scanBtn').style.display = on ? 'none' : '';
    el('stopBtn').style.display = on ? '' : 'none';
    el('prog').classList.toggle('on', on);
    ['step', 'th', 't0', 't1', 'cropBtn', 'reScanBtn'].forEach((id) => { el(id).disabled = on; });
  }

  function updateProgress(count, total, t0) {
    const p = total ? count / total : 0;
    el('bar').style.width = (p * 100).toFixed(1) + '%';
    const elapsed = (performance.now() - t0) / 1000;
    const eta = count > 2 ? (elapsed / count) * (total - count) : 0;
    el('ptxt1').textContent = `抽帧 ${count}/${total}`;
    el('ptxt2').textContent = `已识别 ${state.segments.length} 页 · 剩余约 ${formatTime(eta)}`;
  }

  // ---------------- 缩略图 ----------------
  function renderThumbs() {
    const box = el('thumbs');
    box.innerHTML = '';
    if (!state.segments.length) {
      box.innerHTML = `<div class="bp-empty">还没有识别结果<br>点「开始扫描」，工具会自动逐段抽帧并找出每一页 PPT</div>`;
    } else {
      state.segments.forEach((s) => appendThumb(s));
    }
    updateCounts();
  }

  function appendThumb(seg) {
    const box = el('thumbs');
    const empty = box.querySelector('.bp-empty');
    if (empty) empty.remove();
    const d = document.createElement('div');
    d.className = 'bp-thumb ' + (seg.checked ? 'on' : 'off');
    d.innerHTML = `<img src="${seg.url}" alt="">
      <span class="bp-check">${ICON_CHECK}</span>
      <span class="bp-tc num">${formatTime(seg.time)}</span>
      ${seg.dupOf >= 0 ? '<span class="bp-dup">疑似重复</span>' : ''}`;
    d.addEventListener('click', () => {
      seg.checked = !seg.checked;
      d.className = 'bp-thumb ' + (seg.checked ? 'on' : 'off');
      updateCounts();
    });
    d.addEventListener('dblclick', (e) => { e.preventDefault(); openLightbox(seg); });
    // 单击 = 勾选/取消勾选；双击 = 放大预览
    box.appendChild(d);
    updateCounts();
  }

  function openLightbox(seg) {
    el('lightboxImg').src = seg.url;
    el('lightbox').classList.add('open');
  }

  function updateCounts() {
    el('cntTotal').textContent = state.segments.length;
    el('cntSel').textContent = state.segments.filter((s) => s.checked).length;
  }

  function setAllChecked(v) {
    for (const s of state.segments) s.checked = v;
    renderThumbs();
  }

  // ---------------- Toast ----------------
  let toastTimer = null;
  function toast(msg) {
    const t = el('toast');
    if (!t) return;
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), 2600);
  }

  // ---------------- 裁剪编辑器 ----------------
  let cropRect = null; // {x,y,w,h} 归一化
  let cropDrag = null;

  function openCropEditor() {
    const v = state.video;
    if (!v || !v.videoWidth) { toast('请先播放视频到有画面的位置'); return; }
    cropRect = state.crop ? { ...state.crop } : null;
    cropDrag = null;

    const canvas = el('cropCanvas');
    const scale = Math.min(1, 398 / v.videoWidth);
    canvas.width = Math.max(2, Math.round(v.videoWidth * scale));
    canvas.height = Math.max(2, Math.round(v.videoHeight * scale));
    try { v.pause(); } catch (e) { /* ignore */ }
    const ctx = canvas.getContext('2d');

    // 底图快照（每次重绘都从它拷贝，避免叠加残影）
    const base = document.createElement('canvas');
    base.width = canvas.width; base.height = canvas.height;
    try { base.getContext('2d').drawImage(v, 0, 0, base.width, base.height); }
    catch (e) { toast('画面抓取失败，请重试'); return; }

    const redraw = () => {
      ctx.drawImage(base, 0, 0);
      if (!cropRect) return;
      const x = cropRect.x * canvas.width, y = cropRect.y * canvas.height;
      const w = cropRect.w * canvas.width, h = cropRect.h * canvas.height;
      ctx.fillStyle = 'rgba(0,0,0,.55)';
      ctx.fillRect(0, 0, canvas.width, y);
      ctx.fillRect(0, y + h, canvas.width, canvas.height - y - h);
      ctx.fillRect(0, y, x, h);
      ctx.fillRect(x + w, y, canvas.width - x - w, h);
      ctx.strokeStyle = '#fb7299';
      ctx.lineWidth = 2;
      ctx.setLineDash([7, 5]);
      ctx.strokeRect(x, y, w, h);
      ctx.setLineDash([]);
    };
    redraw();

    const pos = (e) => {
      const r = canvas.getBoundingClientRect();
      return {
        x: clamp((e.clientX - r.left) / r.width, 0, 1),
        y: clamp((e.clientY - r.top) / r.height, 0, 1),
      };
    };
    canvas.onpointerdown = (e) => {
      e.preventDefault();
      try { canvas.setPointerCapture(e.pointerId); } catch (err) { /* ignore */ }
      const p = pos(e);
      cropDrag = { x0: p.x, y0: p.y };
      cropRect = null;
      redraw();
    };
    canvas.onpointermove = (e) => {
      if (!cropDrag) return;
      e.preventDefault();
      const p = pos(e);
      cropRect = {
        x: Math.min(cropDrag.x0, p.x), y: Math.min(cropDrag.y0, p.y),
        w: Math.abs(p.x - cropDrag.x0), h: Math.abs(p.y - cropDrag.y0),
      };
      redraw();
    };
    canvas.onpointerup = () => { cropDrag = null; };
    canvas.onpointercancel = () => { cropDrag = null; };

    el('cropModal').classList.add('open');
  }

  function applyCrop() {
    if (cropRect && cropRect.w > 0.05 && cropRect.h > 0.05) {
      state.crop = { x: cropRect.x, y: cropRect.y, w: cropRect.w, h: cropRect.h };
      persist();
      el('cropLabel').textContent = '已设置';
      toast('裁剪已应用');
    } else {
      toast('请先拖拽框选一个区域');
      return;
    }
    el('cropModal').classList.remove('open');
  }

  // ---------------- 视频监听 ----------------
  function watchVideo() {
    const v = findBestVideo();
    if (v !== state.video) {
      state.video = v;
      if (v) {
        v.addEventListener('loadstart', () => {
          if (!state.scanning) { clearResults(); }
          updateInfo();
        });
      }
      updateInfo();
    }
    if (state.panel && state.panel.classList.contains('open') && !state.scanning) updateInfo();
  }

  // ---------------- 启动 ----------------
  buildUI();
  window.__BILI_PPT_TOGGLE__ = () => {
    state.panel.classList.toggle('open');
    updateInfo();
  };
  setInterval(watchVideo, 2000);
  watchVideo();
})();
