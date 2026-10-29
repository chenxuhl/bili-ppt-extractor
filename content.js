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

  // 扩展自身数据走 chrome.storage.local（不污染页面源，也不受站点清理影响）
  function loadSaved() {
    return new Promise((resolve) => {
      try {
        if (!chrome.storage || !chrome.storage.local) { resolve(null); return; }
        chrome.storage.local.get(['settings', 'crop'], (res) => {
          if (chrome.runtime.lastError) { resolve(null); return; }
          resolve(res || null);
        });
      } catch (e) { resolve(null); }
    });
  }

  function persist() {
    try {
      if (chrome.storage && chrome.storage.local) {
        chrome.storage.local.set({ settings: state.settings, crop: state.crop || null });
      }
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

  // 自动检测视频帧的暗边(letterbox/pillarbox 黑边)并返回有效内容区域。
  // 以原始分辨率逐列/行检测——超细黑边(1-2px)在任何缩放画布上都会被
  // 稀释到均值采不到,必须全分辨率。只读边缘条带,代价可忽略。
  // 返回 {sx,sy,sw,sh} 视频源坐标;无暗边时返回 null。
  function detectDarkEdges(v) {
    const vw = v.videoWidth, vh = v.videoHeight;
    if (!vw || !vh) return null;
    // 检测画布:原始分辨率(上限 1920 防极端)
    const W = Math.min(vw, 1920), H = Math.round(W * vh / vw);
    const c = document.createElement('canvas');
    c.width = W; c.height = H;
    const ctx = c.getContext('2d', { willReadFrequently: true });
    let px;
    try {
      ctx.drawImage(v, 0, 0, W, H);
      px = ctx.getImageData(0, 0, W, H).data;
    } catch (e) { return null; }
    const lum = (x, y) => {
      const o = (y * W + x) * 4;
      return (px[o] * 299 + px[o + 1] * 587 + px[o + 2] * 114) / 1000;
    };

    const DARK = 60;               // 深色列/行阈值(实测黑边亮度 17-97,内容白底 249)
    const MIN_INNER = 0.45;        // 裁后至少保留 45% 内容,防误裁深色课件

    // 边缘列/行:采样该列/行全部像素,均值 < DARK 即为暗边列/行。
    // 逐列推进(1px 步进),超细边也能检出。
    const colMean = (x) => {
      let s = 0, n = 0;
      for (let y = 0; y < H; y += 2) { s += lum(x, y); n++; }
      return s / n;
    };
    const rowMean = (y) => {
      let s = 0, n = 0;
      for (let x = 0; x < W; x += 2) { s += lum(x, y); n++; }
      return s / n;
    };
    // 边缘容差:JPEG 边缘振铃会让最外 1 列/行偏亮(如右缘 x=639 均值 85
    // 而真正的 1px 黑边在 x=638 均值 17)。所以每边看最外 3 列/行中
    // 是否存在暗列,存在则继续向内推进。
    const EDGE_TOLERANCE = 3;
    function darkAheadCol(base, dir) {
      for (let k = 1; k <= EDGE_TOLERANCE; k++) {
        const x = base + dir * k;
        if (x >= 0 && x < W && colMean(x) < DARK) return true;
      }
      return false;
    }
    function darkAheadRow(base, dir) {
      for (let k = 1; k <= EDGE_TOLERANCE; k++) {
        const y = base + dir * k;
        if (y >= 0 && y < H && rowMean(y) < DARK) return true;
      }
      return false;
    }
    let left = 0, top = 0, right = 0, bottom = 0;
    while (left < W * (1 - MIN_INNER)) {
      if (colMean(left) < DARK || darkAheadCol(left, 1)) { left++; continue; }
      break;
    }
    while (top < H * (1 - MIN_INNER)) {
      if (rowMean(top) < DARK || darkAheadRow(top, 1)) { top++; continue; }
      break;
    }
    while (right < W * (1 - MIN_INNER)) {
      if (colMean(W - 1 - right) < DARK || darkAheadCol(W - 1 - right, -1)) { right++; continue; }
      break;
    }
    while (bottom < H * (1 - MIN_INNER)) {
      if (rowMean(H - 1 - bottom) < DARK || darkAheadRow(H - 1 - bottom, -1)) { bottom++; continue; }
      break;
    }

    if (left === 0 && right === 0 && top === 0 && bottom === 0) return null;
    // 换算回视频源坐标(W 可能小于 vw),补 1px 余量宁可多裁不残留黑线
    const k = vw / W;
    const rect = {
      sx: Math.max(0, Math.floor(left * k) - (left > 0 ? 1 : 0)),
      sy: Math.max(0, Math.floor(top * (vh / H)) - (top > 0 ? 1 : 0)),
      sw: 0, sh: 0,
    };
    rect.sw = Math.min(vw - rect.sx, Math.round((W - left - right) * k) + (left > 0 ? 1 : 0) + (right > 0 ? 1 : 0));
    rect.sh = Math.min(vh - rect.sy, Math.round((H - top - bottom) * (vh / H)) + (top > 0 ? 1 : 0) + (bottom > 0 ? 1 : 0));
    if (rect.sw < vw * MIN_INNER || rect.sh < vh * MIN_INNER) return null;
    return rect;
  }

  function grabFrame(v) {
    let { sx, sy, sw, sh } = regionOf(v);
    // 自动去黑边:检测到的暗边进一步收缩源区域(在用户裁剪之后叠加)
    // 以 videoWidth/Height 为缓存 key:切换清晰度后重检
    const edgeKey = v.videoWidth + 'x' + v.videoHeight;
    if (!state._edge || state._edgeKey !== edgeKey) {
      state._edge = detectDarkEdges(v);
      state._edgeKey = edgeKey;
    }
    if (state._edge) {
      const e = state._edge;
      // 与用户裁剪区域求交集
      const x0 = Math.max(sx, e.sx), y0 = Math.max(sy, e.sy);
      const x1 = Math.min(sx + sw, e.sx + e.sw), y1 = Math.min(sy + sh, e.sy + e.sh);
      if (x1 - x0 > 16 && y1 - y0 > 16) { sx = x0; sy = y0; sw = x1 - x0; sh = y1 - y0; }
    }
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
        blob.arrayBuffer().then((ab) => resolve({ sig, bytes: new Uint8Array(ab), w: cw, h: ch, content: contentOf(sig) }));
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
      time: frame.t,
      // 疑似重复仍默认勾选(宁多勿少):回翻页用户可能也要;
      // 徽章提示存在,导出前用户可自行取消
      checked: true, dupOf,
      content: frame.content,
      url: URL.createObjectURL(new Blob([frame.bytes], { type: 'image/jpeg' })),
    };
    state.segments.push(seg);
    appendThumb(seg);
  }

  // 帧内容量：签名方差。空白/纯色/转场过渡帧方差极低，动画构建中间态也偏低
  function contentOf(sig) {
    let mean = 0;
    for (let i = 0; i < sig.length; i++) mean += sig[i];
    mean /= sig.length;
    let v = 0;
    for (let i = 0; i < sig.length; i++) { const d = sig[i] - mean; v += d * d; }
    return Math.sqrt(v / sig.length);
  }

  // 判定一帧是否值得作为候选：内容量过低 = 空白/纯色/淡入淡出中间态
  function isBlank(sig) {
    return contentOf(sig) < 3;
  }

  // 收尾合并:渐进构建兜底。
  // 双阈值已把多数动画中间态并入同段;若个别中间态差异恰好超过翻页阈值被切段,
  // 判定条件(两个都要满足,防误杀内容量天然偏低的正常页):
  //   a) 前段内容量明显低于后段(画面更空 < 60%)
  //   b) 两段画面"同源"——签名差异在翻页量级以下(是同一页的渐进变体),
  //      而不是两张完全不同的正常页
  function mergeBuildups() {
    const segs = state.segments;
    let changed = false;
    for (let i = 0; i < segs.length - 1; i++) {
      const cur = segs[i], nxt = segs[i + 1];
      if (!cur.checked || !nxt) continue;
      const sameFamily = diff(cur.sig, nxt.sig) < state.settings.threshold * 3;
      if (sameFamily &&
          cur.content != null && nxt.content != null &&
          cur.content < nxt.content * 0.6) {
        cur.checked = false;
        cur.buildupOf = i + 1;
        changed = true;
      }
    }
    if (changed) renderThumbs();
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
    let count = 0, prevSig = null, lastFrame = null, anchorSig = null, missed = 0;

    // 三区间判定:
    //   d > hardTh (剧变)      → 立即翻页(整页换内容)
    //   softTh < d <= hardTh    → 中间区,再看"与段锚帧 anchorSig"的累计差异:
    //                              动画:每步小,画面虽变但累计仍接近锚(内容在锚上叠加) → 同段;
    //                              渐进翻页:累计已远离锚 → 翻页
    //   d <= softTh             → 同段
    // 锚帧 = 段内画面与上一切段点最接近的帧(动画起点),它随段缓慢漂移到
    // 画面稳定处,但只有切段才重置。
    // 这样既不把动画中间态当独立页(空白页),也不吞掉渐进式翻页(丢页)。
    const softTh = threshold;
    const hardTh = threshold * 3;

    for (let t = start; t < end && !state.abort; t += step) {
      const cap = await captureAt(v, Math.min(t, end - 0.05));
      count++;
      if (cap) {
        // 空白/纯色/转场过渡帧:完全跳过——不更新代表帧也不更新比较基准,
        // 否则(a)转场黑帧会被导出成空白页,(b)黑帧到内容帧的差异被误判为翻页
        if (isBlank(cap.sig)) {
          updateProgress(count, total, t0);
          continue;
        }
        if (prevSig) {
          const d = diff(prevSig, cap.sig);
          if (!anchorSig) anchorSig = cap.sig;
          const dFromAnchor = diff(anchorSig, cap.sig);
          if (d > hardTh || (d > softTh && dFromAnchor > hardTh)) {
            finalizeSegment(lastFrame); // 上一段最后一帧（已稳定）
            lastFrame = cap;
            anchorSig = cap.sig;
          } else {
            lastFrame = cap; // 同段内(含动画步进)持续后移，取靠近翻页前的稳定帧
            // 段内画面稳定后(小差异)漂移锚,避免锚过老导致长段累积误判
            if (d <= softTh) anchorSig = cap.sig;
          }
        } else {
          lastFrame = cap;
          anchorSig = cap.sig;
        }
        prevSig = cap.sig;
      } else {
        missed++;
      }
      updateProgress(count, total, t0);
    }
    if (state.abort) { if (lastFrame && prevSig) finalizeSegment(lastFrame); }
    else finalizeSegment(lastFrame);
    mergeBuildups();

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
  // 剪辑室美学: 深炭底 + 胶片齿孔 + 取景框角标 + 等宽时间码;
  // B站粉只做"录制灯"点缀, 主色调是暗房里的中性灰。
  const CSS = `
    :host { all: initial; }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: "Segoe UI Variable", "Segoe UI", "Microsoft YaHei UI", "Microsoft YaHei", sans-serif; }
    .num { font-family: "Cascadia Code", ui-monospace, Consolas, monospace; font-feature-settings: "tnum"; }

    /* ---------- 悬浮按钮: 胶片吊牌 ---------- */
    .bp-fab {
      position: fixed; right: 22px; bottom: 110px; z-index: 2147483000;
      display: flex; align-items: center; gap: 10px;
      padding: 12px 18px 12px 14px; border: none; cursor: pointer;
      color: #f2f0ec; font-size: 13px; font-weight: 600; letter-spacing: 2.5px;
      background: #14140f;
      clip-path: polygon(10px 0, 100% 0, 100% calc(100% - 10px), calc(100% - 10px) 100%, 0 100%, 0 10px);
      box-shadow: 0 8px 24px rgba(0,0,0,.55);
      transition: transform .18s cubic-bezier(.34,1.56,.64,1), background .18s;
      animation: bp-in .45s cubic-bezier(.34,1.56,.64,1);
    }
    .bp-fab::before { /* 左缘齿孔(在 clip 内) */
      content: ""; width: 8px; align-self: stretch; margin-left: -6px;
      background:
        radial-gradient(circle 2px at 50% 25%, rgba(242,240,236,.8) 97%, transparent),
        radial-gradient(circle 2px at 50% 50%, rgba(242,240,236,.8) 97%, transparent),
        radial-gradient(circle 2px at 50% 75%, rgba(242,240,236,.8) 97%, transparent);
      background-size: 8px 100%; background-color: #050503; background-repeat: no-repeat;
    }
    .bp-fab:hover { transform: translateY(-3px); background: #1e1e15; }
    .bp-fab svg { width: 17px; height: 17px; color: #fb7299; }
    .bp-fab.ready svg { color: #fb7299; filter: drop-shadow(0 0 5px rgba(251,114,153,.8)); }
    @keyframes bp-in { from { opacity: 0; transform: translateY(16px); } to { opacity: 1; transform: none; } }

    /* ---------- 主面板: 暗房取景器 ---------- */
    .bp-panel {
      position: fixed; right: 16px; top: 64px; bottom: 16px; width: 426px; z-index: 2147483001;
      display: none; flex-direction: row;
      background: #17171a;
      border: 1px solid #2c2c33; border-radius: 6px;
      box-shadow: 0 32px 80px rgba(0,0,0,.65), inset 0 0 0 1px rgba(255,255,255,.02);
      color: #d8d6d0; font-size: 13px; overflow: hidden;
    }
    .bp-panel.open { display: flex; animation: bp-slide .34s cubic-bezier(.22,1.2,.36,1); }
    @keyframes bp-slide { from { opacity: 0; transform: translateX(28px); } to { opacity: 1; transform: none; } }

    /* 左缘胶片齿孔轨 */
    .bp-strip {
      width: 22px; flex: none; background: #0c0c0e;
      border-right: 1px solid #232329;
      background-image: radial-gradient(circle 4px at 50% 50%, #17171a 97%, transparent);
      background-size: 22px 26px;
      animation: bp-feed 14s linear infinite;
    }
    @keyframes bp-feed { from { background-position: 0 0; } to { background-position: 0 260px; } }
    .bp-main { flex: 1; min-width: 0; display: flex; flex-direction: column; position: relative; }
    /* 取景框角标 */
    .bp-main::before, .bp-main::after {
      content: ""; position: absolute; width: 14px; height: 14px; z-index: 5; pointer-events: none;
      border: 1.5px solid rgba(251,114,153,.5);
    }
    .bp-main::before { top: 8px; left: 8px; border-right: none; border-bottom: none; }
    .bp-main::after { bottom: 46px; right: 8px; border-left: none; border-top: none; }

    .bp-head {
      display: flex; align-items: center; gap: 12px;
      padding: 15px 18px 13px;
      border-bottom: 1px solid #26262d;
      background: linear-gradient(180deg, #1c1c21, #17171a);
    }
    .bp-logo {
      width: 34px; height: 34px; flex: none; display: grid; place-items: center;
      background: #0c0c0e; border: 1px solid #33333b; border-radius: 4px;
      box-shadow: inset 0 0 10px rgba(0,0,0,.6);
    }
    .bp-logo svg { width: 17px; height: 17px; }
    .bp-title { font-weight: 300; font-size: 16.5px; flex: 1; letter-spacing: 3px; color: #f2f0ec; }
    .bp-title small {
      display: block; font-weight: 400; font-size: 10px; color: #71717c;
      margin-top: 3px; letter-spacing: 1.5px;
      font-family: "Cascadia Code", ui-monospace, Consolas, monospace;
    }
    .bp-title small::before { content: "● "; color: #fb7299; animation: bp-blink 2.4s steps(1) infinite; }
    @keyframes bp-blink { 0%, 70% { opacity: 1; } 71%, 100% { opacity: .15; } }
    .bp-x {
      width: 28px; height: 28px; border: 1px solid #33333b; border-radius: 4px; cursor: pointer;
      background: transparent; color: #8a8a94; font-size: 14px; line-height: 1;
      transition: color .15s, border-color .15s;
    }
    .bp-x:hover { color: #f2f0ec; border-color: #fb7299; }

    .bp-body { flex: 1; overflow-y: auto; padding: 14px 16px 16px; }
    .bp-body::-webkit-scrollbar { width: 9px; }
    .bp-body::-webkit-scrollbar-track { background: #131317; border-left: 1px solid #232329; }
    .bp-body::-webkit-scrollbar-thumb { background: #38383f; border-radius: 0; border: 2px solid #131317; }
    .bp-body::-webkit-scrollbar-thumb:hover { background: #fb7299; }

    .bp-info {
      display: flex; gap: 10px; align-items: center;
      padding: 10px 12px; margin-bottom: 14px;
      background: #101014; border: 1px solid #26262d; border-radius: 4px;
      font-size: 12px; color: #9a98a0;
      font-family: "Cascadia Code", ui-monospace, Consolas, monospace;
    }
    .bp-info b { color: #fb7299; font-weight: 700; }
    .bp-dot { width: 7px; height: 7px; border-radius: 50%; background: #44444c; flex: none; }
    .bp-dot.ok { background: #35c26e; box-shadow: 0 0 9px rgba(53,194,110,.8); animation: bp-blink 2.4s steps(1) infinite; }

    .bp-sec { margin-bottom: 14px; }
    .bp-sec-title {
      font-size: 10px; font-weight: 600; letter-spacing: 3px; color: #62626e;
      text-transform: uppercase; margin: 0 0 10px 2px;
      font-family: "Cascadia Code", ui-monospace, Consolas, monospace;
    }
    .bp-sec-title::after { content: " ——————"; color: #2c2c33; letter-spacing: 0; }
    .bp-row { display: flex; align-items: center; gap: 10px; margin-bottom: 9px; }
    .bp-row label { width: 60px; flex: none; color: #b8b6ae; font-size: 12px; }
    .bp-row .val { width: 52px; text-align: right; color: #fb7299; font-size: 12px; font-weight: 700; flex: none; }
    input[type="range"] {
      flex: 1; appearance: none; height: 3px;
      background: repeating-linear-gradient(90deg, #3a3a42 0 6px, transparent 6px 12px);
      outline: none; cursor: pointer;
    }
    input[type="range"]::-webkit-slider-thumb {
      appearance: none; width: 6px; height: 16px; border-radius: 1px;
      background: #fb7299; cursor: ew-resize;
      box-shadow: 0 0 8px rgba(251,114,153,.7);
    }
    .bp-tinputs { display: flex; gap: 9px; }
    .bp-tinputs input {
      flex: 1; min-width: 0; padding: 8px 11px; border-radius: 3px;
      border: 1px solid #2c2c33; background: #101014;
      color: #e8e6e0; font-size: 12.5px; outline: none;
      font-family: "Cascadia Code", ui-monospace, Consolas, monospace;
    }
    .bp-tinputs input::placeholder { color: #4e4e58; }
    .bp-tinputs input:focus { border-color: #fb7299; box-shadow: 0 0 0 1px rgba(251,114,153,.25); }
    .bp-hint { font-size: 11px; color: #62626e; margin-top: 6px; line-height: 1.7; }

    /* ---------- 按钮: 剪辑台实体键 ---------- */
    .bp-btn {
      border: 1px solid #33333b; border-radius: 3px; cursor: pointer; font-weight: 600;
      background: linear-gradient(180deg, #23232a, #1a1a1f);
      color: #c8c6c0; letter-spacing: 1px;
      transition: transform .12s, border-color .15s, color .15s, box-shadow .15s;
      box-shadow: 0 2px 0 #0a0a0c;
    }
    .bp-btn:active { transform: translateY(1px); box-shadow: 0 1px 0 #0a0a0c; }
    .bp-btn:disabled { opacity: .4; cursor: not-allowed; }
    .bp-btn.primary {
      background: linear-gradient(180deg, #2a1418, #1e0e12);
      border-color: #6e2c3d; color: #ff8fb0;
      padding: 11px 18px; font-size: 13.5px; letter-spacing: 2px;
    }
    .bp-btn.primary:hover:not(:disabled) {
      border-color: #fb7299; color: #ffb3c8;
      box-shadow: 0 0 16px rgba(251,114,153,.28), 0 2px 0 #0a0a0c;
    }
    .bp-btn.ghost { padding: 11px 14px; font-size: 12.5px; }
    .bp-btn.ghost:hover:not(:disabled) { border-color: #55555f; color: #e8e6e0; }
    .bp-btn.small { padding: 6px 12px; font-size: 11.5px; background: #1d1d22; }
    .bp-btn.small.accent { border-color: #6e2c3d; color: #ff8fb0; background: linear-gradient(180deg, #2a1418, #1e0e12); }
    .bp-btn.small.accent:hover { border-color: #fb7299; }
    .bp-btn.stop {
      background: linear-gradient(180deg, #2d1113, #200c0e); border-color: #7a3038;
      color: #ff8a8a; padding: 11px 14px; font-size: 12.5px; letter-spacing: 3px;
    }
    .bp-btn.stop::before { content: "■ "; font-size: 9px; }

    .bp-croprow { margin-top: 9px; }
    .bp-croplabel { font-size: 11px; color: #71717c; font-family: "Cascadia Code", ui-monospace, Consolas, monospace; }
    .bp-scanrow { display: flex; gap: 9px; }
    .bp-scanrow .bp-btn.primary { flex: 1.5; }
    .bp-scan-hint { margin: 8px 2px 14px; }
    .bp-res-ops { display: flex; gap: 7px; }
    .bp-warn { color: #ffb027; }
    .hidden { display: none !important; }

    /* ---------- 高级选项: 折叠收纳 ---------- */
    .bp-adv { margin-bottom: 14px; }
    .bp-adv summary {
      list-style: none; cursor: pointer; user-select: none;
      font-size: 11px; color: #71717c; letter-spacing: 2px;
      padding: 7px 2px;
      font-family: "Cascadia Code", ui-monospace, Consolas, monospace;
      display: flex; align-items: center; gap: 8px;
      border-top: 1px dashed #2c2c33;
    }
    .bp-adv summary::-webkit-details-marker { display: none; }
    .bp-adv summary::before {
      content: "▸"; color: #fb7299; font-size: 10px;
      transition: transform .2s;
    }
    .bp-adv[open] summary::before { transform: rotate(90deg); }
    .bp-adv summary:hover { color: #a0a0aa; }
    .bp-adv summary::after {
      content: ""; flex: 1; border-bottom: 1px dashed #2c2c33;
      margin-left: 4px;
    }
    .bp-adv[open] summary { margin-bottom: 10px; }
    .bp-adv .bp-sec { margin-bottom: 0; }

    /* ---------- 进度: 走带扫描线 ---------- */
    .bp-progress { display: none; margin-bottom: 13px; }
    .bp-progress.on { display: block; }
    .bp-bar {
      height: 5px; background: #101014; border: 1px solid #26262d;
      position: relative; overflow: hidden;
    }
    .bp-bar > div {
      height: 100%; width: 0%;
      background: repeating-linear-gradient(90deg, #fb7299 0 10px, #c4556f 10px 20px);
      transition: width .35s;
      position: relative;
    }
    .bp-bar > div::after { /* 扫描线 */
      content: ""; position: absolute; inset: 0;
      background: linear-gradient(90deg, transparent 45%, rgba(255,255,255,.5) 50%, transparent 55%);
      background-size: 200% 100%;
      animation: bp-sweep 1.1s linear infinite;
    }
    @keyframes bp-sweep { from { background-position: 200% 0; } to { background-position: -200% 0; } }
    .bp-ptext {
      margin-top: 7px; font-size: 11px; color: #71717c;
      display: flex; justify-content: space-between;
      font-family: "Cascadia Code", ui-monospace, Consolas, monospace;
    }

    /* ---------- 结果: 胶片帧 ---------- */
    .bp-res-head { display: flex; align-items: center; justify-content: space-between; margin: 6px 0 10px; }
    .bp-res-head b { color: #f2f0ec; font-weight: 600; }
    .bp-res-head b em { color: #fb7299; font-style: normal; }
    .bp-thumbs { display: grid; grid-template-columns: 1fr 1fr; gap: 11px; }
    .bp-thumb {
      position: relative; cursor: pointer; background: #000;
      border: 1px solid #2c2c33; border-radius: 2px;
      transition: transform .15s, border-color .15s, box-shadow .15s;
      animation: bp-pop .3s cubic-bezier(.34,1.4,.64,1);
      padding: 5px 0; /* 上下齿孔边距 */
      background-image:
        radial-gradient(circle 2px at 5px 50%, #2c2c33 97%, transparent),
        radial-gradient(circle 2px at calc(100% - 5px) 50%, #2c2c33 97%, transparent);
      background-size: 100% 10px; background-repeat: repeat-y;
      background-color: #0c0c0e;
    }
    @keyframes bp-pop { from { opacity: 0; transform: scale(.88); } to { opacity: 1; transform: none; } }
    .bp-thumb:hover { transform: translateY(-2px); border-color: #55555f; }
    .bp-thumb img { display: block; width: 100%; aspect-ratio: 16/9; object-fit: cover; padding: 0 7px; width: calc(100% - 14px); }
    .bp-thumb.off { opacity: .35; filter: grayscale(.8); }
    .bp-thumb.on { border-color: #6e2c3d; }
    .bp-thumb.on:hover { border-color: #fb7299; box-shadow: 0 4px 16px rgba(251,114,153,.2); }
    .bp-check {
      position: absolute; top: 5px; right: 5px; width: 17px; height: 17px;
      border-radius: 2px; border: 1.5px solid rgba(242,240,236,.75);
      display: grid; place-items: center; pointer-events: none;
      background: rgba(0,0,0,.5);
    }
    .bp-check svg { width: 10px; height: 10px; opacity: 0; }
    .bp-thumb.on .bp-check { background: #fb7299; border-color: #fb7299; box-shadow: 0 0 8px rgba(251,114,153,.6); }
    .bp-thumb.on .bp-check svg { opacity: 1; }
    .bp-tc {
      position: absolute; left: 5px; bottom: 5px;
      padding: 1px 6px; border-radius: 2px;
      background: rgba(0,0,0,.78); color: #ffd7e2;
      font-size: 10px; font-weight: 600; letter-spacing: 1px;
      font-family: "Cascadia Code", ui-monospace, Consolas, monospace;
      border: 1px solid rgba(251,114,153,.3);
    }
    .bp-dup {
      position: absolute; top: 5px; left: 5px;
      padding: 1px 6px; border-radius: 2px;
      background: rgba(255,170,0,.92); color: #42310a;
      font-size: 9.5px; font-weight: 800; letter-spacing: 1px;
    }
    .bp-dup.bp-buildup { background: rgba(120,160,255,.88); color: #101a38; }
    .bp-empty {
      grid-column: 1 / -1; text-align: center; color: #55555f; padding: 30px 12px;
      font-size: 12px; line-height: 2; border: 1px dashed #2c2c33; border-radius: 4px;
    }
    .bp-empty::before { content: "⌗ STANDBY"; display: block; font-size: 10px; letter-spacing: 4px; color: #3d3d46; margin-bottom: 8px; font-family: "Cascadia Code", ui-monospace, Consolas, monospace; }

    .bp-actions { display: flex; gap: 9px; margin-top: 15px; }
    .bp-actions .bp-btn { flex: 1; }

    .bp-foot {
      padding: 9px 16px; border-top: 1px solid #26262d; font-size: 10.5px; color: #55555f; line-height: 1.7;
      background: #101014;
    }

    .bp-toast {
      position: fixed; left: 50%; bottom: 42px; transform: translateX(-50%) translateY(16px);
      z-index: 2147483002; background: #0c0c0e; color: #f2f0ec;
      padding: 10px 20px; border-radius: 3px; font-size: 12.5px; font-weight: 600;
      border: 1px solid #33333b; border-left: 3px solid #fb7299;
      box-shadow: 0 12px 36px rgba(0,0,0,.6);
      opacity: 0; pointer-events: none; transition: opacity .25s, transform .25s;
      max-width: 340px; text-align: center; letter-spacing: .5px;
    }
    .bp-toast.show { opacity: 1; transform: translateX(-50%) translateY(0); }

    .bp-modal {
      position: fixed; inset: 0; z-index: 2147483003; display: none;
      align-items: center; justify-content: center;
      background: rgba(0,0,0,.68); backdrop-filter: blur(5px) grayscale(.4);
    }
    .bp-modal.open { display: flex; }
    .bp-modal-box {
      background: #17171a; border: 1px solid #2c2c33; border-radius: 5px;
      padding: 18px; width: 440px; max-width: calc(100vw - 32px);
      box-shadow: 0 32px 80px rgba(0,0,0,.7);
      position: relative;
    }
    .bp-modal-box::before, .bp-modal-box::after {
      content: ""; position: absolute; width: 12px; height: 12px; pointer-events: none;
      border: 1.5px solid rgba(251,114,153,.55);
    }
    .bp-modal-box::before { top: 7px; left: 7px; border-right: none; border-bottom: none; }
    .bp-modal-box::after { bottom: 7px; right: 7px; border-left: none; border-top: none; }
    .bp-modal-box h3 { font-size: 14px; margin-bottom: 11px; color: #f2f0ec; font-weight: 600; letter-spacing: 2px; }
    .bp-modal-box .bp-hint { margin-bottom: 11px; }
    .bp-cropwrap { position: relative; border-radius: 3px; overflow: hidden; background: #000; border: 1px solid #2c2c33; }
    .bp-cropwrap canvas { display: block; width: 100%; cursor: crosshair; touch-action: none; }
    .bp-modal-actions { display: flex; gap: 8px; justify-content: flex-end; margin-top: 13px; }

    .bp-lightbox {
      position: fixed; inset: 0; z-index: 2147483004; display: none;
      align-items: center; justify-content: center;
      background: rgba(4,4,5,.9); cursor: zoom-out;
    }
    .bp-lightbox.open { display: flex; }
    .bp-lightbox img { max-width: 92vw; max-height: 88vh; border-radius: 3px; box-shadow: 0 16px 60px rgba(0,0,0,.85); }
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
      <div class="bp-strip" aria-hidden="true"></div>
      <div class="bp-main">
      <div class="bp-head">
        <div class="bp-logo">${ICON_SLIDES}</div>
        <div class="bp-title">PPT 提取器<small>BILIBILI VIDEO · FRAME EXTRACT</small></div>
        <button class="bp-x" title="收起">✕</button>
      </div>
      <div class="bp-body">
        <div class="bp-info"><span class="bp-dot" id="dot"></span><span id="vinfo">正在检测播放器…</span></div>

        <div class="bp-sec bp-scanrow">
          <button class="bp-btn primary" id="scanBtn">开始扫描</button>
          <button class="bp-btn stop hidden" id="stopBtn">停止</button>
          <button class="bp-btn ghost" id="reScanBtn" title="重新扫描（清空当前结果）">重扫</button>
        </div>
        <div class="bp-hint bp-scan-hint">点「开始扫描」即可，识别参数全自动。片头太长、只想扫某一段时再展开高级选项。</div>

        <div class="bp-progress" id="prog">
          <div class="bp-bar"><div id="bar"></div></div>
          <div class="bp-ptext"><span id="ptxt1">准备中…</span><span class="num" id="ptxt2"></span></div>
        </div>

        <details class="bp-adv">
          <summary>高级选项</summary>
          <div class="bp-sec">
            <div class="bp-row"><label>采样间隔</label><input type="range" id="step" min="1" max="10" step="1"><span class="val num" id="stepV"></span></div>
            <div class="bp-row"><label>翻页阈值</label><input type="range" id="th" min="2" max="20" step="1"><span class="val num" id="thV"></span></div>
            <div class="bp-tinputs">
              <input id="t0" placeholder="起点 0:00" title="扫描起始时间，支持 90 / 1:30 / 1:00:30">
              <input id="t1" placeholder="终点（默认结尾）" title="扫描结束时间，留空=视频结尾">
            </div>
            <div class="bp-row bp-croprow">
              <label>画面裁剪</label>
              <button class="bp-btn small" id="cropBtn">选取区域…</button>
              <span id="cropLabel" class="bp-croplabel">全画面</span>
            </div>
            <div class="bp-hint">翻页识别不准时再调：阈值越小越容易判定为“新的一页”。同页动画（逐块浮现/高亮）不会切页，取动画放完的完整画面；摄像头/水印干扰时用裁剪只保留 PPT 区域。</div>
          </div>
        </details>

        <div class="bp-res-head">
          <span>识别结果 <b><em id="cntTotal">0</em></b> 页 · 已选 <b id="cntSel">0</b></span>
          <span class="bp-res-ops">
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
            <button class="bp-btn small accent" id="cropApply">应用</button>
          </div>
        </div>
      </div>

      <div class="bp-lightbox" id="lightbox"><img id="lightboxImg" alt="预览"></div>
      </div>
    `;
    shadow.appendChild(panel);

    state.host = host;
    state.shadow = shadow;
    state.panel = panel;
    state.fab = fab;
    wireEvents(shadow);
    // 设置由启动流程在 loadSaved() 完成后统一 syncSettingsToUI()
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
      (v.videoWidth && v.videoWidth < 720 ? ' · <span class="bp-warn">清晰度偏低，建议调高</span>' : '');
  }

  function setScanUI(on) {
    el('scanBtn').classList.toggle('hidden', on);
    el('stopBtn').classList.toggle('hidden', !on);
    el('prog').classList.toggle('on', on);
    ['step', 'th', 't0', 't1', 'cropBtn', 'reScanBtn'].forEach((id) => { el(id).disabled = on; });
    const adv = state.shadow.querySelector('.bp-adv');
    if (adv) {
      if (on) adv.removeAttribute('open');
      adv.style.pointerEvents = on ? 'none' : '';
      adv.style.opacity = on ? '.45' : '';
    }
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
      ${seg.dupOf >= 0 ? '<span class="bp-dup">疑似重复</span>' : ''}
      ${seg.buildupOf != null ? '<span class="bp-dup bp-buildup">构建中间态</span>' : ''}`;
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

  // 设置持久化在 chrome.storage.local，异步加载后再同步 UI
  loadSaved().then((saved) => {
    if (saved) {
      if (saved.settings) state.settings = Object.assign(state.settings, saved.settings);
      if (saved.crop) state.crop = saved.crop;
    }
    syncSettingsToUI();
  });

  const watchTimer = setInterval(watchVideo, 2000);
  watchVideo();

  // 孤儿脚本检测：扩展被重载/更新后，旧 content script 的 chrome API 会失效，
  // 此时自动卸载 UI，避免残留出无效面板（下次注入会因开关被清空而重新初始化）
  const orphanTimer = setInterval(() => {
    let orphaned = false;
    try {
      if (!chrome.runtime || !chrome.runtime.id) orphaned = true;
    } catch (e) { orphaned = true; }
    if (!orphaned) return;
    clearInterval(orphanTimer);
    clearInterval(watchTimer);
    state.abort = true; // 让进行中的扫描循环尽快退出
    try { clearResults(); } catch (e) { /* ignore */ }
    if (state.host && state.host.parentNode) state.host.remove();
    window.__BILI_PPT_TOGGLE__ = null;
  }, 5000);
})();
