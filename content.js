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
    hideFab: false, // 用户偏好:隐藏页面悬浮按钮
    settings: { step: 3, threshold: 5, startStr: '', endStr: '' },
  };

  // 扩展自身数据走 chrome.storage.local（不污染页面源，也不受站点清理影响）
  function loadSaved() {
    return new Promise((resolve) => {
      try {
        if (!chrome.storage || !chrome.storage.local) { resolve(null); return; }
        chrome.storage.local.get(['settings', 'crop', 'hideFab'], (res) => {
          if (chrome.runtime.lastError) { resolve(null); return; }
          resolve(res || null);
        });
      } catch (e) { resolve(null); }
    });
  }

  function persist() {
    try {
      if (chrome.storage && chrome.storage.local) {
        chrome.storage.local.set({ settings: state.settings, crop: state.crop || null, hideFab: !!state.hideFab });
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
  // 暗边探测画布:复用避免逐帧分配(改尺寸会自动清空,每帧重绘)
  const probeCanvas = document.createElement('canvas');
  const probeCtx = probeCanvas.getContext('2d', { willReadFrequently: true });

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
    // 检测画布:原始分辨率(上限 1920 防极端),复用探测画布
    const W = Math.min(vw, 1920), H = Math.round(W * vh / vw);
    let px;
    try {
      probeCanvas.width = W; probeCanvas.height = H; // 改尺寸自动清空
      probeCtx.drawImage(v, 0, 0, W, H);
      px = probeCtx.getImageData(0, 0, W, H).data;
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
      sx: Math.max(0, Math.ceil(left * k) + (left > 0 ? 1 : 0)),
      sy: Math.max(0, Math.ceil(top * (vh / H)) + (top > 0 ? 1 : 0)),
      sw: 0, sh: 0,
    };
    rect.sw = Math.max(16, Math.min(vw - rect.sx, Math.floor((W - left - right) * k) - (right > 0 ? 1 : 0)));
    rect.sh = Math.max(16, Math.min(vh - rect.sy, Math.floor((H - top - bottom) * (vh / H)) - (bottom > 0 ? 1 : 0)));
    if (rect.sw < vw * MIN_INNER || rect.sh < vh * MIN_INNER) return null;
    return rect;
  }

  // 输出画布上的精裁:源坐标裁剪后仍会有亚像素黑边经缩放混进输出边缘
  // (实测:636px 输出的最外列/行亮度 112/180,不是内容白底)。
  // 直接在最终画布上按整数像素再裁一刀,量少(1-2px)且精确。
  // 返回 {top,bottom,left,right},无残边返回 null。
  // 残边判定必须是"相对"的(绝对阈值版本曾把深色背景课件整体亮度 ~31
  // 的四边都裁到 10% 上限):
  //   一行/列是残边 = 绝对偏暗(< 190) 且 明显暗于同侧内缩 3px 处(差 > 50)。
  // 白底课件残边(112-180 vs 内容 249)满足;深色底课件边缘与内部同暗
  // (差 ~0)不满足;边缘有亮内容贴边的行均值被拉高也不满足。
  // 深色底的真黑边由源级 detectDarkEdges 负责,此处只管亚像素残带,
  // CAP=4 上限足够且防误裁扩大(大 CAP + 自适应参考曾在深色课件标题带
  // 上裁掉 8px 真内容——参考点落在亮带上,落差被凭空制造出来)。
  function trimOutputEdges(ctx, w, h) {
    let px;
    try { px = ctx.getImageData(0, 0, w, h).data; } catch (e) { return null; }
    const lum = (x, y) => {
      const o = (y * w + x) * 4;
      return (px[o] * 299 + px[o + 1] * 587 + px[o + 2] * 114) / 1000;
    };
    const EDGE_MAX = 190, GAP = 50, INNER_OFF = 3;
    const CAP = 4; // 残边是 1-2px 亚像素重分配产物,4px 上限足够且防误裁扩大
    const rowMean = (y) => { let s = 0, n = 0; for (let x = 0; x < w; x += 2) { s += lum(x, y); n++; } return s / n; };
    const colMean = (x) => { let s = 0, n = 0; for (let y = 0; y < h; y += 2) { s += lum(x, y); n++; } return s / n; };
    let top = 0, bottom = 0, left = 0, right = 0;
    while (top < CAP) {
      const edge = rowMean(top), inner = rowMean(Math.min(top + INNER_OFF, h - 1));
      if (edge < EDGE_MAX && inner - edge > GAP) top++; else break;
    }
    while (bottom < CAP) {
      const edge = rowMean(h - 1 - bottom), inner = rowMean(Math.max(h - 1 - bottom - INNER_OFF, 0));
      if (edge < EDGE_MAX && inner - edge > GAP) bottom++; else break;
    }
    while (left < CAP) {
      const edge = colMean(left), inner = colMean(Math.min(left + INNER_OFF, w - 1));
      if (edge < EDGE_MAX && inner - edge > GAP) left++; else break;
    }
    while (right < CAP) {
      const edge = colMean(w - 1 - right), inner = colMean(Math.max(w - 1 - right - INNER_OFF, 0));
      if (edge < EDGE_MAX && inner - edge > GAP) right++; else break;
    }
    if (!top && !bottom && !left && !right) return null;
    return { top, bottom, left, right };
  }

  function grabFrame(v) {
    const base = regionOf(v);
    // 自动去黑边:检测到的暗边进一步收缩源区域(在用户裁剪之后叠加)。
    // 逐帧探测,不缓存:黑边宽度会随时间漂移(实测 BV1SKkcBxEXr 前段
    // L87/R88 → 后段 L82/R82,底部黑边只在部分片段出现),按分辨率缓存
    // 的矩形只对一个时间点正确——变宽则残留黑边,变窄则裁掉真内容;
    // 且深色帧会污染缓存,残留检测只认"变宽"方向,变窄永远修不回来。
    // 全分辨率探测每帧 ~10ms,相对每次 seek 的数百毫秒可忽略
    // (准确度优先)。探测画布复用,避免逐帧分配。
    const edge = detectDarkEdges(v);
    let { sx, sy, sw, sh } = base;
    if (edge) {
      // 与用户裁剪区域求交集
      const x0 = Math.max(sx, edge.sx), y0 = Math.max(sy, edge.sy);
      const x1 = Math.min(sx + sw, edge.sx + edge.sw), y1 = Math.min(sy + sh, edge.sy + edge.sh);
      if (x1 - x0 > 16 && y1 - y0 > 16) { sx = x0; sy = y0; sw = x1 - x0; sh = y1 - y0; }
    }
    const scale = Math.min(1, 1920 / sw);
    const cw = Math.max(2, Math.round(sw * scale));
    const ch = Math.max(2, Math.round(sh * scale));
    capCanvas.width = cw; capCanvas.height = ch;
    try {
      capCtx.drawImage(v, sx, sy, sw, sh, 0, 0, cw, ch);
      // 精裁:源裁剪+缩放后边缘仍会残留亚像素黑边,在输出画布上直接按
      // 整数像素裁掉。用临时画布中转(画布不能自引用裁剪,改尺寸会清空)。
      const trim = trimOutputEdges(capCtx, cw, ch);
      if (trim) {
        const tw = cw - trim.left - trim.right, th = ch - trim.top - trim.bottom;
        if (tw > 16 && th > 16) {
          const tmp = document.createElement('canvas');
          tmp.width = tw; tmp.height = th;
          tmp.getContext('2d').drawImage(capCanvas, trim.left, trim.top, tw, th, 0, 0, tw, th);
          capCanvas.width = tw; capCanvas.height = th;
          capCtx.drawImage(tmp, 0, 0);
        }
      }
      // 签名与导出用同一画面(trim 后):黑边会稀释签名的内容分辨率
      // (实测黑边占宽 26% 时,48px 签名只有 35px 有效内容),也压低
      // isBlank 的方差。从输出画布取签名,保证检测与导出一致。
      sigCtx.drawImage(capCanvas, 0, 0, capCanvas.width, capCanvas.height, 0, 0, SIG_W, SIG_H);
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
        blob.arrayBuffer().then((ab) => resolve({ sig, bytes: new Uint8Array(ab), w: capCanvas.width, h: capCanvas.height, content: contentOf(sig) }));
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
    const frame = await grabFrame(v);
    if (frame) frame.t = t; // 记录时间戳:缩略图时间码与导出文件名都依赖它
    return frame;
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
      // 回翻/重复页默认不勾选(实测判定准确,重复多):徽章标注来源页,
      // 想要回翻页点缩略图勾回即可
      checked: dupOf < 0, dupOf,
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

  // 收尾合并:渐进构建识别。
  // 切段用宁多勿少的低阈值,多切出来的动画中间态在这里统一处理。
  // 判定某段是"构建中间态"(默认不勾选)的条件(全部满足):
  //   a) 与后续某段(不限于紧邻)画面"同族"——签名差异低于阈值×1.4
  //      (与去重同一量级:同一页的不同构建步骤,画面仍有明显重合)
  //   b) 内容量比该后续段明显低(画面更空 < 65%):中间态只可能比成品空
  //   c) 时间上在其之前(构建先于成品)
  // 动画步进序列 [空白页 → 半成品 → 完整页] 会把前两个标成中间态,
  // 保留完整的。真正的独立页(分时/实时)与任何后续段都不同族,不受影响。
  function mergeBuildups() {
    const segs = state.segments;
    let changed = false;
    const th = state.settings.threshold;
    for (let i = 0; i < segs.length; i++) {
      const cur = segs[i];
      if (!cur.checked) continue;
      for (let j = i + 1; j < segs.length; j++) {
        const nxt = segs[j];
        if (!nxt) continue;
        if (diff(cur.sig, nxt.sig) > th * 1.4) continue;       // a) 同族
        if (cur.content == null || nxt.content == null) continue;
        if (cur.content >= nxt.content * 0.65) continue;        // b) 明显更空
        // c) 顺序天然满足(i<j)。命中:cur 是 j 的构建中间态
        cur.checked = false;
        cur.buildupOf = j;
        changed = true;
        break;
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
    let count = 0, prevSig = null, lastFrame = null, missed = 0;

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
          if (d > threshold) {
            // 切段宁多勿少:超过用户阈值即翻页。多切出来的动画中间态
            // 由扫描结束后的 mergeBuildups() 统一识别并默认不勾选。
            // 历史教训:提高切段阈值(×3)会吞掉"每步差异小"的渐进式翻页,
            // 丢真页的代价远大于多出几页中间态。
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
  // 主流产品暗色设计语言(参考 Chrome DevTools / Notion / Linear):
  // 系统字体、8px 栅格、6-8px 圆角、中性灰阶层级、
  // 单一品牌色(#fb7299 B站粉)仅用于主操作与关键状态。
  const CSS = `
    :host { all: initial; }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, "Segoe UI", "Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC", sans-serif; }
    .num { font-variant-numeric: tabular-nums; }

    /* ---------- 悬浮按钮 ---------- */
    .bp-fab {
      position: fixed; right: 20px; bottom: 96px; z-index: 2147483000;
      display: flex; align-items: center; gap: 8px;
      padding: 10px 16px; border: none; border-radius: 8px; cursor: pointer;
      background: #fb7299; color: #fff;
      font-size: 13px; font-weight: 600;
      box-shadow: 0 4px 12px rgba(0,0,0,.25);
      transition: background .15s, transform .15s, box-shadow .15s;
      animation: bp-in .3s ease-out;
    }
    .bp-fab:hover { background: #ff87a9; box-shadow: 0 6px 16px rgba(0,0,0,.3); }
    .bp-fab:active { transform: scale(.97); }
    .bp-fab svg { width: 16px; height: 16px; }
    .bp-fab.ready svg { }
    @keyframes bp-in { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: none; } }

    /* ---------- 主面板 ---------- */
    .bp-panel {
      position: fixed; right: 16px; top: 64px; bottom: 16px; width: 420px; z-index: 2147483001;
      display: none; flex-direction: column;
      background: #202124;
      border: 1px solid #3c4043; border-radius: 8px;
      box-shadow: 0 8px 32px rgba(0,0,0,.45);
      color: #e8eaed; font-size: 13px; overflow: hidden;
    }
    .bp-panel.open { display: flex; animation: bp-slide .25s ease-out; }
    @keyframes bp-slide { from { opacity: 0; transform: translateY(-8px); } to { opacity: 1; transform: none; } }

    .bp-head {
      display: flex; align-items: center; gap: 10px;
      padding: 14px 16px;
      border-bottom: 1px solid #3c4043;
    }
    .bp-logo {
      width: 32px; height: 32px; flex: none; display: grid; place-items: center;
      background: #fb7299; border-radius: 6px;
    }
    .bp-logo svg { width: 16px; height: 16px; }
    .bp-title { font-weight: 600; font-size: 14px; flex: 1; color: #e8eaed; }
    .bp-title small { display: block; font-weight: 400; font-size: 11px; color: #9aa0a6; margin-top: 1px; }
    .bp-x {
      width: 28px; height: 28px; border: none; border-radius: 6px; cursor: pointer;
      background: transparent; color: #9aa0a6; font-size: 15px; line-height: 1;
      display: grid; place-items: center;
      transition: background .15s, color .15s;
    }
    .bp-x:hover { background: #3c4043; color: #e8eaed; }

    .bp-body { flex: 1; overflow-y: auto; padding: 16px; }
    .bp-body::-webkit-scrollbar { width: 8px; }
    .bp-body::-webkit-scrollbar-track { background: transparent; }
    .bp-body::-webkit-scrollbar-thumb { background: #5f6368; border-radius: 4px; }
    .bp-body::-webkit-scrollbar-thumb:hover { background: #80868b; }

    .bp-info {
      display: flex; gap: 8px; align-items: center;
      padding: 10px 12px; margin-bottom: 16px;
      background: #292a2d; border-radius: 6px;
      font-size: 12px; color: #9aa0a6;
    }
    .bp-info b { color: #e8eaed; font-weight: 600; }
    .bp-dot { width: 8px; height: 8px; border-radius: 50%; background: #5f6368; flex: none; }
    .bp-dot.ok { background: #34a853; }

    .bp-sec { margin-bottom: 16px; }
    .bp-sec-title {
      font-size: 12px; font-weight: 600; color: #e8eaed;
      margin: 0 0 8px;
    }
    .bp-row { display: flex; align-items: center; gap: 10px; margin-bottom: 8px; }
    .bp-row label { width: 60px; flex: none; color: #9aa0a6; font-size: 12px; }
    .bp-row .val { width: 40px; text-align: right; color: #e8eaed; font-size: 12px; flex: none; }
    input[type="range"] {
      flex: 1; appearance: none; height: 4px; border-radius: 2px;
      background: linear-gradient(90deg, #fb7299 var(--p, 30%), #5f6368 var(--p, 30%));
      outline: none; cursor: pointer;
    }
    input[type="range"]::-webkit-slider-thumb {
      appearance: none; width: 14px; height: 14px; border-radius: 50%;
      background: #fff; border: 2px solid #fb7299; cursor: pointer;
      box-shadow: 0 1px 3px rgba(0,0,0,.4);
    }
    .bp-tinputs { display: flex; gap: 8px; }
    .bp-tinputs input {
      flex: 1; min-width: 0; padding: 7px 10px; border-radius: 6px;
      border: 1px solid #3c4043; background: #292a2d;
      color: #e8eaed; font-size: 12.5px; outline: none;
    }
    .bp-tinputs input::placeholder { color: #6b6e72; }
    .bp-tinputs input:focus { border-color: #fb7299; box-shadow: 0 0 0 1px rgba(251,114,153,.3); }
    .bp-hint { font-size: 11px; color: #9aa0a6; margin-top: 4px; line-height: 1.6; }

    /* ---------- 按钮 ---------- */
    .bp-btn {
      border: 1px solid transparent; border-radius: 6px; cursor: pointer; font-weight: 500;
      background: #3c4043; color: #e8eaed;
      transition: background .15s, border-color .15s, box-shadow .15s;
      font-family: inherit;
    }
    .bp-btn:active { transform: scale(.98); }
    .bp-btn:disabled { opacity: .4; cursor: not-allowed; }
    .bp-btn.primary {
      background: #fb7299; color: #fff; border-color: #fb7299;
      padding: 10px 16px; font-size: 13px; font-weight: 600;
    }
    .bp-btn.primary:hover:not(:disabled) { background: #ff87a9; border-color: #ff87a9; }
    .bp-btn.ghost {
      background: transparent; color: #e8eaed; border-color: #3c4043;
      padding: 10px 14px; font-size: 12.5px;
    }
    .bp-btn.ghost:hover:not(:disabled) { background: #292a2d; }
    .bp-btn.small { padding: 5px 10px; font-size: 11.5px; background: #3c4043; }
    .bp-btn.small:hover:not(:disabled) { background: #4a4e51; }
    .bp-btn.small.accent { background: #fb7299; color: #fff; border-color: #fb7299; }
    .bp-btn.small.accent:hover { background: #ff87a9; }
    .bp-btn.stop { background: #c5221f; color: #fff; border-color: #c5221f; padding: 10px 14px; font-size: 12.5px; }
    .bp-btn.stop:hover:not(:disabled) { background: #d93025; }

    .bp-croprow { margin-top: 8px; }
    .bp-croplabel { font-size: 11px; color: #9aa0a6; }
    .bp-scanrow { display: flex; gap: 8px; }
    .bp-scanrow .bp-btn.primary { flex: 1.5; }
    .bp-scan-hint { margin: 8px 0 16px; }
    .bp-res-ops { display: flex; gap: 6px; }
    .bp-warn { color: #fdd663; }
    .hidden { display: none !important; }

    /* ---------- 高级选项 ---------- */
    .bp-adv { margin-bottom: 16px; }
    .bp-adv summary {
      list-style: none; cursor: pointer; user-select: none;
      font-size: 12px; color: #9aa0a6;
      padding: 8px 0;
      display: flex; align-items: center; gap: 6px;
      border-top: 1px solid #3c4043;
    }
    .bp-adv summary::-webkit-details-marker { display: none; }
    .bp-adv summary::before {
      content: "›"; color: #9aa0a6; font-size: 14px; line-height: 1;
      transition: transform .2s;
    }
    .bp-adv[open] summary::before { transform: rotate(90deg); }
    .bp-adv summary:hover { color: #e8eaed; }
    .bp-adv[open] summary { margin-bottom: 10px; }
    .bp-adv .bp-sec { margin-bottom: 0; }

    /* ---------- 进度 ---------- */
    .bp-progress { display: none; margin-bottom: 16px; }
    .bp-progress.on { display: block; }
    .bp-bar {
      height: 6px; border-radius: 3px; background: #292a2d; overflow: hidden;
    }
    .bp-bar > div {
      height: 100%; width: 0%; border-radius: 3px;
      background: #fb7299;
      transition: width .3s ease;
    }
    .bp-ptext {
      margin-top: 6px; font-size: 11px; color: #9aa0a6;
      display: flex; justify-content: space-between;
    }

    /* ---------- 结果缩略图 ---------- */
    .bp-res-head { display: flex; align-items: center; justify-content: space-between; margin: 4px 0 8px; }
    .bp-res-head b { color: #e8eaed; font-weight: 600; }
    .bp-res-head b em { color: #fb7299; font-style: normal; }
    .bp-thumbs { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    .bp-thumb {
      position: relative; cursor: pointer;
      border: 2px solid #3c4043; border-radius: 6px; overflow: hidden;
      background: #000;
      transition: border-color .15s;
      animation: bp-pop .25s ease-out;
    }
    @keyframes bp-pop { from { opacity: 0; transform: scale(.94); } to { opacity: 1; transform: none; } }
    .bp-thumb:hover { border-color: #80868b; }
    .bp-thumb img { display: block; width: 100%; aspect-ratio: 16/9; object-fit: cover; }
    .bp-thumb.off { opacity: .4; }
    .bp-thumb.on { border-color: #fb7299; }
    .bp-check {
      position: absolute; top: 6px; right: 6px; width: 18px; height: 18px;
      border-radius: 4px; border: 1.5px solid rgba(255,255,255,.85);
      display: grid; place-items: center; pointer-events: none;
      background: rgba(0,0,0,.35);
    }
    .bp-check svg { width: 11px; height: 11px; opacity: 0; }
    .bp-thumb.on .bp-check { background: #fb7299; border-color: #fb7299; }
    .bp-thumb.on .bp-check svg { opacity: 1; }
    .bp-tc {
      position: absolute; left: 6px; bottom: 6px;
      padding: 2px 6px; border-radius: 4px;
      background: rgba(0,0,0,.7); color: #e8eaed;
      font-size: 10px; font-weight: 600;
    }
    .bp-dup {
      position: absolute; top: 6px; left: 6px;
      padding: 2px 6px; border-radius: 4px;
      background: rgba(251,114,153,.85); color: #fff;
      font-size: 10px; font-weight: 600;
    }
    .bp-dup.bp-buildup { background: rgba(154,160,166,.85); color: #202124; }
    .bp-empty {
      grid-column: 1 / -1; text-align: center; color: #6b6e72; padding: 32px 12px;
      font-size: 12px; line-height: 1.8; border: 1px dashed #3c4043; border-radius: 6px;
    }

    .bp-actions { display: flex; gap: 8px; margin-top: 16px; }
    .bp-actions .bp-btn { flex: 1; }

    .bp-foot {
      padding: 10px 16px; border-top: 1px solid #3c4043;
      font-size: 11px; color: #9aa0a6; line-height: 1.6;
      background: #292a2d;
    }
    .bp-foot-opts { margin-top: 6px; }
    .bp-fabtoggle { display: flex; align-items: center; gap: 6px; cursor: pointer; user-select: none; color: #9aa0a6; }
    .bp-fabtoggle:hover { color: #e8eaed; }
    .bp-fabtoggle input { accent-color: #fb7299; width: 13px; height: 13px; cursor: pointer; }

    .bp-toast {
      position: fixed; left: 50%; bottom: 48px; transform: translateX(-50%) translateY(12px);
      z-index: 2147483002; background: #3c4043; color: #e8eaed;
      padding: 10px 18px; border-radius: 8px; font-size: 13px; font-weight: 500;
      box-shadow: 0 4px 16px rgba(0,0,0,.4);
      opacity: 0; pointer-events: none; transition: opacity .2s, transform .2s;
      max-width: 340px; text-align: center;
    }
    .bp-toast.show { opacity: 1; transform: translateX(-50%) translateY(0); }

    .bp-modal {
      position: fixed; inset: 0; z-index: 2147483003; display: none;
      align-items: center; justify-content: center;
      background: rgba(0,0,0,.6);
    }
    .bp-modal.open { display: flex; }
    .bp-modal-box {
      background: #292a2d; border: 1px solid #3c4043; border-radius: 8px;
      padding: 18px; width: 440px; max-width: calc(100vw - 32px);
      box-shadow: 0 16px 48px rgba(0,0,0,.5);
    }
    .bp-modal-box h3 { font-size: 15px; margin-bottom: 10px; color: #e8eaed; font-weight: 600; }
    .bp-modal-box .bp-hint { margin-bottom: 10px; }
    .bp-cropwrap { position: relative; border-radius: 6px; overflow: hidden; background: #000; border: 1px solid #3c4043; }
    .bp-cropwrap canvas { display: block; width: 100%; cursor: crosshair; touch-action: none; }
    .bp-modal-actions { display: flex; gap: 8px; justify-content: flex-end; margin-top: 14px; }

    .bp-lightbox {
      position: fixed; inset: 0; z-index: 2147483004; display: none;
      align-items: center; justify-content: center;
      background: rgba(0,0,0,.85); cursor: zoom-out;
    }
    .bp-lightbox.open { display: flex; }
    .bp-lightbox img { max-width: 92vw; max-height: 88vh; border-radius: 4px; box-shadow: 0 8px 32px rgba(0,0,0,.6); }
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
        <div class="bp-title">PPT 提取器<small>B站视频 · 自动识别翻页</small></div>
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
      <div class="bp-foot">提示：扫描期间请停留在本页面、不要切换清晰度；建议先把播放器切到最高清晰度再开始。<br>导出的 PPTX 每页即视频中出现的一帧幻灯片。
        <div class="bp-foot-opts"><label class="bp-fabtoggle"><input type="checkbox" id="fabToggle"> 在视频页显示悬浮按钮</label></div>
      </div>

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
    el('fabToggle').addEventListener('change', (e) => {
      state.hideFab = !e.target.checked;
      persist();
      applyFabVisibility();
      if (state.hideFab) toast('悬浮按钮已隐藏，用浏览器工具栏的扩展图标唤出面板');
    });
    function applyFabVisibility() {
      state.fab.style.display = state.hideFab ? 'none' : '';
    }
    // 由启动流程在 loadSaved() 后同步勾选状态并应用
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
      ${seg.dupOf >= 0 ? `<span class="bp-dup" title="与第 ${seg.dupOf + 1} 页画面相同">回翻→${seg.dupOf + 1}</span>` : ''}
      ${seg.buildupOf != null ? `<span class="bp-dup bp-buildup" title="第 ${seg.buildupOf + 1} 页的构建中间态">中间态</span>` : ''}`;
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
      if (saved.hideFab) state.hideFab = true;
    }
    syncSettingsToUI();
    const fabToggle = state.shadow.getElementById('fabToggle');
    if (fabToggle) {
      fabToggle.checked = !state.hideFab;
      state.fab.style.display = state.hideFab ? 'none' : '';
    }
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
