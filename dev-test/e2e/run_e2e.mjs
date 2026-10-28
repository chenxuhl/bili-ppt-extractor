/**
 * B站视频PPT提取器 — 真实端到端测试驱动（零依赖：Node 内置 fetch + WebSocket 走 CDP）
 *
 * 流程：
 *   1. 连接浏览器级 CDP → 设置下载目录
 *   2. 打开真实B站视频页 → 等待插件注入 + 播放器就绪
 *   3. 探测阶段：直接 seek <video> 采样帧间差异，找出"静态PPT最集中"的时间窗
 *      （同时提前验证程序化 seek 可行）
 *   4. 通过真实 UI（shadow DOM 按钮/输入框）设置起止时间 → 点击「开始扫描」
 *   5. 等待扫描完成 → 读页数 → 点击「导出PPTX」→ 等待文件落盘
 *
 * 用法：node run_e2e.mjs <videoUrl> <downloadDir>
 */
import { readdirSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { setTimeout as sleep } from 'node:timers/promises';

const DEBUG = process.env.E2E_DEBUG === '1';
const CDP_PORT = process.env.CDP_PORT || '9223';
const VIDEO_URL = process.argv[2];
const DL_DIR = process.argv[3];

if (!VIDEO_URL || !DL_DIR) {
  console.error('usage: node run_e2e.mjs <videoUrl> <downloadDir>');
  process.exit(2);
}

const log = (...a) => console.log('[e2e]', ...a);
const dbg = (...a) => { if (DEBUG) console.log('[dbg]', ...a); };

// ---------------- CDP 基础 ----------------
class Cdp {
  constructor(ws) { this.ws = ws; this.id = 0; this.pending = new Map(); this.handlers = [];
    ws.addEventListener('message', (ev) => {
      const msg = JSON.parse(ev.data);
      if (msg.id && this.pending.has(msg.id)) {
        const { resolve, reject } = this.pending.get(msg.id);
        this.pending.delete(msg.id);
        if (msg.error) reject(new Error(`${msg.error.message} ${msg.error.data || ''}`));
        else resolve(msg.result);
      } else {
        for (const h of this.handlers) h(msg);
      }
    });
  }
  static async connect(url) {
    const ws = new WebSocket(url);
    await new Promise((res, rej) => { ws.addEventListener('open', res); ws.addEventListener('error', rej); });
    return new Cdp(ws);
  }
  send(method, params = {}) {
    const id = ++this.id;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.ws.send(JSON.stringify({ id, method, params }));
    });
  }
  async eval(expression, awaitPromise = false) {
    const r = await this.send('Runtime.evaluate', { expression, returnByValue: true, awaitPromise });
    if (r.exceptionDetails) {
      throw new Error('page eval failed: ' + (r.exceptionDetails.exception?.description || r.exceptionDetails.text));
    }
    return r.result.value;
  }
}

async function getJson(path) {
  const res = await fetch(`http://127.0.0.1:${CDP_PORT}${path}`);
  if (!res.ok) throw new Error(`CDP http ${res.status}`);
  return res.json();
}

// ---------------- 工具 ----------------
const shadowExpr = (inner) => `(()=>{const s=document.querySelector('#bili-ppt-host').shadowRoot;${inner}})()`;

async function waitFor(page, expr, timeoutMs, label, everyMs = 2000) {
  const t0 = Date.now();
  for (;;) {
    let v = null;
    try { v = await page.eval(expr); } catch (e) { dbg(label, 'eval err', e.message); }
    if (v) return v;
    if (Date.now() - t0 > timeoutMs) throw new Error(`timeout waiting: ${label}`);
    await sleep(everyMs);
  }
}

async function listDir(dir) {
  try { return readdirSync(dir); } catch (e) { return []; }
}

// best-effort 面板截图（用于 README 配图，失败不影响测试结果）
async function shot(page, name) {
  try {
    const r = await page.send('Page.captureScreenshot', { format: 'png' });
    const p = join(DL_DIR, name);
    writeFileSync(p, Buffer.from(r.data, 'base64'));
    log('截图:', p);
  } catch (e) { dbg('screenshot failed:', e.message); }
}

async function waitForDownload(dir, timeoutMs) {
  const t0 = Date.now();
  let lastSize = -1, lastFile = null, stable = 0;
  for (;;) {
    const names = (await listDir(dir)).filter((n) => /\.pptx$|\.crdownload$/i.test(n));
    if (names.length) {
      const f = names[0];
      const st = statSync(join(dir, f));
      dbg('downloading:', f, st.size);
      if (f.endsWith('.pptx') && st.size === lastSize && st.size > 1000) {
        stable++;
        if (stable >= 2) return { file: join(dir, f), size: st.size };
      } else { stable = 0; }
      lastSize = st.size; lastFile = f;
    }
    if (Date.now() - t0 > timeoutMs) throw new Error(`timeout waiting download (last: ${lastFile})`);
    await sleep(1500);
  }
}

// ---------------- 探测：找静态PPT时间窗 ----------------
// 每5s取一帧灰度签名，返回时间序列 → Node 侧找低差异窗口
const PROBE_EXPR = `
(async () => {
  const v = document.querySelector('.bpx-player-video-wrap video') || document.querySelector('video');
  if (!v) throw new Error('no video');
  const c = document.createElement('canvas'); c.width = 32; c.height = 18;
  const ctx = c.getContext('2d', { willReadFrequently: true });
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));
  const grab = (t) => new Promise((res) => {
    let done = false;
    const fin = () => {
      if (done) return; done = true;
      v.removeEventListener('seeked', fin);
      clearTimeout(timer); clearInterval(poll);
      sleep(80).then(() => {
        try {
          ctx.drawImage(v, 0, 0, 32, 18);
          const d = ctx.getImageData(0, 0, 32, 18).data;
          const g = [];
          for (let i = 0; i < d.length; i += 4) g.push(d[i]);
          res({ t, g });
        } catch (e) { res({ t, g: null }); }
      });
    };
    const poll = setInterval(() => { if (v.readyState >= 2) fin(); }, 60);
    const timer = setTimeout(fin, 3000);
    v.addEventListener('seeked', fin);
    try { v.currentTime = t; } catch (e) { fin(); }
  });
  const dur = v.duration;
  const out = [];
  for (let t = 5; t <= Math.min(dur - 2, 480); t += 5) out.push(await grab(t));
  return JSON.stringify({ dur, samples: out });
})()
`;

function findBestWindow(samples, dur) {
  // 计算相邻帧平均差
  const diffs = [];
  for (let i = 1; i < samples.length; i++) {
    const a = samples[i - 1].g, b = samples[i].g;
    if (!a || !b) { diffs.push(999); continue; }
    let s = 0;
    for (let j = 0; j < a.length; j++) s += Math.abs(a[j] - b[j]);
    diffs.push(s / a.length);
  }
  // 找最长低差异游程（静态画面）
  const T = 10;
  let best = { len: 0, startIdx: 0 };
  let runStart = -1;
  for (let i = 0; i < diffs.length; i++) {
    if (diffs[i] < T) {
      if (runStart < 0) runStart = i;
      const len = i - runStart;
      if (len > best.len) best = { len, startIdx: runStart };
    } else runStart = -1;
  }
  log('探测: 帧间差序列=', JSON.stringify(diffs.map((d) => (d > 999 ? -1 : Math.round(d)))));
  const startSample = samples[best.startIdx < 1 ? 0 : best.startIdx];
  let startT = startSample ? startSample.t : 30;
  startT = Math.max(5, startT);
  let endT = Math.min(startT + 150, dur - 2);
  if (endT - startT < 60) { startT = Math.max(5, dur * 0.2); endT = Math.min(startT + 150, dur - 2); }
  log(`探测: 最长静态游程 ${best.len * 5}s，选扫描窗 ${startT}s ~ ${endT}s`);
  return { startT: Math.round(startT), endT: Math.round(endT) };
}

const fmt = (sec) => { const m = Math.floor(sec / 60); return `${m}:${String(Math.round(sec % 60)).padStart(2, '0')}`; };

// ---------------- 主流程 ----------------
async function main() {
  // 1. 连接 CDP
  const ver = await getJson('/json/version');
  const browser = await Cdp.connect(ver.webSocketDebuggerUrl);
  await browser.send('Browser.setDownloadBehavior', {
    behavior: 'allow', downloadPath: DL_DIR, eventsEnabled: true,
  });
  log('浏览器:', ver.Browser);

  // headless + --load-extension 时不能在命令行带初始URL（Multiple targets 报错），
  // 也不一定存在 page target → 需要时在此创建一个
  let target = null;
  for (let i = 0; i < 30; i++) {
    const list = await getJson('/json/list');
    const page = list.find((t) => t.type === 'page');
    if (page) { target = page; break; }
    if (i === 5) {
      log('无 page target，通过 Target.createTarget 创建…');
      await browser.send('Target.createTarget', { url: 'about:blank' });
    }
    await sleep(1000);
  }
  if (!target) throw new Error('no page target');
  const page = await Cdp.connect(target.webSocketDebuggerUrl);
  await page.send('Runtime.enable');
  await page.send('Page.enable');
  log('已连接页面目标');

  // 2. 打开视频页
  log('导航到:', VIDEO_URL);
  await page.send('Page.navigate', { url: VIDEO_URL });
  await waitFor(page, `!!document.querySelector('#bili-ppt-host')`, 60000, '插件注入');
  log('插件已注入');

  const videoReady = `(()=>{const v=document.querySelector('video');return !!v && isFinite(v.duration) && v.duration>0})()`;
  await waitFor(page, videoReady, 90000, '视频就绪(duration)');
  const dur = await page.eval(`document.querySelector('video').duration`);
  const res2 = await page.eval(`(()=>{const v=document.querySelector('video');return v.videoWidth+'x'+v.videoHeight})()`);
  log(`视频就绪: 时长 ${Math.round(dur)}s, 分辨率 ${res2}`);

  // 3. 探测找静态窗口
  log('探测阶段（程序化seek采样）…');
  const probeStr = await page.eval(PROBE_EXPR, true);
  const probe = JSON.parse(probeStr);
  const { startT, endT } = findBestWindow(probe.samples, probe.dur);

  // 4. 驱动真实 UI
  await page.eval(shadowExpr(`s.querySelector('.bp-fab').click(); return true;`));
  await sleep(500);
  const info = await page.eval(shadowExpr(`return s.getElementById('vinfo').textContent;`));
  log('面板信息:', info);
  await shot(page, 'panel-open.png');
  const scanDisabled = await page.eval(shadowExpr(`return s.getElementById('scanBtn').disabled;`));
  if (scanDisabled) throw new Error('scanBtn 被禁用: ' + info);

  // 设置起止时间（走真实输入框 + change 事件）
  await page.eval(shadowExpr(`
    const a = s.getElementById('t0'), b = s.getElementById('t1');
    a.value = ${JSON.stringify(fmt(startT))}; a.dispatchEvent(new Event('change'));
    b.value = ${JSON.stringify(fmt(endT))}; b.dispatchEvent(new Event('change'));
    return true;
  `));
  log(`已设置扫描范围: ${fmt(startT)} → ${fmt(endT)}`);

  // 开始扫描
  await page.eval(shadowExpr(`s.getElementById('scanBtn').click(); return true;`));
  log('已点击「开始扫描」，等待完成…');

  const scanHiddenExpr = shadowExpr(`return s.getElementById('scanBtn').classList.contains('hidden');`);
  // 等扫描开始（scanBtn hidden）
  let sawScanning = false;
  const tScan0 = Date.now();
  while (Date.now() - tScan0 < 480000) {
    const hidden = await page.eval(scanHiddenExpr);
    if (hidden) sawScanning = true;
    if (sawScanning && !hidden) break;
    if (!sawScanning) {
      const prog = await page.eval(shadowExpr(`return s.getElementById('ptxt1').textContent;`));
      dbg('progress:', prog);
    }
    await sleep(3000);
  }
  if (!sawScanning) throw new Error('扫描未启动');
  log('扫描完成');
  await shot(page, 'scan-done.png');

  const cntTotal = await page.eval(shadowExpr(`return s.getElementById('cntTotal').textContent;`));
  const cntSel = await page.eval(shadowExpr(`return s.getElementById('cntSel').textContent;`));
  const progText = await page.eval(shadowExpr(`return s.getElementById('ptxt1').textContent + ' | ' + s.getElementById('ptxt2').textContent;`));
  log(`识别结果: 共 ${cntTotal} 页, 已选(勾选) ${cntSel} 页`);
  dbg('progress text:', progText);

  if (parseInt(cntTotal, 10) === 0) throw new Error('识别出 0 页 — 扫描逻辑未产出结果');

  // 5. 导出 PPTX
  log('点击「导出 PPTX」…');
  await page.eval(shadowExpr(`s.getElementById('pptxBtn').click(); return true;`));
  const dl = await waitForDownload(DL_DIR, 240000);
  log(`已下载: ${dl.file} (${dl.size} bytes)`);

  console.log(JSON.stringify({
    ok: true,
    video: VIDEO_URL,
    durationSec: Math.round(dur),
    scanRange: [fmt(startT), fmt(endT)],
    pagesDetected: parseInt(cntTotal, 10),
    pagesSelected: parseInt(cntSel, 10),
    pptx: dl.file,
    pptxSize: dl.size,
  }, null, 2));
}

main().catch((e) => { console.error('[e2e] FAILED:', e.message); process.exit(1); });
