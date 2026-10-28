#!/usr/bin/env node
/**
 * E2E 真实验证：启动加载了扩展的 Edge → 打开真实 B 站视频 →
 * 打开扩展面板 → 设定时间范围自动扫描 → 导出 PPTX → 验证产物。
 *
 * 用法: node run-e2e.mjs [视频URL] [--start=秒] [--end=秒] [--headless]
 */
import { spawn } from 'node:child_process';
import { mkdirSync, existsSync, readdirSync, writeFileSync } from 'node:fs';
import { join, dirname, resolve, basename } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(__dirname, '..', '..'); // 仓库根 = 扩展目录
const OUT = join(__dirname, 'downloads');
const PROFILE = join(__dirname, process.env.E2E_PROFILE || 'e2e-profile');
const EDGE = process.env.E2E_BROWSER ||
  'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const PORT = 9333;

const args = process.argv.slice(2);
const headless = args.includes('--headless');
const startArg = getArg(args, '--start') || '0';
const endArg = getArg(args, '--end') || '';
const videoUrl = args.find((a) => !a.startsWith('--')) ||
  'https://www.bilibili.com/video/BV1YE411D7nH/?p=4';

function getArg(list, name) {
  const hit = list.find((a) => a.startsWith(name + '='));
  return hit ? hit.slice(name.length + 1) : null;
}

mkdirSync(OUT, { recursive: true });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ---------------- CDP ----------------
let ws = null;
let msgId = 0;
const pending = new Map();
const consoleLogs = [];
const netLogs = [];

function wsConnect(url) {
  return new Promise((ok, fail) => {
    const s = new WebSocket(url);
    s.onopen = () => { ws = s; ok(s); };
    s.onerror = () => fail(new Error('ws connect fail ' + url));
    s.onmessage = (m) => {
      const d = JSON.parse(m.data);
      if (d.id && pending.has(d.id)) {
        const { res, rej } = pending.get(d.id);
        pending.delete(d.id);
        d.error ? rej(new Error(d.error.message)) : res(d.result);
      } else if (d.method === 'Log.entryAdded') {
        consoleLogs.push(d.params.entry.text?.slice(0, 200));
      } else if (d.method === 'Network.loadingFailed') {
        const e = d.params;
        netLogs.push(`FAILED ${e.type} ${e.errorText} canceled=${e.canceled} ${e.blockedReason || ''}`);
      } else if (d.method === 'Network.requestWillBeSent') {
        const u = d.params.request.url;
        if (/up[cv]+d?fl|\.m4s|\.mp4|akamai|bilivideo|hdslb/.test(u)) netLogs.push('REQ ' + u.slice(0, 150));
      }
    };
  });
}

async function send(method, params = {}, sessionId) {
  const id = ++msgId;
  const msg = { id, method, params };
  if (sessionId) msg.sessionId = sessionId;
  const done = new Promise((res, rej) => pending.set(id, { res, rej }));
  ws.send(JSON.stringify(msg));
  return done;
}

async function evalInPage(S, expression) {
  const r = await S('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true });
  if (r.exceptionDetails) throw new Error('page eval failed: ' + JSON.stringify(r.exceptionDetails.exception?.description || r.exceptionDetails.text));
  return r.result?.value;
}

async function waitFor(S, expr, label, timeoutSec) {
  for (let i = 0; i < timeoutSec * 2; i++) {
    try {
      if (await evalInPage(S, expr) === true) return true;
    } catch (e) { /* nav in flight */ }
    await sleep(500);
  }
  console.error('[e2e] TIMEOUT waiting for:', label);
  try {
    const diag = await evalInPage(S, `(()=>{
      const v=document.querySelector('video');
      return JSON.stringify({url:location.href, hasVideo:!!v,
        readyState:v?v.readyState:null, duration:v&&isFinite(v.duration)?v.duration:null,
        videoWidth:v?v.videoWidth:null, err:v? (v.error? v.error.code : null):null})})()`);
    console.error('[e2e] 诊断:', diag);
  } catch { console.error('[e2e] 诊断失败(页面 eval 抛错)'); }
  return false;
}

async function screenshot(S, name) {
  const { data } = await S('Page.captureScreenshot', { format: 'png' });
  const p = join(OUT, name);
  writeFileSync(p, Buffer.from(data, 'base64'));
  console.log('[e2e] screenshot:', basename(p));
}

async function fetchJson(pathname) {
  const r = await fetch(`http://127.0.0.1:${PORT}${pathname}`);
  return r.json();
}

// ---------------- main ----------------
async function main() {
  console.log('[e2e] 扩展目录:', ROOT);
  console.log('[e2e] 视频:', videoUrl);
  console.log('[e2e] 扫描范围:', `start=${startArg}s end=${endArg || '结尾'}`);

  const flags = [
    `--remote-debugging-port=${PORT}`,
    `--user-data-dir=${PROFILE}`,
    `--load-extension=${ROOT}`,
    `--disable-extensions-except=${ROOT}`,
    '--no-first-run', '--no-default-browser-check',
    '--autoplay-policy=no-user-gesture-required',
    '--mute-audio', '--window-size=1440,900',
  ];
  if (headless) flags.push('--headless=new');
  else flags.push(videoUrl);

  console.log('[e2e] 启动 Edge...');
  const edge = spawn(EDGE, flags, { stdio: 'ignore' });
  edge.on('exit', (c) => console.log('[e2e] Edge 退出码:', c));
  const cleanup = () => { try { edge.kill(); } catch {} };
  process.on('exit', cleanup);
  process.on('SIGINT', () => { cleanup(); process.exit(1); });

  // 等 DevTools
  let ok = false;
  for (let i = 0; i < 40 && !ok; i++) {
    try { ok = (await fetchJson('/json/list')).length > 0; } catch { /* retry */ }
    if (!ok) await sleep(500);
  }
  if (!ok) throw new Error('DevTools 端口未就绪');
  console.log('[e2e] DevTools 端口已就绪');

  // 等视频页 target
  let target = null;
  for (let i = 0; i < 60 && !target; i++) {
    try {
      const list = await fetchJson('/json/list');
      target = list.find((t) => t.type === 'page' && t.url.includes('bilibili.com/video'));
    } catch { /* retry */ }
    if (!target) await sleep(1000);
  }
  if (!target) throw new Error('视频页 target 未出现');
  console.log('[e2e] 页面:', target.title || target.url);

  // 直接连接页面 target 的 WS,不走 browser 级 Target.* API
  await wsConnect(target.webSocketDebuggerUrl);
  const S = (m, p) => send(m, p);
  await S('Page.enable');
  await S('Runtime.enable');
  await S('Log.enable');
  await S('Network.enable');

  // 1. 等 video 元素出现,然后主动触发播放,等流数据就绪
  console.log('[e2e] 等待 <video> 元素...');
  if (!await waitFor(S, `!!document.querySelector('video')`, 'video element', 60))
    throw new Error('<video> 元素未出现');
  // 用 CDP 派发受信任的鼠标点击到播放器中央,触发 B 站播放器真正启动
  console.log('[e2e] 点击播放器中央(受信任事件)...');
  const rect = await evalInPage(S, `(()=>{
    const p=document.querySelector('#bilibili-player, .bpx-player-container, .player-area');
    const r=p?p.getBoundingClientRect():null;
    return r?JSON.stringify({x:r.x+r.width/2,y:r.y+r.height/2,w:r.width,h:r.height}):null})()`);
  if (rect) {
    const { x, y } = JSON.parse(rect);
    console.log('[e2e] player center:', x, y);
    await S('Input.dispatchMouseEvent', { type: 'mousePressed', x, y, button: 'left', clickCount: 1 });
    await S('Input.dispatchMouseEvent', { type: 'mouseReleased', x, y, button: 'left', clickCount: 1 });
  } else {
    console.log('[e2e] 未找到播放器容器,再点击 video 自身');
    const vr = await evalInPage(S, `(()=>{const r=document.querySelector('video').getBoundingClientRect();return JSON.stringify({x:r.x+r.width/2,y:r.y+r.height/2})})()`);
    const { x, y } = JSON.parse(vr);
    await S('Input.dispatchMouseEvent', { type: 'mousePressed', x, y, button: 'left', clickCount: 1 });
    await S('Input.dispatchMouseEvent', { type: 'mouseReleased', x, y, button: 'left', clickCount: 1 });
  }
  console.log('[e2e] 触发播放...');
  let ready = false;
  for (let i = 0; i < 60 && !ready; i++) {
    await evalInPage(S, `(()=>{const v=document.querySelector('video');
      if(!v) return 'novideo';
      v.muted=true; v.volume=0;
      try{ v.play().catch(()=>{}); }catch(e){}
      return v.readyState+'/'+(isFinite(v.duration)?v.duration:'inf')})()`).then(r => { if (i % 10 === 0) console.log('[e2e] 播放器状态:', r); });
    ready = await evalInPage(S, `(()=>{const v=document.querySelector('video');
      return !!v && v.readyState>=2 && isFinite(v.duration) && v.duration>0})()`);
    if (!ready) await sleep(2000);
  }
  if (!ready) {
    // 篡诊断:截图 + 播放器深状态
    await screenshot(S, '00-diag-not-ready.png');
    try {
      const diag = await evalInPage(S, `(()=>{
        const v=document.querySelector('video');
        const txt=document.body.innerText.slice(0,300);
        return JSON.stringify({
          currentSrc:v?v.currentSrc:null, networkState:v?v.networkState:null,
          playinfo:!!window.__playinfo__, title:document.title,
          centerBtn:!!document.querySelector('.bpx-player-center-click-wrap, .squirtle-video-start'),
          bodySnippet:txt.replace(/\\s+/g,' ').slice(0,200)})})()`);
      console.error('[e2e] 深诊断:', diag);
    } catch (e) { console.error('[e2e] 深诊断失败:', e.message); }
    // 尝试点中央播放按钮再等一轮
    console.log('[e2e] 尝试点击播放器中央按钮并 reload...');
    await evalInPage(S, `(()=>{const b=document.querySelector('.bpx-player-center-click-wrap, .squirtle-video-start');if(b)b.click();return 'clicked:'+!!b})()`);
    await sleep(8000);
    ready = await evalInPage(S, `(()=>{const v=document.querySelector('video');return !!v&&v.readyState>=2&&isFinite(v.duration)&&v.duration>0})()`);
    if (!ready) {
      await evalInPage(S, `location.reload()`);
      await sleep(12000);
      for (let i = 0; i < 30 && !ready; i++) {
        await evalInPage(S, `(()=>{const v=document.querySelector('video');if(v){v.muted=true;try{v.play().catch(()=>{})}catch(e){}}})()`);
        ready = await evalInPage(S, `(()=>{const v=document.querySelector('video');return !!v&&v.readyState>=2&&isFinite(v.duration)&&v.duration>0})()`);
        if (!ready) await sleep(2000);
      }
    }
  }
  if (!ready) {
    await screenshot(S, '00-diag-final.png');
    console.error('[e2e] === 网络日志(流请求/失败) ===');
    netLogs.slice(-40).forEach((l) => console.error('  ', l));
    console.error('[e2e] === 浏览器日志(尾 20 条) ===');
    consoleLogs.slice(-20).forEach((l) => console.error('  ', l));
    throw new Error('<video> 流数据未就绪(播放触发失败)');
  }
  const dur = await evalInPage(S, `document.querySelector('video').duration`);
  console.log(`[e2e] 视频就绪,时长 ${dur.toFixed(0)}s`);

  // 2. 等扩展注入
  console.log('[e2e] 等待扩展注入...');
  if (!await waitFor(S, `!!document.querySelector('#bili-ppt-host')?.shadowRoot`, 'extension FAB', 60))
    throw new Error('扩展未注入(CI 环境扩展加载失败?)');
  console.log('[e2e] 扩展已加载 ✓');
  await screenshot(S, '01-page-loaded.png');

  // 3. 播放 3 秒触发清晰度提升(从 360p 升到更高)再暂停
  console.log('[e2e] 播放 5s 让播放器提升画质...');
  await evalInPage(S, `(()=>{const v=document.querySelector('video');v.muted=true;v.play();return 'playing'})()`);
  await sleep(5000);
  await evalInPage(S, `document.querySelector('video').pause()`);
  const res = await evalInPage(S, `(()=>{const v=document.querySelector('video');return v.videoWidth+'x'+v.videoHeight})()`);
  console.log('[e2e] 当前分辨率:', res);
  await screenshot(S, '02-after-play.png');

  // 4. 打开面板 + 设置时间范围
  console.log('[e2e] 打开扩展面板...');
  await evalInPage(S, `document.querySelector('#bili-ppt-host').shadowRoot.querySelector('.bp-fab').click()`);
  if (!await waitFor(S, `document.querySelector('#bili-ppt-host').shadowRoot.querySelector('.bp-panel')?.classList.contains('open')`, 'panel open', 15))
    throw new Error('面板未打开');
  console.log('[e2e] 面板已打开 ✓');
  await screenshot(S, '03-panel-open.png');

  // 设置扫描范围
  await evalInPage(S, `(()=>{const sh=document.querySelector('#bili-ppt-host').shadowRoot;
    sh.querySelector('#t0').value='${startArg}';
    sh.querySelector('#t0').dispatchEvent(new Event('change',{bubbles:true}));
    sh.querySelector('#t1').value='${endArg}';
    sh.querySelector('#t1').dispatchEvent(new Event('change',{bubbles:true}));
    return 'range set'})()`);

  // 5. 开始扫描
  console.log('[e2e] 点击开始扫描...');
  await evalInPage(S, `document.querySelector('#bili-ppt-host').shadowRoot.querySelector('#scanBtn').click()`);

  // 6. 轮询扫描进度直到完成(扫描中「停止」按钮可见)
  console.log('[e2e] 等待扫描完成...');
  await sleep(2000); // 让扫描先启动
  const scanTimeoutSec = 25 * 60; // 最长 25 分钟
  let done = false, lastProgress = '';
  for (let i = 0; i < scanTimeoutSec * 2 && !done; i++) {
    const st = await evalInPage(S, `(()=>{const sh=document.querySelector('#bili-ppt-host').shadowRoot;
      const scanning=!sh.querySelector('#stopBtn').classList.contains('hidden');
      const prog=sh.querySelector('#prog')?.textContent||'';
      const thumbs=sh.querySelectorAll('.bp-thumb').length;
      return JSON.stringify({scanning,prog,thumbs})})()`);
    if (st !== lastProgress) { lastProgress = st; console.log('[e2e]', st); }
    done = !JSON.parse(st).scanning;
    if (!done) await sleep(1000);
  }
  if (!done) throw new Error('扫描超时');
  await screenshot(S, '04-scan-done.png');
  const pages = await evalInPage(S, `document.querySelectorAll('#bili-ppt-host').shadowRoot.querySelectorAll('.bp-thumb').length`);
  console.log(`[e2e] 扫描完成,识别 ${pages} 页 ✅`);
  if (!pages || Number(pages) === 0) throw new Error('扫描完成但 0 页 — 检测失败');

  // 7. 导出 PPTX
  console.log('[e2e] 点击导出 PPTX...');
  const before = listDownloads();
  await evalInPage(S, `document.querySelector('#bili-ppt-host').shadowRoot.querySelector('#pptxBtn').click()`);
  let pptxFile = null;
  for (let i = 0; i < 60 && !pptxFile; i++) {
    await sleep(1000);
    pptxFile = listDownloads().find((f) => f.endsWith('.pptx') && !before.includes(f));
  }
  await screenshot(S, '05-exported.png');
  if (!pptxFile) throw new Error('PPTX 未下载');
  console.log('[e2e] PPTX 已下载:', pptxFile);

  console.log('[e2e] ===== E2E 全流程成功 =====');
  console.log('[e2e] 产物: ' + pptxFile);
  cleanup();
  process.exit(0);
}

function listDownloads() {
  const sys = join(process.env.USERPROFILE || '', 'Downloads');
  const dir = existsSync(sys) ? sys : OUT;
  try { return readdirSync(dir).map((f) => join(dir, f)); } catch { return []; }
}

main().catch((e) => { console.error('[e2e] FATAL:', e.message); process.exit(1); });
