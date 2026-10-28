/**
 * Node 环境测试：用真实的 pptx-builder.js 生成 PPTX。
 * 用生产的同一份代码（不是复制品）来验证，保证浏览器里跑的就是验证过的逻辑。
 */
const fs = require('fs');
const path = require('path');
const builder = require('../pptx-builder.js');

const SLIDES_DIR = path.join(__dirname, 'slides');
const OUT = path.join(__dirname, 'test-output.pptx');

const slides = [1, 2, 3, 4].map((n) => {
  const bytes = new Uint8Array(fs.readFileSync(path.join(SLIDES_DIR, `slide${n}.png`)));
  return { bytes, w: 1600, h: 900, ext: 'png' };
});

const pptx = builder.buildPptx(slides);
fs.writeFileSync(OUT, Buffer.from(pptx));

console.log(`test-output.pptx OK, ${pptx.length} bytes, ${slides.length} slides`);

// 顺带验证 buildZip：把 4 张图打包成 zip
const zip = builder.buildZip(slides.map((s, i) => ({ name: `p${i + 1}.png`, data: s.bytes })));
fs.writeFileSync(path.join(__dirname, 'test-output.zip'), Buffer.from(zip));
console.log(`test-output.zip OK, ${zip.length} bytes`);
