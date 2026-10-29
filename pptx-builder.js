/*!
 * BiliPPT pptx-builder — 在浏览器/Node 中生成最小化但合法的 PPTX（Office Open XML）。
 * 内置 CRC32、Store-only ZIP 写入器与 PPTX 模板，零第三方依赖。
 *
 * 对外 API（挂到全局 BiliPptx）：
 *   buildPptx(slides, opts?) -> Uint8Array
 *     slides: [{ bytes: Uint8Array, w: number, h: number, ext?: 'jpg'|'png' }]
 *   buildZip(files) -> Uint8Array
 *     files:  [{ name: string, data: Uint8Array }]
 */
(function (global) {
  'use strict';

  var NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  var NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  var REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

  var XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

  // ---------------- CRC32 ----------------
  var CRC_TABLE = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(data) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < data.length; i++) c = CRC_TABLE[(c ^ data[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  // ---------------- Store-only ZIP ----------------
  // 图片(jpg/png)本身已压缩，用 STORE 方式打包完全合法，且实现极简、体积无差。
  function buildZip(files) {
    var enc = new TextEncoder();
    var DOS_TIME = 0; // 00:00:00
    var DOS_DATE = ((2024 - 1980) << 9) | (1 << 5) | 1; // 2024-01-01

    var parts = [];   // all byte chunks in order
    var central = []; // central directory entries
    var offset = 0;

    function push(buf) {
      parts.push(buf);
      offset += buf.length;
    }

    for (var fi = 0; fi < files.length; fi++) {
      var f = files[fi];
      var nameBytes = enc.encode(f.name);
      var data = f.data;
      var crc = crc32(data);

      var local = new Uint8Array(30 + nameBytes.length);
      var dv = new DataView(local.buffer);
      dv.setUint32(0, 0x04034b50, true);      // local file header signature
      dv.setUint16(4, 20, true);               // version needed to extract
      dv.setUint16(6, 0x0800, true);           // flags: UTF-8 filename
      dv.setUint16(8, 0, true);               // method: store
      dv.setUint16(10, DOS_TIME, true);
      dv.setUint16(12, DOS_DATE, true);
      dv.setUint32(14, crc, true);
      dv.setUint32(18, data.length, true);     // compressed size
      dv.setUint32(22, data.length, true);     // uncompressed size
      dv.setUint16(26, nameBytes.length, true);
      dv.setUint16(28, 0, true);               // extra field length
      local.set(nameBytes, 30);

      push(local);
      push(data);

      central.push({
        nameBytes: nameBytes,
        crc: crc,
        size: data.length,
        offset: offset - local.length - data.length,
      });
    }

    var cdStart = offset;
    var cdCount = central.length;

    for (var ci = 0; ci < cdCount; ci++) {
      var e = central[ci];
      var entry = new Uint8Array(46 + e.nameBytes.length);
      var ev = new DataView(entry.buffer);
      ev.setUint32(0, 0x02014b50, true);      // central directory signature
      ev.setUint16(4, 20, true);               // version made by
      ev.setUint16(6, 20, true);               // version needed
      ev.setUint16(8, 0x0800, true);           // flags
      ev.setUint16(10, 0, true);               // method
      ev.setUint16(12, DOS_TIME, true);
      ev.setUint16(14, DOS_DATE, true);
      ev.setUint32(16, e.crc, true);
      ev.setUint32(20, e.size, true);
      ev.setUint32(24, e.size, true);
      ev.setUint16(28, e.nameBytes.length, true);
      ev.setUint16(30, 0, true);               // extra len
      ev.setUint16(32, 0, true);               // comment len
      ev.setUint16(34, 0, true);               // disk number start
      ev.setUint16(36, 0, true);               // internal attrs
      ev.setUint32(38, 0, true);               // external attrs
      ev.setUint32(42, e.offset, true);        // relative offset of local header
      entry.set(e.nameBytes, 46);
      push(entry);
    }

    var eocd = new Uint8Array(22);
    var eov = new DataView(eocd.buffer);
    eov.setUint32(0, 0x06054b50, true);
    eov.setUint16(4, 0, true);          // disk number
    eov.setUint16(6, 0, true);          // disk with central dir
    eov.setUint16(8, cdCount, true);
    eov.setUint16(10, cdCount, true);
    eov.setUint32(12, offset - cdStart, true); // size of central dir
    eov.setUint32(16, cdStart, true);           // offset of central dir
    eov.setUint16(20, 0, true);                 // comment len
    push(eocd);

    // concat
    var total = offset;
    var out = new Uint8Array(total);
    var pos = 0;
    for (var pi = 0; pi < parts.length; pi++) {
      out.set(parts[pi], pos);
      pos += parts[pi].length;
    }
    return out;
  }

  // ---------------- XML 模板 ----------------
  function esc(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function contentTypesXml(usedExt, slideCount) {
    var s = XML_DECL +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>';
    if (usedExt.has('jpg') || usedExt.has('jpeg')) {
      s += '<Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="jpeg" ContentType="image/jpeg"/>';
    }
    if (usedExt.has('png')) {
      s += '<Default Extension="png" ContentType="image/png"/>';
    }
    s += '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>';
    s += '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>';
    s += '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>';
    s += '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>';
    for (var i = 1; i <= slideCount; i++) {
      s += '<Override PartName="/ppt/slides/slide' + i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
    }
    s += '</Types>';
    return s;
  }

  function relsXml(rels) {
    var s = XML_DECL + '<Relationships xmlns="' + REL_NS + '">';
    for (var i = 0; i < rels.length; i++) {
      var r = rels[i];
      s += '<Relationship Id="' + r.id + '" Type="' + r.type + '" Target="' + esc(r.target) + '"/>';
    }
    return s + '</Relationships>';
  }

  var T = NS_R + '/';

  function presentationXml(slideCount, cx, cy) {
    var s = XML_DECL +
      '<p:presentation xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '" saveSubsetFonts="1">' +
      '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
      '<p:sldIdLst>';
    for (var i = 0; i < slideCount; i++) {
      s += '<p:sldId id="' + (256 + i) + '" r:id="rId' + (2 + i) + '"/>';
    }
    s += '</p:sldIdLst>';
    s += '<p:sldSz cx="' + cx + '" cy="' + cy + '"/>';
    s += '<p:notesSz cx="6858000" cy="9144000"/>';
    s += '</p:presentation>';
    return s;
  }

  function presentationRelsXml(slideCount) {
    var rels = [{ id: 'rId1', type: T + 'slideMaster', target: 'slideMasters/slideMaster1.xml' }];
    for (var i = 0; i < slideCount; i++) {
      rels.push({ id: 'rId' + (2 + i), type: T + 'slide', target: 'slides/slide' + (i + 1) + '.xml' });
    }
    rels.push({ id: 'rId' + (2 + slideCount), type: T + 'theme', target: 'theme/theme1.xml' });
    return relsXml(rels);
  }

  var MASTER_RELS_XML = relsXml([
    { id: 'rId1', type: T + 'slideLayout', target: '../slideLayouts/slideLayout1.xml' },
    { id: 'rId2', type: T + 'theme', target: '../theme/theme1.xml' },
  ]);

  var LAYOUT_RELS_XML = relsXml([
    { id: 'rId1', type: T + 'slideMaster', target: '../slideMasters/slideMaster1.xml' },
  ]);

  function emptySpTree(cx, cy) {
    return '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="' + cx + '" cy="' + cy + '"/></a:xfrm></p:grpSpPr>';
  }

  function slideMasterXml(cx, cy) {
    return XML_DECL +
      '<p:sldMaster xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '">' +
      '<p:cSld>' +
      '<p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>' +
      '<p:spTree>' + emptySpTree(cx, cy) + '</p:spTree>' +
      '</p:cSld>' +
      '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
      '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
      '</p:sldMaster>';
  }

  function slideLayoutXml(cx, cy) {
    return XML_DECL +
      '<p:sldLayout xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '" type="blank" preserve="1">' +
      '<p:cSld name="Blank">' +
      '<p:spTree>' + emptySpTree(cx, cy) + '</p:spTree>' +
      '</p:cSld>' +
      '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
      '</p:sldLayout>';
  }

  function slideXml(cx, cy) {
    return XML_DECL +
      '<p:sld xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '">' +
      '<p:cSld>' +
      '<p:spTree>' + emptySpTree(cx, cy) +
      '<p:pic>' +
      '<p:nvPicPr><p:cNvPr id="2" name="Slide"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
      '</p:pic>' +
      '</p:spTree>' +
      '</p:cSld>' +
      '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
      '</p:sld>';
  }

  function slideRelsXml(n, ext) {
    return relsXml([
      { id: 'rId1', type: T + 'slideLayout', target: '../slideLayouts/slideLayout1.xml' },
      { id: 'rId2', type: T + 'image', target: '../media/image' + n + '.' + ext },
    ]);
  }

  var THEME_XML = XML_DECL +
    '<a:theme xmlns:a="' + NS_A + '" name="BiliPPT Theme">' +
    '<a:themeElements>' +
    '<a:clrScheme name="Office">' +
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
    '<a:dk2><a:srgbClr val="44546A"/></a:dk2>' +
    '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
    '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>' +
    '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
    '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>' +
    '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
    '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>' +
    '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
    '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>' +
    '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
    '</a:clrScheme>' +
    '<a:fontScheme name="Office">' +
    '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
    '</a:fontScheme>' +
    '<a:fmtScheme name="Office">' +
    '<a:fillStyleLst>' +
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="100000"/><a:satMod val="180000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="78000"/><a:satMod val="180000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="86000"/><a:satMod val="180000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>' +
    '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:shade val="83000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="73000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>' +
    '</a:fillStyleLst>' +
    '<a:lnStyleLst>' +
    '<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>' +
    '<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>' +
    '<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>' +
    '</a:lnStyleLst>' +
    '<a:effectStyleLst>' +
    '<a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>' +
    '<a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="42000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>' +
    '<a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="42000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle>' +
    '</a:effectStyleLst>' +
    '<a:bgFillStyleLst>' +
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="180000"/></a:schemeClr></a:solidFill>' +
    '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="100000"/><a:satMod val="180000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="73000"/><a:satMod val="180000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="76000"/><a:satMod val="180000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>' +
    '</a:bgFillStyleLst>' +
    '</a:fmtScheme>' +
    '</a:themeElements>' +
    '</a:theme>';

  // ---------------- PPTX ----------------
  function buildPptx(slides, opts) {
    if (!slides || !slides.length) throw new Error('buildPptx: slides 为空');
    for (var i = 0; i < slides.length; i++) {
      var s = slides[i];
      if (!s || !s.bytes) throw new Error('buildPptx: 第 ' + (i + 1) + ' 页缺少图片数据');
    }
    opts = opts || {};

    var first = slides[0];
    var cx, cy;
    if (first.w >= first.h) {
      cx = 12192000;
      cy = Math.round(12192000 * first.h / first.w);
    } else {
      cy = 9144000;
      cx = Math.round(9144000 * first.w / first.h);
    }
    if (!isFinite(cx) || !isFinite(cy) || cx <= 0 || cy <= 0) throw new Error('buildPptx: 图片尺寸非法');

    var usedExt = new Set();
    for (var j = 0; j < slides.length; j++) {
      usedExt.add((slides[j].ext || 'jpg').toLowerCase());
    }

    var enc = new TextEncoder();
    function str(x) { return enc.encode(x); }

    var files = [];
    files.push({ name: '[Content_Types].xml', data: str(contentTypesXml(usedExt, slides.length)) });
    files.push({ name: '_rels/.rels', data: str(relsXml([
      { id: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', target: 'ppt/presentation.xml' },
    ])) });
    files.push({ name: 'ppt/presentation.xml', data: str(presentationXml(slides.length, cx, cy)) });
    files.push({ name: 'ppt/_rels/presentation.xml.rels', data: str(presentationRelsXml(slides.length)) });
    files.push({ name: 'ppt/slideMasters/slideMaster1.xml', data: str(slideMasterXml(cx, cy)) });
    files.push({ name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', data: str(MASTER_RELS_XML) });
    files.push({ name: 'ppt/slideLayouts/slideLayout1.xml', data: str(slideLayoutXml(cx, cy)) });
    files.push({ name: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', data: str(LAYOUT_RELS_XML) });
    files.push({ name: 'ppt/theme/theme1.xml', data: str(THEME_XML) });
    for (var k = 0; k < slides.length; k++) {
      var sl = slides[k];
      var n = k + 1;
      var ext = (sl.ext || 'jpg').toLowerCase();
      if (ext === 'jpeg') ext = 'jpg';
      files.push({ name: 'ppt/slides/slide' + n + '.xml', data: str(slideXml(cx, cy)) });
      files.push({ name: 'ppt/slides/_rels/slide' + n + '.xml.rels', data: str(slideRelsXml(n, ext)) });
      files.push({ name: 'ppt/media/image' + n + '.' + ext, data: sl.bytes });
    }
    return buildZip(files);
  }

  var api = { buildPptx: buildPptx, buildZip: buildZip, crc32: crc32 };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  global.BiliPptx = api;
})(typeof self !== 'undefined' ? self : globalThis);
