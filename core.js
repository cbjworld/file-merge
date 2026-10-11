/* 병합 코어 (브라우저/Node 공용). 전역 의존: JSZip, ExcelJS, DOMParser, XMLSerializer */
const Merge = (() => {
  'use strict';

  class MergeError extends Error {}

  const HH = 'http://www.hancom.co.kr/hwpml/2011/head';
  const OPF = 'http://www.idpf.org/2007/opf/';
  const NONE_REF = '4294967295';
  const LANGS = ['HANGUL', 'LATIN', 'HANJA', 'JAPANESE', 'OTHER', 'SYMBOL', 'USER'];
  const LIST_ORDER = ['fontfaces', 'borderFills', 'charProperties', 'tabProperties', 'numberings',
    'bullets', 'paraProperties', 'styles', 'memoProperties', 'trackChanges', 'trackChangeAuthors'];
  const ATTR_KIND = {
    borderFillIDRef: 'borderFill', charPrIDRef: 'charPr', tabPrIDRef: 'tabPr', paraPrIDRef: 'paraPr',
    styleIDRef: 'style', outlineShapeIDRef: 'numbering', memoShapeIDRef: 'memoPr',
  };

  /* ---------------- 공용 DOM 도우미 ---------------- */
  const parseXml = (text) => {
    const d = new DOMParser().parseFromString(text.replace(/^﻿/, ''), 'application/xml');
    if (d.getElementsByTagName('parsererror').length) throw new MergeError('XML을 해석하지 못했습니다.');
    return d;
  };
  const serializeXml = (doc) =>
    '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>' +
    new XMLSerializer().serializeToString(doc.documentElement);
  const kids = (el, local) => Array.from(el.childNodes).filter((n) => n.nodeType === 1 && n.localName === local);
  const kid = (el, local) => kids(el, local)[0] || null;
  const allEls = (root) => [root, ...Array.from(root.getElementsByTagName('*'))];
  const canon = (el, drop = ['id']) => {
    const c = el.cloneNode(true);
    drop.forEach((a) => c.removeAttribute(a));
    return new XMLSerializer().serializeToString(c);
  };
  const dirOf = (p) => (p.includes('/') ? p.slice(0, p.lastIndexOf('/') + 1) : '');
  const baseOf = (p) => p.slice(p.lastIndexOf('/') + 1);

  /* ================= HWPX ================= */

  async function loadPkg(file) {
    let zip;
    try {
      zip = await JSZip.loadAsync(file.data);
    } catch (e) {
      throw new MergeError('HWPX(.hwpx) 형식이 아닙니다. .hwp 파일은 한글에서 "HWPX로 저장"한 뒤 사용하세요.');
    }
    const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
    const text = (n) => zip.file(n).async('string');
    const bytes = (n) => zip.file(n).async('uint8array');

    let hpfPath = null;
    if (names.includes('META-INF/container.xml')) {
      const c = parseXml(await text('META-INF/container.xml'));
      for (const e of allEls(c.documentElement)) {
        if (e.localName === 'rootfile' && names.includes(e.getAttribute('full-path'))) {
          hpfPath = e.getAttribute('full-path');
          break;
        }
      }
    }
    if (!hpfPath) hpfPath = names.find((n) => /content\.hpf$/i.test(n)) || null;
    if (!hpfPath) throw new MergeError('HWPX(.hwpx) 형식이 아닙니다. (content.hpf 없음)');

    const hpf = parseXml(await text(hpfPath));
    const resolve = (href) => {
      if (names.includes(href)) return href;
      const alt = dirOf(hpfPath) + href;
      return names.includes(alt) ? alt : href;
    };
    const items = new Map();
    for (const it of Array.from(hpf.getElementsByTagNameNS(OPF, 'item'))) items.set(it.getAttribute('id'), it);
    const spine = Array.from(hpf.getElementsByTagNameNS(OPF, 'itemref')).map((r) => r.getAttribute('idref'));

    let headerHref = null;
    for (const [, el] of items) {
      const h = el.getAttribute('href') || '';
      if (/header\.xml$/i.test(h)) { headerHref = resolve(h); break; }
    }
    if (!headerHref) throw new MergeError('header.xml을 찾지 못했습니다.');

    let sectionHrefs = [];
    for (const id of spine) {
      const el = items.get(id);
      const h = el && el.getAttribute('href');
      if (h && /section\d+\.xml$/i.test(h)) sectionHrefs.push(resolve(h));
    }
    if (!sectionHrefs.length) {
      sectionHrefs = names.filter((n) => /Contents\/section\d+\.xml$/i.test(n))
        .sort((a, b) => parseInt(a.match(/(\d+)\.xml$/)[1], 10) - parseInt(b.match(/(\d+)\.xml$/)[1], 10));
    }
    if (!sectionHrefs.length) throw new MergeError('본문(section)을 찾지 못했습니다.');
    return { zip, names, text, bytes, hpfPath, hpf, items, resolve, headerHref, sectionHrefs };
  }

  function remapRefs(root, maps, binmap) {
    for (const e of allEls(root)) {
      for (const attr in ATTR_KIND) {
        const v = e.getAttribute(attr);
        if (v !== null && v !== NONE_REF) {
          const nv = maps[ATTR_KIND[attr]] && maps[ATTR_KIND[attr]][v];
          if (nv !== undefined) e.setAttribute(attr, nv);
        }
      }
      const b = e.getAttribute('binaryItemIDRef');
      if (b !== null && binmap[b] !== undefined) e.setAttribute('binaryItemIDRef', binmap[b]);
    }
  }

  function insertList(refList, listName, incList) {
    const doc = refList.ownerDocument;
    const shell = doc.createElementNS(incList.namespaceURI, incList.prefix ? incList.prefix + ':' + listName : listName);
    for (const a of Array.from(incList.attributes)) shell.setAttribute(a.name, a.value);
    const my = LIST_ORDER.indexOf(listName);
    const children = Array.from(refList.childNodes).filter((n) => n.nodeType === 1);
    let before = null;
    for (const c of children) {
      const i = LIST_ORDER.indexOf(c.localName);
      if (i > -1 && i > (my < 0 ? LIST_ORDER.length : my)) { before = c; break; }
    }
    refList.insertBefore(shell, before);
    return shell;
  }

  const maxId = (els) => Math.max(-1, ...els.map((e) => parseInt(e.getAttribute('id'), 10)).filter((n) => !isNaN(n)));

  function mergeHeader(baseRoot, incRoot, idx, binmap) {
    const baseRef = kid(baseRoot, 'refList');
    const incRef = kid(incRoot, 'refList');
    if (!baseRef || !incRef) throw new MergeError('header.xml 구조(refList)를 인식하지 못했습니다.');
    const baseDoc = baseRoot.ownerDocument;

    const maps = { borderFill: {}, tabPr: {}, charPr: {}, numbering: {}, bullet: {}, paraPr: {}, style: {}, memoPr: {} };
    const fontmaps = {};
    LANGS.forEach((l) => { fontmaps[l] = {}; });

    // 1) 글꼴
    let baseFF = kid(baseRef, 'fontfaces');
    const incFF = kid(incRef, 'fontfaces');
    if (incFF && !baseFF) baseFF = insertList(baseRef, 'fontfaces', incFF);
    if (incFF) {
      for (const incFace of kids(incFF, 'fontface')) {
        const lang = incFace.getAttribute('lang');
        let baseFace = kids(baseFF, 'fontface').find((f) => f.getAttribute('lang') === lang);
        if (!baseFace) {
          baseFace = baseDoc.createElementNS(incFace.namespaceURI, incFace.prefix ? incFace.prefix + ':fontface' : 'fontface');
          for (const a of Array.from(incFace.attributes)) baseFace.setAttribute(a.name, a.value);
          baseFF.appendChild(baseFace);
        }
        const keys = new Map(kids(baseFace, 'font').map((f) => [canon(f), f.getAttribute('id')]));
        let nextId = maxId(kids(baseFace, 'font')) + 1;
        for (const f of kids(incFace, 'font')) {
          const k = canon(f);
          const old = f.getAttribute('id');
          if (keys.has(k)) {
            if (fontmaps[lang]) fontmaps[lang][old] = keys.get(k);
          } else {
            const nf = baseDoc.importNode(f, true);
            nf.setAttribute('id', String(nextId));
            baseFace.appendChild(nf);
            keys.set(k, String(nextId));
            if (fontmaps[lang]) fontmaps[lang][old] = String(nextId);
            nextId++;
          }
        }
        baseFace.setAttribute('fontCnt', String(kids(baseFace, 'font').length));
      }
      baseFF.setAttribute('itemCnt', String(kids(baseFF, 'fontface').length));
    }

    function mergeList(listName, itemTag, kind, prep) {
      const incList = kid(incRef, listName);
      if (!incList) return;
      let baseList = kid(baseRef, listName);
      if (!baseList) baseList = insertList(baseRef, listName, incList);
      const keys = new Map();
      for (const it of kids(baseList, itemTag)) {
        const k = canon(it);
        if (!keys.has(k)) keys.set(k, it.getAttribute('id'));
      }
      let nextId = maxId(kids(baseList, itemTag)) + 1;
      for (const it of kids(incList, itemTag)) {
        const old = it.getAttribute('id');
        const ni = baseDoc.importNode(it, true);
        remapRefs(ni, maps, binmap);
        if (prep) prep(ni);
        const k = canon(ni);
        if (keys.has(k)) { maps[kind][old] = keys.get(k); continue; }
        const newId = String(nextId++);
        ni.setAttribute('id', newId);
        baseList.appendChild(ni);
        keys.set(k, newId);
        maps[kind][old] = newId;
      }
      baseList.setAttribute('itemCnt', String(kids(baseList, itemTag).length));
    }

    const fontPrep = (charPr) => {
      const fr = kid(charPr, 'fontRef');
      if (!fr) return;
      for (const lang of LANGS) {
        const v = fr.getAttribute(lang.toLowerCase());
        if (v !== null && fontmaps[lang][v] !== undefined) fr.setAttribute(lang.toLowerCase(), fontmaps[lang][v]);
      }
    };
    const headingPrep = (paraPr) => {
      for (const hd of Array.from(paraPr.getElementsByTagNameNS(HH, 'heading'))) {
        const t = hd.getAttribute('type');
        const v = hd.getAttribute('idRef');
        if ((t === 'OUTLINE' || t === 'NUMBER') && maps.numbering[v] !== undefined) hd.setAttribute('idRef', maps.numbering[v]);
        else if (t === 'BULLET' && maps.bullet[v] !== undefined) hd.setAttribute('idRef', maps.bullet[v]);
      }
    };

    mergeList('borderFills', 'borderFill', 'borderFill');
    mergeList('tabProperties', 'tabPr', 'tabPr');
    mergeList('charProperties', 'charPr', 'charPr', fontPrep);
    mergeList('numberings', 'numbering', 'numbering');
    mergeList('bullets', 'bullet', 'bullet');
    mergeList('paraProperties', 'paraPr', 'paraPr', headingPrep);
    mergeList('memoProperties', 'memoPr', 'memoPr');

    // 스타일 (다음 스타일 참조가 앞/뒤를 가리킬 수 있어 2단계)
    const incStyles = kid(incRef, 'styles');
    if (incStyles) {
      let baseStyles = kid(baseRef, 'styles');
      if (!baseStyles) baseStyles = insertList(baseRef, 'styles', incStyles);
      const dropKeys = ['id', 'nextStyleIDRef'];
      const keys = new Map();
      for (const st of kids(baseStyles, 'style')) {
        const k = canon(st, dropKeys);
        if (!keys.has(k)) keys.set(k, st.getAttribute('id'));
      }
      const names = new Set(kids(baseStyles, 'style').map((s) => s.getAttribute('name')));
      let nextId = maxId(kids(baseStyles, 'style')) + 1;
      const appended = [];
      for (const st of kids(incStyles, 'style')) {
        const old = st.getAttribute('id');
        const ns = baseDoc.importNode(st, true);
        remapRefs(ns, maps, binmap);
        const k = canon(ns, dropKeys);
        if (keys.has(k)) { maps.style[old] = keys.get(k); continue; }
        const newId = String(nextId++);
        ns.setAttribute('id', newId);
        if (names.has(ns.getAttribute('name'))) {
          ns.setAttribute('name', `${ns.getAttribute('name')}_${idx}`);
          if (ns.getAttribute('engName')) ns.setAttribute('engName', `${ns.getAttribute('engName')}_${idx}`);
        }
        names.add(ns.getAttribute('name'));
        baseStyles.appendChild(ns);
        keys.set(k, newId);
        maps.style[old] = newId;
        appended.push([st.getAttribute('nextStyleIDRef'), ns]);
      }
      for (const [oldNext, ns] of appended) {
        if (oldNext !== null && maps.style[oldNext] !== undefined) ns.setAttribute('nextStyleIDRef', maps.style[oldNext]);
      }
      baseStyles.setAttribute('itemCnt', String(kids(baseStyles, 'style').length));
    }
    return maps;
  }

  /**
   * files: [{name, data(ArrayBuffer|Uint8Array)}]  (목록 순서 = 문서 순서)
   * 반환: {data: Uint8Array, sections, warnings}
   */
  async function mergeHwpx(files, opts = {}) {
    const log = opts.log || (() => {});
    const progress = opts.progress || (() => {});
    if (files.length < 2) throw new MergeError('병합하려면 파일이 2개 이상 필요합니다.');
    const total = files.length;
    const warnings = [];

    let base;
    try { base = await loadPkg(files[0]); }
    catch (e) { throw new MergeError(`기준 파일(${files[0].name}): ${e.message}`); }
    const baseHeader = parseXml(await base.text(base.headerHref));
    const baseHpf = base.hpf;
    const manifest = baseHpf.getElementsByTagNameNS(OPF, 'manifest')[0];
    const spine = baseHpf.getElementsByTagNameNS(OPF, 'spine')[0];

    const newFiles = new Map();
    const existing = new Set(base.names);
    let sectionCount = base.sectionHrefs.length;
    let nextSec = 1 + Math.max(...base.sectionHrefs.map((h) => parseInt(h.match(/(\d+)\.xml$/)[1], 10)));

    log(`기준 파일: ${files[0].name} (구역 ${sectionCount}개)`);
    await progress(1, total);

    for (let idx = 1; idx < total; idx++) {
      const f = files[idx];
      try {
        const pkg = await loadPkg(f);
        const incHeader = parseXml(await pkg.text(pkg.headerHref));

        const binmap = {};
        const extras = [];
        const skip = new Set([...pkg.sectionHrefs, pkg.headerHref]);
        for (const [iid, el] of pkg.items) {
          const href = el.getAttribute('href');
          if (!href) continue;
          const real = pkg.resolve(href);
          if (skip.has(real) || !pkg.names.includes(real)) continue;
          if (/(settings|content)\.(xml|hpf)$/i.test(real)) continue;
          const newReal = dirOf(real) + `m${idx}_` + baseOf(real);
          const newId = `m${idx}_${iid}`;
          binmap[iid] = newId;
          const ne = baseHpf.importNode(el, true);
          ne.setAttribute('id', newId);
          ne.setAttribute('href', dirOf(href) + `m${idx}_` + baseOf(href));
          manifest.appendChild(ne);
          extras.push({ newReal, data: await pkg.bytes(real), media: el.getAttribute('media-type') || '' });
        }

        const maps = mergeHeader(baseHeader.documentElement, incHeader.documentElement, idx, binmap);

        const fixIdRef = (root) => {
          for (const e of allEls(root)) {
            const v = e.getAttribute('idRef');
            if (v !== null && binmap[v] !== undefined) e.setAttribute('idRef', binmap[v]);
          }
        };
        for (const x of extras) {
          let data = x.data;
          if (/xml/i.test(x.media) && /\.xml$/i.test(x.newReal)) {
            const d = parseXml(new TextDecoder('utf-8').decode(data));
            remapRefs(d.documentElement, maps, binmap);
            fixIdRef(d.documentElement);
            data = serializeXml(d);
          }
          newFiles.set(x.newReal, data);
        }

        for (const secHref of pkg.sectionHrefs) {
          const sec = parseXml(await pkg.text(secHref));
          remapRefs(sec.documentElement, maps, binmap);
          fixIdRef(sec.documentElement);
          let name = `Contents/section${nextSec}.xml`;
          while (existing.has(name) || newFiles.has(name)) { nextSec++; name = `Contents/section${nextSec}.xml`; }
          newFiles.set(name, serializeXml(sec));
          const item = baseHpf.createElementNS(OPF, 'opf:item');
          item.setAttribute('id', `section${nextSec}`);
          item.setAttribute('href', name);
          item.setAttribute('media-type', 'application/xml');
          manifest.appendChild(item);
          const ref = baseHpf.createElementNS(OPF, 'opf:itemref');
          ref.setAttribute('idref', `section${nextSec}`);
          ref.setAttribute('linear', 'yes');
          spine.appendChild(ref);
          nextSec++;
          sectionCount++;
        }
        log(`  OK ${f.name} - 구역 ${pkg.sectionHrefs.length}개 추가`);
      } catch (e) {
        const msg = e instanceof MergeError ? e.message : `처리 실패 (${e.message})`;
        warnings.push(`[제외됨] ${f.name}: ${msg}`);
        log(`  X ${f.name} - ${msg}`);
      }
      await progress(idx + 1, total);
    }

    baseHeader.documentElement.setAttribute('secCnt', String(sectionCount));
    newFiles.set(base.headerHref, serializeXml(baseHeader));
    newFiles.set(base.hpfPath, serializeXml(baseHpf));

    const out = new JSZip();
    if (base.names.includes('mimetype')) out.file('mimetype', await base.bytes('mimetype'), { compression: 'STORE' });
    for (const n of base.names) {
      if (n === 'mimetype') continue;
      out.file(n, newFiles.has(n) ? newFiles.get(n) : await base.bytes(n), { compression: 'DEFLATE' });
    }
    for (const [n, data] of newFiles) {
      if (!base.names.includes(n)) out.file(n, data, { compression: 'DEFLATE' });
    }
    const data = await out.generateAsync({ type: 'uint8array' });
    log(`\n완료: ${total}개 파일, 구역 ${sectionCount}개`);
    return { data, sections: sectionCount, warnings };
  }

  /* ================= 엑셀 ================= */

  const NAME_KEYWORDS = ['성명', '이름', '명단', 'NAME'];
  const isPlaceholder = (v) => v === null || v === undefined || (typeof v === 'string' && v.trim() === '') ||
    (typeof v === 'number' && v === 0);

  function cellVal(cell) {
    if (!cell || cell.type === ExcelJS.ValueType.Merge) return null;
    let v = cell.value;
    if (v === null || v === undefined) return null;
    if (typeof v === 'object' && !(v instanceof Date)) {
      if ('result' in v) v = v.result;
      else if (v.richText) v = v.richText.map((t) => t.text).join('');
      else if (v.text !== undefined) v = v.text;
      else if (v.error) v = null;
      else if (v.formula) v = null;
    }
    if (v && typeof v === 'object' && !(v instanceof Date)) {
      if (v.richText) v = v.richText.map((t) => t.text).join('');
      else v = null;
    }
    return v === undefined ? null : v;
  }

  const colNum = (letters) => letters.split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  function mergeRanges(ws) {
    const out = [];
    for (const m of ws.model.merges || []) {
      const r = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(m);
      if (r) out.push({ left: colNum(r[1]), top: +r[2], right: colNum(r[3]), bottom: +r[4] });
    }
    return out;
  }

  function detectLayout(ws) {
    const ncol = ws.columnCount;
    const merges = mergeRanges(ws);
    const rowCells = (r) => { const a = []; for (let c = 1; c <= ncol; c++) a.push(cellVal(ws.getCell(r, c))); return a; };
    const nonEmpty = (r) => rowCells(r).filter((v) => !isPlaceholder(v)).length;

    let headerEnd = 0;
    for (const m of merges) {
      const span = m.bottom - m.top;
      if (m.top <= 8 && span >= 1 && span <= 3) headerEnd = Math.max(headerEnd, m.bottom);
    }
    if (!headerEnd) {
      for (let r = 1; r <= 8; r++) if (nonEmpty(r) >= 2) { headerEnd = r; break; }
    }
    // 키워드(성명/이름) 행이 있으면 그 행(병합이면 병합 끝 행)까지를 헤더로 본다
    let keyCol = 0;
    outer: for (let r = 1; r <= Math.max(headerEnd, 1) + 2 && r <= 10; r++) {
      for (let c = 1; c <= ncol; c++) {
        const v = cellVal(ws.getCell(r, c));
        if (v !== null && NAME_KEYWORDS.some((kw) => String(v).toUpperCase().includes(kw))) {
          keyCol = c - 1;
          const m = merges.find((x) => x.top === r && x.left === c);
          headerEnd = Math.max(headerEnd, m ? m.bottom : r);
          break outer;
        }
      }
    }
    return { headerEnd: headerEnd || 1, keyCol, ncol };
  }

  function readTable(ws, layout) {
    const rows = [];
    for (let r = layout.headerEnd + 1; r <= ws.rowCount; r++) {
      const key = cellVal(ws.getCell(r, layout.keyCol + 1));
      if (key === null || String(key).trim() === '') break;
      const vals = [];
      for (let c = 1; c <= layout.ncol; c++) vals.push(cellVal(ws.getCell(r, c)));
      rows.push({ row: r, vals });
    }
    return rows;
  }

  const headerOf = (ws, layout) => {
    const a = [];
    for (let c = 1; c <= layout.ncol; c++) { const v = cellVal(ws.getCell(layout.headerEnd, c)); a.push(v === null ? '' : String(v).trim()); }
    return JSON.stringify(a);
  };

  async function mergeExcel(files, opts = {}) {
    const log = opts.log || (() => {});
    const progress = opts.progress || (() => {});
    if (files.length < 2) throw new MergeError('취합하려면 파일이 2개 이상 필요합니다.');
    const total = files.length;
    const warnings = [];

    const wb = new ExcelJS.Workbook();
    try { await wb.xlsx.load(files[0].data); }
    catch (e) { throw new MergeError(`기준 파일(${files[0].name})을 열 수 없습니다. (.xlsx 형식만 가능)`); }
    if (!wb.worksheets.length) throw new MergeError('기준 파일에 시트가 없습니다.');

    // 시트(탭)마다 따로 기준 상태를 만든다: 첫 번째 탭은 첫 번째 탭끼리, 두 번째 탭은 두 번째 탭끼리 합친다.
    const sheets = [];
    wb.worksheets.forEach((ws, idx) => {
      const layout = detectLayout(ws);
      const baseTable = readTable(ws, layout);
      if (!baseTable.length) { log(`  - 시트 '${ws.name}': 데이터 표가 없어 건너뜁니다.`); return; }
      const keyToRow = new Map();
      for (const { row, vals } of baseTable) {
        const k = String(vals[layout.keyCol]).trim();
        if (!keyToRow.has(k)) keyToRow.set(k, row);
      }
      const tableEnd = baseTable[baseTable.length - 1].row;
      sheets.push({ idx, ws, layout, colCount: layout.ncol, refHeader: headerOf(ws, layout), keyToRow,
        hasFooter: ws.rowCount > tableEnd, added: 0 });
      log(`기준 파일 시트 ${idx + 1} '${ws.name}': 헤더 ${layout.headerEnd}행까지, 키 열 ${layout.keyCol + 1}번째, 항목 ${keyToRow.size}개`);
    });
    if (!sheets.length) throw new MergeError('기준 파일에서 합칠 수 있는 데이터 표를 찾지 못했습니다.');
    log(`기준 파일: ${files[0].name} (시트 ${wb.worksheets.length}개 중 ${sheets.length}개 대상)`);
    await progress(1, total);

    for (let i = 1; i < total; i++) {
      const f = files[i];
      try {
        const w2 = new ExcelJS.Workbook();
        try { await w2.xlsx.load(f.data); } catch (e) { throw new MergeError('엑셀(.xlsx) 형식이 아닙니다.'); }
        if (!w2.worksheets.length) throw new MergeError('시트가 없습니다.');
        let any = false;
        for (const st of sheets) {
          const s2 = w2.worksheets[st.idx];
          if (!s2) { warnings.push(`[시트 없음] ${f.name}: ${st.idx + 1}번째 시트가 없어 '${st.ws.name}' 시트는 건너뜁니다.`); continue; }
          if (s2.name !== st.ws.name) warnings.push(`[시트 이름 다름] ${f.name}: ${st.idx + 1}번째 시트 이름이 '${s2.name}'입니다 (기준: '${st.ws.name}'). 순서 기준으로 합쳤습니다.`);
          const l2 = detectLayout(s2);
          const rows = readTable(s2, l2);
          if (!rows.length) { log(`  - ${f.name} 시트 ${st.idx + 1}: 데이터 행 없음`); continue; }
          any = true;
          if (headerOf(s2, l2) !== st.refHeader) warnings.push(`[열 불일치 경고] ${f.name} / ${s2.name}: 헤더가 기준 파일과 다릅니다.`);

          let filled = 0, added = 0, conflicts = 0;
          for (const { vals } of rows) {
            const keyVal = vals[l2.keyCol];
            if (keyVal === null || String(keyVal).trim() === '') continue;
            const key = String(keyVal).trim();
            const meaningful = vals.some((v, ci) => ci !== l2.keyCol && !isPlaceholder(v));
            if (st.keyToRow.has(key)) {
              if (!meaningful) continue;
              const r = st.keyToRow.get(key);
              for (let ci = 0; ci < Math.min(st.colCount, vals.length); ci++) {
                if (ci === l2.keyCol) continue;
                const val = vals[ci];
                if (isPlaceholder(val)) continue;
                const cell = st.ws.getCell(r, ci + 1);
                if (cell.type === ExcelJS.ValueType.Merge || cell.type === ExcelJS.ValueType.Formula) continue;
                if (isPlaceholder(cell.value)) { cell.value = val; filled++; }
                else if (String(cell.value) !== String(val)) {
                  conflicts++;
                  warnings.push(`[값 충돌 경고] '${st.ws.name}' 시트 '${key}' 항목 값이 파일마다 다릅니다 (기존: ${cell.value}, ${f.name}: ${val}). 기존 값을 유지했습니다.`);
                }
              }
            } else {
              const padded = vals.slice(0, st.colCount);
              while (padded.length < st.colCount) padded.push(null);
              const row = st.ws.addRow(padded);
              st.keyToRow.set(key, row.number);
              added++;
            }
          }
          st.added += added;
          log(`  OK ${f.name} [${s2.name}] - 값 반영 ${filled}건, 신규 항목 ${added}건` + (conflicts ? `, 충돌 ${conflicts}건` : ''));
        }
        if (!any) throw new MergeError('합칠 데이터 행을 찾지 못했습니다.');
      } catch (e) {
        const msg = e instanceof MergeError ? e.message : `처리 실패 (${e.message})`;
        warnings.push(`[제외됨] ${f.name}: ${msg}`);
        log(`  X ${f.name} - ${msg}`);
      }
      await progress(i + 1, total);
    }
    for (const st of sheets) {
      if (st.added && st.hasFooter) {
        warnings.push(`[참고] '${st.ws.name}' 시트에 새로 추가된 항목 ${st.added}건은 표 맨 아래(합계 행 뒤)에 들어갔습니다. 합계 수식 범위에는 포함되지 않습니다.`);
      }
    }
    const rowsTotal = sheets.reduce((n, s) => n + s.keyToRow.size, 0);
    const buf = await wb.xlsx.writeBuffer();
    log(`\n완료: ${total}개 파일, 시트 ${sheets.length}개, 항목 ${rowsTotal}개`);
    return { data: new Uint8Array(buf), rows: rowsTotal, sheets: sheets.length, warnings };
  }

  return { mergeHwpx, mergeExcel, MergeError };
})();
if (typeof module !== 'undefined' && module.exports) module.exports = Merge;
