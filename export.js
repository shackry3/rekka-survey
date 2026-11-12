/* 公共施設 劣化調査ノート  —  export.js
   既定の様式（劣化度調査票／劣化状況写真／バリアフリー状況）を ExcelJS で再現して出力する。
   app.js の photoNoteLines・wrapText・facLabel・blobToDataURL を使う。 */
'use strict';
(function () {
  const T = { style: 'thin' }, H = { style: 'hair' };
  const FILL_LABEL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCFECF7' } };   // 見出し（水色）
  const FILL_HEAD = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD0DFE6' } };    // 表頭
  const FILL_GRAY = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
  const PX_PER_PT = 96 / 72;
  const colPx = w => Math.floor(w * 8 + 5);       // 列幅（文字数）→ ピクセル（標準フォント 游ゴシック11pt の数字幅 8px）

  const ymd = t => { const d = new Date(t); return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())); };
  const isNum = s => /^[\d.,]+$/.test(String(s || '').trim());
  const yearOf = s => { const m = String(s || '').match(/(\d{4})/); return m ? +m[1] : null; };

  function box(ws, r1, c1, r2, c2) {        // 範囲の外周に細線
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
      const cell = ws.getCell(r, c), b = Object.assign({}, cell.border);
      if (r === r1) b.top = T; if (r === r2) b.bottom = T; if (c === c1) b.left = T; if (c === c2) b.right = T;
      cell.border = b;
    }
  }
  function grid(ws, r1, c1, r2, c2) {       // 範囲内の全セルに細線
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) ws.getCell(r, c).border = { top: T, left: T, bottom: T, right: T };
  }
  function put(ws, addr, value, o = {}) {
    if (o.merge) ws.mergeCells(o.merge);
    const c = ws.getCell(addr);
    c.value = value;
    c.font = Object.assign({ name: o.font, size: 11 }, o.f || {});
    c.alignment = Object.assign({ vertical: 'middle', horizontal: o.h || 'center', wrapText: true }, o.a || {});
    if (o.fill) c.fill = o.fill;
    if (o.numFmt) c.numFmt = o.numFmt;
    return c;
  }
  function sheetName(base, used) {
    let n = base.replace(/[\[\]:*?\/\\]/g, '').slice(0, 31), i = 2;
    while (used.has(n)) n = base.slice(0, 28) + '(' + i++ + ')';
    used.add(n); return n;
  }

  /* ============================ 劣化度調査票 ============================ */
  function sheetChosa(ws, f, recs, pj) {
    const FONT = 'BIZ UDゴシック', CFG = window.SURVEY_CONFIG;
    [1.75, 9, 6.875, 6.875, 6.875, 9, 7.125, 7.125, 9, 14.625, 9.125, 1.25].forEach((w, i) => ws.getColumn(i + 1).width = w);
    const surveyT = recs.length ? recs[0].takenAt : null;
    put(ws, 'B1', '■劣化状況調査票', { font: FONT, h: 'left', f: { size: 12 }, merge: 'B1:K1' });
    [2, 3, 4].forEach(r => ws.getRow(r).height = 24);
    const by = yearOf(f.builtYear), sy = surveyT ? new Date(surveyT).getFullYear() : new Date().getFullYear();
    const L = { font: FONT, fill: FILL_LABEL };
    put(ws, 'B2', '施設番号', L); put(ws, 'C2', isNum(f.no) ? +f.no : (f.no || ''), { font: FONT, merge: 'C2:E2' });
    put(ws, 'F2', '調査日', L); put(ws, 'G2', surveyT ? ymd(surveyT) : '', { font: FONT, merge: 'G2:H2', numFmt: 'yyyy/m/d' });
    put(ws, 'I2', '記入者', L); put(ws, 'J2', pj.surveyor || '', { font: FONT, merge: 'J2:K2' });
    put(ws, 'B3', '建物名', L); put(ws, 'C3', facLabel(f), { font: FONT, merge: 'C3:E3', a: { shrinkToFit: true } });
    put(ws, 'F3', '建築年度', L); put(ws, 'G3', f.builtYear ? (isNum(f.builtYear) ? f.builtYear + '年' : f.builtYear) : '', { font: FONT, merge: 'G3:H3' });
    put(ws, 'I3', '築年数', L); put(ws, 'J3', by ? `${sy - by}年` : '', { font: FONT, merge: 'J3:K3' });
    put(ws, 'B4', '構造', L); put(ws, 'C4', f.structure || '', { font: FONT, merge: 'C4:E4' });
    put(ws, 'F4', '延床面積', L);
    put(ws, 'G4', f.floorArea ? (isNum(f.floorArea) ? Number(String(f.floorArea).replace(/,/g, '')).toLocaleString('ja-JP', { maximumFractionDigits: 2 }) + '㎡' : f.floorArea) : '', { font: FONT, merge: 'G4:H4' });
    put(ws, 'I4', '階数', L); put(ws, 'J4', f.floors ? (isNum(f.floors) ? f.floors + '階' : f.floors) : '', { font: FONT, merge: 'J4:K4' });
    grid(ws, 2, 2, 4, 11);
    ws.getRow(5).height = 8.1;
    // 総評（判定・本文は事務所で記入）
    ws.getRow(6).height = 72.75;
    put(ws, 'B6', '総評', L); put(ws, 'C6', '', { font: FONT }); put(ws, 'D6', '', { font: FONT, h: 'left', merge: 'D6:K6' });
    grid(ws, 6, 2, 6, 11);
    ws.getRow(7).height = 17.25;
    // 部位別
    ws.getRow(8).height = 52.5;
    const HD = { font: FONT, fill: FILL_HEAD };
    put(ws, 'B8', '部位', Object.assign({ merge: 'B8:C8' }, HD));
    put(ws, 'D8', '主な工事\n履歴', Object.assign({ merge: 'D8:E8' }, HD));
    put(ws, 'F8', '劣化状況', Object.assign({ merge: 'F8:J8' }, HD));
    put(ws, 'K8', '判定', HD);
    const nos = new Map(recs.map((r, i) => [r.id, i + 1]));
    CFG.sheetParts.forEach((p, i) => {
      const r = 9 + i;
      const lines = recs.filter(x => x.part === p).map(x => {
        const where = [x.floor, x.room, x.detail].filter(Boolean).join('　');
        const m = String(x.memo || '').split(/\n+/).map(t => t.trim().replace(/。$/, '')).filter(Boolean).join('、') || (x.phenomena || []).join('、');
        const sp = x.spread ? `【${x.spread}】` : '';
        return `・${where}${where && m ? '：' : ''}${m}${sp}（写真${nos.get(x.id)}）`;
      });
      const txt = lines.join('\n');
      const nLines = lines.reduce((a, l) => a + Math.max(1, wrapText(l, 23).length), 0);
      ws.getRow(r).height = Math.max(67.5, nLines * 14 + 8);
      put(ws, `B${r}`, p, { font: FONT, merge: `B${r}:C${r}` });
      put(ws, `D${r}`, (f.history || {})[p] || '', { font: FONT, merge: `D${r}:E${r}`, f: { size: 9 } });
      put(ws, `F${r}`, txt || '-', { font: FONT, h: txt ? 'left' : 'center', merge: `F${r}:J${r}` });
      put(ws, `K${r}`, '', { font: FONT });
    });
    grid(ws, 8, 2, 8 + CFG.sheetParts.length, 11);
    const rN = 9 + CFG.sheetParts.length;
    ws.getRow(rN).height = 19.5;
    put(ws, `B${rN}`, '■特記事項', { font: FONT, h: 'left', merge: `B${rN}:K${rN}` });
    ws.getRow(rN + 1).height = 72.75;
    put(ws, `B${rN + 1}`, f.note || '', { font: FONT, h: 'left', merge: `B${rN + 1}:K${rN + 1}`, a: { vertical: 'top' } });
    box(ws, rN + 1, 2, rN + 1, 11);
    ws.getRow(rN + 2).height = 75;
    put(ws, `B${rN + 2}`, CFG.legend, { font: 'ＭＳ ゴシック', h: 'left', merge: `B${rN + 2}:K${rN + 2}`, f: { size: 10 } });
    ws.pageSetup = { paperSize: 9, orientation: 'portrait', scale: 89, printArea: `A1:L${rN + 3}`,
      margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 } };
  }

  /* ============================ 劣化状況写真 ============================ */
  async function sheetPhoto(ws, wb, f, recs, tick) {
    const FONT = 'ＭＳ Ｐ明朝';
    const W = [0.75, 0.375, 54.625, 0.375, 1.625, 9, 27.625, 0.75];
    W.forEach((w, i) => ws.getColumn(i + 1).width = w);
    const BH = [3, 12.95, ...Array(14).fill(15.95), 12.95, 3];   // 1ブロック18行
    const LINES0 = 7, NL = window.SURVEY_CONFIG.photoNote.lines;
    const pages = Math.max(1, Math.ceil(recs.length / 3));
    const font = { name: FONT, size: 11 };
    for (let p = 0; p < pages; p++) {
      const R0 = 1 + p * 61;
      ws.getRow(R0).height = 14.25;
      if (p === 0) { const c = ws.getCell(R0, 3); c.value = '■劣化状況写真'; c.font = font; }
      for (let j = 0; j < 3; j++) {
        const i = p * 3 + j, t = R0 + 1 + j * 20;
        BH.forEach((h, k) => ws.getRow(t + k).height = h);
        ws.getRow(t + 18).height = 15.95; ws.getRow(t + 19).height = 15.95;
        const r = recs[i] || null;            // 写真がない欄も枠と項目名は残す
        // 枠線：写真枠（B〜D）と項目表（F〜G）
        for (let k = 0; k < 18; k++) {
          const top = k === 0 ? T : undefined, bottom = k === 17 ? T : undefined;
          ws.getCell(t + k, 2).border = { left: T, top, bottom };
          ws.getCell(t + k, 3).border = { top, bottom };
          ws.getCell(t + k, 4).border = { right: T, top, bottom };
          const inner = k >= 1 && k <= 15 ? H : bottom;
          ws.getCell(t + k, 6).border = { left: T, top, bottom: inner };
          ws.getCell(t + k, 7).border = { left: k >= 1 && k <= 16 ? H : undefined, right: T, top, bottom: inner };
        }
        const label = (k, s) => { const c = ws.getCell(t + k, 6); c.value = s; c.font = font; c.alignment = { horizontal: 'center', vertical: 'middle' }; };
        const val = (k, v, o = {}) => { const c = ws.getCell(t + k, 7); c.value = v; c.font = font; c.alignment = Object.assign({ vertical: 'middle', horizontal: 'left', indent: 0 }, o); return c; };
        label(1, '写真　№'); label(2, '撮影日'); label(4, '階'); label(5, '場所'); label(6, '部位');
        if (!r) continue;
        val(1, i + 1, { horizontal: 'center' });
        val(2, ymd(r.takenAt), { horizontal: 'center' }).numFmt = 'yyyy/m/d';
        val(4, r.floor || '');
        val(5, r.room || '', { shrinkToFit: true });
        val(6, r.detail || (r.part === '全景' ? '' : r.part || ''), { shrinkToFit: true });
        const pn = photoNoteLines(r), nl = pn.lines, C = pn.cap;
        const lines = nl.length > C ? [...nl.slice(0, C - 1), nl.slice(C - 1).join('')] : nl;
        lines.forEach((s, k) => val(LINES0 + k, s, { shrinkToFit: k === C - 1 && nl.length > C }));
        if (pn.note) val(LINES0 + NL - 1, pn.note, { shrinkToFit: true });   // 範囲の注釈は所見欄の最後の行
        // 写真：C列の幅いっぱいに配置（縦長などで枠の高さを超える場合は高さに合わせて中央配置）
        // 写真の大きさは cm で固定し、縦横比は必ず元のまま（config.js の photoBox で調整）
        const PB = window.SURVEY_CONFIG.photoBox, cm2px = c => c / 2.54 * 96;
        const cw = cm2px(PB.columnCm), maxW = cm2px(PB.widthCm);
        const blockPx = BH.reduce((a, b) => a + b, 0) * PX_PER_PT;  // 枠の高さ（px）
        if (tick) tick();
        let buf = null;
        try { buf = r.broken ? null : (r.photo instanceof Blob ? await r.photo.arrayBuffer() : r.photo); } catch (e) { buf = null; }
        if (!buf || !buf.byteLength) {        // 写真データが読めない場合は欄に注記して続行
          r.broken = true;
          const c = ws.getCell(t + 8, 3); c.value = '（写真データなし）'; c.font = font; c.alignment = { horizontal: 'center', vertical: 'middle' };
          continue;
        }
        const imgId = wb.addImage({ buffer: new Uint8Array(buf), extension: 'jpeg' });
        let w = maxW, h = maxW * r.h / r.w;
        if (h > blockPx - 6) { h = blockPx - 6; w = h * r.w / r.h; }
        let yPt = (BH.reduce((a, b) => a + b, 0) - h / PX_PER_PT) / 2, k2 = 0;
        while (k2 < 17 && yPt >= BH[k2]) { yPt -= BH[k2]; k2++; }
        ws.addImage(imgId, {
          tl: { nativeCol: 2, nativeColOff: Math.max(0, Math.round((cw - w) / 2 * 9525)), nativeRow: t - 1 + k2, nativeRowOff: Math.round(yPt * 12700) },
          ext: { width: w, height: h }, editAs: 'oneCell'
        });
      }
      if (p < pages - 1) ws.getRow(R0 + 60).addPageBreak();
    }
    ws.pageSetup = { paperSize: 9, orientation: 'portrait', scale: 88, printArea: `A1:H${pages * 61}`,
      margins: { left: 0.394, right: 0.709, top: 0.787, bottom: 0.394, header: 0.315, footer: 0.315 } };
  }

  /* ============================ バリアフリー状況 ============================ */
  function sheetBF(ws, f, pj) {
    const FONT = 'ＭＳ ゴシック', items = window.SURVEY_CONFIG.barrierFree, bf = f.bf || {};
    ws.getColumn(1).width = 1.75; ws.getColumn(2).width = 15.125;
    for (let c = 3; c <= 13; c++) ws.getColumn(c).width = 11.5;
    ws.getRow(1).height = 13.5; ws.getRow(2).height = 30;
    const UF = { name: 'BIZ UDゴシック', size: 14 };
    const c2 = ws.getCell('B2'); c2.value = `${pj.city || ''}公共施設　バリアフリー調査`; c2.font = UF; c2.alignment = { vertical: 'middle' };
    put(ws, 'G2', '施設名', { font: 'BIZ UDゴシック', f: { size: 14 } });
    put(ws, 'H2', facLabel(f), { font: 'BIZ UDゴシック', f: { size: 12 }, h: 'left', merge: 'H2:J2', a: { shrinkToFit: true, wrapText: false } });
    put(ws, 'K2', '階数', { font: 'BIZ UDゴシック', f: { size: 14 } });
    put(ws, 'L2', f.floors ? (isNum(f.floors) ? f.floors + '階' : f.floors) : '', { font: 'BIZ UDゴシック', f: { size: 12 }, merge: 'L2:M2' });
    for (let c = 7; c <= 13; c++) ws.getCell(2, c).border = { bottom: T };
    ws.getRow(4).height = 54; ws.getRow(5).height = 40.5; ws.getRow(6).height = 53.25;
    // 表頭（グループ）
    let c = 2;
    const groups = [];
    items.forEach(it => { const g = groups[groups.length - 1]; if (g && g.name === it.group) g.n++; else groups.push({ name: it.group, n: 1 }); });
    groups.forEach(g => {
      const a = ws.getCell(4, c).address, b = ws.getCell(4, c + g.n - 1).address;
      put(ws, a, g.name, { font: FONT, f: { size: 12, bold: true }, fill: FILL_GRAY, merge: g.n > 1 ? `${a}:${b}` : undefined });
      c += g.n;
    });
    items.forEach((it, i) => {
      put(ws, ws.getCell(5, 2 + i).address, it.label, { font: FONT, f: { size: 12 } });
      let v = bf[it.key] || '';
      if (it.type === 'yn+text' && bf[it.key + '_t']) v = (v ? v + '\n' : '') + bf[it.key + '_t'];
      if (it.type === 'num' && v) v = isNum(v) ? `${v}台` : v;
      put(ws, ws.getCell(6, 2 + i).address, v, { font: FONT, f: { size: 12 } });
    });
    grid(ws, 4, 2, 6, 1 + items.length);
    ws.getRow(7).height = 35.25;
    const c7 = ws.getCell('B7'); c7.value = '■その他気づいた点記入'; c7.font = { name: FONT, size: 12, bold: true }; c7.alignment = { vertical: 'middle' };
    ws.getRow(8).height = 226.5;
    const last = ws.getCell(8, 1 + items.length).address;
    put(ws, 'B8', bf.notes || '', { font: FONT, f: { size: 12 }, h: 'left', merge: `B8:${last}`, a: { vertical: 'top' } });
    box(ws, 8, 2, 8, 1 + items.length);
    ws.pageSetup = { paperSize: 9, orientation: 'landscape', scale: 84, printArea: `A1:${ws.getCell(9, 2 + items.length).address}`,
      margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 } };
  }

  /* ============================ 調査記録一覧（全件出力時） ============================ */
  function sheetList(ws, items) {
    const FONT = 'BIZ UDPゴシック';
    ws.columns = [
      { header: '施設番号', width: 9 }, { header: '施設名', width: 24 }, { header: '写真№', width: 7 },
      { header: '部位（区分）', width: 13 }, { header: '階', width: 8 }, { header: '場所', width: 16 }, { header: '部位（詳細）', width: 22 },
      { header: '劣化現象', width: 22 }, { header: '広がり', width: 10 }, { header: '所見メモ', width: 56 }, { header: '撮影日時', width: 16 }
    ];
    items.forEach(({ f, recs }) => recs.forEach((r, i) => ws.addRow([
      f.no || '', facLabel(f), i + 1, r.part || '', r.floor || '', r.room || '', r.detail || '',
      (r.phenomena || []).join('、'), r.spread || '', r.memo || '', fmtDate(r.takenAt)])));
    ws.eachRow((row, i) => row.eachCell({ includeEmpty: true }, c => {
      c.font = { name: FONT, size: 10, bold: i === 1 };
      c.alignment = i === 1 ? { vertical: 'middle', horizontal: 'center' } : { vertical: 'top', wrapText: true };
      c.border = { top: T, left: T, bottom: T, right: T };
      if (i === 1) c.fill = FILL_HEAD;
    }));
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    ws.autoFilter = { from: 'A1', to: 'K1' };
    ws.pageSetup = { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: '1:1' };
  }

  window.buildSurveyWorkbook = async function (pj, items, onProgress) {
    let done = 0; const tick = () => { done++; if (onProgress && done % 5 === 0) onProgress(done); };
    const wb = new ExcelJS.Workbook();
    wb.creator = pj.surveyor || ''; wb.created = new Date();
    const used = new Set(), single = items.length === 1;
    if (!single) sheetList(wb.addWorksheet(sheetName('調査記録一覧', used)), items);
    for (const { f, recs } of items) {
      const pre = single ? '' : `${f.no ? f.no + '_' : ''}${facLabel(f).replace(/\s+/g, '')}`.slice(0, 22) + '_';
      sheetChosa(wb.addWorksheet(sheetName(pre + (single ? '劣化度調査票' : '調査票'), used)), f, recs, pj);
      await sheetPhoto(wb.addWorksheet(sheetName(pre + (single ? '劣化状況写真' : '写真'), used)), wb, f, recs, tick);
      sheetBF(wb.addWorksheet(sheetName(pre + (single ? 'バリアフリー状況' : 'BF'), used)), f, pj);
    }
    // JPEGは圧縮済みなので再圧縮しない（STORE）。iPhoneでの出力時間を大きく短縮する
    const buf = await wb.xlsx.writeBuffer({ zip: { compression: 'STORE' } });
    return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  };
})();
