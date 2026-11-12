/* 公共施設 劣化調査ノート  —  app.js（画面・保存・バックアップ・CSV取込）
   Excel出力は export.js。データは端末内（IndexedDB）に保存する。 */
'use strict';
const CFG = window.SURVEY_CONFIG;
const $ = (s, el = document) => el.querySelector(s);
const view = $('#view');

/* ------------------------------------------------------------------ DB */
const DB_NAME = 'rekka-survey', DB_VER = 1;
let _db;
function db() {
  if (_db) return Promise.resolve(_db);
  return new Promise((res, rej) => {
    const rq = indexedDB.open(DB_NAME, DB_VER);
    rq.onupgradeneeded = () => {
      const d = rq.result;
      d.createObjectStore('kv');
      d.createObjectStore('facilities', { keyPath: 'id' });
      const r = d.createObjectStore('records', { keyPath: 'id' });
      r.createIndex('facilityId', 'facilityId');
    };
    rq.onsuccess = () => { _db = rq.result; res(_db); };
    rq.onerror = () => rej(rq.error);
  });
}
async function tx(store, mode, fn) {
  const d = await db();
  return new Promise((res, rej) => {
    const t = d.transaction(store, mode);
    let out;
    Promise.resolve(fn(t.objectStore(store))).then(v => { out = v; });
    t.oncomplete = () => res(out);
    t.onerror = () => rej(t.error);
    t.onabort = () => rej(t.error || new Error('保存に失敗しました（容量不足の可能性）'));
  });
}
const req = r => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
const DB = {
  get: (st, k) => tx(st, 'readonly', s => req(s.get(k))),
  all: (st) => tx(st, 'readonly', s => req(s.getAll())),
  put: (st, v, k) => tx(st, 'readwrite', s => req(k === undefined ? s.put(v) : s.put(v, k))),
  del: (st, k) => tx(st, 'readwrite', s => req(s.delete(k))),
  byFacility: (fid) => tx('records', 'readonly', s => req(s.index('facilityId').getAll(fid))),
  clear: async () => { for (const st of ['kv', 'facilities', 'records']) await tx(st, 'readwrite', s => req(s.clear())); }
};

/* ------------------------------------------------------------- helpers */
const uid = p => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pad = (n, w = 2) => String(n).padStart(w, '0');
function fmtDate(t) {
  const d = new Date(t);
  return `${d.getFullYear()}/${pad(d.getMonth() + 1)}/${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function toast(msg, ms = 2200) {
  const t = $('#toast'); t.textContent = msg; t.classList.add('show');
  clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.remove('show'), ms);
}
const urls = new Set();
function blobURL(b) { const u = URL.createObjectURL(asBlob(b)); urls.add(u); return u; }
function releaseURLs() { urls.forEach(u => URL.revokeObjectURL(u)); urls.clear(); }
const debounce = (fn, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };
const blobToDataURL = b => new Promise(r => { const f = new FileReader(); f.onload = () => r(f.result); f.readAsDataURL(b); });
const dataURLToBlob = async u => (await fetch(u)).blob();
/* 写真は ArrayBuffer で保存する（iPhoneのSafariでは、IndexedDBに保存したBlobが後で読めなくなる不具合があるため） */
const asBlob = x => x instanceof Blob ? x : (x ? new Blob([x], { type: 'image/jpeg' }) : new Blob([], { type: 'image/jpeg' }));
/* 起動時：古い形式（Blob）の写真を ArrayBuffer に移し替え、読めない写真に印を付ける */
async function migratePhotos() {
  let broken = 0;
  for (const r of await DB.all('records')) {
    let changed = false;
    for (const k of ['photo', 'thumb']) {
      if (r[k] instanceof Blob) {
        try { r[k] = await r[k].arrayBuffer(); }
        catch (e) { r[k] = null; }
        changed = true;
      }
    }
    for (const x of r.extras || []) for (const k of ['photo', 'thumb']) {
      if (x[k] instanceof Blob) { try { x[k] = await x[k].arrayBuffer(); } catch (e) { x[k] = null; } changed = true; }
    }
    if (r.extras && r.extras.some(x => !(x.photo && x.photo.byteLength))) { r.extras = r.extras.filter(x => x.photo && x.photo.byteLength); changed = true; }
    const isBroken = !(r.photo && r.photo.byteLength);
    if (isBroken !== !!r.broken) { r.broken = isBroken; changed = true; }
    if (isBroken) broken++;
    if (changed) await DB.put('records', r);
  }
  return broken;
}
function sortFacilities(list) {
  return list.sort((a, b) => String(a.no || '').localeCompare(String(b.no || ''), 'ja', { numeric: true })
    || String(a.building || '').localeCompare(String(b.building || ''), 'ja') || a.createdAt - b.createdAt);
}
function sortRecords(list) { return list.sort((a, b) => a.takenAt - b.takenAt || a.id.localeCompare(b.id)); }
const facLabel = f => f.name + (f.building ? '　' + f.building : '');
/* 写真番号：Excelに出力する記録だけに 1, 2, 3… と振る */
function photoNumbers(recs) { const m = new Map(); let n = 0; recs.forEach(r => { if (!r.exclude) m.set(r.id, ++n); }); return m; }
async function getProject() { return (await DB.get('kv', 'project')) || { city: '', title: '', surveyor: '' }; }

/* 写真台帳の所見欄：メモを「1行○文字×○行」に折り返す（export.js と共用） */
function charW(ch) { return /[\x20-\x7e｡-ﾟ]/.test(ch) ? 0.5 : 1; }
function wrapText(s, n) {
  const out = []; let line = '', w = 0;
  for (const ch of s) {
    const cw = charW(ch);
    if (w + cw > n && line) { out.push(line); line = ''; w = 0; }
    line += ch; w += cw;
  }
  if (line) out.push(line);
  return out;
}
function photoNoteLines(r) {
  const { charsPerLine: N, lines: L } = CFG.photoNote;
  let paras = String(r.memo || '').split(/\n+/).map(s => s.trim()).filter(Boolean);
  if (!paras.length && (r.phenomena || []).length) paras = [r.phenomena.join('、')];
  if (!paras.length && r.part === '全景') paras = ['全景'];
  const wrapped = paras.map(p => wrapText(p, N));
  let lines = [];
  wrapped.forEach((w, i) => { if (i) lines.push(''); lines.push(...w); });
  if (lines.length > L) lines = wrapped.flat();          // 入りきらなければ段落間の空行を詰める
  return { lines, over: Math.max(0, lines.length - L) };
}

/* ------------------------------------------------------- image handling */
async function loadBitmap(file) {
  if ('createImageBitmap' in window) {
    try { return await createImageBitmap(file, { imageOrientation: 'from-image' }); } catch (e) { /* fallthrough */ }
  }
  return new Promise((res, rej) => {
    const img = new Image(); const u = URL.createObjectURL(file);
    img.onload = () => { res(img); URL.revokeObjectURL(u); }; img.onerror = rej; img.src = u;
  });
}
function drawScaled(src, max, q) {
  const w0 = src.width, h0 = src.height, k = Math.min(1, max / Math.max(w0, h0));
  const c = document.createElement('canvas'); c.width = Math.round(w0 * k); c.height = Math.round(h0 * k);
  c.getContext('2d').drawImage(src, 0, 0, c.width, c.height);
  return new Promise(r => c.toBlob(b => r({ blob: b, w: c.width, h: c.height }), 'image/jpeg', q));
}
async function processImage(file) {
  const bmp = await loadBitmap(file);
  const main = await drawScaled(bmp, CFG.image.maxPx, CFG.image.quality);
  const th = await drawScaled(bmp, 360, 0.7);
  if (bmp.close) bmp.close();
  return { photo: await main.blob.arrayBuffer(), thumb: await th.blob.arrayBuffer(), w: main.w, h: main.h };
}

/* 写真ライブラリから取り込んだJPEGの撮影日時（EXIF DateTimeOriginal）を読む */
async function exifDate(file) {
  try {
    const v = new DataView(await file.slice(0, 196608).arrayBuffer());
    if (v.getUint16(0) !== 0xFFD8) return null;
    let o = 2;
    while (o < v.byteLength - 4) {
      const mk = v.getUint16(o), len = v.getUint16(o + 2);
      if (mk === 0xFFE1 && v.getUint32(o + 4) === 0x45786966) {
        const t = o + 10, le = v.getUint16(t) === 0x4949;
        const u16 = x => v.getUint16(t + x, le), u32 = x => v.getUint32(t + x, le);
        const find = (ifd, tag) => { const n = u16(ifd); for (let i = 0; i < n; i++) { const e = ifd + 2 + i * 12; if (u16(e) === tag) return e; } return null; };
        const ex = find(u32(4), 0x8769); if (ex === null) return null;
        const d = find(u32(ex + 8), 0x9003); if (d === null) return null;
        const off = u32(d + 8); let str = '';
        for (let i = 0; i < 19; i++) str += String.fromCharCode(v.getUint8(t + off + i));
        const m = str.match(/(\d{4}):(\d{2}):(\d{2}) (\d{2}):(\d{2}):(\d{2})/);
        return m ? new Date(+m[1], m[2] - 1, +m[3], +m[4], +m[5], +m[6]).getTime() : null;
      }
      if ((mk & 0xFF00) !== 0xFF00) break;
      o += 2 + len;
    }
  } catch (e) { }
  return null;
}

/* 本体（写真アプリ）への保存：Webアプリからは自動保存できないため、共有シート→「画像を保存」を使う */
const originals = new Map();      // 撮影直後の元画像（このセッション中だけ保持）
async function getSettings() { return Object.assign({ saveToDevice: true }, await DB.get('kv', 'settings')); }
const isMobile = () => /iPhone|iPad|Android/i.test(navigator.userAgent) || (navigator.maxTouchPoints > 1 && /Macintosh/.test(navigator.userAgent));
async function shareImages(files) {
  if (navigator.canShare && navigator.canShare({ files })) {
    try { await navigator.share({ files }); return true; }
    catch (e) {
      if (e.name === 'AbortError') return false;
      console.warn(e); toast('共有画面を開けませんでした：' + e.message, 5000); return false;
    }
  }
  if (isMobile()) { toast('この端末では共有画面を開けません', 4000); return false; }
  for (const f of files) {                       // 共有できない環境（PCなど）はダウンロード
    const a = document.createElement('a'); a.href = URL.createObjectURL(f); a.download = f.name;
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 30000);
  }
  return true;
}
function photoFileName(f, n, ext = 'jpg') { return `${f.no ? f.no + '_' : ''}${facLabel(f).replace(/[\\/:*?"<>|\s]+/g, '')}_写真${n}.${ext}`; }

/* ------------------------------------------------------------- router */
const backBtn = $('#backBtn'), titleEl = $('#title'), fab = $('#fab'), fabInner = $('#fabInner');
function go(hash) { location.hash = hash; }
window.go = go;
window.addEventListener('hashchange', render);
backBtn.onclick = () => go(backBtn.dataset.to || '#/');
$('#menuBtn').onclick = () => go('#/menu');
function setChrome({ title, back, fabHTML }) {
  titleEl.textContent = title; document.title = title;
  backBtn.classList.toggle('hidden', !back); backBtn.dataset.to = back || '';
  fab.classList.toggle('hidden', !fabHTML); fabInner.innerHTML = fabHTML || '';
}
async function render() {
  releaseURLs(); window.scrollTo(0, 0);
  const [, a, b] = (location.hash.replace(/^#/, '') || '/').split('/');
  const id = b && decodeURIComponent(b);
  try {
    if (a === 'f') return await vFacility(id);
    if (a === 'r') return await vRecord(id);
    if (a === 'bf') return await vBarrierFree(id);
    if (a === 'menu') return await vMenu();
    if (a === 'newf') return await vFacilityEdit(null);
    if (a === 'ef') return await vFacilityEdit(id);
    return await vHome();
  } catch (e) { console.error(e); view.innerHTML = `<div class="card">エラー：${esc(e.message)}</div>`; }
}
const enc = encodeURIComponent;

/* --------------------------------------------------------------- home */
async function vHome() {
  const [pj, facs, recs] = [await getProject(), sortFacilities(await DB.all('facilities')), await DB.all('records')];
  const cnt = {}; recs.forEach(r => cnt[r.facilityId] = (cnt[r.facilityId] || 0) + 1);
  setChrome({ title: pj.title || '劣化調査ノート', fabHTML: `<button class="btn pri" onclick="go('#/newf')">＋ 建物を追加</button>` });
  view.innerHTML = `
    <div class="card">
      <div style="font-size:13px;color:var(--sub)">${esc(pj.city) || '（自治体名 未設定）'}</div>
      <div style="font-size:18px;font-weight:700;margin-top:2px">${esc(pj.title) || '調査名 未設定'}</div>
      <div class="muted" style="margin-top:6px">建物 ${facs.length} 件 ／ 写真 ${recs.length} 枚</div>
      ${!pj.title ? `<button class="btn ghost" style="margin-top:10px" onclick="go('#/menu')">調査情報を入力</button>` : ''}
    </div>
    <div class="card">
      <h2>調査対象（調査票1枚＝1件）</h2>
      ${facs.length ? facs.map(f => `
        <div class="list-item" onclick="go('#/f/${enc(f.id)}')">
          <span class="no">${esc(f.no || '—')}</span>
          <span class="name">${esc(f.name)}${f.building ? `<span style="font-weight:400;color:var(--sub)">　${esc(f.building)}</span>` : ''}</span>
          <span class="cnt">${cnt[f.id] || 0} 枚</span>
        </div>`).join('') :
      `<div class="empty">調査対象がまだありません。<br>下の「建物を追加」から登録するか、<br>「出力・設定」から施設リスト（CSV）を取り込んでください。</div>`}
    </div>`;
}

/* --------------------------------------------------- facility edit */
async function vFacilityEdit(id) {
  const f = id ? await DB.get('facilities', id) : { no: '', name: '', building: '', history: {} };
  const back = id ? `#/f/${enc(id)}` : '#/';
  setChrome({ title: id ? '建物情報の編集' : '建物を追加', back });
  const inp = (key, label, ph = '', hint = '') => `<label class="f">${label}</label>
    <input type="text" data-k="${key}" value="${esc(f[key])}" placeholder="${esc(ph)}">${hint ? `<div class="hint">${hint}</div>` : ''}`;
  view.innerHTML = `
    <div class="card">
      <h2>基本情報（劣化度調査票の上段）</h2>
      ${inp('no', '施設番号', '例：12')}
      ${inp('name', '施設名 *', '例：図書館・歴史資料館')}
      ${inp('building', '棟名（1施設に複数棟あり、棟ごとに調査票を作る場合）', '例：体育館')}
      <div class="row" style="gap:10px">
        <div style="flex:1;min-width:130px">${inp('builtYear', '建築年度', '例：2000')}</div>
        <div style="flex:1;min-width:130px">${inp('structure', '構造', '例：RC／S造')}</div>
      </div>
      <div class="row" style="gap:10px">
        <div style="flex:1;min-width:130px">${inp('floorArea', '延床面積（㎡）', '例：3899.27')}</div>
        <div style="flex:1;min-width:130px">${inp('floors', '階数', '例：2')}</div>
      </div>
      ${inp('address', '所在地')}
      <div class="hint">築年数は、建築年度と調査日から自動で計算します。</div>
    </div>
    <div class="card">
      <h2>主な工事履歴（調査票の部位ごと・任意）</h2>
      ${CFG.sheetParts.map(p => `<label class="f">${esc(p)}</label>
        <input type="text" data-h="${esc(p)}" value="${esc((f.history || {})[p])}" placeholder="例：H28 屋上防水改修">`).join('')}
    </div>
    <div class="card">
      <h2>特記事項</h2>
      <textarea data-k="note" placeholder="調査票の「特記事項」欄に入ります">${esc(f.note)}</textarea>
    </div>
    <div class="row"><button class="btn pri" id="saveF">保存</button></div>
    ${id ? `<div style="margin-top:28px"><button class="btn danger wide" id="delF">この建物と写真をすべて削除</button></div>` : ''}`;
  $('#saveF').onclick = async () => {
    const rec = Object.assign({}, f);
    view.querySelectorAll('[data-k]').forEach(el => rec[el.dataset.k] = el.value.trim());
    rec.history = {}; view.querySelectorAll('[data-h]').forEach(el => { if (el.value.trim()) rec.history[el.dataset.h] = el.value.trim(); });
    if (!rec.name) return toast('施設名を入力してください');
    rec.updatedAt = Date.now();
    if (!id) {
      rec.createdAt = Date.now();
      const cand = rec.no ? `f_${rec.no}_${rec.building || ''}` : '';
      rec.id = cand && !(await DB.get('facilities', cand)) ? cand : uid('f_');
    }
    await DB.put('facilities', rec);
    go(`#/f/${enc(rec.id)}`);
  };
  if (id) $('#delF').onclick = async () => {
    const rs = await DB.byFacility(id);
    if (!confirm(`「${facLabel(f)}」と写真 ${rs.length} 枚を削除します。元に戻せません。よろしいですか？`)) return;
    for (const r of rs) await DB.del('records', r.id);
    await DB.del('facilities', id); toast('削除しました'); go('#/');
  };
}

/* ------------------------------------------------------- facility */
let currentFacility = null;
async function vFacility(id) {
  const f = await DB.get('facilities', id);
  if (!f) return go('#/');
  currentFacility = id;
  const recs = sortRecords(await DB.byFacility(id));
  const nums = photoNumbers(recs);
  const bfDone = Object.values(f.bf || {}).filter(v => v).length;
  setChrome({
    title: facLabel(f), back: '#/',
    fabHTML: `<button class="btn ghost" style="flex:0 0 auto;background:var(--card)" onclick="pickLibrary()">🖼 写真から</button>
              <button class="btn acc" onclick="pickCamera()">📷 撮影して記録</button>`
  });
  const info = [f.builtYear && `${f.builtYear}${/^\d+$/.test(f.builtYear) ? '年度' : ''}`, f.structure, f.floorArea && `${f.floorArea}${/^[\d.,]+$/.test(f.floorArea) ? '㎡' : ''}`, f.floors && `${f.floors}${/^\d+$/.test(f.floors) ? '階' : ''}`].filter(Boolean).join('／');
  // 部位ごとの枚数
  const byPart = {}; recs.forEach(r => byPart[r.part || '未設定'] = (byPart[r.part || '未設定'] || 0) + 1);
  view.innerHTML = `
    <div class="card">
      <div style="display:flex;align-items:center;gap:8px">
        <span class="no" style="font-family:ui-monospace,Menlo,monospace;background:var(--chip);border-radius:6px;padding:3px 7px">${esc(f.no || '—')}</span>
        <b style="flex:1;font-size:17px">${esc(facLabel(f))}</b>
      </div>
      <div class="muted" style="margin-top:6px">${info ? esc(info) : '建築年度・構造などは「建物情報」から入力'}</div>
      <div style="margin-top:8px">${CFG.parts.map(p => `<span class="badge" style="${byPart[p] ? '' : 'opacity:.45'}">${esc(p)} ${byPart[p] || 0}</span>`).join('')}
        ${byPart['未設定'] ? `<span class="badge" style="background:#f6dcd6">部位未設定 ${byPart['未設定']}</span>` : ''}</div>
      <div class="row" style="margin-top:12px">
        <button class="btn ghost sm" onclick="go('#/ef/${enc(id)}')">建物情報</button>
        <button class="btn ghost sm" onclick="go('#/bf/${enc(id)}')">バリアフリー${bfDone ? ' ✓' : ''}</button>
        <button class="btn ghost sm" id="xlsOne">Excel出力</button>
      </div>
      ${recs.length ? `<button class="btn ghost sm wide" id="saveAll" style="margin-top:8px">📥 この建物の写真をまとめて本体に保存（${recs.reduce((a, r) => a + 1 + (r.extras || []).length, 0)}枚）</button>` : ''}
    </div>
    ${recs.length ? `<div class="grid">${recs.map((r, i) => {
      const nl = photoNoteLines(r);
      return `<div class="thumb" onclick="go('#/r/${enc(r.id)}')" style="${r.exclude ? 'opacity:.5' : ''}">
        ${r.broken ? `<div style="aspect-ratio:4/3;display:flex;align-items:center;justify-content:center;background:#f6dcd6;color:var(--danger);font-size:13px;text-align:center;padding:8px">写真データが<br>読み込めません<br>（開いて差し替え）</div>`
          : `<img src="${blobURL(r.thumb || r.photo)}" alt="">`}
        <div class="cap"><b>${r.exclude ? '出力しない' : '写真' + nums.get(r.id)}</b> ${esc([r.floor, r.room].filter(Boolean).join(' '))}
          ${(r.extras || []).length ? `<span class="badge">📷×${r.extras.length + 1}</span>` : ''}<br>
          ${r.part ? esc(r.part) : '<span style="color:var(--danger)">部位未入力</span>'}${r.detail ? '／' + esc(r.detail) : ''}
          ${!nl.lines.length ? '<br><span style="color:var(--danger)">所見未入力</span>' : ''}
          ${nl.over ? '<br><span style="color:var(--danger)">所見が長すぎ</span>' : ''}
        </div>
      </div>`;
    }).join('')}</div>`
      : `<div class="empty">まだ写真がありません。<br>下の「撮影して記録」で撮影してください。<br>1枚目は建物の全景がおすすめです。</div>`}`;
  $('#xlsOne').onclick = () => exportExcel([id]);
  if (recs.length) $('#saveAll').onclick = () => saveAllPhotos(f, recs);
}

/* まとめて本体に保存：iPhoneの共有画面は一度に多くの写真を渡すと開かないため、20枚ずつ保存する */
function saveAllPhotos(f, recs) {
  const B = 20, nums = photoNumbers(recs), list = [];
  recs.forEach((r, i) => {
    const base = r.exclude ? `対象外${i + 1}` : nums.get(r.id);
    if (!r.broken) list.push({ photo: r.photo, n: base });
    (r.extras || []).forEach((x, j) => list.push({ photo: x.photo, n: `${base}-${j + 2}` }));
  });
  if (!list.length) return toast('保存できる写真がありません');
  const batches = [];
  for (let i = 0; i < list.length; i += B) batches.push(list.slice(i, i + B));
  let k = 0;
  const dlg = $('#dlg');
  const show = () => {
    const b = batches[k];
    $('#dlgTitle').textContent = `写真を本体に保存（${k + 1} / ${batches.length} 回目）`;
    $('#dlgName').textContent = `写真${b[0].n} 〜 写真${b[b.length - 1].n}（${b.length}枚）`;
    $('#dlgHint').textContent = 'ボタンを押すと共有画面が開きます。その中の「○枚の画像を保存」を押してください。';
    $('#dlgShare').textContent = '📥 この分を保存';
  };
  $('#dlgShare').onclick = async () => {
    const files = batches[k].map(x => new File([x.photo], photoFileName(f, x.n), { type: 'image/jpeg' }));
    const ok = await shareImages(files);
    if (!ok) return;
    k++;
    if (k >= batches.length) { dlg.close(); toast('すべて保存しました'); } else show();
  };
  $('#dlgClose').onclick = () => dlg.close();
  show(); dlg.showModal();
}

/* 同じ箇所の写真を追加（台帳に使うのは1枚。記録画面で選び直せる） */
async function addSamePlace(rid, files, fromCamera) {
  captureTarget = null;
  const r = await DB.get('records', rid); if (!r) return;
  const st = await getSettings();
  toast('写真を保存中…', 60000);
  try {
    r.extras = r.extras || [];
    for (const file of files) {
      const img = await processImage(file);
      r.extras.push({ photo: img.photo, thumb: img.thumb, w: img.w, h: img.h,
        takenAt: fromCamera ? Date.now() : ((await exifDate(file)) || file.lastModified || Date.now()) });
      if (fromCamera && st.saveToDevice) originals.set(rid, file);
    }
    r.updatedAt = Date.now(); await DB.put('records', r);
    toast(`同じ箇所の写真を追加しました（計 ${r.extras.length + 1} 枚）。台帳に使う写真は★で選べます。`, 3500);
    render();
  } catch (e) { console.error(e); toast('写真の保存に失敗しました：' + e.message, 5000); }
}

/* ------------------------------------------------------- capture */
let captureTarget = null;          // 「同じ箇所をもう1枚」のときの追加先レコード
function pickCamera(target) { captureTarget = typeof target === 'string' ? target : null; $('#camInput').value = ''; $('#camInput').click(); }
function pickLibrary(target) { captureTarget = typeof target === 'string' ? target : null; $('#libInput').value = ''; $('#libInput').click(); }
window.pickCamera = pickCamera; window.pickLibrary = pickLibrary;
$('#camInput').onchange = e => addPhotos([...e.target.files], true);
$('#libInput').onchange = e => addPhotos([...e.target.files], false);

async function addPhotos(files, fromCamera) {
  if (!files.length || !currentFacility) return;
  if (captureTarget) return addSamePlace(captureTarget, files, fromCamera);
  const fid = currentFacility;
  const existing = await DB.byFacility(fid);
  const last = (await DB.get('kv', 'last_' + fid)) || {};
  const st = await getSettings();
  toast(files.length > 1 ? `${files.length} 枚を取り込み中…` : '写真を保存中…', 60000);
  let firstId = null;
  try {
    for (const file of files) {
      const img = await processImage(file);
      const first = !existing.length && !firstId;
      const rec = {
        id: uid('r_'), facilityId: fid,
        room: first ? '' : (last.room || ''), extras: [],
        floor: first ? '' : (last.floor || ''), part: first ? '全景' : (fromCamera ? (last.part || '') : ''),
        detail: '', phenomena: [], memo: '',
        takenAt: fromCamera ? Date.now() : ((await exifDate(file)) || file.lastModified || Date.now()),
        createdAt: Date.now(), updatedAt: Date.now(),
        photo: img.photo, thumb: img.thumb, w: img.w, h: img.h
      };
      await DB.put('records', rec);
      if (fromCamera && st.saveToDevice) originals.set(rec.id, file);
      firstId = firstId || rec.id;
    }
    requestPersist();
    toast(files.length > 1 ? `${files.length} 枚を追加しました。各写真を開いて入力してください。` : '保存しました');
    if (files.length === 1) go(`#/r/${enc(firstId)}`); else render();
  } catch (e) {
    console.error(e); toast('写真の保存に失敗しました：' + e.message, 5000);
  }
}
async function requestPersist() {
  try { if (navigator.storage && navigator.storage.persist && !(await navigator.storage.persisted())) await navigator.storage.persist(); } catch (e) { }
}

/* ------------------------------------------------------- record */
async function vRecord(id) {
  const r = await DB.get('records', id);
  if (!r) return go('#/');
  const f = await DB.get('facilities', r.facilityId);
  const recs = sortRecords(await DB.byFacility(r.facilityId));
  const idx = recs.findIndex(x => x.id === id);
  const nums = photoNumbers(recs);
  currentFacility = f.id;
  const back = `#/f/${enc(f.id)}`;
  const label = r.exclude ? '台帳対象外' : `写真${nums.get(id)}`;
  setChrome({
    title: `${label}　${facLabel(f)}`, back,
    fabHTML: `<button class="btn ghost" style="background:var(--card)" onclick="go('${back}')">一覧へ</button>
              <button class="btn acc" onclick="pickCamera()">📷 次の箇所を撮影</button>`
  });
  r.extras = r.extras || [];
  const chip = (val, on) => `<button class="chip ${on ? 'on' : ''}" data-v="${esc(val)}">${esc(val)}</button>`;
  const { charsPerLine: N, lines: L } = CFG.photoNote;
  const orig = originals.get(id);
  // この建物で使った場所（部屋名）の候補
  const rooms = [...new Set(recs.map(x => (x.room || '').trim()).filter(Boolean))].slice(-12);
  view.innerHTML = `
    ${orig ? `<div class="card" id="saveBanner" style="background:#fff6ec;border-color:#e8c9a6;display:flex;align-items:center;gap:10px;padding:10px 12px">
      <span style="flex:1;font-size:14px">撮影した写真（元の画質）を本体の写真アプリにも保存しますか？</span>
      <button class="btn acc" id="saveOrig" style="padding:11px 14px;font-size:15px">📥 保存</button>
      <button class="btn ghost" id="skipOrig" style="padding:11px 12px;font-size:14px">しない</button></div>` : ''}
    ${r.broken ? `<div class="card" style="background:#f6dcd6;border-color:#e3a99c;color:var(--danger)">この写真のデータが端末から読み込めなくなっています。下の「写真を差し替え」から撮り直すか、写真アプリから選び直してください（入力した内容はそのまま残ります）。</div>`
      : `<img class="photo-main" src="${blobURL(r.photo)}" alt="">`}
    ${r.extras.length ? `<div class="card" style="padding:10px">
        <div class="muted" style="margin-bottom:6px">同じ箇所の写真 ${r.extras.length + 1} 枚 ― ★が台帳に使う写真。別の写真をタップすると入れ替わります</div>
        <div style="display:flex;gap:8px;overflow-x:auto">
          ${r.broken ? '' : `<div style="position:relative;flex:0 0 96px"><img src="${blobURL(r.thumb || r.photo)}" style="width:96px;height:72px;object-fit:cover;border-radius:6px;outline:3px solid var(--acc)"><span style="position:absolute;left:4px;top:2px;color:#fff;text-shadow:0 0 3px #000">★</span></div>`}
          ${r.extras.map((x, i) => `<div style="position:relative;flex:0 0 96px">
            <img data-pick="${i}" src="${blobURL(x.thumb || x.photo)}" style="width:96px;height:72px;object-fit:cover;border-radius:6px;cursor:pointer">
            <button data-delx="${i}" style="position:absolute;right:2px;top:2px;border:none;border-radius:50%;width:24px;height:24px;background:rgba(0,0,0,.6);color:#fff;font-size:14px">×</button></div>`).join('')}
        </div></div>` : ''}
    <div class="muted" style="margin:6px 2px 10px;display:flex;justify-content:space-between">
      <span>撮影：${fmtDate(r.takenAt)}</span><span>${idx + 1} / ${recs.length}</span>
    </div>
    <div class="row" style="margin-bottom:10px">
      <button class="btn ghost" id="addSame">📷 同じ箇所をもう1枚</button>
      <button class="btn ghost" id="addSameLib" style="flex:0 0 auto">🖼 写真から</button>
    </div>
    <div class="row" style="margin-bottom:12px">
      <button class="btn ghost" ${idx <= 0 ? 'disabled' : ''} id="prevR">‹ 前の箇所</button>
      <button class="btn ghost" ${idx >= recs.length - 1 ? 'disabled' : ''} id="nextR">次の箇所 ›</button>
    </div>
    <div class="card">
      <label style="display:flex;align-items:center;gap:10px;font-size:15px;font-weight:700">
        <input type="checkbox" id="incl" ${r.exclude ? '' : 'checked'} style="width:22px;height:22px"> Excel（調査票・写真台帳）に出力する
      </label>
      <div class="hint">チェックを外すと、記録は残したまま出力からは外れます（写真番号も詰めて振り直します）。</div>
      <label class="f">部位（調査票の区分）</label>
      <div class="chips" id="cP">${CFG.parts.map(p => chip(p, r.part === p)).join('')}</div>
      <label class="f">階</label>
      <input type="text" id="floor" value="${esc(r.floor)}" placeholder="例：2階">
      <div class="chips" id="cF" style="margin-top:7px">${CFG.floors.map(p => chip(p, r.floor === p)).join('')}</div>
      <label class="f">場所（部屋名など）</label>
      <input type="text" id="room" value="${esc(r.room)}" placeholder="例：児童室／北側階段／屋上東側">
      ${rooms.length ? `<div class="chips" id="cR" style="margin-top:7px">${rooms.map(p => chip(p, r.room === p)).join('')}</div>` : ''}
      <label class="f">部位（詳細）</label>
      <input type="text" id="detail" value="${esc(r.detail)}" placeholder="例：パラペット立上り部／天井ボード">
      <label class="f">劣化現象（複数選択可・メモ作成の補助）</label>
      <div class="chips" id="cH">${CFG.phenomena.map(p => chip(p.name, (r.phenomena || []).includes(p.name))).join('')}</div>
      <label class="f">所見メモ（劣化の状況と、それが何を意味しているか）</label>
      <textarea id="memo" placeholder="例：壁のクラック（多数あり）&#10;雨水が浸入すると鉄筋が錆びるおそれがある">${esc(r.memo)}</textarea>
      <div class="hint" id="cnt"></div>
      <div class="row" style="margin-top:8px">
        <button class="btn ghost" id="insEx" style="font-size:14px;padding:10px">選んだ現象の例文を挿入</button>
      </div>
      <div class="hint">改行ごとに段落になります。キーボードのマイクボタンで音声入力もできます。入力は自動保存されます。</div>
    </div>
    <div class="row" style="margin-top:16px">
      <button class="btn ghost" id="saveDev" ${r.broken ? 'disabled' : ''}>📥 本体に保存</button>
      <button class="btn ghost" id="repPhoto">🔄 写真を差し替え</button>
    </div>
    <button class="btn danger wide" id="delR" style="margin-top:10px">この箇所の記録を削除</button>`;

  const fname = (ext) => photoFileName(f, r.exclude ? `対象外${idx + 1}` : nums.get(id), ext);
  if (orig) {
    $('#saveOrig').onclick = async () => {
      const ext = (orig.type || '').includes('png') ? 'png' : (orig.type || '').includes('heic') ? 'heic' : 'jpg';
      const ok = await shareImages([new File([orig], fname(ext), { type: orig.type || 'image/jpeg' })]);
      if (ok) { originals.delete(id); $('#saveBanner').remove(); }
    };
    $('#skipOrig').onclick = () => { originals.delete(id); $('#saveBanner').remove(); };
  }
  // 同じ箇所の追加撮影・台帳に使う写真の選択
  $('#addSame').onclick = () => pickCamera(id);
  $('#addSameLib').onclick = () => pickLibrary(id);
  view.querySelectorAll('[data-pick]').forEach(el => el.onclick = async () => {
    const i = +el.dataset.pick, x = r.extras[i];
    const cur = { photo: r.photo, thumb: r.thumb, w: r.w, h: r.h, takenAt: r.takenAt };
    Object.assign(r, { photo: x.photo, thumb: x.thumb, w: x.w, h: x.h, broken: false });
    if (cur.photo && cur.photo.byteLength) r.extras[i] = cur; else r.extras.splice(i, 1);
    r.updatedAt = Date.now(); await DB.put('records', r); render();
  });
  view.querySelectorAll('[data-delx]').forEach(el => el.onclick = async ev => {
    ev.stopPropagation();
    if (!confirm('この写真を削除しますか？')) return;
    r.extras.splice(+el.dataset.delx, 1); r.updatedAt = Date.now(); await DB.put('records', r); render();
  });
  $('#repPhoto').onclick = () => { $('#repInput').value = ''; $('#repInput').click(); };
  $('#repInput').onchange = async e => {
    const file = e.target.files[0]; if (!file) return;
    toast('差し替え中…', 30000);
    try {
      const img = await processImage(file);
      Object.assign(r, { photo: img.photo, thumb: img.thumb, w: img.w, h: img.h, broken: false, updatedAt: Date.now() });
      await DB.put('records', r); toast('差し替えました'); render();
    } catch (err) { toast('差し替えに失敗しました：' + err.message, 5000); }
  };
  $('#saveDev').onclick = async () => {
    const src = originals.get(id) || r.photo;
    await shareImages([new File([src], fname('jpg'), { type: 'image/jpeg' })]);
    originals.delete(id); const b = $('#saveBanner'); if (b) b.remove();
  };

  const counter = () => {
    const nl = photoNoteLines({ memo: $('#memo').value, phenomena: r.phenomena, part: r.part });
    $('#cnt').innerHTML = nl.over
      ? `<span style="color:var(--danger)">写真台帳の所見欄（${N}字×${L}行）に入りきりません：${nl.lines.length}行／${L}行</span>`
      : `写真台帳の所見欄：${nl.lines.length}行／${L}行（1行${N}字目安）`;
  };
  const save = async () => {
    r.floor = $('#floor').value.trim(); r.room = $('#room').value.trim(); r.detail = $('#detail').value.trim(); r.memo = $('#memo').value;
    r.updatedAt = Date.now();
    await DB.put('records', r);
    await DB.put('kv', { floor: r.floor, room: r.room, part: r.part === '全景' ? '' : r.part }, 'last_' + f.id);
  };
  const saveSoon = debounce(save, 400);
  $('#detail').oninput = saveSoon;
  $('#memo').oninput = () => { counter(); saveSoon(); };
  $('#incl').onchange = async () => { r.exclude = !$('#incl').checked; await save(); render(); };
  const quick = (inp, box) => {
    inp.oninput = () => { if (box) box.querySelectorAll('.chip').forEach(x => x.classList.toggle('on', x.dataset.v === inp.value)); saveSoon(); };
    if (box) box.querySelectorAll('.chip').forEach(c => c.onclick = () => { inp.value = inp.value === c.dataset.v ? '' : c.dataset.v; inp.oninput(); });
  };
  quick($('#floor'), $('#cF')); quick($('#room'), $('#cR'));
  $('#cP').querySelectorAll('.chip').forEach(c => c.onclick = () => {
    r.part = r.part === c.dataset.v ? '' : c.dataset.v;
    $('#cP').querySelectorAll('.chip').forEach(x => x.classList.toggle('on', x.dataset.v === r.part)); counter(); save();
  });
  $('#cH').querySelectorAll('.chip').forEach(c => c.onclick = () => {
    const v = c.dataset.v, s = new Set(r.phenomena || []);
    s.has(v) ? s.delete(v) : s.add(v); r.phenomena = CFG.phenomena.map(p => p.name).filter(n => s.has(n));
    c.classList.toggle('on', s.has(v)); counter(); save();
  });
  $('#insEx').onclick = () => {
    const ex = CFG.phenomena.filter(p => (r.phenomena || []).includes(p.name) && p.meaning).map(p => p.meaning);
    if (!ex.length) return toast('先に劣化現象を選んでください');
    const m = $('#memo'); const add = ex.filter(t => !m.value.includes(t));
    m.value = [m.value.trim(), ...add].filter(Boolean).join('\n'); counter(); save(); m.focus();
  };
  $('#prevR').onclick = async () => { await save(); go(`#/r/${enc(recs[idx - 1].id)}`); };
  $('#nextR').onclick = async () => { await save(); go(`#/r/${enc(recs[idx + 1].id)}`); };
  $('#delR').onclick = async () => {
    if (!confirm(`この箇所の記録（写真 ${r.extras.length + 1} 枚）を削除します。よろしいですか？`)) return;
    await DB.del('records', id); toast('削除しました'); go(back);
  };
  window.addEventListener('hashchange', () => { if (document.body.contains($('#memo'))) save(); }, { once: true });
  counter();
}

/* ------------------------------------------------------- barrier free */
async function vBarrierFree(id) {
  const f = await DB.get('facilities', id);
  if (!f) return go('#/');
  f.bf = f.bf || {};
  setChrome({ title: `バリアフリー　${facLabel(f)}`, back: `#/f/${enc(id)}` });
  let lastGroup = '';
  const rows = CFG.barrierFree.map(it => {
    const g = it.group.replace(/\n/g, ''), lb = it.label.replace(/\n/g, '').replace(/※.*/, '');
    const head = g !== lastGroup ? `<div style="font-weight:700;margin:16px 0 6px;font-size:14px">${esc(g)}</div>` : '';
    lastGroup = g;
    const v = f.bf[it.key] || '';
    let ctl;
    if (it.type === 'num') ctl = `<input type="number" inputmode="numeric" data-bf="${it.key}" value="${esc(v)}" style="width:110px" placeholder="台">`;
    else ctl = `<span class="chips" data-yn="${it.key}">${['有', '無'].map(x => `<button class="chip ${v === x ? 'on' : ''}" data-v="${x}">${x}</button>`).join('')}</span>
      ${it.type === 'yn+text' ? `<input type="text" data-bf="${it.key}_t" value="${esc(f.bf[it.key + '_t'])}" placeholder="${esc(it.hint || '')}" style="margin-top:7px">` : ''}`;
    return `${head}<div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:4px 0">
      <span style="min-width:96px;font-size:14px;color:var(--sub)">${esc(lb)}</span><span style="flex:1">${ctl}</span></div>`;
  }).join('');
  view.innerHTML = `
    <div class="card">${rows}</div>
    <div class="card"><h2>その他気づいた点</h2>
      <textarea data-bf="notes" style="min-height:140px">${esc(f.bf.notes)}</textarea></div>
    <div class="hint" style="text-align:center">入力は自動保存されます。</div>`;
  const save = debounce(async () => { f.updatedAt = Date.now(); await DB.put('facilities', f); }, 300);
  view.querySelectorAll('[data-bf]').forEach(el => el.oninput = () => { f.bf[el.dataset.bf] = el.value.trim(); save(); });
  view.querySelectorAll('[data-yn]').forEach(box => box.querySelectorAll('.chip').forEach(c => c.onclick = () => {
    const k = box.dataset.yn; f.bf[k] = f.bf[k] === c.dataset.v ? '' : c.dataset.v;
    box.querySelectorAll('.chip').forEach(x => x.classList.toggle('on', x.dataset.v === f.bf[k])); save();
  }));
}

/* ------------------------------------------------------- menu */
async function vMenu() {
  const pj = await getProject();
  const [facs, recs] = [await DB.all('facilities'), await DB.all('records')];
  let usage = '';
  try { const e = await navigator.storage.estimate(); usage = `使用容量 約 ${(e.usage / 1048576).toFixed(0)} MB`; } catch (e) { }
  let persisted = false; try { persisted = await navigator.storage.persisted(); } catch (e) { }
  const st = await getSettings();
  setChrome({ title: '出力・設定', back: '#/' });
  view.innerHTML = `
    <div class="card">
      <h2>調査情報</h2>
      <label class="f">自治体名</label><input type="text" id="pCity" value="${esc(pj.city)}" placeholder="例：○○町">
      <div class="hint">バリアフリー調査票の見出し「○○公共施設　バリアフリー調査」に使います。</div>
      <label class="f">調査名</label><input type="text" id="pTitle" value="${esc(pj.title)}" placeholder="例：公共施設 劣化状況調査">
      <label class="f">記入者</label><input type="text" id="pSur" value="${esc(pj.surveyor)}" placeholder="例：山田、宮川">
      <div class="hint">劣化度調査票の「記入者」欄に入ります。</div>
    </div>
    <div class="card">
      <h2>撮影の設定</h2>
      <label style="display:flex;align-items:center;gap:10px;font-size:15px">
        <input type="checkbox" id="sDev" ${st.saveToDevice ? 'checked' : ''} style="width:22px;height:22px">
        撮影後に「本体の写真アプリにも保存」ボタンを表示する
      </label>
      <div class="hint">Webアプリは写真アプリへ自動保存できないため、撮影直後の画面のボタン →「画像を保存」で、元の画質のまま本体に残せます。アプリ内とExcelには、印刷に十分な長辺${CFG.image.maxPx}pxに圧縮して保存します。</div>
    </div>
    <div class="card">
      <h2>Excel出力</h2>
      <p class="muted" style="margin-top:0">建物ごとに「劣化度調査票」「劣化状況写真」「バリアフリー状況」の3シートを様式どおりに出力します。全件出力では、先頭に「調査記録一覧」シートが付きます。調査票の判定・総評は空欄で出力されるので、事務所で記入してください。</p>
      <button class="btn pri wide" id="xlsAll" ${facs.length ? '' : 'disabled'}>全件をExcel出力（${facs.length}件・写真${recs.length}枚）</button>
    </div>
    <div class="card">
      <h2>バックアップ・統合</h2>
      <p class="muted" style="margin-top:0">写真を含む全データを1つのファイルに書き出します。毎日の調査終了時に書き出しておくと安心です。複数の端末で調査した場合は、それぞれのファイルをこの端末で「読み込み」すると統合されます。</p>
      <div class="row">
        <button class="btn" id="bkOut" ${recs.length || facs.length ? '' : 'disabled'}>書き出し</button>
        <button class="btn" id="bkIn">読み込み（統合）</button>
      </div>
    </div>
    <div class="card">
      <h2>施設リストの取り込み</h2>
      <p class="muted" style="margin-top:0">1行目に列見出し「施設番号」「施設名」を置いたCSVを取り込みます。任意で「棟名」「建築年度」「構造」「延床面積」「階数」「所在地」も読み込みます。1行が調査票1枚になります。ExcelのCSV保存（Shift_JIS／UTF-8）どちらでも読めます。</p>
      <button class="btn wide" id="csvIn">CSVを選んで取り込む</button>
    </div>
    <div class="card">
      <h2>端末内のデータ</h2>
      <div class="muted">建物 ${facs.length} 件 ／ 写真 ${recs.length} 枚 ${usage ? '／ ' + usage : ''}</div>
      <div class="muted" style="margin-top:4px">${persisted ? '✓ ブラウザに永続保存を許可されています' : '※ ブラウザの判断でデータが消える可能性があります。こまめにバックアップしてください。'}</div>
      <button class="btn danger wide" style="margin-top:14px" id="wipe">すべてのデータを削除</button>
    </div>
    <p class="muted" style="text-align:center">公共施設 劣化調査ノート v1.6</p>`;
  const saveP = debounce(async () => {
    await DB.put('kv', { city: $('#pCity').value.trim(), title: $('#pTitle').value.trim(), surveyor: $('#pSur').value.trim() }, 'project');
  }, 300);
  ['#pCity', '#pTitle', '#pSur'].forEach(s => $(s).oninput = saveP);
  $('#sDev').onchange = async () => { await DB.put('kv', Object.assign(st, { saveToDevice: $('#sDev').checked }), 'settings'); };
  $('#xlsAll').onclick = () => exportExcel(null);
  $('#bkOut').onclick = exportBackup;
  $('#bkIn').onclick = () => { $('#bkInput').value = ''; $('#bkInput').click(); };
  $('#csvIn').onclick = () => { $('#csvInput').value = ''; $('#csvInput').click(); };
  $('#wipe').onclick = async () => {
    if (!confirm('端末内のすべての建物・写真・記録を削除します。先にバックアップを書き出しましたか？')) return;
    if (!confirm('本当に削除しますか？元に戻せません。')) return;
    await DB.clear(); toast('削除しました'); go('#/');
  };
}

/* ------------------------------------------------------- file save */
/* 時間のかかる出力の後は、iPhoneの「ユーザー操作の直後でないと共有できない」制約にかからないよう、
   いったんダイアログを出して、ボタンを押してもらってから共有シートを開く */
function saveFile(blob, filename) {
  const file = new File([blob], filename, { type: blob.type });
  const canShare = navigator.canShare && navigator.canShare({ files: [file] });
  const mobile = /iPhone|iPad|Android/i.test(navigator.userAgent) || (navigator.maxTouchPoints > 1 && /Macintosh/.test(navigator.userAgent));
  const download = () => {
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 60000);
  };
  if (!(mobile && canShare)) { download(); return; }
  const dlg = $('#dlg');
  $('#dlgTitle').textContent = '✓ 出力の準備ができました';
  $('#dlgName').textContent = `${filename}（${(blob.size / 1048576).toFixed(1)} MB）`;
  $('#dlgHint').textContent = 'ボタンを押すと共有画面が開きます。「"ファイル"に保存」やAirDropを選んでください。';
  $('#dlgShare').textContent = '📤 保存・共有';
  $('#dlgShare').onclick = async () => {
    try { await navigator.share({ files: [file] }); dlg.close(); toast('保存・共有しました'); }
    catch (e) { if (e.name !== 'AbortError') { console.warn(e); toast('共有できませんでした：' + e.message, 5000); } }
  };
  $('#dlgClose').onclick = () => dlg.close();
  dlg.showModal();
}
function stamp() { const d = new Date(); return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`; }
function safeName(s) { return String(s || '').replace(/[\\/:*?"<>|\s]+/g, '_').slice(0, 40); }

async function exportExcel(facilityIds) {
  if (!window.ExcelJS) return toast('Excel出力の部品を読み込み中です。少し待って再度お試しください。');
  toast('Excelを作成中…', 180000);
  try {
    const pj = await getProject();
    let facs = sortFacilities(await DB.all('facilities'));
    if (facilityIds) facs = facs.filter(f => facilityIds.includes(f.id));
    if (!facs.length) return toast('出力する建物がありません');
    await migratePhotos();
    const items = [];
    for (const f of facs) items.push({ f, recs: sortRecords(await DB.byFacility(f.id)).filter(r => !r.exclude) });
    const total = items.reduce((a, x) => a + x.recs.length, 0);
    const blob = await window.buildSurveyWorkbook(pj, items, n => toast(`Excelを作成中… 写真 ${n} / ${total}`, 180000));
    const label = facs.length === 1 ? `${facs[0].no || ''}${facLabel(facs[0])}` : (pj.city || '全件');
    saveFile(blob, `劣化度調査_${safeName(label)}_${stamp()}.xlsx`);
    const noPart = items.reduce((a, x) => a + x.recs.filter(r => !r.part).length, 0);
    const lost = items.reduce((a, x) => a + x.recs.filter(r => r.broken).length, 0);
    if (lost) { toast(`写真データが読み込めない写真が ${lost} 枚あり、その欄は「写真データなし」で出力しました。`, 7000); return; }
    toast(noPart ? `部位（区分）が未設定の写真が ${noPart} 枚あり、調査票の劣化状況には載っていません。` : 'Excelの準備ができました', noPart ? 6000 : 2200);
  } catch (e) { console.error(e); toast('Excel出力に失敗しました：' + e.message, 6000); }
}

/* ------------------------------------------------------- backup */
async function exportBackup() {
  toast('バックアップを作成中…', 120000);
  try {
    const pj = await getProject();
    const facilities = await DB.all('facilities');
    const records = [];
    for (const r of await DB.all('records')) {
      if (r.broken) continue;
      const extras = [];
      for (const x of r.extras || []) extras.push(Object.assign({}, x, { photo: await blobToDataURL(asBlob(x.photo)), thumb: await blobToDataURL(asBlob(x.thumb)) }));
      records.push(Object.assign({}, r, { photo: await blobToDataURL(asBlob(r.photo)), thumb: await blobToDataURL(asBlob(r.thumb)), extras }));
    }
    const blob = new Blob([JSON.stringify({ app: 'rekka-survey', version: 2, exportedAt: Date.now(), project: pj, facilities, records })], { type: 'application/json' });
    toast('書き出しの準備ができました');
    saveFile(blob, `劣化調査バックアップ_${safeName(pj.city || '')}_${stamp()}.json`);
  } catch (e) { console.error(e); toast('書き出しに失敗しました：' + e.message, 6000); }
}
$('#bkInput').onchange = async e => {
  const file = e.target.files[0]; if (!file) return;
  toast('読み込み中…', 120000);
  try {
    const data = JSON.parse(await file.text());
    if (data.app !== 'rekka-survey') throw new Error('このアプリのバックアップファイルではありません');
    let nf = 0, nr = 0, ur = 0;
    const pj = await getProject();
    if (!pj.title && data.project) await DB.put('kv', data.project, 'project');
    for (const f of data.facilities || []) {
      const cur = await DB.get('facilities', f.id);
      if (!cur) { await DB.put('facilities', f); nf++; }
      else if ((f.updatedAt || 0) > (cur.updatedAt || 0)) await DB.put('facilities', Object.assign({}, cur, f));
    }
    for (const r of data.records || []) {
      const cur = await DB.get('records', r.id);
      if (cur && (cur.updatedAt || 0) >= (r.updatedAt || 0)) continue;
      const buf = async u => (await dataURLToBlob(u)).arrayBuffer();
      const extras = [];
      for (const x of r.extras || []) extras.push(Object.assign({}, x, { photo: await buf(x.photo), thumb: await buf(x.thumb) }));
      await DB.put('records', Object.assign({}, r, { photo: await buf(r.photo), thumb: await buf(r.thumb), extras, broken: false }));
      cur ? ur++ : nr++;
    }
    toast(`統合しました：建物 +${nf}、写真 +${nr}（更新 ${ur}）`, 4000); render();
  } catch (err) { console.error(err); toast('読み込みに失敗しました：' + err.message, 6000); }
};

/* ------------------------------------------------------- CSV import */
function parseCSV(text) {
  const out = []; let row = [], cell = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += ch; }
    else if (ch === '"') q = true;
    else if (ch === ',') { row.push(cell); cell = ''; }
    else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cell); out.push(row); row = []; cell = ''; }
    else cell += ch;
  }
  if (cell || row.length) { row.push(cell); out.push(row); }
  return out.filter(r => r.some(c => c.trim()));
}
async function readText(file) {
  const buf = await file.arrayBuffer();
  try { return new TextDecoder('utf-8', { fatal: true }).decode(buf).replace(/^﻿/, ''); }
  catch (e) { return new TextDecoder('shift_jis').decode(buf); }
}
$('#csvInput').onchange = async e => {
  const file = e.target.files[0]; if (!file) return;
  try {
    const rows = parseCSV(await readText(file));
    const head = rows.shift().map(h => h.trim());
    const col = names => head.findIndex(h => names.includes(h));
    const C = {
      no: col(['施設番号', '番号', 'No', 'NO', 'No.']), name: col(['施設名', '施設名称', '名称']),
      building: col(['棟名', '棟', '建物名']), builtYear: col(['建築年度', '建築年', '竣工年度']),
      structure: col(['構造']), floorArea: col(['延床面積', '延べ床面積', '延床面積（㎡）']),
      floors: col(['階数']), address: col(['所在地', '住所'])
    };
    if (C.name < 0) throw new Error('「施設名」の列が見つかりません');
    const existing = await DB.all('facilities');
    let add = 0, upd = 0;
    for (const r of rows) {
      const v = {}; for (const k in C) v[k] = C[k] >= 0 ? (r[C[k]] || '').trim() : '';
      if (!v.name) continue;
      const cur = existing.find(x => x.name === v.name && (x.building || '') === v.building && (x.no || '') === v.no);
      if (cur) {
        for (const k in v) if (v[k] && !cur[k]) cur[k] = v[k];
        cur.updatedAt = Date.now(); await DB.put('facilities', cur); upd++;
      } else {
        const cand = v.no ? `f_${v.no}_${v.building}` : '';
        const id = cand && !existing.some(x => x.id === cand) ? cand : uid('f_');
        const nf = Object.assign(v, { id, history: {}, note: '', createdAt: Date.now(), updatedAt: Date.now() });
        await DB.put('facilities', nf); existing.push(nf); add++;
      }
    }
    toast(`取り込みました：追加 ${add} 件、更新 ${upd} 件`, 3500); go('#/');
  } catch (err) { console.error(err); toast('取り込みに失敗しました：' + err.message, 6000); }
};

/* ------------------------------------------------------- boot */
if ('serviceWorker' in navigator && location.protocol !== 'file:') {
  navigator.serviceWorker.register('sw.js').catch(e => console.warn('SW', e));
}
(async () => {
  try {
    const broken = await migratePhotos();
    if (broken) toast(`写真データが読み込めない写真が ${broken} 枚あります。一覧の赤い欄から差し替えてください。`, 7000);
  } catch (e) { console.warn(e); }
  render();
})();
