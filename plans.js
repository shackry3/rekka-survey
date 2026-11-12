/* 公共施設 劣化調査ノート  —  plans.js（図面・撮影位置）
   階ごとに図面（画像／PDF）を取り込み、図面上をタップ → カメラ起動 → 記録。
   撮影位置図（図面＋番号付きピン）を画像で出力する。app.js の DB・render 等を使う。 */
'use strict';

let assignRecord = null;          // 既存の記録に位置を付けるとき、その記録のID
async function plansOf(fid) {
  const ps = await DB.plans(fid + ':');
  return ps.sort((a, b) => floorKey(a.floor) - floorKey(b.floor) || (a.createdAt - b.createdAt));
}

/* ---------- 取り込み ---------- */
async function loadPdfJs() {
  if (window.pdfjsLib) return;
  await new Promise((res, rej) => {
    const s = document.createElement('script'); s.src = 'vendor/pdf.min.js'; s.onload = res; s.onerror = () => rej(new Error('PDF読込部品を読み込めません'));
    document.head.appendChild(s);
  });
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';
}
async function canvasToBuf(c, q = 0.85) {
  const b = await new Promise(r => c.toBlob(r, 'image/jpeg', q));
  return { buf: await b.arrayBuffer(), w: c.width, h: c.height };
}
const PLAN_MAX = 3000;                                   // 図面画像の長辺（px）
async function planImageFromFile(file, pageNo) {
  if (/pdf/i.test(file.type) || /\.pdf$/i.test(file.name)) {
    await loadPdfJs();
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
    const n = Math.min(Math.max(1, pageNo || 1), pdf.numPages);
    const page = await pdf.getPage(n);
    const v0 = page.getViewport({ scale: 1 });
    const scale = Math.min(PLAN_MAX / Math.max(v0.width, v0.height), 6);
    const vp = page.getViewport({ scale });
    const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
    const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
    await page.render({ canvasContext: ctx, viewport: vp }).promise;
    const out = await canvasToBuf(c); c.width = c.height = 0; return out;
  }
  const bmp = await loadBitmap(file);
  const k = Math.min(1, PLAN_MAX / Math.max(bmp.width, bmp.height));
  const c = document.createElement('canvas'); c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
  const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
  ctx.drawImage(bmp, 0, 0, c.width, c.height); if (bmp.close) bmp.close();
  return await canvasToBuf(c);
}
async function pdfPageCount(file) {
  if (!(/pdf/i.test(file.type) || /\.pdf$/i.test(file.name))) return 1;
  await loadPdfJs();
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  return pdf.numPages;
}

/* ---------- 図面の一覧 ---------- */
async function vPlans(fid) {
  const f = await DB.get('facilities', fid);
  if (!f) return go('#/');
  currentFacility = fid;
  const plans = await plansOf(fid);
  const recs = await recordsOf(f);
  const cnt = {}; recs.forEach(r => { if (r.pin) cnt[r.pin.planId] = (cnt[r.pin.planId] || 0) + 1; });
  const assigning = assignRecord && recs.some(r => r.id === assignRecord);
  setChrome({ title: `図面　${facLabel(f)}`, back: assigning ? `#/r/${enc(assignRecord)}` : `#/f/${enc(fid)}` });
  view.innerHTML = `
    ${assigning ? `<div class="card" style="background:#fff6ec;border-color:#e8c9a6">位置を付ける図面（階）を選んでください。</div>` : ''}
    <div class="card">
      <h2>図面（階ごとに1枚）</h2>
      ${plans.length ? plans.map(p => `
        <div class="list-item" data-plan="${esc(p.id)}">
          <img src="${blobURL(p.image)}" style="width:72px;height:54px;object-fit:contain;background:#eee;border-radius:6px">
          <span class="name">${esc(p.floor)}</span>
          <span class="cnt">📍 ${cnt[p.id] || 0}</span>
        </div>`).join('') : `<div class="empty">図面がまだありません。<br>下で階の名前を入れて、図面の画像かPDFを選んでください。</div>`}
    </div>
    ${assigning ? '' : `
    <div class="card">
      <h2>図面を追加</h2>
      <label class="f" style="margin-top:0">階</label>
      <input type="text" id="plFloor" placeholder="例：1階／2階／屋上">
      <div class="chips" style="margin-top:7px" id="plChips">${CFG.floors.map(x => `<button class="chip" data-v="${esc(x)}">${esc(x)}</button>`).join('')}</div>
      <button class="btn pri wide" id="plAdd" style="margin-top:12px">図面（画像・PDF）を選ぶ</button>
      <div class="hint">PDFは1ページずつ取り込みます（複数ページの場合はページ番号を聞きます）。同じ階を選ぶと図面を差し替えます（撮影位置はそのまま）。</div>
    </div>
    ${plans.length ? `<div class="card">
      <h2>撮影位置図の出力</h2>
      <p class="muted" style="margin-top:0">図面に写真番号のピンを描いた画像（JPEG）を、階ごとに出力します。番号は一覧の並び順の写真番号です。</p>
      <button class="btn wide" id="plOut">🗺 撮影位置図を出力（${plans.length}枚）</button>
    </div>` : ''}`}`;
  view.querySelectorAll('[data-plan]').forEach(el => el.onclick = () => go(`#/plan/${enc(el.dataset.plan)}`));
  if (assigning) return;
  view.querySelectorAll('#plChips .chip').forEach(c => c.onclick = () => { $('#plFloor').value = c.dataset.v; });
  $('#plAdd').onclick = () => {
    if (!$('#plFloor').value.trim()) return toast('先に階の名前を入れてください');
    const inp = $('#planInput'); inp.value = '';
    inp.onchange = async e => {
      const file = e.target.files[0]; if (!file) return;
      const floor = $('#plFloor').value.trim();
      try {
        toast('図面を読み込み中…', 60000);
        let page = 1;
        const n = await pdfPageCount(file);
        if (n > 1) { const v = prompt(`このPDFは ${n} ページあります。何ページ目を「${floor}」の図面にしますか？`, '1'); if (!v) return toast('取り込みを中止しました'); page = parseInt(v, 10) || 1; }
        const img = await planImageFromFile(file, page);
        const same = plans.find(p => p.floor === floor);
        const p = Object.assign(same || { id: uid('p_'), facilityId: fid, floor, createdAt: Date.now() },
          { image: img.buf, w: img.w, h: img.h, source: file.name + (n > 1 ? ` p.${page}` : ''), updatedAt: Date.now() });
        await DB.putPlan(p);
        requestPersist();
        toast(same ? `「${floor}」の図面を差し替えました` : `「${floor}」の図面を追加しました`);
        go(`#/plan/${enc(p.id)}`);
      } catch (err) { console.error(err); toast('図面を読み込めませんでした：' + err.message, 6000); }
    };
    inp.click();
  };
  if ($('#plOut')) $('#plOut').onclick = async () => {
    toast('撮影位置図を作成中…', 60000);
    try {
      const nums = photoNumbers(recs), items = [];
      for (const p of plans) items.push({ data: await renderPlanImage(f, p, recs, nums), name: `${f.no ? f.no + '_' : ''}${facLabel(f).replace(/\s+/g, '')}_撮影位置図_${p.floor}.jpg` });
      toast('準備ができました', 1500);
      shareInBatches(items, '撮影位置図を保存', 10);
    } catch (err) { console.error(err); toast('作成に失敗しました：' + err.message, 6000); }
  };
}

/* ---------- 撮影位置図の画像 ---------- */
function pinRadius(p) { return Math.max(16, Math.round(Math.max(p.w, p.h) / 75)); }
async function renderPlanImage(f, p, recs, nums) {
  const bmp = await createImageBitmap(asBlob(p.image));
  const head = Math.round(Math.max(p.w, p.h) / 30);
  const c = document.createElement('canvas'); c.width = p.w; c.height = p.h + head;
  const ctx = c.getContext('2d');
  ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
  ctx.drawImage(bmp, 0, head, p.w, p.h); if (bmp.close) bmp.close();
  ctx.fillStyle = '#222'; ctx.font = `bold ${Math.round(head * 0.55)}px sans-serif`; ctx.textBaseline = 'middle';
  ctx.fillText(`撮影位置図　${f.no ? f.no + '　' : ''}${facLabel(f)}　${p.floor}`, head * 0.4, head / 2);
  const R = pinRadius(p);
  for (const r of recs.filter(x => x.pin && x.pin.planId === p.id && !x.exclude)) {
    const x = r.pin.x * p.w, y = head + r.pin.y * p.h;
    ctx.beginPath(); ctx.arc(x, y, R, 0, Math.PI * 2);
    ctx.fillStyle = '#c62828'; ctx.fill(); ctx.lineWidth = Math.max(2, R / 6); ctx.strokeStyle = '#fff'; ctx.stroke();
    ctx.fillStyle = '#fff'; ctx.font = `bold ${Math.round(R * 1.05)}px sans-serif`; ctx.textAlign = 'center';
    ctx.fillText(String(nums.get(r.id)), x, y + 1); ctx.textAlign = 'left';
  }
  const out = await new Promise(res => c.toBlob(res, 'image/jpeg', 0.88));
  c.width = c.height = 0;
  return out;
}

/* ---------- 図面の表示・ピン ---------- */
async function vPlan(pid) {
  const p = await DB.getPlan(pid);
  if (!p) return go('#/');
  const f = await DB.get('facilities', p.facilityId);
  currentFacility = f.id;
  const recs = await recordsOf(f);
  const nums = photoNumbers(recs);
  const pins = recs.filter(r => r.pin && r.pin.planId === pid);
  const assigning = assignRecord && recs.find(r => r.id === assignRecord);
  let placing = !assigning && !!vPlan.placing;
  setChrome({
    title: `${p.floor}　${facLabel(f)}`, back: assigning ? `#/r/${enc(assigning.id)}` : `#/plans/${enc(f.id)}`,
    fabHTML: assigning ? `<button class="btn ghost" id="cancelAs" style="background:var(--card)">やめる</button>`
      : `<button class="btn ghost" style="background:var(--card);flex:0 0 auto" onclick="go('#/f/${enc(f.id)}')">一覧へ</button>
         <button class="btn ${placing ? 'acc' : 'pri'}" id="placeBtn">${placing ? '📍 位置をタップ → 撮影' : '📍 撮影位置を置く'}</button>`
  });
  view.innerHTML = `
    <div class="card" style="padding:8px 10px;margin-bottom:8px;display:flex;align-items:center;gap:8px">
      <span class="muted" style="flex:1" id="plHint">${assigning ? `写真「${esc(assigning.room || assigning.detail || '')}」の位置をタップしてください` : placing ? '図面上の撮影位置をタップするとカメラが起動します' : '2本指で拡大・1本指で移動。ピンをタップすると記録を開きます'}</span>
      <button class="btn ghost sm" id="zIn" style="flex:0 0 auto">＋</button>
      <button class="btn ghost sm" id="zOut" style="flex:0 0 auto">－</button>
      <button class="btn ghost sm" id="zFit" style="flex:0 0 auto">全体</button>
    </div>
    <div id="pv" style="position:relative;overflow:hidden;height:calc(100vh - 290px);min-height:300px;touch-action:none;background:#d9d6cc;border-radius:10px">
      <div id="pvIn" style="position:absolute;left:0;top:0;transform-origin:0 0">
        <img id="pvImg" src="${blobURL(p.image)}" style="display:block;width:${p.w}px;height:${p.h}px;max-width:none;user-select:none;-webkit-user-select:none" draggable="false">
        ${pins.map(r => `<div class="pin" data-rid="${esc(r.id)}" style="position:absolute;left:${r.pin.x * p.w}px;top:${r.pin.y * p.h}px;transform-origin:0 0">
          <div style="transform:translate(-50%,-50%);min-width:30px;height:30px;padding:0 5px;border-radius:15px;background:${r.exclude ? '#888' : '#c62828'};color:#fff;font-weight:700;font-size:15px;display:flex;align-items:center;justify-content:center;border:2px solid #fff;box-shadow:0 1px 4px rgba(0,0,0,.4)">${r.exclude ? '×' : nums.get(r.id)}</div></div>`).join('')}
      </div>
    </div>
    ${assigning ? '' : `<div class="row" style="margin-top:8px">
      <button class="btn ghost sm" id="plRename">✏ 階名変更</button>
      <button class="btn ghost sm" id="plReplace">🔄 差し替え</button>
      <button class="btn danger sm" id="plDelete">🗑 削除</button>
    </div>`}`;
  const pv = $('#pv'), pin = $('#pvIn');
  let s = 1, tx = 0, ty = 0;
  const apply = () => {
    pin.style.transform = `translate(${tx}px,${ty}px) scale(${s})`;
    pin.querySelectorAll('.pin').forEach(el => el.style.transform = `scale(${1 / s})`);
  };
  const fit = () => {
    const W = pv.clientWidth, H = pv.clientHeight;
    s = Math.min(W / p.w, H / p.h); tx = (W - p.w * s) / 2; ty = (H - p.h * s) / 2; apply();
  };
  const zoomAt = (k, cx, cy) => {
    const ns = Math.min(Math.max(s * k, 0.05), 4);
    tx = cx - (cx - tx) * ns / s; ty = cy - (cy - ty) * ns / s; s = ns; apply();
  };
  if (vPlan.view && vPlan.view.pid === pid) ({ s, tx, ty } = vPlan.view), apply(); else fit();
  const remember = () => { vPlan.view = { pid, s, tx, ty }; };
  $('#zIn').onclick = () => { zoomAt(1.5, pv.clientWidth / 2, pv.clientHeight / 2); remember(); };
  $('#zOut').onclick = () => { zoomAt(1 / 1.5, pv.clientWidth / 2, pv.clientHeight / 2); remember(); };
  $('#zFit').onclick = () => { fit(); remember(); };
  // 1本指で移動・2本指で拡大縮小
  const pts = new Map(); let moved = false, last = null;
  pv.addEventListener('pointerdown', e => { pts.set(e.pointerId, { x: e.clientX, y: e.clientY }); if (pts.size === 1) moved = false; last = null; });
  pv.addEventListener('pointermove', e => {
    if (!pts.has(e.pointerId)) return;
    const prev = pts.get(e.pointerId); pts.set(e.pointerId, { x: e.clientX, y: e.clientY });
    if (pts.size === 1) {
      const dx = e.clientX - prev.x, dy = e.clientY - prev.y;
      if (Math.abs(dx) + Math.abs(dy) > 0) { tx += dx; ty += dy; apply(); }
      if (!moved && (Math.abs(dx) > 3 || Math.abs(dy) > 3)) moved = true;
    } else if (pts.size === 2) {
      moved = true;
      const [a, b] = [...pts.values()], rect = pv.getBoundingClientRect();
      const d = Math.hypot(a.x - b.x, a.y - b.y), cx = (a.x + b.x) / 2 - rect.left, cy = (a.y + b.y) / 2 - rect.top;
      if (last) { zoomAt(d / last.d, cx, cy); tx += cx - last.cx; ty += cy - last.cy; apply(); }
      last = { d, cx, cy };
    }
  });
  const up = e => { pts.delete(e.pointerId); if (pts.size < 2) last = null; remember(); };
  pv.addEventListener('pointerup', up); pv.addEventListener('pointercancel', up);
  pv.addEventListener('wheel', e => { e.preventDefault(); const r = pv.getBoundingClientRect(); zoomAt(e.deltaY < 0 ? 1.15 : 1 / 1.15, e.clientX - r.left, e.clientY - r.top); remember(); }, { passive: false });
  // タップ（click）：ピン→記録を開く／配置モード→位置を決めてカメラ起動
  pv.addEventListener('click', async e => {
    if (moved) { moved = false; return; }
    const pinEl = e.target.closest('.pin');
    if (pinEl && !assigning) { go(`#/r/${enc(pinEl.dataset.rid)}`); return; }
    const rect = pv.getBoundingClientRect();
    const x = ((e.clientX - rect.left) - tx) / s / p.w, y = ((e.clientY - rect.top) - ty) / s / p.h;
    if (x < 0 || y < 0 || x > 1 || y > 1) return;
    if (assigning) {
      assigning.pin = { planId: pid, x, y }; assigning.updatedAt = Date.now();
      if (!assigning.floor) assigning.floor = p.floor;
      const rid = assigning.id; assignRecord = null;
      await DB.put('records', assigning); toast('撮影位置を付けました'); go(`#/r/${enc(rid)}`);
      return;
    }
    if (placing) pickCamera(null, { planId: pid, x, y, floor: p.floor });   // タップ直後にカメラを開く（iPhoneの制約）
  });
  if (assigning) { $('#cancelAs').onclick = () => { const rid = assignRecord; assignRecord = null; go(`#/r/${enc(rid)}`); }; return; }
  $('#placeBtn').onclick = () => { vPlan.placing = !placing; render(true); };
  // 図面の管理
  $('#plRename').onclick = async () => {
    const v = prompt('階の名前', p.floor); if (v === null || !v.trim() || v.trim() === p.floor) return;
    p.floor = v.trim(); p.updatedAt = Date.now(); await DB.putPlan(p);
    if (confirm(`この図面に撮影位置がある写真（${pins.length}枚）の「階」も「${p.floor}」に変えますか？`))
      for (const r of pins) { r.floor = p.floor; r.updatedAt = Date.now(); await DB.put('records', r); }
    toast('階名を変更しました'); render(true);
  };
  $('#plReplace').onclick = () => {
    const inp = $('#planInput'); inp.value = '';
    inp.onchange = async e => {
      const file = e.target.files[0]; if (!file) return;
      try {
        toast('図面を読み込み中…', 60000);
        let page = 1; const n = await pdfPageCount(file);
        if (n > 1) { const v = prompt(`このPDFは ${n} ページあります。何ページ目を使いますか？`, '1'); if (!v) return toast('中止しました'); page = parseInt(v, 10) || 1; }
        const img = await planImageFromFile(file, page);
        Object.assign(p, { image: img.buf, w: img.w, h: img.h, source: file.name, updatedAt: Date.now() });
        await DB.putPlan(p); vPlan.view = null;
        toast('図面を差し替えました（撮影位置はそのままの相対位置で残ります）', 4000); render();
      } catch (err) { console.error(err); toast('図面を読み込めませんでした：' + err.message, 6000); }
    };
    inp.click();
  };
  $('#plDelete').onclick = async () => {
    const msg = pins.length
      ? `「${p.floor}」の図面を削除します。\nこの図面に付けた撮影位置（${pins.length}箇所）も外れます（写真と入力内容は残ります）。よろしいですか？`
      : `「${p.floor}」の図面を削除します。よろしいですか？`;
    if (!confirm(msg)) return;
    for (const r of pins) { delete r.pin; r.updatedAt = Date.now(); await DB.put('records', r); }
    await DB.delPlan(p); vPlan.view = null; vPlan.placing = false;
    toast('図面を削除しました'); go(`#/plans/${enc(f.id)}`);
  };
}

/* ---------- 複数ファイルを少しずつ共有（iPhoneは一度に多くを渡すと共有画面が開かない） ---------- */
function shareInBatches(items, title, B = 20) {
  if (!items.length) return toast('保存できるものがありません');
  if (!isMobile()) { shareImages(items.map(x => new File([x.data], x.name, { type: x.data.type || 'image/jpeg' }))); return; }
  const batches = [];
  for (let i = 0; i < items.length; i += B) batches.push(items.slice(i, i + B));
  let k = 0;
  const dlg = $('#dlg');
  const show = () => {
    const b = batches[k];
    $('#dlgTitle').textContent = `${title}（${k + 1} / ${batches.length} 回目）`;
    $('#dlgName').textContent = b.length === 1 ? b[0].name : `${b[0].name} ほか ${b.length - 1} 件`;
    $('#dlgHint').textContent = 'ボタンを押すと共有画面が開きます。その中の「画像を保存」（写真アプリ）か「"ファイル"に保存」を押してください。';
    $('#dlgShare').textContent = '📥 この分を保存';
  };
  $('#dlgShare').onclick = async () => {
    const files = batches[k].map(x => new File([x.data], x.name, { type: x.data.type || 'image/jpeg' }));
    if (!(await shareImages(files))) return;
    k++;
    if (k >= batches.length) { dlg.close(); toast('保存しました'); } else show();
  };
  $('#dlgClose').onclick = () => dlg.close();
  show(); dlg.showModal();
}
