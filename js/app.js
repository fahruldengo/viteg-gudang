/* ============================================================
   VITEG Gudang — aplikasi frontend
   ============================================================ */
const App = { user: null, cabang: [], models: [] };
const CHUNK_MASUK = 25;   // batas aman panjang URL GET
const CHUNK_KELUAR = 50;

/* ---------------- util ---------------- */
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
const esc = v => String(v == null ? '' : v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const isUtama = () => App.user && App.user.role === 'UTAMA';
const today = () => new Date().toLocaleDateString('sv-SE');
const daysAgo = n => { const d = new Date(); d.setDate(d.getDate() - n); return d.toLocaleDateString('sv-SE'); };
const fmt = n => Number(n || 0).toLocaleString('id-ID');

function toast(msg, bad) {
  const t = $('#toast');
  t.textContent = msg;
  t.className = 'toast show' + (bad ? ' bad' : '');
  clearTimeout(toast._t);
  toast._t = setTimeout(() => { t.className = 'toast'; }, 3200);
}
const loadingHtml = (txt = 'Memuat data…') => `<div class="loading"><div class="spin"></div>${esc(txt)}</div>`;
function errBox(e, extra = '') {
  const list = e && e.errors && e.errors.length
    ? `<ul>${e.errors.map(x => `<li>${esc(x)}</li>`).join('')}</ul>` : '';
  return `<div class="alert bad"><strong>${esc(e.message || e)}</strong>${extra}${list}</div>`;
}
const okBox = m => `<div class="alert ok">${m}</div>`;
const badge = s => `<span class="badge b-${esc(s)}">${esc(s)}</span>`;
const cabNama = k => { const c = App.cabang.find(x => x.kode === k); return c ? c.nama : (k || '-'); };
function opsiCabang(sel, semua, hanyaAktif = true) {
  let h = semua ? `<option value="">${esc(semua)}</option>` : '';
  App.cabang.filter(c => !hanyaAktif || c.aktif === '1').forEach(c => {
    h += `<option value="${esc(c.kode)}" ${c.kode === sel ? 'selected' : ''}>${esc(c.nama)} (${esc(c.kode)})</option>`;
  });
  return h;
}

/* normalisasi — sama dengan backend */
const normImei = v => String(v || '').replace(/\D/g, '');
const normSn = v => String(v || '').toUpperCase().replace(/\s+/g, '');
const normMac = v => { const h = String(v || '').toUpperCase().replace(/[^0-9A-F]/g, ''); return h.length === 12 ? h.match(/../g).join(':') : ''; };
const MAC_SEP = /^([0-9a-f]{2}[:-]){5}[0-9a-f]{2}$/i;

/** Pecah satu baris tempelan menjadi [imei, sn, mac]; urutan kolom dideteksi otomatis. */
function parseLine(line) {
  // Copy dari Excel dipisah Tab -> pisah hanya di Tab (agar "3,57E+14" tetap utuh dan terdeteksi)
  const tok = (line.includes('\t') ? line.split('\t') : line.split(/[,;|]+|\s+/)).map(s => s.trim()).filter(Boolean);
  if (!tok.length) return null;
  let imei = '', mac = '';
  const rest = [];
  tok.forEach(t => {
    if (!imei && /^\d{15}$/.test(t)) imei = t;
    else if (!mac && MAC_SEP.test(t)) mac = t;
    else rest.push(t);
  });
  if (!imei && rest.length) imei = rest.shift();
  const sn = rest.shift() || '';
  if (!mac && rest.length) mac = rest.shift();
  return [imei, sn, mac];
}

function loadScript(src) {
  return new Promise((ok, no) => { const s = document.createElement('script'); s.src = src; s.onload = ok; s.onerror = no; document.head.appendChild(s); });
}

const XLSX_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
async function loadXLSX() {
  if (!window.XLSX) await loadScript(XLSX_URL);
  return window.XLSX;
}
/** Nilai sel Excel -> teks; angka besar (IMEI) tanpa notasi ilmiah. */
const cellText = v => typeof v === 'number' ? (Number.isInteger(v) ? v.toFixed(0) : String(v)) : String(v == null ? '' : v).trim();
const isSci = s => /^\d([.,]\d+)?E\+\d+$/i.test(String(s).trim());

/** Template Excel untuk barang masuk: kolom IMEI/SN/MAC berformat Teks. */
async function downloadTemplateMasuk() {
  try {
    const X = await loadXLSX();
    const ws = X.utils.aoa_to_sheet([['IMEI', 'SN', 'MAC']]);
    const ROWS = 1000;
    for (let r = 1; r <= ROWS; r++) for (let c = 0; c < 3; c++) ws[X.utils.encode_cell({ r, c })] = { t: 's', v: '', z: '@' };
    ws['!ref'] = 'A1:C' + (ROWS + 1);
    ws['!cols'] = [{ wch: 20 }, { wch: 22 }, { wch: 20 }];
    const info = X.utils.aoa_to_sheet([
      ['PETUNJUK PENGISIAN — isi data di sheet "Data", jangan ubah judul kolom'],
      [''],
      ['Kolom', 'Aturan', 'Contoh'],
      ['IMEI', 'Wajib, 15 digit angka', '356789012345671'],
      ['SN', 'Wajib, serial number perangkat', 'ZTEGC1234567'],
      ['MAC', 'Wajib, 12 karakter hex; boleh pakai : atau - atau tanpa pemisah', 'AA:BB:CC:11:22:33'],
      [''],
      ['Kolom di sheet Data sudah berformat Teks supaya IMEI tidak berubah jadi 3,57E+14.'],
      ['Satu baris = satu unit. Model/tipe dan cabang dipilih di aplikasi saat impor.']
    ]);
    info['!cols'] = [{ wch: 12 }, { wch: 60 }, { wch: 22 }];
    const wb = X.utils.book_new();
    X.utils.book_append_sheet(wb, ws, 'Data');
    X.utils.book_append_sheet(wb, info, 'Petunjuk');
    X.writeFile(wb, 'template-barang-masuk.xlsx');
  } catch (e) { toast('Gagal membuat template. Periksa koneksi internet.', true); }
}

/** Baca file Excel/CSV -> [[imei, sn, mac], ...]. Kolom dicari dari judul; kalau tak ada judul, pakai urutan A-B-C. */
async function readMasukFile(file) {
  const X = await loadXLSX();
  const wb = X.read(await file.arrayBuffer(), { type: 'array', raw: false, cellText: false });
  const name = wb.SheetNames.includes('Data') ? 'Data' : wb.SheetNames[0];
  const rows = X.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '' });
  let hi = rows.findIndex(r => r.some(c => /imei/i.test(String(c))));
  let ci = { imei: 0, sn: 1, mac: 2 };
  if (hi >= 0) {
    const h = rows[hi].map(c => String(c).trim().toLowerCase());
    const find = re => h.findIndex(c => re.test(c));
    ci = { imei: find(/imei/), sn: find(/^(sn|s\/n|serial|no\.?\s*seri|nomor seri)/), mac: find(/mac/) };
    if (ci.sn < 0 || ci.mac < 0) throw new Error('Judul kolom SN atau MAC tidak ditemukan. Gunakan template.');
  }
  return rows.slice(hi + 1)
    .map(r => [cellText(r[ci.imei]), cellText(r[ci.sn]), cellText(r[ci.mac])])
    .filter(r => r.some(Boolean));
}

/** Ekspor ke Excel (.xlsx); IMEI tetap teks. Cadangan: CSV. */
async function exportExcel(nama, cols, rows) {
  const aoa = [cols.map(c => c[1])].concat(rows.map(r => cols.map(c => String(r[c[0]] == null ? '' : r[c[0]]))));
  try {
    await loadXLSX();
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = cols.map(c => ({ wch: c[2] || 16 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');
    XLSX.writeFile(wb, nama + '.xlsx');
  } catch (e) {
    const csv = aoa.map(r => r.map(v => `"${v.replace(/"/g, '""')}"`).join(';')).join('\r\n');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob(['﻿' + csv], { type: 'text/csv' }));
    a.download = nama + '.csv';
    a.click();
  }
}

/* ---------------- modal ---------------- */
function modal(html) {
  $('#modalCard').innerHTML = html;
  $('#modal').classList.remove('hidden');
  const f = $('#modalCard input:not([disabled]), #modalCard select');
  if (f) f.focus();
}
function closeModal() { $('#modal').classList.add('hidden'); }

/* ---------------- menu & router ---------------- */
const MENU = [
  { g: 'Transaksi' },
  { id: 'dashboard', label: 'Dashboard', ic: '◫' },
  { id: 'masuk', label: 'Barang Masuk', ic: '↓' },
  { id: 'keluar', label: 'Barang Keluar', ic: '↑' },
  { id: 'mutasi', label: 'Mutasi Cabang', ic: '⇄', utama: true },
  { g: 'Data' },
  { id: 'stok', label: 'Stok Unit', ic: '▤' },
  { id: 'riwayat', label: 'Riwayat Transaksi', ic: '☰' },
  { id: 'lacak', label: 'Lacak Unit', ic: '⌕' },
  { g: 'Pengaturan' },
  { id: 'cabang', label: 'Cabang', ic: '⌂', utama: true },
  { id: 'pengguna', label: 'Pengguna', ic: '☺', utama: true },
  { id: 'password', label: 'Ganti Password', ic: '⚿' }
];

function renderNav() {
  $('#nav').innerHTML = MENU.filter(m => !m.utama || isUtama()).map(m =>
    m.g ? `<div class="group">${m.g}</div>`
        : `<a href="#${m.id}" data-id="${m.id}"><span class="ic">${m.ic}</span>${m.label}</a>`
  ).join('');
}

async function route() {
  let id = location.hash.replace('#', '') || 'dashboard';
  const m = MENU.find(x => x.id === id && (!x.utama || isUtama()));
  if (!m) { id = 'dashboard'; }
  const item = MENU.find(x => x.id === id);
  $$('#nav a').forEach(a => a.classList.toggle('active', a.dataset.id === id));
  $('#pageTitle').textContent = item.label;
  document.body.classList.remove('nav-open');
  const v = $('#view');
  v.innerHTML = loadingHtml();
  try { await PAGES[id](v); }
  catch (e) { v.innerHTML = errBox(e); }
}

/* ---------------- boot & login ---------------- */
async function boot() {
  document.title = CONFIG.APP_NAME;
  $('#loginTitle').textContent = $('#appName').textContent = CONFIG.APP_NAME;
  $('#loginOrg').textContent = $('#appOrg').textContent = CONFIG.ORG_NAME || '';
  if (Auth.get()) await showApp(); else showLogin();
}

function showLogin() {
  $('#app').classList.add('hidden');
  $('#login').classList.remove('hidden');
  $('#loginForm [name=username]').focus();
}

async function showApp() {
  $('#login').classList.add('hidden');
  $('#app').classList.remove('hidden');
  $('#view').innerHTML = loadingHtml();
  try {
    const d = await API.call('init');
    App.user = d.user; App.cabang = d.cabang; App.models = d.models;
  } catch (e) {
    if (e.code === 'AUTH') { Auth.clear(); showLogin(); return; }
    $('#view').innerHTML = errBox(e);
    return;
  }
  $('#userName').textContent = App.user.nama;
  $('#userRole').textContent = (isUtama() ? 'Admin Utama' : 'Admin Cabang') + ' · ' + cabNama(App.user.cabang);
  renderNav();
  route();
}

$('#loginForm').addEventListener('submit', async e => {
  e.preventDefault();
  const f = e.target, btn = f.querySelector('button');
  $('#loginError').textContent = '';
  btn.disabled = true; btn.textContent = 'Memeriksa…';
  try {
    const r = await API.call('login', { username: f.username.value.trim(), hash: await sha256(f.password.value) });
    Auth.set(r);
    f.password.value = '';
    await showApp();
  } catch (err) {
    $('#loginError').textContent = err.message;
  } finally {
    btn.disabled = false; btn.textContent = 'Masuk';
  }
});

$('#btnLogout').onclick = async () => {
  try { await API.call('logout'); } catch (e) {}
  Auth.clear(); App.user = null; location.hash = ''; showLogin();
};
$('#btnMenu').onclick = () => document.body.classList.toggle('nav-open');
$('#scrim').onclick = () => document.body.classList.remove('nav-open');
$('#modal').addEventListener('click', e => { if (e.target.id === 'modal') closeModal(); });
window.addEventListener('hashchange', () => { if (App.user) route(); });

/* ============================================================
   HALAMAN
   ============================================================ */

/* ---------------- Dashboard ---------------- */
async function pageDashboard(v) {
  const d = await API.call('dashboard');
  const tbl = (obj, head, label) => {
    const rows = Object.keys(obj).sort((a, b) => obj[b] - obj[a]);
    return `<div class="table-wrap"><table><thead><tr><th>${head}</th><th class="num">Unit</th></tr></thead><tbody>
      ${rows.length ? rows.map(k => `<tr><td>${esc(label(k))}</td><td class="num">${fmt(obj[k])}</td></tr>`).join('')
        : '<tr><td colspan="2" class="empty">Belum ada stok</td></tr>'}</tbody></table></div>`;
  };
  v.innerHTML = `
    <div class="stats">
      <div class="stat hi"><small>Stok di gudang</small><b>${fmt(d.stok)}</b></div>
      <div class="stat"><small>Masuk hari ini</small><b>${fmt(d.hariIni.MASUK)}</b></div>
      <div class="stat"><small>Keluar hari ini</small><b>${fmt(d.hariIni.KELUAR)}</b></div>
      <div class="stat"><small>Retur hari ini</small><b>${fmt(d.hariIni.RETUR)}</b></div>
      <div class="stat"><small>Total sudah keluar</small><b>${fmt(d.keluar)}</b></div>
    </div>
    <div class="grid2" style="margin-bottom:18px">
      <div class="card"><h3>Stok per cabang</h3>${tbl(d.perCabang, 'Cabang', cabNama)}</div>
      <div class="card"><h3>Stok per model</h3>${tbl(d.perModel, 'Model', k => k)}</div>
    </div>
    <div class="card"><h3>Transaksi terakhir <a class="btn sm" href="#riwayat">Lihat semua</a></h3>
      ${trxTable(d.terakhir)}</div>`;
}

function trxTable(rows) {
  return `<div class="table-wrap"><table><thead><tr>
    <th>Waktu</th><th>Jenis</th><th>IMEI</th><th>SN</th><th>MAC</th><th>Model</th><th>Cabang</th><th>Penerima</th><th>Referensi</th><th>Keterangan</th><th>Oleh</th>
    </tr></thead><tbody>${rows.length ? rows.map(t => `<tr>
      <td class="mono">${esc(t.waktu)}</td><td>${badge(t.jenis)}</td>
      <td class="mono">${esc(t.imei)}</td><td class="mono">${esc(t.sn)}</td><td class="mono">${esc(t.mac)}</td>
      <td>${esc(t.model)}</td>
      <td>${t.jenis === 'MUTASI' ? esc(t.cabang_asal) + ' → ' + esc(t.cabang) : esc(t.cabang)}</td>
      <td>${esc(t.penerima)}</td><td>${esc(t.referensi)}</td><td>${esc(t.keterangan)}</td><td>${esc(t.oleh)}</td>
    </tr>`).join('') : '<tr><td colspan="11" class="empty">Belum ada transaksi</td></tr>'}</tbody></table></div>`;
}

/* ---------------- Barang Masuk ---------------- */
async function pageMasuk(v) {
  const list = [];
  v.innerHTML = `
  <div class="card">
    <h3>Informasi penerimaan</h3>
    <div class="form-grid">
      ${isUtama()
        ? `<label>Masuk ke cabang<select id="mCab">${opsiCabang(App.user.cabang)}</select></label>`
        : `<label>Cabang<input value="${esc(cabNama(App.user.cabang))}" disabled></label>`}
      <label>Model / tipe<input id="mModel" list="dlModel" placeholder="mis. Modem FWA MyRepublic">
        <datalist id="dlModel">${App.models.map(m => `<option value="${esc(m)}">`).join('')}</datalist></label>
      <label>No. referensi<input id="mRef" placeholder="No. DO / surat jalan"></label>
      <label>Keterangan<input id="mKet" placeholder="Opsional"></label>
    </div>
  </div>
  <div class="card">
    <h3>Daftar unit</h3>
    <div class="tabs"><button data-t="scan" class="on">Scan satu per satu</button><button data-t="paste">Tempel massal</button><button data-t="file">Impor file Excel</button></div>
    <div id="tScan">
      <div class="row scan-row">
        <label>IMEI<input id="sImei" inputmode="numeric" placeholder="15 digit" autocomplete="off"></label>
        <label>SN<input id="sSn" autocomplete="off"></label>
        <label>MAC address<input id="sMac" placeholder="AA:BB:CC:DD:EE:FF" autocomplete="off"></label>
        <button class="btn" id="sAdd">Tambah</button>
      </div>
      <p class="hint">Tekan Enter di tiap kolom untuk pindah ke kolom berikutnya — cocok untuk scanner barcode.</p>
    </div>
    <div id="tPaste" class="hidden">
      <div class="format-box">
        <strong>Format:</strong> satu unit per baris, urutan <b>IMEI · SN · MAC</b>, dipisah Tab (hasil copy dari Excel), koma, titik koma, atau spasi.
        <pre>356789012345671	ZTEGC1234567	AA:BB:CC:11:22:33
356789012345672,ZTEGC1234568,AABBCC112234
356789012345673;ZTEGC1234569;aa-bb-cc-11-22-35</pre>
        Baris judul (IMEI, SN, MAC) otomatis dilewati. Kalau copy dari Excel, pastikan kolom IMEI berformat <b>Teks</b>,
        bukan Angka, supaya tidak tersalin sebagai 3,57E+14.
      </div>
      <textarea id="pText" placeholder="Tempel di sini…"></textarea>
      <div class="row" style="margin-top:8px"><button class="btn" id="pAdd">Tambahkan ke daftar</button></div>
    </div>
    <div id="tFile" class="hidden">
      <div class="format-box">
        <strong>Langkah:</strong> 1) Unduh template · 2) Isi kolom IMEI, SN, MAC di sheet <b>Data</b> (satu baris = satu unit) ·
        3) Simpan lalu pilih filenya di bawah. Model dan cabang dipilih di atas. File .xlsx, .xls, dan .csv juga bisa,
        asalkan baris judulnya memuat kata IMEI, SN, dan MAC.
      </div>
      <div class="row">
        <button class="btn" id="fTpl">⇩ Unduh template Excel</button>
        <label class="btn primary file-btn">Pilih file Excel…<input type="file" id="fFile" accept=".xlsx,.xls,.csv" hidden></label>
      </div>
    </div>
    <div id="mMsg" style="margin-top:14px"></div>
    <div class="table-wrap"><table>
      <thead><tr><th>#</th><th>IMEI</th><th>SN</th><th>MAC</th><th>Cek</th><th></th></tr></thead>
      <tbody id="mBody"></tbody></table></div>
    <div class="summary">
      <span><b id="mCount">0</b> unit dalam daftar <span id="mErr" style="color:var(--bad)"></span></span>
      <div class="row"><button class="btn ghost" id="mClear">Kosongkan</button>
        <button class="btn primary" id="mSave">Simpan barang masuk</button></div>
    </div>
  </div>`;

  const check = () => {
    const cnt = {};
    list.forEach(x => ['i' + x.imei, 's' + x.sn, 'm' + x.mac].forEach(k => cnt[k] = (cnt[k] || 0) + 1));
    list.forEach(x => {
      const e = [];
      if (x.sci) e.push('IMEI tersalin sebagai format ilmiah (E+14) — ubah kolom Excel ke Teks lalu salin ulang');
      else if (x.imei.length !== 15) e.push('IMEI harus 15 digit');
      if (!x.sn) e.push('SN kosong');
      if (!x.mac) e.push('MAC tidak valid');
      if (x.imei && cnt['i' + x.imei] > 1) e.push('IMEI dobel');
      if (x.sn && cnt['s' + x.sn] > 1) e.push('SN dobel');
      if (x.mac && cnt['m' + x.mac] > 1) e.push('MAC dobel');
      if (x.srv) e.push(x.srv);
      x.err = e.join(', ');
    });
  };
  const render = () => {
    check();
    $('#mBody').innerHTML = list.length ? list.map((x, i) => `
      <tr class="${x.err ? 'err' : ''}"><td>${i + 1}</td>
        <td class="mono">${esc(x.imei)}</td><td class="mono">${esc(x.sn)}</td><td class="mono">${esc(x.mac || x.rawMac)}</td>
        <td>${x.err ? `<small style="color:var(--bad)">${esc(x.err)}</small>` : '<span class="badge b-on">OK</span>'}</td>
        <td><button class="btn ghost sm" data-del="${i}" title="Hapus">✕</button></td></tr>`).join('')
      : '<tr><td colspan="6" class="empty">Belum ada unit. Scan atau tempel data di atas.</td></tr>';
    const nErr = list.filter(x => x.err).length;
    $('#mCount').textContent = list.length;
    $('#mErr').textContent = nErr ? `· ${nErr} bermasalah` : '';
  };
  const add = (imei, sn, mac) => {
    const sci = isSci(imei);
    list.push({ imei: sci ? String(imei).trim() : normImei(imei), sn: normSn(sn), mac: normMac(mac), rawMac: mac || '', sci });
  };

  $$('.tabs button', v).forEach(b => b.onclick = () => {
    $$('.tabs button', v).forEach(x => x.classList.toggle('on', x === b));
    $('#tScan').classList.toggle('hidden', b.dataset.t !== 'scan');
    $('#tPaste').classList.toggle('hidden', b.dataset.t !== 'paste');
    $('#tFile').classList.toggle('hidden', b.dataset.t !== 'file');
  });

  $('#fTpl').onclick = downloadTemplateMasuk;
  $('#fFile').onchange = async e => {
    const file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    try {
      const rows = await readMasukFile(file);
      if (!rows.length) { toast('File tidak berisi data', true); return; }
      rows.forEach(r => add(r[0], r[1], r[2]));
      render();
      const nErr = list.filter(x => x.err).length;
      toast(`${rows.length} unit dibaca dari ${file.name}` + (nErr ? ` · ${nErr} bermasalah` : ''), nErr > 0);
    } catch (err) {
      $('#mMsg').innerHTML = errBox({ message: 'Gagal membaca file: ' + err.message });
    }
  };

  const addScan = () => {
    const i = $('#sImei').value, s = $('#sSn').value, m = $('#sMac').value;
    if (!i && !s && !m) return;
    add(i, s, m);
    $('#sImei').value = $('#sSn').value = $('#sMac').value = '';
    render();
    $('#sImei').focus();
  };
  $('#sImei').onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); $('#sSn').focus(); } };
  $('#sSn').onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); $('#sMac').focus(); } };
  $('#sMac').onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); addScan(); } };
  $('#sAdd').onclick = addScan;

  $('#pAdd').onclick = () => {
    let n = 0;
    $('#pText').value.split(/\r?\n/).forEach(line => {
      const p = parseLine(line);
      if (!p) return;
      if (/^imei$/i.test(p[0])) return; // lewati baris judul
      add(p[0], p[1], p[2]); n++;
    });
    $('#pText').value = '';
    render();
    toast(n + ' baris ditambahkan');
  };
  $('#mBody').onclick = e => {
    const b = e.target.closest('[data-del]');
    if (b) { list.splice(+b.dataset.del, 1); render(); }
  };
  $('#mClear').onclick = () => { if (!list.length || confirm('Kosongkan daftar?')) { list.length = 0; $('#mMsg').innerHTML = ''; render(); } };

  $('#mSave').onclick = async () => {
    const msg = $('#mMsg');
    msg.innerHTML = '';
    list.forEach(x => x.srv = '');
    render();
    const model = $('#mModel').value.trim();
    if (!model) { toast('Model/tipe wajib diisi', true); $('#mModel').focus(); return; }
    if (!list.length) { toast('Daftar unit masih kosong', true); return; }
    if (list.some(x => x.err)) { toast('Perbaiki baris yang bermasalah dulu', true); return; }
    const base = {
      cabang: isUtama() ? $('#mCab').value : App.user.cabang,
      model, referensi: $('#mRef').value.trim(), keterangan: $('#mKet').value.trim()
    };
    const btn = $('#mSave');
    btn.disabled = true;
    let masuk = 0, retur = 0;
    try {
      while (list.length) {
        btn.textContent = `Menyimpan… (${list.length} tersisa)`;
        const chunk = list.slice(0, CHUNK_MASUK);
        const r = await API.call('masuk', Object.assign({}, base, { items: chunk.map(x => [x.imei, x.sn, x.mac]) }));
        masuk += r.masuk; retur += r.retur;
        list.splice(0, chunk.length);
        render();
      }
      if (!App.models.includes(model)) App.models.push(model);
      msg.innerHTML = okBox(`<strong>Tersimpan.</strong> ${masuk} unit baru${retur ? `, ${retur} unit retur` : ''} masuk ke ${esc(cabNama(base.cabang))}.`);
      toast('Barang masuk tersimpan');
      $('#mRef').value = $('#mKet').value = '';
    } catch (e) {
      (e.errors || []).forEach(s => {
        const m = s.match(/^IMEI (\d{15}): (.*)$/);
        if (m) { const x = list.find(i => i.imei === m[1]); if (x) x.srv = m[2]; }
      });
      render();
      msg.innerHTML = errBox(e, masuk + retur ? `<div>${masuk + retur} unit sebelumnya sudah tersimpan dan dihapus dari daftar.</div>` : '');
    } finally {
      btn.disabled = false; btn.textContent = 'Simpan barang masuk';
    }
  };
  render();
  $('#sImei').focus();
}

/* ---------------- Barang Keluar & Mutasi (pilih unit dari stok) ---------------- */
const pageKeluar = v => pagePilih(v, 'keluar');
const pageMutasi = v => pagePilih(v, 'mutasi');

async function pagePilih(v, mode) {
  const stok = await API.call('stok', { status: 'GUDANG' });
  const ix = { imei: {}, sn: {}, mac: {} };
  stok.forEach(u => { ix.imei[u.imei] = u; ix.sn[u.sn] = u; ix.mac[u.mac] = u; });
  const find = k => {
    k = String(k || '').trim();
    if (!k) return null;
    const d = normImei(k);
    if (d.length === 15 && /^[\d\s-]+$/.test(k) && ix.imei[d]) return ix.imei[d];
    return ix.sn[normSn(k)] || ix.mac[normMac(k)] || null;
  };
  const list = [];
  const isK = mode === 'keluar';

  v.innerHTML = `
  <div class="card">
    <h3>${isK ? 'Informasi pengeluaran' : 'Informasi mutasi'}</h3>
    <div class="form-grid">
      ${isK
        ? `<label>Penerima<input id="kPen" placeholder="Nama pelanggan / teknisi"></label>
           <label>No. referensi<input id="kRef" placeholder="No. WO / order"></label>`
        : `<label>Cabang tujuan<select id="kTuj">${opsiCabang('', '— pilih cabang —')}</select></label>
           <label>No. referensi<input id="kRef" placeholder="No. surat jalan"></label>`}
      <label>Keterangan<input id="kKet" placeholder="Opsional"></label>
    </div>
  </div>
  <div class="card">
    <h3>Unit ${isK ? 'keluar' : 'dipindahkan'} <small class="muted">${fmt(stok.length)} unit tersedia di stok${isUtama() ? ' semua cabang' : ''}</small></h3>
    <div class="tabs"><button data-t="scan" class="on">Scan</button><button data-t="paste">Tempel massal</button></div>
    <div id="tScan" class="row scan-row">
      <label>Scan IMEI / SN / MAC<input id="kScan" autocomplete="off" placeholder="Scan atau ketik lalu Enter"></label>
      <button class="btn" id="kAdd">Tambah</button>
    </div>
    <div id="tPaste" class="hidden">
      <textarea id="kText" placeholder="Satu IMEI / SN / MAC per baris"></textarea>
      <div class="row" style="margin-top:8px"><button class="btn" id="kPaste">Tambahkan ke daftar</button></div>
    </div>
    <div id="kMsg" style="margin-top:14px"></div>
    <div class="table-wrap"><table>
      <thead><tr><th>#</th><th>IMEI</th><th>SN</th><th>MAC</th><th>Model</th><th>Cabang</th><th></th></tr></thead>
      <tbody id="kBody"></tbody></table></div>
    <div class="summary">
      <span><b id="kCount">0</b> unit dipilih</span>
      <div class="row"><button class="btn ghost" id="kClear">Kosongkan</button>
        <button class="btn primary" id="kSave">${isK ? 'Simpan barang keluar' : 'Simpan mutasi'}</button></div>
    </div>
  </div>`;

  const render = () => {
    $('#kBody').innerHTML = list.length ? list.map((u, i) => `<tr>
      <td>${i + 1}</td><td class="mono">${esc(u.imei)}</td><td class="mono">${esc(u.sn)}</td><td class="mono">${esc(u.mac)}</td>
      <td>${esc(u.model)}</td><td>${esc(cabNama(u.cabang))}</td>
      <td><button class="btn ghost sm" data-del="${i}">✕</button></td></tr>`).join('')
      : '<tr><td colspan="7" class="empty">Belum ada unit dipilih.</td></tr>';
    $('#kCount').textContent = list.length;
  };
  const addKey = k => {
    const u = find(k);
    if (!u) return 'tidak ada di stok gudang';
    if (list.includes(u)) return 'sudah ada di daftar';
    list.push(u);
    return '';
  };

  $$('.tabs button', v).forEach(b => b.onclick = () => {
    $$('.tabs button', v).forEach(x => x.classList.toggle('on', x === b));
    $('#tScan').classList.toggle('hidden', b.dataset.t !== 'scan');
    $('#tPaste').classList.toggle('hidden', b.dataset.t !== 'paste');
  });
  const scan = () => {
    const k = $('#kScan').value.trim();
    if (!k) return;
    const e = addKey(k);
    if (e) toast(k + ': ' + e, true);
    $('#kScan').value = '';
    $('#kScan').focus();
    render();
  };
  $('#kScan').onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); scan(); } };
  $('#kAdd').onclick = scan;
  $('#kPaste').onclick = () => {
    const gagal = [];
    let n = 0;
    $('#kText').value.split(/\r?\n/).map(s => s.trim()).filter(Boolean).forEach(k => {
      const e = addKey(k);
      if (e) gagal.push(k + ': ' + e); else n++;
    });
    $('#kText').value = gagal.map(g => g.split(':')[0]).join('\n');
    $('#kMsg').innerHTML = gagal.length ? errBox({ message: `${gagal.length} baris tidak bisa ditambahkan (dibiarkan di kotak tempel)`, errors: gagal }) : '';
    render();
    toast(n + ' unit ditambahkan');
  };
  $('#kBody').onclick = e => { const b = e.target.closest('[data-del]'); if (b) { list.splice(+b.dataset.del, 1); render(); } };
  $('#kClear').onclick = () => { list.length = 0; $('#kMsg').innerHTML = ''; render(); };

  $('#kSave').onclick = async () => {
    const msg = $('#kMsg');
    msg.innerHTML = '';
    if (!list.length) { toast('Belum ada unit dipilih', true); return; }
    const base = { referensi: $('#kRef').value.trim(), keterangan: $('#kKet').value.trim() };
    if (isK) {
      base.penerima = $('#kPen').value.trim();
      if (!base.penerima) { toast('Penerima wajib diisi', true); $('#kPen').focus(); return; }
    } else {
      base.tujuan = $('#kTuj').value;
      if (!base.tujuan) { toast('Pilih cabang tujuan', true); return; }
    }
    const btn = $('#kSave');
    btn.disabled = true;
    let done = 0;
    try {
      while (list.length) {
        btn.textContent = `Menyimpan… (${list.length} tersisa)`;
        const chunk = list.slice(0, CHUNK_KELUAR);
        await API.call(mode, Object.assign({}, base, { keys: chunk.map(u => u.imei) }));
        chunk.forEach(u => { delete ix.imei[u.imei]; delete ix.sn[u.sn]; delete ix.mac[u.mac]; });
        done += chunk.length;
        list.splice(0, chunk.length);
        render();
      }
      msg.innerHTML = okBox(isK
        ? `<strong>Tersimpan.</strong> ${done} unit keluar ke ${esc(base.penerima)}.`
        : `<strong>Tersimpan.</strong> ${done} unit dipindahkan ke ${esc(cabNama(base.tujuan))}.`);
      toast('Transaksi tersimpan');
    } catch (e) {
      msg.innerHTML = errBox(e, done ? `<div>${done} unit sebelumnya sudah tersimpan.</div>` : '');
    } finally {
      btn.disabled = false; btn.textContent = isK ? 'Simpan barang keluar' : 'Simpan mutasi';
    }
  };
  render();
  $('#kScan').focus();
}

/* ---------------- Stok ---------------- */
async function pageStok(v) {
  const all = await API.call('stok', {});
  const models = [...new Set(all.map(u => u.model))].sort();
  const PAGE = 100;
  let page = 0, rows = [];
  v.innerHTML = `
  <div class="card">
    <div class="toolbar">
      ${isUtama() ? `<label>Cabang<select id="fCab">${opsiCabang('', 'Semua cabang', false)}</select></label>` : ''}
      <label>Status<select id="fSt"><option value="GUDANG">Di gudang</option><option value="KELUAR">Sudah keluar</option><option value="">Semua</option></select></label>
      <label>Model<select id="fMod"><option value="">Semua model</option>${models.map(m => `<option>${esc(m)}</option>`).join('')}</select></label>
      <label class="grow">Cari<input id="fQ" placeholder="IMEI, SN, MAC, penerima, referensi…"></label>
      <button class="btn" id="fExp">Ekspor Excel</button>
    </div>
    <div class="table-wrap"><table>
      <thead><tr><th>IMEI</th><th>SN</th><th>MAC</th><th>Model</th><th>Status</th><th>Cabang</th><th>Tgl masuk</th><th>Tgl keluar</th><th>Penerima</th><th>Referensi</th></tr></thead>
      <tbody id="sBody"></tbody></table></div>
    <div class="summary"><span id="sInfo"></span>
      <div class="row"><button class="btn sm" id="pPrev">‹ Sebelumnya</button><button class="btn sm" id="pNext">Berikutnya ›</button></div></div>
  </div>`;

  const apply = () => {
    const cab = isUtama() ? $('#fCab').value : '';
    const st = $('#fSt').value, mod = $('#fMod').value, q = $('#fQ').value.trim().toUpperCase();
    rows = all.filter(u => (!cab || u.cabang === cab) && (!st || u.status === st) && (!mod || u.model === mod) &&
      (!q || [u.imei, u.sn, u.mac, u.mac.replace(/:/g, ''), u.penerima, u.referensi].join(' ').toUpperCase().includes(q)));
    page = 0; draw();
  };
  const draw = () => {
    const slice = rows.slice(page * PAGE, page * PAGE + PAGE);
    $('#sBody').innerHTML = slice.length ? slice.map(u => `<tr>
      <td class="mono"><a href="#lacak" data-q="${esc(u.imei)}">${esc(u.imei)}</a></td><td class="mono">${esc(u.sn)}</td><td class="mono">${esc(u.mac)}</td>
      <td>${esc(u.model)}</td><td>${badge(u.status)}</td><td>${esc(cabNama(u.cabang))}</td>
      <td class="mono">${esc(u.tgl_masuk.slice(0, 16))}</td><td class="mono">${esc(u.tgl_keluar.slice(0, 16))}</td>
      <td>${esc(u.penerima)}</td><td>${esc(u.referensi)}</td></tr>`).join('')
      : '<tr><td colspan="10" class="empty">Tidak ada data</td></tr>';
    const from = rows.length ? page * PAGE + 1 : 0;
    $('#sInfo').innerHTML = `Menampilkan <b>${fmt(from)}–${fmt(Math.min(rows.length, (page + 1) * PAGE))}</b> dari <b>${fmt(rows.length)}</b> unit`;
    $('#pPrev').disabled = page === 0;
    $('#pNext').disabled = (page + 1) * PAGE >= rows.length;
  };
  $$('#fCab, #fSt, #fMod', v).forEach(el => el.onchange = apply);
  $('#fQ').oninput = apply;
  $('#pPrev').onclick = () => { page--; draw(); };
  $('#pNext').onclick = () => { page++; draw(); };
  $('#sBody').onclick = e => { const a = e.target.closest('[data-q]'); if (a) sessionStorage.setItem('vg_lacak', a.dataset.q); };
  $('#fExp').onclick = () => exportExcel('stok-' + today(), [
    ['imei', 'IMEI', 18], ['sn', 'SN', 18], ['mac', 'MAC', 19], ['model', 'Model', 24], ['status', 'Status', 10],
    ['cabang', 'Cabang', 10], ['tgl_masuk', 'Tgl Masuk', 19], ['tgl_keluar', 'Tgl Keluar', 19],
    ['penerima', 'Penerima', 22], ['referensi', 'Referensi', 18], ['keterangan', 'Keterangan', 24]
  ], rows);
  apply();
}

/* ---------------- Riwayat ---------------- */
async function pageRiwayat(v) {
  let data = [], rows = [];
  v.innerHTML = `
  <div class="card">
    <div class="toolbar">
      <label>Dari<input type="date" id="rDari" value="${daysAgo(30)}"></label>
      <label>Sampai<input type="date" id="rSampai" value="${today()}"></label>
      <button class="btn" id="rLoad">Tampilkan</button>
      <label>Jenis<select id="rJenis"><option value="">Semua</option>${['MASUK', 'KELUAR', 'RETUR', 'MUTASI', 'HAPUS'].map(j => `<option>${j}</option>`).join('')}</select></label>
      ${isUtama() ? `<label>Cabang<select id="rCab">${opsiCabang('', 'Semua cabang', false)}</select></label>` : ''}
      <label class="grow">Cari<input id="rQ" placeholder="IMEI, SN, MAC, penerima, referensi…"></label>
      <button class="btn" id="rExp">Ekspor Excel</button>
    </div>
    <div id="rOut">${loadingHtml()}</div>
  </div>`;
  const apply = () => {
    const j = $('#rJenis').value, cab = isUtama() ? $('#rCab').value : '', q = $('#rQ').value.trim().toUpperCase();
    rows = data.filter(t => (!j || t.jenis === j) && (!cab || t.cabang === cab || t.cabang_asal === cab) &&
      (!q || [t.imei, t.sn, t.mac, t.mac.replace(/:/g, ''), t.penerima, t.referensi, t.keterangan, t.oleh].join(' ').toUpperCase().includes(q)));
    const shown = rows.slice(0, 500);
    $('#rOut').innerHTML = trxTable(shown) +
      `<div class="summary"><span><b>${fmt(rows.length)}</b> transaksi${rows.length > 500 ? ' (500 teratas ditampilkan; ekspor untuk semua)' : ''}</span></div>`;
  };
  const load = async () => {
    $('#rOut').innerHTML = loadingHtml();
    try {
      data = await API.call('riwayat', { dari: $('#rDari').value, sampai: $('#rSampai').value });
      apply();
    } catch (e) { $('#rOut').innerHTML = errBox(e); }
  };
  $('#rLoad').onclick = load;
  $$('#rJenis, #rCab', v).forEach(el => el.onchange = apply);
  $('#rQ').oninput = () => data.length && apply();
  $('#rExp').onclick = () => exportExcel(`riwayat-${$('#rDari').value}_${$('#rSampai').value}`, [
    ['id', 'ID Transaksi', 22], ['waktu', 'Waktu', 19], ['jenis', 'Jenis', 9], ['imei', 'IMEI', 18], ['sn', 'SN', 18],
    ['mac', 'MAC', 19], ['model', 'Model', 24], ['cabang_asal', 'Cabang Asal', 12], ['cabang', 'Cabang', 10],
    ['penerima', 'Penerima', 22], ['referensi', 'Referensi', 18], ['keterangan', 'Keterangan', 24], ['oleh', 'Oleh', 12]
  ], rows);
  load();
}

/* ---------------- Lacak unit ---------------- */
async function pageLacak(v) {
  v.innerHTML = `
  <div class="card">
    <div class="row">
      <label>IMEI / SN / MAC<input id="lQ" autocomplete="off" placeholder="Scan atau ketik lalu Enter"></label>
      <button class="btn primary" id="lGo">Lacak</button>
    </div>
  </div>
  <div id="lOut"></div>`;
  const go = async () => {
    const q = $('#lQ').value.trim();
    if (!q) return;
    $('#lOut').innerHTML = loadingHtml('Mencari…');
    try {
      const r = await API.call('lacak', { q });
      const u = r.unit;
      $('#lOut').innerHTML = (u ? `
        <div class="card"><h3>Unit ditemukan ${badge(u.status)}
          ${isUtama() && u.status === 'GUDANG' ? `<button class="btn danger sm" id="lDel">Hapus (salah input)</button>` : ''}</h3>
          <div class="unit-card">
            <div><small>IMEI</small><span class="mono">${esc(u.imei)}</span></div>
            <div><small>SN</small><span class="mono">${esc(u.sn)}</span></div>
            <div><small>MAC</small><span class="mono">${esc(u.mac)}</span></div>
            <div><small>Model</small>${esc(u.model)}</div>
            <div><small>Cabang</small>${esc(cabNama(u.cabang))}</div>
            <div><small>Tgl masuk</small>${esc(u.tgl_masuk)}</div>
            ${u.status === 'KELUAR' ? `<div><small>Tgl keluar</small>${esc(u.tgl_keluar)}</div><div><small>Penerima</small>${esc(u.penerima)}</div>` : ''}
            <div><small>Referensi</small>${esc(u.referensi || '-')}</div>
          </div></div>`
        : `<div class="alert bad">Unit tidak ditemukan di stok${r.riwayat.length ? ' (ada riwayat lama di bawah)' : ''}.</div>`) +
        (r.riwayat.length ? `<div class="card"><h3>Riwayat</h3><ul class="timeline">${r.riwayat.map(t => `
          <li><span class="mono">${esc(t.waktu)}</span><span>${badge(t.jenis)}</span>
          <span>${esc(t.jenis === 'MUTASI' ? cabNama(t.cabang_asal) + ' → ' + cabNama(t.cabang) : cabNama(t.cabang))}
          ${t.penerima ? ' · ' + esc(t.penerima) : ''}${t.referensi ? ' · Ref ' + esc(t.referensi) : ''}${t.keterangan ? ' · ' + esc(t.keterangan) : ''}
          <small class="muted"> — ${esc(t.oleh)}</small></span></li>`).join('')}</ul></div>` : '');
      const del = $('#lDel');
      if (del) del.onclick = () => hapusUnit(u, go);
    } catch (e) { $('#lOut').innerHTML = errBox(e); }
  };
  $('#lQ').onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); go(); } };
  $('#lGo').onclick = go;
  const pre = sessionStorage.getItem('vg_lacak');
  if (pre) { sessionStorage.removeItem('vg_lacak'); $('#lQ').value = pre; go(); }
  $('#lQ').focus();
}

function hapusUnit(u, after) {
  modal(`<h3>Hapus unit ${esc(u.imei)}?</h3>
    <p class="muted" style="margin:0">Gunakan hanya untuk data salah input. Penghapusan tetap tercatat di riwayat.</p>
    <label>Alasan<input id="hAl" placeholder="mis. salah scan IMEI"></label>
    <p class="form-error" id="hErr"></p>
    <div class="modal-actions"><button class="btn" id="hNo">Batal</button><button class="btn primary" id="hYes">Hapus</button></div>`);
  $('#hNo').onclick = closeModal;
  $('#hYes').onclick = async () => {
    try {
      await API.call('hapusUnit', { keys: [u.imei], alasan: $('#hAl').value.trim() });
      closeModal(); toast('Unit dihapus'); after();
    } catch (e) { $('#hErr').textContent = e.message; }
  };
}

/* ---------------- Cabang ---------------- */
async function pageCabang(v) {
  App.cabang = await API.call('cabang');
  v.innerHTML = `<div class="card"><h3>Daftar cabang <button class="btn primary sm" id="cNew">+ Tambah cabang</button></h3>
    <div class="table-wrap"><table><thead><tr><th>Kode</th><th>Nama</th><th>Alamat</th><th>Status</th><th></th></tr></thead>
    <tbody>${App.cabang.map((c, i) => `<tr><td class="mono">${esc(c.kode)}</td><td>${esc(c.nama)}</td><td>${esc(c.alamat)}</td>
      <td>${c.aktif === '1' ? '<span class="badge b-on">AKTIF</span>' : '<span class="badge b-off">NONAKTIF</span>'}</td>
      <td><button class="btn sm" data-i="${i}">Ubah</button></td></tr>`).join('')}</tbody></table></div></div>`;
  const form = c => {
    const baru = !c;
    c = c || { kode: '', nama: '', alamat: '', aktif: '1' };
    modal(`<h3>${baru ? 'Tambah' : 'Ubah'} cabang</h3>
      <label>Kode<input id="cKode" value="${esc(c.kode)}" ${baru ? '' : 'disabled'} placeholder="mis. GTO-KOTA"></label>
      <label>Nama<input id="cNama" value="${esc(c.nama)}"></label>
      <label>Alamat<input id="cAlm" value="${esc(c.alamat)}"></label>
      <label>Status<select id="cAkt"><option value="1">Aktif</option><option value="0" ${c.aktif === '0' ? 'selected' : ''}>Nonaktif</option></select></label>
      <p class="form-error" id="cErr"></p>
      <div class="modal-actions"><button class="btn" id="cNo">Batal</button><button class="btn primary" id="cYes">Simpan</button></div>`);
    $('#cNo').onclick = closeModal;
    $('#cYes').onclick = async () => {
      try {
        App.cabang = await API.call('simpanCabang', { baru, kode: $('#cKode').value, nama: $('#cNama').value, alamat: $('#cAlm').value, aktif: $('#cAkt').value });
        closeModal(); toast('Cabang tersimpan'); pageCabang(v);
      } catch (e) { $('#cErr').textContent = e.message; }
    };
  };
  $('#cNew').onclick = () => form(null);
  $$('[data-i]', v).forEach(b => b.onclick = () => form(App.cabang[+b.dataset.i]));
}

/* ---------------- Pengguna ---------------- */
async function pagePengguna(v) {
  const users = await API.call('users');
  v.innerHTML = `<div class="card"><h3>Daftar pengguna <button class="btn primary sm" id="uNew">+ Tambah pengguna</button></h3>
    <div class="table-wrap"><table><thead><tr><th>Username</th><th>Nama</th><th>Role</th><th>Cabang</th><th>Status</th><th></th></tr></thead>
    <tbody>${users.map((u, i) => `<tr><td class="mono">${esc(u.username)}</td><td>${esc(u.nama)}</td>
      <td>${u.role === 'UTAMA' ? 'Admin Utama' : 'Admin Cabang'}</td><td>${esc(cabNama(u.cabang))}</td>
      <td>${u.aktif === '1' ? '<span class="badge b-on">AKTIF</span>' : '<span class="badge b-off">NONAKTIF</span>'}</td>
      <td><button class="btn sm" data-i="${i}">Ubah</button></td></tr>`).join('')}</tbody></table></div></div>`;
  const form = u => {
    const baru = !u;
    u = u || { username: '', nama: '', role: 'CABANG', cabang: '', aktif: '1' };
    modal(`<h3>${baru ? 'Tambah' : 'Ubah'} pengguna</h3>
      <label>Username<input id="uUser" value="${esc(u.username)}" ${baru ? '' : 'disabled'} autocomplete="off"></label>
      <label>Nama lengkap<input id="uNama" value="${esc(u.nama)}"></label>
      <label>Role<select id="uRole"><option value="CABANG">Admin Cabang</option><option value="UTAMA" ${u.role === 'UTAMA' ? 'selected' : ''}>Admin Utama</option></select></label>
      <label>Cabang<select id="uCab">${opsiCabang(u.cabang)}</select></label>
      <label>${baru ? 'Password' : 'Reset password (kosongkan jika tidak diubah)'}<input id="uPass" type="password" autocomplete="new-password"></label>
      <label>Status<select id="uAkt"><option value="1">Aktif</option><option value="0" ${u.aktif === '0' ? 'selected' : ''}>Nonaktif</option></select></label>
      <p class="form-error" id="uErr"></p>
      <div class="modal-actions"><button class="btn" id="uNo">Batal</button><button class="btn primary" id="uYes">Simpan</button></div>`);
    $('#uNo').onclick = closeModal;
    $('#uYes').onclick = async () => {
      const pass = $('#uPass').value;
      if (pass && pass.length < 6) { $('#uErr').textContent = 'Password minimal 6 karakter.'; return; }
      try {
        await API.call('simpanUser', {
          baru, username: $('#uUser').value, nama: $('#uNama').value, role: $('#uRole').value,
          cabang: $('#uCab').value, aktif: $('#uAkt').value, hash: pass ? await sha256(pass) : ''
        });
        closeModal(); toast('Pengguna tersimpan'); pagePengguna(v);
      } catch (e) { $('#uErr').textContent = e.message; }
    };
  };
  $('#uNew').onclick = () => form(null);
  $$('[data-i]', v).forEach(b => b.onclick = () => form(users[+b.dataset.i]));
}

/* ---------------- Ganti password ---------------- */
async function pagePassword(v) {
  v.innerHTML = `<div class="card" style="max-width:420px"><form id="pwForm" style="display:flex;flex-direction:column;gap:14px">
    <label>Password lama<input type="password" name="lama" required autocomplete="current-password"></label>
    <label>Password baru<input type="password" name="baru" required minlength="6" autocomplete="new-password"></label>
    <label>Ulangi password baru<input type="password" name="ulang" required autocomplete="new-password"></label>
    <p class="form-error" id="pwErr"></p>
    <button class="btn primary">Simpan password</button></form></div>`;
  $('#pwForm').onsubmit = async e => {
    e.preventDefault();
    const f = e.target;
    if (f.baru.value !== f.ulang.value) { $('#pwErr').textContent = 'Ulangan password tidak sama.'; return; }
    try {
      await API.call('gantiPassword', { lama: await sha256(f.lama.value), baru: await sha256(f.baru.value) });
      f.reset(); $('#pwErr').textContent = ''; toast('Password berhasil diganti');
    } catch (err) { $('#pwErr').textContent = err.message; }
  };
}

const PAGES = {
  dashboard: pageDashboard, masuk: pageMasuk, keluar: pageKeluar, mutasi: pageMutasi,
  stok: pageStok, riwayat: pageRiwayat, lacak: pageLacak,
  cabang: pageCabang, pengguna: pagePengguna, password: pagePassword
};

boot();
