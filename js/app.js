/* ============================================================
   VITEG Gudang — aplikasi frontend
   ============================================================ */
const App = { user: null, cabang: [], models: [], sales: [], pelanggan: [] };
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
const kat = u => u.kategori === 'DEMO' ? 'DEMO' : 'JUAL';
const salesNama = (id, dflt = '-') => { const x = App.sales.find(s => s.id === id); return x ? x.nama : (id || dflt); };
const pelLabel = p => `${p.nama} · ${p.no_ktp}`;
const mapsUrl = m => /^https?:\/\//i.test(m) ? m : 'https://www.google.com/maps?q=' + encodeURIComponent(m);
function opsiSales(sel) {
  return '<option value="">— pilih sales —</option>' + App.sales.filter(s => s.aktif === '1')
    .map(s => `<option value="${esc(s.id)}" ${s.id === sel ? 'selected' : ''}>${esc(s.nama)}${isUtama() ? ' · ' + esc(s.cabang) : ''}</option>`).join('');
}
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

/** Pecah satu baris tempelan menjadi [imei, sn]; IMEI (15 digit) dideteksi otomatis di kolom mana pun. */
function parseLine(line) {
  // Copy dari Excel dipisah Tab -> pisah hanya di Tab (agar "3,57E+14" tetap utuh dan terdeteksi)
  const tok = (line.includes('\t') ? line.split('\t') : line.split(/[,;|]+|\s+/)).map(s => s.trim()).filter(Boolean);
  if (!tok.length) return null;
  let imei = '';
  const rest = [];
  tok.forEach(t => {
    if (!imei && /^\d{15}$/.test(t)) imei = t;
    else rest.push(t);
  });
  if (!imei && rest.length) imei = rest.shift();
  return [imei, rest.shift() || ''];
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

/** Template Excel untuk barang masuk: kolom IMEI/SN berformat Teks. */
async function downloadTemplateMasuk() {
  try {
    const X = await loadXLSX();
    const ws = X.utils.aoa_to_sheet([['IMEI', 'SN']]);
    const ROWS = 1000;
    for (let r = 1; r <= ROWS; r++) for (let c = 0; c < 2; c++) ws[X.utils.encode_cell({ r, c })] = { t: 's', v: '', z: '@' };
    ws['!ref'] = 'A1:B' + (ROWS + 1);
    ws['!cols'] = [{ wch: 20 }, { wch: 24 }];
    const info = X.utils.aoa_to_sheet([
      ['PETUNJUK PENGISIAN — isi data di sheet "Data", jangan ubah judul kolom'],
      [''],
      ['Kolom', 'Aturan', 'Contoh'],
      ['IMEI', 'Wajib, 15 digit angka', '356789012345671'],
      ['SN', 'Wajib, serial number perangkat', 'ZTEGC1234567'],
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

/** Baca file Excel/CSV -> [[imei, sn], ...]. Kolom dicari dari judul; kalau tak ada judul, pakai urutan A-B. */
async function readMasukFile(file) {
  const X = await loadXLSX();
  const wb = X.read(await file.arrayBuffer(), { type: 'array', raw: false, cellText: false });
  const name = wb.SheetNames.includes('Data') ? 'Data' : wb.SheetNames[0];
  const rows = X.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '' });
  let hi = rows.findIndex(r => r.some(c => /imei/i.test(String(c))));
  let ci = { imei: 0, sn: 1 };
  if (hi >= 0) {
    const h = rows[hi].map(c => String(c).trim().toLowerCase());
    const find = re => h.findIndex(c => re.test(c));
    ci = { imei: find(/imei/), sn: find(/^(sn|s\/n|serial|no\.?\s*seri|nomor seri)/) };
    if (ci.sn < 0) throw new Error('Judul kolom SN tidak ditemukan. Gunakan template.');
  }
  return rows.slice(hi + 1)
    .map(r => [cellText(r[ci.imei]), cellText(r[ci.sn])])
    .filter(r => r.some(Boolean));
}

/** Ekspor beberapa sheet sekaligus. sheets = [{ name, cols: [[key, judul, lebar]], rows, numeric }] */
async function exportExcelMulti(nama, sheets) {
  try {
    const X = await loadXLSX();
    const wb = X.utils.book_new();
    sheets.forEach(sh => {
      const aoa = [sh.cols.map(c => c[1])].concat(sh.rows.map(r => sh.cols.map(c => {
        const val = r[c[0]];
        if (sh.numeric && typeof val === 'number') return val;
        return String(val == null ? '' : val);
      })));
      const ws = X.utils.aoa_to_sheet(aoa);
      ws['!cols'] = sh.cols.map(c => ({ wch: c[2] || 16 }));
      X.utils.book_append_sheet(wb, ws, sh.name.slice(0, 31));
    });
    X.writeFile(wb, nama + '.xlsx');
  } catch (e) { toast('Gagal ekspor Excel. Periksa koneksi internet.', true); }
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
function closeModal() { $('#modal').classList.add('hidden'); $('#modalCard').classList.remove('wide'); }

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
  { g: 'Stock Opname' },
  { id: 'opname', label: 'Hitung Fisik', ic: '☑' },
  { id: 'laporan', label: 'Laporan Bulanan', ic: '▥' },
  { id: 'cocok', label: 'Cocokkan MyRepublic', ic: '⇆' },
  { g: 'Master Data' },
  { id: 'pelanggan', label: 'Pelanggan', ic: '◉' },
  { id: 'sales', label: 'Sales', ic: '★' },
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
  closeModal();
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
    App.user = d.user; App.cabang = d.cabang; App.models = d.models; App.sales = d.sales || [];
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
      <div class="stat"><small>Terjual hari ini</small><b>${fmt(d.hariIni.KELUAR)}</b></div>
      <div class="stat"><small>Retur hari ini</small><b>${fmt(d.hariIni.RETUR)}</b></div>
      <div class="stat"><small>Total terjual</small><b>${fmt(d.keluar)}</b></div>
      <div class="stat demo"><small>Demo di gudang</small><b>${fmt(d.stokDemo)}</b></div>
      <div class="stat demo"><small>Demo dibawa sales</small><b>${fmt(d.demoKeluar)}</b></div>
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
    <th>Waktu</th><th>Jenis</th><th>IMEI</th><th>SN</th><th>Model</th><th>Cabang</th><th>Pelanggan / penerima</th><th>Sales</th><th>Referensi</th><th>Keterangan</th><th>Oleh</th>
    </tr></thead><tbody>${rows.length ? rows.map(t => `<tr>
      <td class="mono">${esc(t.waktu)}</td><td>${badge(t.jenis)}</td>
      <td class="mono">${esc(t.imei)}</td><td class="mono">${esc(t.sn)}</td>
      <td>${esc(t.model)}</td>
      <td>${t.jenis === 'MUTASI' ? esc(t.cabang_asal) + ' → ' + esc(t.cabang) : esc(t.cabang)}</td>
      <td>${esc(t.penerima)}</td><td>${t.sales_id ? esc(salesNama(t.sales_id)) : ''}</td><td>${esc(t.referensi)}</td><td>${esc(t.keterangan)}</td><td>${esc(t.oleh)}</td>
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
      <label>Kategori<select id="mKat"><option value="JUAL">Barang jual</option><option value="DEMO">Barang demo (tidak dijual)</option></select></label>
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
        <button class="btn" id="sAdd">Tambah</button>
      </div>
      <p class="hint">Scan IMEI lalu Enter, scan SN lalu Enter — unit langsung masuk daftar. Cocok untuk scanner barcode.</p>
    </div>
    <div id="tPaste" class="hidden">
      <div class="format-box">
        <strong>Format:</strong> satu unit per baris, urutan <b>IMEI · SN</b>, dipisah Tab (hasil copy dari Excel), koma, titik koma, atau spasi.
        <pre>356789012345671	ZTEGC1234567
356789012345672,ZTEGC1234568
356789012345673;ZTEGC1234569</pre>
        Baris judul (IMEI, SN) otomatis dilewati. Kalau copy dari Excel, pastikan kolom IMEI berformat <b>Teks</b>,
        bukan Angka, supaya tidak tersalin sebagai 3,57E+14.
      </div>
      <textarea id="pText" placeholder="Tempel di sini…"></textarea>
      <div class="row" style="margin-top:8px"><button class="btn" id="pAdd">Tambahkan ke daftar</button></div>
    </div>
    <div id="tFile" class="hidden">
      <div class="format-box">
        <strong>Langkah:</strong> 1) Unduh template · 2) Isi kolom IMEI dan SN di sheet <b>Data</b> (satu baris = satu unit) ·
        3) Simpan lalu pilih filenya di bawah. Model dan cabang dipilih di atas. File .xlsx, .xls, dan .csv juga bisa,
        asalkan baris judulnya memuat kata IMEI dan SN.
      </div>
      <div class="row">
        <button class="btn" id="fTpl">⇩ Unduh template Excel</button>
        <label class="btn primary file-btn">Pilih file Excel…<input type="file" id="fFile" accept=".xlsx,.xls,.csv" hidden></label>
      </div>
    </div>
    <div id="mMsg" style="margin-top:14px"></div>
    <div class="table-wrap"><table>
      <thead><tr><th>#</th><th>IMEI</th><th>SN</th><th>Cek</th><th></th></tr></thead>
      <tbody id="mBody"></tbody></table></div>
    <div class="summary">
      <span><b id="mCount">0</b> unit dalam daftar <span id="mErr" style="color:var(--bad)"></span></span>
      <div class="row"><button class="btn ghost" id="mClear">Kosongkan</button>
        <button class="btn primary" id="mSave">Simpan barang masuk</button></div>
    </div>
  </div>`;

  const check = () => {
    const cnt = {};
    list.forEach(x => ['i' + x.imei, 's' + x.sn].forEach(k => cnt[k] = (cnt[k] || 0) + 1));
    list.forEach(x => {
      const e = [];
      if (x.sci) e.push('IMEI tersalin sebagai format ilmiah (E+14) — ubah kolom Excel ke Teks lalu salin ulang');
      else if (x.imei.length !== 15) e.push('IMEI harus 15 digit');
      if (!x.sn) e.push('SN kosong');
      if (x.imei && cnt['i' + x.imei] > 1) e.push('IMEI dobel');
      if (x.sn && cnt['s' + x.sn] > 1) e.push('SN dobel');
      if (x.srv) e.push(x.srv);
      x.err = e.join(', ');
    });
  };
  const render = () => {
    check();
    $('#mBody').innerHTML = list.length ? list.map((x, i) => `
      <tr class="${x.err ? 'err' : ''}"><td>${i + 1}</td>
        <td class="mono">${esc(x.imei)}</td><td class="mono">${esc(x.sn)}</td>
        <td>${x.err ? `<small style="color:var(--bad)">${esc(x.err)}</small>` : '<span class="badge b-on">OK</span>'}</td>
        <td><button class="btn ghost sm" data-del="${i}" title="Hapus">✕</button></td></tr>`).join('')
      : '<tr><td colspan="5" class="empty">Belum ada unit. Scan, tempel, atau impor data di atas.</td></tr>';
    const nErr = list.filter(x => x.err).length;
    $('#mCount').textContent = list.length;
    $('#mErr').textContent = nErr ? `· ${nErr} bermasalah` : '';
  };
  const add = (imei, sn) => {
    const sci = isSci(imei);
    list.push({ imei: sci ? String(imei).trim() : normImei(imei), sn: normSn(sn), sci });
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
      rows.forEach(r => add(r[0], r[1]));
      render();
      const nErr = list.filter(x => x.err).length;
      toast(`${rows.length} unit dibaca dari ${file.name}` + (nErr ? ` · ${nErr} bermasalah` : ''), nErr > 0);
    } catch (err) {
      $('#mMsg').innerHTML = errBox({ message: 'Gagal membaca file: ' + err.message });
    }
  };

  const addScan = () => {
    const i = $('#sImei').value, s = $('#sSn').value;
    if (!i && !s) return;
    add(i, s);
    $('#sImei').value = $('#sSn').value = '';
    render();
    $('#sImei').focus();
  };
  $('#sImei').onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); $('#sSn').focus(); } };
  $('#sSn').onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); addScan(); } };
  $('#sAdd').onclick = addScan;

  $('#pAdd').onclick = () => {
    let n = 0;
    $('#pText').value.split(/\r?\n/).forEach(line => {
      const p = parseLine(line);
      if (!p) return;
      if (/^imei$/i.test(p[0])) return; // lewati baris judul
      add(p[0], p[1]); n++;
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
      model, kategori: $('#mKat').value, referensi: $('#mRef').value.trim(), keterangan: $('#mKet').value.trim()
    };
    const btn = $('#mSave');
    btn.disabled = true;
    let masuk = 0, retur = 0;
    try {
      while (list.length) {
        btn.textContent = `Menyimpan… (${list.length} tersisa)`;
        const chunk = list.slice(0, CHUNK_MASUK);
        const r = await API.call('masuk', Object.assign({}, base, { items: chunk.map(x => [x.imei, x.sn]) }));
        masuk += r.masuk; retur += r.retur;
        list.splice(0, chunk.length);
        render();
      }
      if (!App.models.includes(model)) App.models.push(model);
      msg.innerHTML = okBox(`<strong>Tersimpan.</strong> ${masuk} unit baru${base.kategori === 'DEMO' && masuk ? ' (barang demo)' : ''}${retur ? `, ${retur} unit retur` : ''} masuk ke ${esc(cabNama(base.cabang))}.`);
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
  const isK = mode === 'keluar';
  const [stok, pelanggan] = await Promise.all([
    API.call('stok', { status: 'GUDANG' }),
    isK ? API.call('pelanggan') : Promise.resolve([])
  ]);
  App.pelanggan = pelanggan;
  const ix = { imei: {}, sn: {} };
  stok.forEach(u => { ix.imei[u.imei] = u; ix.sn[u.sn] = u; });
  const find = k => {
    k = String(k || '').trim();
    if (!k) return null;
    const d = normImei(k);
    if (d.length === 15 && /^[\d\s-]+$/.test(k) && ix.imei[d]) return ix.imei[d];
    return ix.sn[normSn(k)] || null;
  };
  const list = [];
  let tipe = 'JUAL';
  let pelSel = null;
  const nDemo = stok.filter(u => kat(u) === 'DEMO').length;

  v.innerHTML = `
  ${isK ? `<div class="seg" id="kTipe">
      <button data-v="JUAL" class="on">Penjualan<small>Barang jual → pelanggan</small></button>
      <button data-v="DEMO">Barang demo<small>Dibawa sales, tidak dijual</small></button>
    </div>` : ''}
  <div class="card">
    <h3>${isK ? 'Informasi pengeluaran' : 'Informasi mutasi'}</h3>
    <div class="form-grid">
      ${isK ? `
        <label id="lblPel"><span>Pelanggan <em id="pelOpt" class="muted hidden">(opsional)</em></span>
          <div class="input-btn">
            <input id="kPel" list="dlPel" placeholder="Cari nama / No. KTP" autocomplete="off">
            <button class="btn" type="button" id="kPelNew" title="Tambah pelanggan baru">+ Baru</button>
          </div>
          <datalist id="dlPel"></datalist>
          <small class="pel-info" id="kPelInfo"></small>
        </label>
        <label>Sales <span id="salesHint" class="muted"></span><select id="kSales">${opsiSales()}</select></label>
        <label>No. referensi<input id="kRef" placeholder="No. WO / order"></label>`
      : `<label>Cabang tujuan<select id="kTuj">${opsiCabang('', '— pilih cabang —')}</select></label>
         <label>No. referensi<input id="kRef" placeholder="No. surat jalan"></label>`}
      <label>Keterangan<input id="kKet" placeholder="Opsional"></label>
    </div>
  </div>
  <div class="card">
    <h3>Unit ${isK ? 'keluar' : 'dipindahkan'} <small class="muted" id="kAvail"></small></h3>
    <div class="tabs"><button data-t="scan" class="on">Scan</button><button data-t="paste">Tempel massal</button></div>
    <div id="tScan" class="row scan-row">
      <label>Scan IMEI / SN<input id="kScan" autocomplete="off" placeholder="Scan atau ketik lalu Enter"></label>
      <button class="btn" id="kAdd">Tambah</button>
    </div>
    <div id="tPaste" class="hidden">
      <textarea id="kText" placeholder="Satu IMEI atau SN per baris"></textarea>
      <div class="row" style="margin-top:8px"><button class="btn" id="kPaste">Tambahkan ke daftar</button></div>
    </div>
    <div id="kMsg" style="margin-top:14px"></div>
    <div class="table-wrap"><table>
      <thead><tr><th>#</th><th>IMEI</th><th>SN</th><th>Model</th><th>Kategori</th><th>Cabang</th><th></th></tr></thead>
      <tbody id="kBody"></tbody></table></div>
    <div class="summary">
      <span><b id="kCount">0</b> unit dipilih</span>
      <div class="row"><button class="btn ghost" id="kClear">Kosongkan</button>
        <button class="btn primary" id="kSave"></button></div>
    </div>
  </div>`;

  const saveLabel = () => isK ? (tipe === 'JUAL' ? 'Simpan penjualan' : 'Simpan barang demo keluar') : 'Simpan mutasi';
  const render = () => {
    $('#kBody').innerHTML = list.length ? list.map((u, i) => `<tr>
      <td>${i + 1}</td><td class="mono">${esc(u.imei)}</td><td class="mono">${esc(u.sn)}</td>
      <td>${esc(u.model)}</td><td>${badge(kat(u))}</td><td>${esc(cabNama(u.cabang))}</td>
      <td><button class="btn ghost sm" data-del="${i}">✕</button></td></tr>`).join('')
      : '<tr><td colspan="7" class="empty">Belum ada unit dipilih.</td></tr>';
    $('#kCount').textContent = list.length;
    $('#kSave').textContent = saveLabel();
    const tersedia = isK ? (tipe === 'JUAL' ? stok.length - nDemo : nDemo) : stok.length;
    $('#kAvail').textContent = `${fmt(tersedia)} unit ${isK ? (tipe === 'JUAL' ? 'barang jual' : 'barang demo') : ''} tersedia${isUtama() ? ' (semua cabang)' : ''}`;
  };
  const addKey = k => {
    const u = find(k);
    if (!u) return 'tidak ada di stok gudang';
    if (list.includes(u)) return 'sudah ada di daftar';
    if (isK && tipe === 'JUAL' && kat(u) === 'DEMO') return 'barang DEMO tidak boleh dijual';
    if (isK && tipe === 'DEMO' && kat(u) !== 'DEMO') return 'bukan barang demo';
    list.push(u);
    return '';
  };

  if (isK) {
    const fillPel = () => {
      $('#dlPel').innerHTML = App.pelanggan.map(p => `<option value="${esc(pelLabel(p))}">`).join('');
    };
    const showPel = () => {
      $('#kPelInfo').innerHTML = pelSel
        ? `✓ ${esc(pelSel.id)} · ${esc(pelSel.alamat || 'alamat belum diisi')}${pelSel.maps ? ` · <a href="${esc(mapsUrl(pelSel.maps))}" target="_blank" rel="noopener">lokasi</a>` : ''}`
        : '';
    };
    fillPel();
    $('#kPel').oninput = () => {
      const val = $('#kPel').value.trim();
      pelSel = App.pelanggan.find(p => pelLabel(p) === val) ||
        (/^\d{16}$/.test(val) ? App.pelanggan.find(p => p.no_ktp === val) : null) || null;
      showPel();
    };
    $('#kPelNew').onclick = () => pelangganForm(null, p => {
      App.pelanggan.push(p); fillPel();
      pelSel = p; $('#kPel').value = pelLabel(p); showPel();
    });
    $$('#kTipe button').forEach(b => b.onclick = () => {
      tipe = b.dataset.v;
      $$('#kTipe button').forEach(x => x.classList.toggle('on', x === b));
      $('#pelOpt').classList.toggle('hidden', tipe === 'JUAL');
      $('#salesHint').textContent = tipe === 'DEMO' ? '(pemegang demo)' : '';
      const before = list.length;
      for (let i = list.length - 1; i >= 0; i--) if ((kat(list[i]) === 'DEMO') !== (tipe === 'DEMO')) list.splice(i, 1);
      if (before !== list.length) toast(`${before - list.length} unit dikeluarkan dari daftar karena kategorinya tidak sesuai`);
      render();
    });
    if (!App.sales.some(s => s.aktif === '1')) {
      $('#kMsg').innerHTML = `<div class="alert bad">Belum ada data sales aktif. Tambahkan dulu di menu <a href="#sales">Sales</a>.</div>`;
    }
  }

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
      base.tipe = tipe;
      base.sales_id = $('#kSales').value;
      if ($('#kPel').value.trim() && !pelSel) { toast('Pelanggan tidak dikenal — pilih dari daftar atau klik + Baru', true); $('#kPel').focus(); return; }
      if (tipe === 'JUAL' && !pelSel) { toast('Pilih pelanggan', true); $('#kPel').focus(); return; }
      if (!base.sales_id) { toast('Pilih sales', true); $('#kSales').focus(); return; }
      base.pelanggan_id = pelSel ? pelSel.id : '';
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
        chunk.forEach(u => { delete ix.imei[u.imei]; delete ix.sn[u.sn]; stok.splice(stok.indexOf(u), 1); });
        done += chunk.length;
        list.splice(0, chunk.length);
        render();
      }
      msg.innerHTML = okBox(!isK
        ? `<strong>Tersimpan.</strong> ${done} unit dipindahkan ke ${esc(cabNama(base.tujuan))}.`
        : tipe === 'JUAL'
          ? `<strong>Penjualan tersimpan.</strong> ${done} unit ke ${esc(pelSel.nama)} · sales ${esc(salesNama(base.sales_id))}.`
          : `<strong>Tersimpan.</strong> ${done} unit demo dibawa ${esc(salesNama(base.sales_id))}${pelSel ? ' untuk ' + esc(pelSel.nama) : ''}.`);
      toast('Transaksi tersimpan');
    } catch (e) {
      msg.innerHTML = errBox(e, done ? `<div>${done} unit sebelumnya sudah tersimpan.</div>` : '');
    } finally {
      btn.disabled = false; btn.textContent = saveLabel();
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
      <label>Kategori<select id="fKat"><option value="">Semua</option><option value="JUAL">Barang jual</option><option value="DEMO">Barang demo</option></select></label>
      <label>Model<select id="fMod"><option value="">Semua model</option>${models.map(m => `<option>${esc(m)}</option>`).join('')}</select></label>
      <label class="grow">Cari<input id="fQ" placeholder="IMEI, SN, pelanggan, sales, referensi…"></label>
      <button class="btn" id="fExp">Ekspor Excel</button>
    </div>
    <div class="table-wrap"><table>
      <thead><tr><th>IMEI</th><th>SN</th><th>Model</th><th>Kategori</th><th>Status</th><th>Cabang</th><th>Tgl masuk</th><th>Tgl keluar</th><th>Pelanggan / penerima</th><th>Sales</th><th>Referensi</th></tr></thead>
      <tbody id="sBody"></tbody></table></div>
    <div class="summary"><span id="sInfo"></span>
      <div class="row"><button class="btn sm" id="pPrev">‹ Sebelumnya</button><button class="btn sm" id="pNext">Berikutnya ›</button></div></div>
  </div>`;

  all.forEach(u => { u.kategori = kat(u); u.sales = salesNama(u.sales_id, ''); });
  const apply = () => {
    const cab = isUtama() ? $('#fCab').value : '';
    const st = $('#fSt').value, kt = $('#fKat').value, mod = $('#fMod').value, q = $('#fQ').value.trim().toUpperCase();
    rows = all.filter(u => (!cab || u.cabang === cab) && (!st || u.status === st) && (!kt || u.kategori === kt) &&
      (!mod || u.model === mod) &&
      (!q || [u.imei, u.sn, u.penerima, u.sales, u.referensi].join(' ').toUpperCase().includes(q)));
    page = 0; draw();
  };
  const draw = () => {
    const slice = rows.slice(page * PAGE, page * PAGE + PAGE);
    $('#sBody').innerHTML = slice.length ? slice.map(u => `<tr>
      <td class="mono"><a href="#lacak" data-q="${esc(u.imei)}">${esc(u.imei)}</a></td><td class="mono">${esc(u.sn)}</td>
      <td>${esc(u.model)}</td><td>${badge(u.kategori)}</td><td>${badge(u.status)}</td><td>${esc(cabNama(u.cabang))}</td>
      <td class="mono">${esc(u.tgl_masuk.slice(0, 16))}</td><td class="mono">${esc(u.tgl_keluar.slice(0, 16))}</td>
      <td>${esc(u.penerima)}</td><td>${esc(u.sales)}</td><td>${esc(u.referensi)}</td></tr>`).join('')
      : '<tr><td colspan="11" class="empty">Tidak ada data</td></tr>';
    const from = rows.length ? page * PAGE + 1 : 0;
    $('#sInfo').innerHTML = `Menampilkan <b>${fmt(from)}–${fmt(Math.min(rows.length, (page + 1) * PAGE))}</b> dari <b>${fmt(rows.length)}</b> unit`;
    $('#pPrev').disabled = page === 0;
    $('#pNext').disabled = (page + 1) * PAGE >= rows.length;
  };
  $$('#fCab, #fSt, #fKat, #fMod', v).forEach(el => el.onchange = apply);
  $('#fQ').oninput = apply;
  $('#pPrev').onclick = () => { page--; draw(); };
  $('#pNext').onclick = () => { page++; draw(); };
  $('#sBody').onclick = e => { const a = e.target.closest('[data-q]'); if (a) sessionStorage.setItem('vg_lacak', a.dataset.q); };
  $('#fExp').onclick = () => exportExcel('stok-' + today(), [
    ['imei', 'IMEI', 18], ['sn', 'SN', 18], ['model', 'Model', 24], ['kategori', 'Kategori', 9], ['status', 'Status', 10],
    ['cabang', 'Cabang', 10], ['tgl_masuk', 'Tgl Masuk', 19], ['tgl_keluar', 'Tgl Keluar', 19],
    ['pelanggan_id', 'ID Pelanggan', 12], ['penerima', 'Pelanggan / Penerima', 24], ['sales', 'Sales', 20],
    ['referensi', 'Referensi', 18], ['keterangan', 'Keterangan', 24]
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
      <label>Jenis<select id="rJenis"><option value="">Semua</option>${[['MASUK', 'Masuk'], ['KELUAR', 'Keluar (penjualan)'], ['DEMO', 'Keluar demo'], ['RETUR', 'Retur'], ['MUTASI', 'Mutasi'], ['KATEGORI', 'Ubah kategori'], ['HAPUS', 'Hapus']].map(j => `<option value="${j[0]}">${j[1]}</option>`).join('')}</select></label>
      ${isUtama() ? `<label>Cabang<select id="rCab">${opsiCabang('', 'Semua cabang', false)}</select></label>` : ''}
      <label class="grow">Cari<input id="rQ" placeholder="IMEI, SN, pelanggan, sales, referensi…"></label>
      <button class="btn" id="rExp">Ekspor Excel</button>
    </div>
    <div id="rOut">${loadingHtml()}</div>
  </div>`;
  const apply = () => {
    const j = $('#rJenis').value, cab = isUtama() ? $('#rCab').value : '', q = $('#rQ').value.trim().toUpperCase();
    rows = data.filter(t => (!j || t.jenis === j) && (!cab || t.cabang === cab || t.cabang_asal === cab) &&
      (!q || [t.imei, t.sn, t.penerima, salesNama(t.sales_id, ''), t.referensi, t.keterangan, t.oleh].join(' ').toUpperCase().includes(q)));
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
  $('#rExp').onclick = () => { rows.forEach(t => { t.sales = salesNama(t.sales_id, ''); }); exportExcel(`riwayat-${$('#rDari').value}_${$('#rSampai').value}`, [
    ['id', 'ID Transaksi', 22], ['waktu', 'Waktu', 19], ['jenis', 'Jenis', 9], ['imei', 'IMEI', 18], ['sn', 'SN', 18],
    ['model', 'Model', 24], ['cabang_asal', 'Cabang Asal', 12], ['cabang', 'Cabang', 10],
    ['kategori', 'Kategori', 9], ['pelanggan_id', 'ID Pelanggan', 12], ['penerima', 'Pelanggan / Penerima', 24], ['sales', 'Sales', 20],
    ['referensi', 'Referensi', 18], ['keterangan', 'Keterangan', 24], ['oleh', 'Oleh', 12]
  ], rows); };
  load();
}

/* ---------------- Lacak unit ---------------- */
async function pageLacak(v) {
  v.innerHTML = `
  <div class="card">
    <div class="row">
      <label>IMEI / SN<input id="lQ" autocomplete="off" placeholder="Scan atau ketik lalu Enter"></label>
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
        <div class="card"><h3><span>Unit ditemukan ${badge(u.status)} ${badge(kat(u))}</span>
          ${isUtama() && u.status === 'GUDANG' ? `<span class="row"><button class="btn sm" id="lKat">Ubah ke ${kat(u) === 'DEMO' ? 'barang jual' : 'barang demo'}</button><button class="btn danger sm" id="lDel">Hapus (salah input)</button></span>` : ''}</h3>
          <div class="unit-card">
            <div><small>IMEI</small><span class="mono">${esc(u.imei)}</span></div>
            <div><small>SN</small><span class="mono">${esc(u.sn)}</span></div>
            <div><small>Model</small>${esc(u.model)}</div>
            <div><small>Cabang</small>${esc(cabNama(u.cabang))}</div>
            <div><small>Tgl masuk</small>${esc(u.tgl_masuk)}</div>
            ${u.status === 'KELUAR' ? `<div><small>Tgl keluar</small>${esc(u.tgl_keluar)}</div>
              <div><small>${r.pelanggan ? 'Pelanggan' : 'Penerima'}</small>${esc(u.penerima)}</div>
              <div><small>Sales</small>${esc(salesNama(u.sales_id))}</div>` : ''}
            <div><small>Referensi</small>${esc(u.referensi || '-')}</div>
          </div>
          ${r.pelanggan ? `<div class="pel-box">
            <div><small>Pelanggan</small><strong>${esc(r.pelanggan.nama)}</strong> <span class="muted mono">${esc(r.pelanggan.id)}</span></div>
            <div><small>No. KTP</small><span class="mono">${esc(r.pelanggan.no_ktp)}</span></div>
            <div><small>Alamat</small>${esc(r.pelanggan.alamat || '-')}</div>
            <div><small>Lokasi</small>${r.pelanggan.maps ? `<a href="${esc(mapsUrl(r.pelanggan.maps))}" target="_blank" rel="noopener">Buka di Google Maps ↗</a>` : '-'}</div>
          </div>` : ''}</div>`
        : `<div class="alert bad">Unit tidak ditemukan di stok${r.riwayat.length ? ' (ada riwayat lama di bawah)' : ''}.</div>`) +
        (r.riwayat.length ? `<div class="card"><h3>Riwayat</h3><ul class="timeline">${r.riwayat.map(t => `
          <li><span class="mono">${esc(t.waktu)}</span><span>${badge(t.jenis)}</span>
          <span>${esc(t.jenis === 'MUTASI' ? cabNama(t.cabang_asal) + ' → ' + cabNama(t.cabang) : cabNama(t.cabang))}
          ${t.penerima ? ' · ' + esc(t.penerima) : ''}${t.sales_id ? ' · Sales ' + esc(salesNama(t.sales_id)) : ''}${t.referensi ? ' · Ref ' + esc(t.referensi) : ''}${t.keterangan ? ' · ' + esc(t.keterangan) : ''}
          <small class="muted"> — ${esc(t.oleh)}</small></span></li>`).join('')}</ul></div>` : '');
      const del = $('#lDel');
      if (del) del.onclick = () => hapusUnit(u, go);
      const kb = $('#lKat');
      if (kb) kb.onclick = () => ubahKategori(u, go);
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

/* ---------------- Pelanggan ---------------- */
function pelangganForm(p, onSaved) {
  const baru = !p;
  p = p || { id: '', no_ktp: '', nama: '', alamat: '', maps: '', cabang: App.user.cabang };
  modal(`<h3>${baru ? 'Tambah' : 'Ubah'} pelanggan</h3>
    <label>No. KTP (NIK)<input id="pKtp" value="${esc(p.no_ktp)}" inputmode="numeric" placeholder="16 digit"></label>
    <label>Nama lengkap<input id="pNama" value="${esc(p.nama)}"></label>
    <label>Alamat<textarea id="pAlm" style="min-height:60px;font-family:inherit;font-size:14px">${esc(p.alamat)}</textarea></label>
    <label>Lokasi maps<input id="pMaps" value="${esc(p.maps)}" placeholder="Link Google Maps atau 0.5435, 123.0568"></label>
    <div class="row" style="margin-top:-6px">
      <button class="btn sm" type="button" id="pGps">⌖ Pakai lokasi saya sekarang</button>
      <small class="muted" id="pGpsInfo"></small>
    </div>
    ${isUtama() ? `<label>Cabang<select id="pCab">${opsiCabang(p.cabang)}</select></label>` : ''}
    <p class="form-error" id="pErr"></p>
    <div class="modal-actions"><button class="btn" id="pNo">Batal</button><button class="btn primary" id="pYes">Simpan</button></div>`);
  $('#pKtp').oninput = e => { e.target.value = e.target.value.replace(/\D/g, '').slice(0, 16); };
  $('#pGps').onclick = () => {
    if (!navigator.geolocation) { $('#pGpsInfo').textContent = 'Perangkat tidak mendukung GPS'; return; }
    $('#pGpsInfo').textContent = 'Mengambil lokasi…';
    navigator.geolocation.getCurrentPosition(
      pos => { $('#pMaps').value = pos.coords.latitude.toFixed(6) + ', ' + pos.coords.longitude.toFixed(6); $('#pGpsInfo').textContent = `Akurasi ±${Math.round(pos.coords.accuracy)} m`; },
      err => { $('#pGpsInfo').textContent = 'Gagal: ' + err.message; },
      { enableHighAccuracy: true, timeout: 15000 });
  };
  $('#pNo').onclick = closeModal;
  $('#pYes').onclick = async () => {
    const btn = $('#pYes');
    btn.disabled = true;
    try {
      const saved = await API.call('simpanPelanggan', {
        id: p.id, no_ktp: $('#pKtp').value, nama: $('#pNama').value, alamat: $('#pAlm').value,
        maps: $('#pMaps').value, cabang: isUtama() ? $('#pCab').value : ''
      });
      closeModal(); toast('Pelanggan tersimpan');
      onSaved(saved);
    } catch (e) { $('#pErr').textContent = e.message; btn.disabled = false; }
  };
}

async function pelangganDetail(id) {
  modal(loadingHtml());
  $('#modalCard').classList.add('wide');
  try {
    const r = await API.call('pelangganDetail', { id });
    const p = r.pelanggan;
    $('#modalCard').innerHTML = `
      <h3>${esc(p.nama)} <small class="muted mono">${esc(p.id)}</small></h3>
      <div class="unit-card">
        <div><small>No. KTP</small><span class="mono">${esc(p.no_ktp)}</span></div>
        <div><small>Cabang</small>${esc(cabNama(p.cabang))}</div>
        <div style="grid-column:1/-1"><small>Alamat</small>${esc(p.alamat || '-')}</div>
        <div style="grid-column:1/-1"><small>Lokasi</small>${p.maps ? `<a href="${esc(mapsUrl(p.maps))}" target="_blank" rel="noopener">${esc(p.maps)} ↗</a>` : '-'}</div>
      </div>
      <h3 style="margin:6px 0 0">Unit di pelanggan ini (${r.units.length})</h3>
      <div class="table-wrap"><table><thead><tr><th>IMEI</th><th>SN</th><th>Model</th><th>Kategori</th><th>Tgl keluar</th><th>Sales</th></tr></thead>
        <tbody>${r.units.length ? r.units.map(u => `<tr><td class="mono">${esc(u.imei)}</td><td class="mono">${esc(u.sn)}</td>
          <td>${esc(u.model)}</td><td>${badge(kat(u))}</td><td class="mono">${esc(u.tgl_keluar.slice(0, 16))}</td><td>${esc(salesNama(u.sales_id))}</td></tr>`).join('')
          : '<tr><td colspan="6" class="empty">Tidak ada unit</td></tr>'}</tbody></table></div>
      ${r.riwayat.length ? `<details><summary class="muted" style="cursor:pointer">Riwayat transaksi (${r.riwayat.length})</summary>${trxTable(r.riwayat)}</details>` : ''}
      <div class="modal-actions"><button class="btn" id="dClose">Tutup</button></div>`;
    $('#dClose').onclick = closeModal;
  } catch (e) { $('#modalCard').innerHTML = errBox(e) + '<div class="modal-actions"><button class="btn" onclick="closeModal()">Tutup</button></div>'; }
}

async function pagePelanggan(v) {
  App.pelanggan = await API.call('pelanggan');
  v.innerHTML = `<div class="card">
    <div class="toolbar">
      ${isUtama() ? `<label>Cabang<select id="plCab">${opsiCabang('', 'Semua cabang', false)}</select></label>` : ''}
      <label class="grow">Cari<input id="plQ" placeholder="Nama, No. KTP, alamat, ID…"></label>
      <button class="btn" id="plExp">Ekspor Excel</button>
      <button class="btn primary" id="plNew">+ Tambah pelanggan</button>
    </div>
    <div class="table-wrap"><table><thead><tr><th>ID</th><th>Nama</th><th>No. KTP</th><th>Alamat</th><th>Lokasi</th><th>Cabang</th><th></th></tr></thead>
      <tbody id="plBody"></tbody></table></div>
    <div class="summary"><span id="plInfo"></span></div>
  </div>`;
  let rows = [];
  const draw = () => {
    const cab = isUtama() ? $('#plCab').value : '', q = $('#plQ').value.trim().toUpperCase();
    rows = App.pelanggan.filter(p => (!cab || p.cabang === cab) &&
      (!q || [p.id, p.nama, p.no_ktp, p.alamat].join(' ').toUpperCase().includes(q)));
    const shown = rows.slice(0, 300);
    $('#plBody').innerHTML = shown.length ? shown.map(p => `<tr>
      <td class="mono">${esc(p.id)}</td><td><strong>${esc(p.nama)}</strong></td><td class="mono">${esc(p.no_ktp)}</td>
      <td>${esc(p.alamat)}</td>
      <td>${p.maps ? `<a href="${esc(mapsUrl(p.maps))}" target="_blank" rel="noopener">Buka maps ↗</a>` : '<span class="muted">-</span>'}</td>
      <td>${esc(cabNama(p.cabang))}</td>
      <td class="nowrap"><button class="btn sm" data-det="${esc(p.id)}">Unit</button> <button class="btn sm" data-edit="${esc(p.id)}">Ubah</button></td></tr>`).join('')
      : '<tr><td colspan="7" class="empty">Belum ada pelanggan</td></tr>';
    $('#plInfo').innerHTML = `<b>${fmt(rows.length)}</b> pelanggan${rows.length > 300 ? ' (300 pertama ditampilkan, gunakan pencarian)' : ''}`;
  };
  $('#plQ').oninput = draw;
  if (isUtama()) $('#plCab').onchange = draw;
  $('#plNew').onclick = () => pelangganForm(null, p => { App.pelanggan.push(p); draw(); });
  $('#plBody').onclick = e => {
    const d = e.target.closest('[data-det]'), ed = e.target.closest('[data-edit]');
    if (d) pelangganDetail(d.dataset.det);
    if (ed) {
      const p = App.pelanggan.find(x => x.id === ed.dataset.edit);
      pelangganForm(p, s => { Object.assign(p, s); draw(); });
    }
  };
  $('#plExp').onclick = () => exportExcel('pelanggan-' + today(), [
    ['id', 'ID', 11], ['nama', 'Nama', 26], ['no_ktp', 'No. KTP', 19], ['alamat', 'Alamat', 40],
    ['maps', 'Lokasi', 30], ['cabang', 'Cabang', 10], ['dibuat', 'Dibuat', 19]
  ], rows);
  draw();
}

/* ---------------- Sales ---------------- */
async function pageSales(v) {
  App.sales = await API.call('sales');
  v.innerHTML = `<div class="card"><h3>Daftar sales <button class="btn primary sm" id="slNew">+ Tambah sales</button></h3>
    <div class="table-wrap"><table><thead><tr><th>ID</th><th>Nama</th><th>No. HP</th><th>Cabang</th><th>Status</th><th></th></tr></thead>
    <tbody>${App.sales.length ? App.sales.map((s, i) => `<tr><td class="mono">${esc(s.id)}</td><td>${esc(s.nama)}</td>
      <td class="mono">${esc(s.no_hp)}</td><td>${esc(cabNama(s.cabang))}</td>
      <td>${s.aktif === '1' ? '<span class="badge b-on">AKTIF</span>' : '<span class="badge b-off">NONAKTIF</span>'}</td>
      <td><button class="btn sm" data-i="${i}">Ubah</button></td></tr>`).join('')
      : '<tr><td colspan="6" class="empty">Belum ada sales</td></tr>'}</tbody></table></div></div>`;
  const form = s => {
    const baru = !s;
    s = s || { id: '', nama: '', no_hp: '', cabang: App.user.cabang, aktif: '1' };
    modal(`<h3>${baru ? 'Tambah' : 'Ubah'} sales</h3>
      <label>Nama<input id="sNama" value="${esc(s.nama)}"></label>
      <label>No. HP<input id="sHp" value="${esc(s.no_hp)}" inputmode="tel" placeholder="08xx"></label>
      ${isUtama() ? `<label>Cabang<select id="sCab">${opsiCabang(s.cabang)}</select></label>` : ''}
      <label>Status<select id="sAkt"><option value="1">Aktif</option><option value="0" ${s.aktif === '0' ? 'selected' : ''}>Nonaktif</option></select></label>
      <p class="form-error" id="sErr"></p>
      <div class="modal-actions"><button class="btn" id="sNo">Batal</button><button class="btn primary" id="sYes">Simpan</button></div>`);
    $('#sNo').onclick = closeModal;
    $('#sYes').onclick = async () => {
      try {
        App.sales = await API.call('simpanSales', { id: s.id, nama: $('#sNama').value, no_hp: $('#sHp').value,
          cabang: isUtama() ? $('#sCab').value : '', aktif: $('#sAkt').value });
        closeModal(); toast('Sales tersimpan'); pageSales(v);
      } catch (e) { $('#sErr').textContent = e.message; }
    };
  };
  $('#slNew').onclick = () => form(null);
  $$('[data-i]', v).forEach(b => b.onclick = () => form(App.sales[+b.dataset.i]));
}

function ubahKategori(u, after) {
  const ke = kat(u) === 'DEMO' ? 'JUAL' : 'DEMO';
  modal(`<h3>Ubah kategori ${esc(u.imei)}</h3>
    <p style="margin:0">${badge(kat(u))} → ${badge(ke)}</p>
    <label>Alasan<input id="kaAl" placeholder="${ke === 'JUAL' ? 'mis. masa demo selesai, dijual' : 'mis. dipakai untuk demo'}"></label>
    <p class="form-error" id="kaErr"></p>
    <div class="modal-actions"><button class="btn" id="kaNo">Batal</button><button class="btn primary" id="kaYes">Ubah ke ${ke}</button></div>`);
  $('#kaNo').onclick = closeModal;
  $('#kaYes').onclick = async () => {
    try {
      await API.call('ubahKategori', { keys: [u.imei], kategori: ke, alasan: $('#kaAl').value.trim() });
      closeModal(); toast('Kategori diubah'); after();
    } catch (e) { $('#kaErr').textContent = e.message; }
  };
}

/* ============================================================
   STOCK OPNAME
   ============================================================ */
const HASIL = {
  COCOK: ['Cocok', 'b-on'],
  HILANG: ['Tidak ditemukan fisik', 'b-HAPUS'],
  TIDAK_TERDAFTAR: ['Tidak terdaftar di sistem', 'b-RETUR'],
  CABANG_LAIN: ['Tercatat di cabang lain', 'b-MUTASI'],
  SUDAH_KELUAR: ['Tercatat sudah keluar', 'b-KELUAR'],
  DOBEL: ['Sudah discan', 'b-JUAL']
};
const hasilBadge = h => `<span class="badge ${(HASIL[h] || ['', 'b-JUAL'])[1]}">${esc((HASIL[h] || [h])[0])}</span>`;

/* ---------------- Hitung fisik ---------------- */
async function pageOpname(v) {
  const list = await API.call('opnameList');
  const draftOf = c => list.find(o => o.status === 'DRAFT' && o.cabang === c);
  v.innerHTML = `
  <div class="card">
    <h3>Hitung fisik stok</h3>
    <p class="muted" style="margin:-6px 0 14px">Scan semua unit yang ada secara fisik di gudang. Sistem membandingkannya dengan stok tercatat:
      unit yang <b>tidak ditemukan</b>, unit yang <b>tidak terdaftar</b>, dan unit yang tercatat <b>sudah keluar</b> atau berada <b>di cabang lain</b>.
      Opname tidak mengubah stok; selisih ditindaklanjuti lewat menu transaksi.</p>
    <div class="row">
      ${isUtama() ? `<label style="max-width:280px">Cabang<select id="oCab">${opsiCabang(App.user.cabang)}</select></label>` : ''}
      <label>Catatan<input id="oCat" placeholder="mis. Opname akhir November 2026"></label>
      <button class="btn primary" id="oStart"></button>
    </div>
  </div>
  <div class="card"><h3>Riwayat opname</h3>
    <div class="table-wrap"><table><thead><tr><th>ID</th><th>Cabang</th><th>Mulai</th><th>Selesai</th><th>Status</th>
      <th class="num">Stok sistem</th><th class="num">Cocok</th><th class="num">Tidak ditemukan</th><th class="num">Selisih lain</th><th></th></tr></thead>
    <tbody>${list.length ? list.map(o => `<tr>
      <td class="mono">${esc(o.id)}</td><td>${esc(cabNama(o.cabang))}</td><td class="mono">${esc(o.dibuat.slice(0, 16))}</td>
      <td class="mono">${esc((o.selesai || '').slice(0, 16))}</td>
      <td>${o.status === 'DRAFT' ? '<span class="badge b-RETUR">BERJALAN</span>' : '<span class="badge b-on">SELESAI</span>'}</td>
      <td class="num">${o.status === 'DRAFT' ? '' : fmt(o.jml_sistem)}</td><td class="num">${o.status === 'DRAFT' ? '' : fmt(o.jml_cocok)}</td>
      <td class="num">${o.status === 'DRAFT' ? '' : `<b style="color:${+o.jml_hilang ? 'var(--bad)' : 'inherit'}">${fmt(o.jml_hilang)}</b>`}</td>
      <td class="num">${o.status === 'DRAFT' ? '' : fmt(o.jml_selisih)}</td>
      <td><button class="btn sm" data-o="${esc(o.id)}">${o.status === 'DRAFT' ? 'Lanjutkan' : 'Lihat hasil'}</button></td></tr>`).join('')
      : '<tr><td colspan="10" class="empty">Belum pernah opname</td></tr>'}</tbody></table></div></div>`;
  const cab = () => isUtama() ? $('#oCab').value : App.user.cabang;
  const label = () => { $('#oStart').textContent = draftOf(cab()) ? 'Lanjutkan opname berjalan' : 'Mulai opname'; };
  if (isUtama()) $('#oCab').onchange = label;
  label();
  $('#oStart').onclick = async () => {
    try {
      const o = await API.call('opnameMulai', { cabang: cab(), catatan: $('#oCat').value.trim() });
      opnameSesi(v, o.id);
    } catch (e) { toast(e.message, true); }
  };
  $$('[data-o]', v).forEach(b => b.onclick = () => {
    const o = list.find(x => x.id === b.dataset.o);
    if (o.status === 'DRAFT') opnameSesi(v, o.id); else opnameHasil(v, o.id);
  });
}

async function opnameSesi(v, id) {
  v.innerHTML = loadingHtml('Menyiapkan opname…');
  const [det, units] = await Promise.all([API.call('opnameDetail', { id }), API.call('stok', {})]);
  const o = det.opname;
  const ix = { imei: {}, sn: {} };
  units.forEach(u => { ix.imei[u.imei] = u; ix.sn[u.sn] = u; });
  const find = k => {
    const d = normImei(k);
    if (d.length === 15 && /^[\d\s-]+$/.test(k) && ix.imei[d]) return ix.imei[d];
    return ix.sn[normSn(k)] || null;
  };
  const sistem = units.filter(u => u.status === 'GUDANG' && u.cabang === o.cabang);
  const scanned = {};      // id (imei / SN mentah) -> { kunci, hasil, unit }
  det.items.forEach(it => { scanned[it.imei || normSn(it.kunci)] = { kunci: it.kunci, hasil: it.hasil, unit: ix.imei[it.imei], ket: it.keterangan }; });
  let pending = [], syncing = false, timer = null, tab = 'belum';

  const evalKey = k => {
    const x = find(k);
    const key = x ? x.imei : normSn(k);
    if (scanned[key]) return { key, hasil: 'DOBEL', unit: x };
    let h, ket = '';
    if (!x) h = 'TIDAK_TERDAFTAR';
    else if (x.status !== 'GUDANG') { h = 'SUDAH_KELUAR'; ket = 'Sistem: keluar ' + x.tgl_keluar.slice(0, 10) + ' ke ' + x.penerima; }
    else if (x.cabang !== o.cabang) { h = 'CABANG_LAIN'; ket = 'Sistem: stok ' + cabNama(x.cabang); }
    else h = 'COCOK';
    return { key, hasil: h, unit: x, ket };
  };

  v.innerHTML = `
  <div class="card op-head">
    <div><small class="muted">Opname berjalan · ${esc(o.id)}</small><h3 style="margin:2px 0 0">${esc(cabNama(o.cabang))}</h3>
      <small class="muted">Dimulai ${esc(o.dibuat)} oleh ${esc(o.dibuat_oleh)}${o.catatan ? ' · ' + esc(o.catatan) : ''}</small></div>
    <div class="row"><button class="btn ghost" id="oBack">‹ Kembali</button><button class="btn danger" id="oCancel">Batalkan</button>
      <button class="btn primary" id="oDone">Selesaikan opname</button></div>
  </div>
  <div class="stats">
    <div class="stat"><small>Stok tercatat di sistem</small><b id="cSis">0</b></div>
    <div class="stat hi"><small>Sudah discan &amp; cocok</small><b id="cCok">0</b></div>
    <div class="stat"><small>Belum discan</small><b id="cBel">0</b></div>
    <div class="stat"><small>Selisih (fisik ada, sistem beda)</small><b id="cSel" style="color:var(--bad)">0</b></div>
  </div>
  <div class="progress"><div id="oBar"></div></div>
  <div class="card">
    <div class="row scan-row">
      <label>Scan IMEI / SN<input id="oScan" autocomplete="off" placeholder="Scan unit satu per satu, Enter"></label>
      <button class="btn" id="oSync">Simpan progres</button>
    </div>
    <div id="oLast" class="last-scan"></div>
    <p class="hint" id="oPend"></p>
    <details style="margin-top:6px"><summary class="muted" style="cursor:pointer">Tempel banyak sekaligus</summary>
      <textarea id="oText" placeholder="Satu IMEI atau SN per baris" style="margin-top:8px"></textarea>
      <button class="btn" id="oPaste" style="margin-top:8px">Proses</button></details>
  </div>
  <div class="card">
    <div class="tabs" id="oTabs"><button data-t="belum" class="on">Belum discan</button><button data-t="selisih">Selisih</button><button data-t="cocok">Cocok</button></div>
    <div id="oList"></div>
  </div>`;

  const draw = () => {
    const vals = Object.values(scanned);
    const cocok = vals.filter(s => s.hasil === 'COCOK');
    const selisih = vals.filter(s => s.hasil !== 'COCOK');
    const belum = sistem.filter(u => !scanned[u.imei]);
    $('#cSis').textContent = fmt(sistem.length);
    $('#cCok').textContent = fmt(cocok.length);
    $('#cBel').textContent = fmt(belum.length);
    $('#cSel').textContent = fmt(selisih.length);
    $('#oBar').style.width = (sistem.length ? Math.min(100, cocok.length / sistem.length * 100) : 0) + '%';
    $$('#oTabs button').forEach(b => {
      const n = { belum: belum.length, selisih: selisih.length, cocok: cocok.length }[b.dataset.t];
      b.textContent = { belum: 'Belum discan', selisih: 'Selisih', cocok: 'Cocok' }[b.dataset.t] + ` (${n})`;
    });
    const unitRows = arr => arr.slice(0, 500).map(u => `<tr><td class="mono">${esc(u.imei)}</td><td class="mono">${esc(u.sn)}</td><td>${esc(u.model)}</td><td>${badge(kat(u))}</td></tr>`).join('');
    if (tab === 'belum') {
      $('#oList').innerHTML = `<div class="table-wrap"><table><thead><tr><th>IMEI</th><th>SN</th><th>Model</th><th>Kategori</th></tr></thead>
        <tbody>${belum.length ? unitRows(belum) : '<tr><td colspan="4" class="empty">Semua stok tercatat sudah discan 🎉</td></tr>'}</tbody></table></div>`;
    } else {
      const arr = tab === 'cocok' ? cocok : selisih;
      $('#oList').innerHTML = `<div class="table-wrap"><table><thead><tr><th>Discan</th><th>IMEI</th><th>SN</th><th>Model</th><th>Hasil</th><th>Keterangan</th></tr></thead>
        <tbody>${arr.length ? arr.slice().reverse().slice(0, 500).map(s => `<tr><td class="mono">${esc(s.kunci)}</td>
          <td class="mono">${esc(s.unit ? s.unit.imei : '')}</td><td class="mono">${esc(s.unit ? s.unit.sn : '')}</td>
          <td>${esc(s.unit ? s.unit.model : '')}</td><td>${hasilBadge(s.hasil)}</td><td>${esc(s.ket || '')}</td></tr>`).join('')
          : '<tr><td colspan="6" class="empty">Kosong</td></tr>'}</tbody></table></div>`;
    }
    $('#oPend').textContent = pending.length ? `${pending.length} scan belum tersimpan ke server${syncing ? ' — menyimpan…' : ''}` : 'Semua scan tersimpan ✓';
  };
  const sync = async () => {
    clearTimeout(timer);
    if (syncing || !pending.length) return;
    syncing = true; draw();
    const batch = pending.slice(0, 100);
    try {
      const res = await API.call('opnameScan', { id, keys: batch });
      pending = pending.slice(batch.length);
      // hasil server adalah acuan (mis. unit cabang lain tidak terlihat oleh admin cabang)
      res.hasil.forEach(h => {
        if (h.hasil === 'DOBEL') return;
        const ent = Object.values(scanned).find(s => s.kunci === h.kunci);
        if (ent && ent.hasil !== h.hasil) { ent.hasil = h.hasil; ent.ket = ent.ket || (HASIL[h.hasil] || [''])[0]; }
      });
    } catch (e) { toast('Gagal menyimpan scan: ' + e.message, true); }
    syncing = false; draw();
    if (pending.length) timer = setTimeout(sync, 500);
  };
  const add = k => {
    k = String(k || '').trim();
    if (!k) return null;
    const r = evalKey(k);
    if (r.hasil !== 'DOBEL') {
      scanned[r.key] = { kunci: k, hasil: r.hasil, unit: r.unit, ket: r.ket };
      pending.push(k);
    }
    return r;
  };
  const showLast = (k, r) => {
    const el = $('#oLast');
    el.className = 'last-scan ' + (r.hasil === 'COCOK' ? 'ok' : r.hasil === 'DOBEL' ? 'dup' : 'bad');
    el.innerHTML = `<b>${esc(k)}</b> ${hasilBadge(r.hasil)} ${r.unit ? esc(r.unit.model) : ''} ${r.ket ? '· ' + esc(r.ket) : ''}`;
  };
  $('#oScan').onkeydown = e => {
    if (e.key !== 'Enter') return;
    e.preventDefault();
    const k = $('#oScan').value.trim();
    $('#oScan').value = '';
    const r = add(k);
    if (!r) return;
    showLast(k, r);
    draw();
    clearTimeout(timer);
    timer = setTimeout(sync, pending.length >= 25 ? 0 : 6000);
  };
  $('#oPaste').onclick = () => {
    const lines = $('#oText').value.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    let n = 0;
    lines.forEach(k => { const r = add(k); if (r && r.hasil !== 'DOBEL') n++; });
    $('#oText').value = '';
    toast(`${n} unit diproses`);
    draw(); sync();
  };
  $('#oSync').onclick = sync;
  $$('#oTabs button').forEach(b => b.onclick = () => { tab = b.dataset.t; $$('#oTabs button').forEach(x => x.classList.toggle('on', x === b)); draw(); });
  const guard = e => { if (pending.length) { e.preventDefault(); e.returnValue = ''; } };
  window.addEventListener('beforeunload', guard);
  const leave = () => { window.removeEventListener('beforeunload', guard); };
  $('#oBack').onclick = async () => { await sync(); leave(); pageOpname(v); };
  $('#oCancel').onclick = () => {
    modal(`<h3>Batalkan opname?</h3><p class="muted" style="margin:0">Semua hasil scan opname ini akan diabaikan.</p>
      <div class="modal-actions"><button class="btn" id="cNo">Tidak</button><button class="btn danger" id="cYes">Ya, batalkan</button></div>`);
    $('#cNo').onclick = closeModal;
    $('#cYes').onclick = async () => {
      try { await API.call('opnameBatal', { id }); closeModal(); leave(); toast('Opname dibatalkan'); pageOpname(v); }
      catch (e) { toast(e.message, true); }
    };
  };
  $('#oDone').onclick = async () => {
    await sync();
    if (pending.length) { toast('Masih ada scan yang belum tersimpan, coba lagi', true); return; }
    const belum = sistem.filter(u => !scanned[u.imei]).length;
    modal(`<h3>Selesaikan opname?</h3>
      <p style="margin:0">${belum ? `<b style="color:var(--bad)">${fmt(belum)} unit</b> tercatat di gudang tetapi belum discan dan akan dicatat sebagai <b>tidak ditemukan</b>.`
        : 'Semua stok tercatat sudah discan.'}</p>
      <p class="muted" style="margin:0">Setelah selesai, opname tidak bisa discan lagi.</p>
      <div class="modal-actions"><button class="btn" id="cNo">Lanjut scan</button><button class="btn primary" id="cYes">Selesaikan</button></div>`);
    $('#cNo').onclick = closeModal;
    $('#cYes').onclick = async () => {
      $('#cYes').disabled = true;
      try { await API.call('opnameSelesai', { id }); closeModal(); leave(); toast('Opname selesai'); opnameHasil(v, id); }
      catch (e) { toast(e.message, true); $('#cYes').disabled = false; }
    };
  };
  draw();
  $('#oScan').focus();
}

async function opnameHasil(v, id) {
  v.innerHTML = loadingHtml();
  const r = await API.call('opnameDetail', { id });
  const o = r.opname, items = r.items;
  const grp = h => items.filter(i => i.hasil === h);
  const sel = items.filter(i => i.hasil !== 'COCOK' && i.hasil !== 'HILANG');
  const tbl = (arr, kunci) => `<div class="table-wrap"><table><thead><tr>${kunci ? '<th>Discan</th>' : ''}<th>IMEI</th><th>SN</th><th>Model</th><th>Kategori</th><th>Hasil</th><th>Keterangan</th></tr></thead>
    <tbody>${arr.length ? arr.map(i => `<tr>${kunci ? `<td class="mono">${esc(i.kunci)}</td>` : ''}<td class="mono">${esc(i.imei)}</td><td class="mono">${esc(i.sn)}</td>
      <td>${esc(i.model)}</td><td>${i.kategori ? badge(i.kategori) : ''}</td><td>${hasilBadge(i.hasil)}</td><td>${esc(i.keterangan)}</td></tr>`).join('')
      : `<tr><td colspan="${kunci ? 7 : 6}" class="empty">Tidak ada</td></tr>`}</tbody></table></div>`;
  v.innerHTML = `
  <div class="card op-head">
    <div><small class="muted">Hasil opname · ${esc(o.id)}</small><h3 style="margin:2px 0 0">${esc(cabNama(o.cabang))}</h3>
      <small class="muted">${esc(o.dibuat)} s/d ${esc(o.selesai)} · ${esc(o.selesai_oleh)}${o.catatan ? ' · ' + esc(o.catatan) : ''}</small></div>
    <div class="row"><button class="btn ghost" id="hBack">‹ Kembali</button><button class="btn" id="hExp">Ekspor Excel</button></div>
  </div>
  <div class="stats">
    <div class="stat"><small>Stok tercatat</small><b>${fmt(o.jml_sistem)}</b></div>
    <div class="stat hi"><small>Cocok</small><b>${fmt(o.jml_cocok)}</b></div>
    <div class="stat"><small>Tidak ditemukan fisik</small><b style="color:${+o.jml_hilang ? 'var(--bad)' : 'inherit'}">${fmt(o.jml_hilang)}</b></div>
    <div class="stat"><small>Selisih lain</small><b>${fmt(o.jml_selisih)}</b></div>
    <div class="stat"><small>Akurasi</small><b>${o.jml_sistem > 0 ? (o.jml_cocok / o.jml_sistem * 100).toFixed(1) + '%' : '-'}</b></div>
  </div>
  ${+o.jml_hilang || sel.length ? `<div class="alert bad"><b>Tindak lanjut:</b> unit <i>tidak ditemukan</i> dicari ulang atau dihapus dengan alasan
    (Lacak Unit → Hapus). Unit <i>tidak terdaftar</i> diinput lewat Barang Masuk. Unit <i>sudah keluar</i> dicek transaksinya,
    lalu diretur lewat Barang Masuk bila memang kembali. Unit <i>cabang lain</i> dirapikan lewat Mutasi.</div>`
    : `<div class="alert ok"><b>Stok fisik sama persis dengan sistem.</b></div>`}
  <div class="card"><h3>Tidak ditemukan fisik (${grp('HILANG').length})</h3>${tbl(grp('HILANG'))}</div>
  <div class="card"><h3>Selisih: fisik ada, data sistem berbeda (${sel.length})</h3>${tbl(sel, true)}</div>
  <div class="card"><details><summary style="cursor:pointer;font-weight:600">Cocok (${grp('COCOK').length})</summary><div style="margin-top:12px">${tbl(grp('COCOK'))}</div></details></div>`;
  $('#hBack').onclick = () => pageOpname(v);
  const cols = [['imei', 'IMEI', 18], ['sn', 'SN', 18], ['model', 'Model', 22], ['kategori', 'Kategori', 9], ['hasil_teks', 'Hasil', 26], ['keterangan', 'Keterangan', 40], ['kunci', 'Discan', 18]];
  items.forEach(i => { i.hasil_teks = (HASIL[i.hasil] || [i.hasil])[0]; });
  $('#hExp').onclick = () => exportExcelMulti(`opname-${o.cabang}-${o.dibuat.slice(0, 10)}`, [
    { name: 'Ringkasan', cols: [['k', 'Keterangan', 28], ['v', 'Nilai', 30]], rows: [
      { k: 'ID Opname', v: o.id }, { k: 'Cabang', v: cabNama(o.cabang) }, { k: 'Mulai', v: o.dibuat }, { k: 'Selesai', v: o.selesai },
      { k: 'Petugas', v: o.selesai_oleh }, { k: 'Catatan', v: o.catatan }, { k: 'Stok tercatat', v: o.jml_sistem },
      { k: 'Cocok', v: o.jml_cocok }, { k: 'Tidak ditemukan fisik', v: o.jml_hilang }, { k: 'Selisih lain', v: o.jml_selisih }] },
    { name: 'Tidak Ditemukan', cols, rows: grp('HILANG') },
    { name: 'Selisih', cols, rows: sel },
    { name: 'Cocok', cols, rows: grp('COCOK') }
  ]);
}

/* ---------------- Laporan bulanan ---------------- */
async function pageLaporan(v) {
  const bulanIni = today().slice(0, 7);
  v.innerHTML = `
  <div class="card">
    <div class="toolbar">
      <label>Bulan<input type="month" id="lpBln" value="${bulanIni}" max="${bulanIni}"></label>
      ${isUtama() ? `<label>Cabang<select id="lpCab">${opsiCabang('', 'Semua cabang', false)}</select></label>` : ''}
      <button class="btn primary" id="lpGo">Tampilkan</button>
      <span class="spacer"></span>
      <button class="btn" id="lpExp" disabled>Ekspor Excel</button>
    </div>
    <p class="hint" style="margin:0">Stok awal &amp; akhir dihitung ulang dari riwayat transaksi, jadi laporan bulan mana pun selalu bisa ditarik.
      Gunakan file ekspor untuk pencocokan dengan logistik MyRepublic.</p>
  </div>
  <div id="lpOut"></div>`;
  let data = null;
  const NAMA_BLN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  const blnTeks = b => NAMA_BLN[+b.slice(5, 7) - 1] + ' ' + b.slice(0, 4);
  const COLS = [['awal', 'Stok awal'], ['masuk', 'Masuk'], ['retur', 'Retur'], ['mutasi_masuk', 'Mutasi masuk'],
    ['terjual', 'Terjual'], ['demo_keluar', 'Demo keluar'], ['mutasi_keluar', 'Mutasi keluar'], ['hapus', 'Dihapus'],
    ['ubah_kategori', 'Ubah kategori'], ['akhir', 'Stok akhir']];
  const load = async () => {
    $('#lpOut').innerHTML = loadingHtml('Menghitung laporan…');
    try {
      data = await API.call('laporanBulanan', { bulan: $('#lpBln').value, cabang: isUtama() ? $('#lpCab').value : '' });
    } catch (e) { $('#lpOut').innerHTML = errBox(e); return; }
    const R = data.ringkasan;
    const tot = {};
    COLS.forEach(c => { tot[c[0]] = R.reduce((s, r) => s + r[c[0]], 0); });
    const sum = k => fmt(tot[k]);
    const cell = (k, n) => `<td class="num${k === 'akhir' || k === 'awal' ? ' strong' : ''}">${n ? (k === 'ubah_kategori' && n > 0 ? '+' : '') + fmt(n) : '<span class="muted">·</span>'}</td>`;
    let tab = 'terjual';
    $('#lpOut').innerHTML = `
      <div class="stats">
        <div class="stat"><small>Stok awal ${esc(blnTeks(data.bulan))}</small><b>${sum('awal')}</b></div>
        <div class="stat"><small>Masuk + retur</small><b>${fmt(tot.masuk + tot.retur)}</b></div>
        <div class="stat"><small>Terjual</small><b>${sum('terjual')}</b></div>
        <div class="stat demo"><small>Demo keluar</small><b>${sum('demo_keluar')}</b></div>
        <div class="stat hi"><small>Stok akhir</small><b>${sum('akhir')}</b></div>
      </div>
      <div class="card"><h3>Ringkasan per model — ${esc(blnTeks(data.bulan))} · ${esc(data.cabang ? cabNama(data.cabang) : 'Semua cabang')}</h3>
        <div class="table-wrap"><table class="lp"><thead><tr><th>Model</th><th>Kategori</th>${COLS.map(c => `<th class="num">${c[1]}</th>`).join('')}</tr></thead>
        <tbody>${R.length ? R.map(r => `<tr><td>${esc(r.model)}</td><td>${badge(r.kategori)}</td>${COLS.map(c => cell(c[0], r[c[0]])).join('')}</tr>`).join('')
          : `<tr><td colspan="${COLS.length + 2}" class="empty">Tidak ada data pada bulan ini</td></tr>`}</tbody>
        ${R.length ? `<tfoot><tr><td colspan="2"><b>Total</b></td>${COLS.map(c => `<td class="num"><b>${(c[0] === 'ubah_kategori' && tot[c[0]] > 0 ? '+' : '') + fmt(tot[c[0]])}</b></td>`).join('')}</tr></tfoot>` : ''}
        </table></div>
        <p class="hint">Stok akhir = stok awal + masuk + retur + mutasi masuk − terjual − demo keluar − mutasi keluar − dihapus ± ubah kategori.</p></div>
      <div class="card">
        <div class="tabs" id="lpTabs"><button data-t="terjual" class="on">Terjual (${data.terjual.length})</button>
          <button data-t="demo">Demo keluar (${data.demo.length})</button><button data-t="sisa">Sisa stok akhir bulan (${data.sisa.length})</button></div>
        <div id="lpList"></div></div>`;
    const drawList = () => {
      if (tab === 'sisa') {
        $('#lpList').innerHTML = `<div class="table-wrap"><table><thead><tr><th>IMEI</th><th>SN</th><th>Model</th><th>Kategori</th><th>Cabang</th></tr></thead>
          <tbody>${data.sisa.length ? data.sisa.slice(0, 500).map(u => `<tr><td class="mono">${esc(u.imei)}</td><td class="mono">${esc(u.sn)}</td>
            <td>${esc(u.model)}</td><td>${badge(u.kategori)}</td><td>${esc(cabNama(u.cabang))}</td></tr>`).join('') : '<tr><td colspan="5" class="empty">Kosong</td></tr>'}</tbody></table></div>
          ${data.sisa.length > 500 ? '<p class="hint">500 baris pertama ditampilkan; ekspor untuk semua.</p>' : ''}`;
      } else {
        const L = tab === 'terjual' ? data.terjual : data.demo;
        $('#lpList').innerHTML = `<div class="table-wrap"><table><thead><tr><th>Tanggal</th><th>IMEI</th><th>SN</th><th>Model</th><th>Pelanggan</th><th>Sales</th><th>Cabang</th><th>Referensi</th></tr></thead>
          <tbody>${L.length ? L.slice(0, 500).map(t => `<tr><td class="mono">${esc(t.waktu.slice(0, 16))}</td><td class="mono">${esc(t.imei)}</td><td class="mono">${esc(t.sn)}</td>
            <td>${esc(t.model)}</td><td>${esc(t.pelanggan)}</td><td>${esc(t.sales)}</td><td>${esc(cabNama(t.cabang))}</td><td>${esc(t.referensi)}</td></tr>`).join('')
            : '<tr><td colspan="8" class="empty">Kosong</td></tr>'}</tbody></table></div>`;
      }
    };
    $$('#lpTabs button').forEach(b => b.onclick = () => { tab = b.dataset.t; $$('#lpTabs button').forEach(x => x.classList.toggle('on', x === b)); drawList(); });
    drawList();
    $('#lpExp').disabled = false;
  };
  $('#lpGo').onclick = load;
  $('#lpExp').onclick = () => {
    if (!data) return;
    const trxCols = [['waktu', 'Tanggal', 19], ['imei', 'IMEI', 18], ['sn', 'SN', 18], ['model', 'Model', 22], ['cabang', 'Cabang', 10],
      ['pelanggan', 'Pelanggan', 24], ['no_ktp', 'No. KTP', 19], ['sales', 'Sales', 18], ['referensi', 'Referensi', 16]];
    exportExcelMulti(`laporan-${data.bulan}-${data.cabang || 'semua'}`, [
      { name: 'Ringkasan', cols: [['model', 'Model', 22], ['kategori', 'Kategori', 9]].concat(COLS.map(c => [c[0], c[1], 12])),
        rows: data.ringkasan, numeric: true },
      { name: 'Terjual', cols: trxCols, rows: data.terjual },
      { name: 'Demo Keluar', cols: trxCols, rows: data.demo },
      { name: 'Sisa Stok', cols: [['imei', 'IMEI', 18], ['sn', 'SN', 18], ['model', 'Model', 22], ['kategori', 'Kategori', 9], ['cabang', 'Cabang', 10]], rows: data.sisa }
    ]);
  };
  load();
}

/* ---------------- Cocokkan data MyRepublic ---------------- */
const MR = {
  COCOK: ['Cocok', 'b-on'],
  BEDA_STATUS: ['Status berbeda', 'b-RETUR'],
  BEDA_SN: ['SN berbeda', 'b-RETUR'],
  HANYA_MYREP: ['Hanya di data MyRepublic', 'b-HAPUS'],
  HANYA_KITA: ['Hanya di data kita', 'b-MUTASI'],
  DOBEL: ['IMEI dobel di file MyRepublic', 'b-HAPUS']
};
/** Status di file MyRepublic -> STOK | TERJUAL | DEMO | '' | teks asli */
function mrStatus(s) {
  const t = String(s || '').trim().toLowerCase();
  if (!t) return '';
  if (/demo/.test(t)) return 'DEMO';
  if (/(belum|not|un).{0,6}(jual|sold|install|aktif|activ)/.test(t)) return 'STOK';
  if (/(terjual|sold|jual|aktif|activ|install|terpasang|used|pelanggan|customer|keluar|delivered)/.test(t)) return 'TERJUAL';
  if (/(stok|stock|gudang|warehouse|available|tersedia|ready|sisa|on ?hand)/.test(t)) return 'STOK';
  return String(s).trim().toUpperCase();
}
const ourStatus = u => kat(u) === 'DEMO' ? 'DEMO' : (u.status === 'GUDANG' ? 'STOK' : 'TERJUAL');

async function pageCocok(v) {
  v.innerHTML = `
  <div class="card">
    <h3>Cocokkan dengan data logistik MyRepublic</h3>
    <div class="format-box">
      Unggah file Excel/CSV dari MyRepublic. Kolom yang dikenali dari judulnya: <b>IMEI</b> (wajib), <b>SN</b> / Serial (opsional),
      <b>Status</b> (opsional, mis. <i>Stock, Available, Sold, Installed, Terjual</i>). Data hanya dibandingkan di browser dan tidak disimpan.
    </div>
    <div class="row">
      ${isUtama() ? `<label style="max-width:260px">Bandingkan dengan cabang<select id="mrCab">${opsiCabang('', 'Semua cabang', false)}</select></label>` : ''}
      <label style="max-width:300px">Unit kita yang ikut dibandingkan<select id="mrScope">
        <option value="all">Semua unit (stok + terjual + demo)</option><option value="STOK">Hanya stok di gudang</option></select></label>
      <label class="btn primary file-btn">Pilih file MyRepublic…<input type="file" id="mrFile" accept=".xlsx,.xls,.csv" hidden></label>
    </div>
  </div>
  <div id="mrOut"></div>`;
  let ours = null, last = null;
  const rerun = () => { if (last) renderCocok.apply(null, last); };
  $('#mrScope').onchange = rerun;
  if (isUtama()) $('#mrCab').onchange = rerun;
  $('#mrFile').onchange = async e => {
    const file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    $('#mrOut').innerHTML = loadingHtml('Membaca & mencocokkan…');
    try {
      const [X, data] = await Promise.all([loadXLSX(), ours ? Promise.resolve(ours) : API.call('stok', {})]);
      ours = data;
      const wb = X.read(await file.arrayBuffer(), { type: 'array' });
      const rows = X.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: '' });
      const hi = rows.findIndex(r => r.some(c => /imei/i.test(String(c))));
      if (hi < 0) throw new Error('Kolom berjudul IMEI tidak ditemukan di sheet pertama.');
      const h = rows[hi].map(c => String(c).trim().toLowerCase());
      const ci = { imei: h.findIndex(c => /imei/.test(c)), sn: h.findIndex(c => /^(sn|s\/n|serial|no\.?\s*seri)/.test(c)), st: h.findIndex(c => /(status|kondisi|keterangan|state)/.test(c)) };
      const theirs = rows.slice(hi + 1).map((r, i) => ({
        baris: hi + i + 2, imei: normImei(cellText(r[ci.imei])), sn: ci.sn >= 0 ? normSn(cellText(r[ci.sn])) : '',
        status_asli: ci.st >= 0 ? cellText(r[ci.st]) : ''
      })).filter(r => r.imei);
      theirs.forEach(r => { r.status_mr = mrStatus(r.status_asli); });
      last = [file.name, theirs, ci];
      renderCocok(file.name, theirs, ci);
    } catch (err) { $('#mrOut').innerHTML = errBox({ message: 'Gagal: ' + err.message }); }
  };
  const renderCocok = (fname, theirs, ci) => {
    const cab = isUtama() ? $('#mrCab').value : '';
    const scope = $('#mrScope').value;
    const byImei = {};
    ours.forEach(u => { byImei[u.imei] = u; });
    const res = [], seen = {};
    theirs.forEach(t => {
      const u = byImei[t.imei];
      const row = { imei: t.imei, sn_mr: t.sn, status_asli: t.status_asli, status_mr: t.status_mr, baris: t.baris,
        sn: u ? u.sn : '', status_kita: u ? ourStatus(u) : '', cabang: u ? u.cabang : '', model: u ? u.model : '',
        penerima: u ? u.penerima : '', tgl_keluar: u ? u.tgl_keluar : '' };
      if (seen[t.imei]) { row.hasil = 'DOBEL'; res.push(row); return; }
      seen[t.imei] = true;
      if (!u) row.hasil = 'HANYA_MYREP';
      else if (t.sn && t.sn !== u.sn) row.hasil = 'BEDA_SN';
      else if (t.status_mr && t.status_mr !== row.status_kita) row.hasil = 'BEDA_STATUS';
      else row.hasil = 'COCOK';
      res.push(row);
    });
    ours.filter(u => !seen[u.imei] && (!cab || u.cabang === cab) && (scope === 'all' || ourStatus(u) === 'STOK')).forEach(u => res.push({
      imei: u.imei, sn: u.sn, status_kita: ourStatus(u), cabang: u.cabang, model: u.model, penerima: u.penerima,
      tgl_keluar: u.tgl_keluar, hasil: 'HANYA_KITA', sn_mr: '', status_mr: '', status_asli: ''
    }));
    const n = k => res.filter(r => r.hasil === k).length;
    let tab = n('HANYA_MYREP') ? 'HANYA_MYREP' : n('BEDA_STATUS') ? 'BEDA_STATUS' : 'HANYA_KITA';
    const ket = r => ({
      HANYA_MYREP: 'Belum diinput di sistem kita (Barang Masuk) atau IMEI salah ketik',
      HANYA_KITA: r.status_kita === 'STOK' ? 'Ada di stok kita, tidak ada di data MyRepublic' : 'Tercatat ' + r.status_kita.toLowerCase() + ' di sistem kita',
      BEDA_STATUS: `MyRepublic: ${r.status_asli} · kita: ${r.status_kita}${r.penerima ? ' (' + r.penerima + ')' : ''}`,
      BEDA_SN: `SN MyRepublic ${r.sn_mr} ≠ SN kita ${r.sn}`,
      DOBEL: 'IMEI muncul lebih dari sekali di file (baris ' + r.baris + ')',
      COCOK: ''
    }[r.hasil]);
    res.forEach(r => { r.keterangan = ket(r); r.hasil_teks = MR[r.hasil][0]; r.cabang_nama = r.cabang ? cabNama(r.cabang) : ''; });
    const tStok = theirs.filter(t => t.status_mr === 'STOK').length, tJual = theirs.filter(t => t.status_mr === 'TERJUAL').length;
    $('#mrOut').innerHTML = `
      <div class="alert ${n('COCOK') === res.length ? 'ok' : 'bad'}" style="background:#f8f9fb;color:var(--ink-2)">
        File <b>${esc(fname)}</b>: ${fmt(theirs.length)} baris IMEI${ci.st >= 0 ? ` · menurut MyRepublic ${fmt(tStok)} stok, ${fmt(tJual)} terjual` : ' · tanpa kolom status'}${ci.sn < 0 ? ' · tanpa kolom SN' : ''}.
      </div>
      <div class="stats">
        <div class="stat hi"><small>Cocok</small><b>${fmt(n('COCOK'))}</b></div>
        <div class="stat"><small>Status berbeda</small><b>${fmt(n('BEDA_STATUS'))}</b></div>
        <div class="stat"><small>Hanya di MyRepublic</small><b>${fmt(n('HANYA_MYREP'))}</b></div>
        <div class="stat"><small>Hanya di data kita</small><b>${fmt(n('HANYA_KITA'))}</b></div>
        <div class="stat"><small>SN beda / dobel</small><b>${fmt(n('BEDA_SN') + n('DOBEL'))}</b></div>
      </div>
      <div class="card">
        <div class="toolbar" style="margin-bottom:0"><div class="tabs" id="mrTabs" style="flex:1;margin:0">
          ${['BEDA_STATUS', 'HANYA_MYREP', 'HANYA_KITA', 'BEDA_SN', 'DOBEL', 'COCOK'].map(k => `<button data-t="${k}" class="${k === tab ? 'on' : ''}">${MR[k][0]} (${n(k)})</button>`).join('')}
        </div><button class="btn" id="mrExp">Ekspor hasil</button></div>
        <div id="mrList" style="margin-top:14px"></div>
      </div>`;
    const draw = () => {
      const L = res.filter(r => r.hasil === tab);
      $('#mrList').innerHTML = `<div class="table-wrap"><table><thead><tr><th>IMEI</th><th>SN kita</th><th>Model</th><th>Cabang</th><th>Status kita</th><th>Status MyRepublic</th><th>Keterangan</th></tr></thead>
        <tbody>${L.length ? L.slice(0, 500).map(r => `<tr><td class="mono"><a href="#lacak" data-q="${esc(r.imei)}">${esc(r.imei)}</a></td>
          <td class="mono">${esc(r.sn)}</td><td>${esc(r.model)}</td><td>${esc(r.cabang_nama)}</td>
          <td>${r.status_kita ? badge(r.status_kita) : ''}</td>
          <td>${esc(r.status_asli)}</td><td>${esc(r.keterangan)}</td></tr>`).join('')
          : '<tr><td colspan="7" class="empty">Tidak ada</td></tr>'}</tbody></table></div>
        ${L.length > 500 ? '<p class="hint">500 baris pertama ditampilkan; ekspor untuk semua.</p>' : ''}`;
    };
    $$('#mrTabs button').forEach(b => b.onclick = () => { tab = b.dataset.t; $$('#mrTabs button').forEach(x => x.classList.toggle('on', x === b)); draw(); });
    $('#mrList').onclick = e => { const a = e.target.closest('[data-q]'); if (a) sessionStorage.setItem('vg_lacak', a.dataset.q); };
    const cols = [['imei', 'IMEI', 18], ['sn', 'SN Kita', 18], ['sn_mr', 'SN MyRepublic', 18], ['model', 'Model', 22], ['cabang_nama', 'Cabang', 14],
      ['status_kita', 'Status Kita', 11], ['status_asli', 'Status MyRepublic', 16], ['penerima', 'Pelanggan', 22], ['tgl_keluar', 'Tgl Keluar', 19],
      ['hasil_teks', 'Hasil', 24], ['keterangan', 'Keterangan', 44]];
    $('#mrExp').onclick = () => exportExcelMulti('pencocokan-myrepublic-' + today(), [
      { name: 'Ringkasan', cols: [['k', 'Hasil', 30], ['v', 'Jumlah', 10]], rows: Object.keys(MR).map(k => ({ k: MR[k][0], v: n(k) })), numeric: true },
      { name: 'Perlu Dicek', cols, rows: res.filter(r => r.hasil !== 'COCOK') },
      { name: 'Cocok', cols, rows: res.filter(r => r.hasil === 'COCOK') }
    ]);
    draw();
  };
}

const PAGES = {
  dashboard: pageDashboard, masuk: pageMasuk, keluar: pageKeluar, mutasi: pageMutasi,
  stok: pageStok, riwayat: pageRiwayat, lacak: pageLacak, pelanggan: pagePelanggan, sales: pageSales,
  opname: pageOpname, laporan: pageLaporan, cocok: pageCocok,
  cabang: pageCabang, pengguna: pagePengguna, password: pagePassword
};

boot();
