// Lapisan komunikasi ke Apps Script (GET saja, tanpa preflight CORS)
const Auth = {
  get() {
    try { return JSON.parse(localStorage.getItem('vg_auth') || 'null'); } catch (e) { return null; }
  },
  set(v) { try { localStorage.setItem('vg_auth', JSON.stringify(v)); } catch (e) {} },
  clear() { try { localStorage.removeItem('vg_auth'); } catch (e) {} },
  token() { const a = Auth.get(); return a ? a.token : ''; }
};

class ApiError extends Error {
  constructor(msg, errors, code) { super(msg); this.errors = errors || []; this.code = code; }
}

const API = {
  async call(action, data) {
    if (!CONFIG.API_URL || CONFIG.API_URL.indexOf('http') !== 0) {
      throw new ApiError('API_URL belum diisi di js/config.js');
    }
    const url = new URL(CONFIG.API_URL);
    url.searchParams.set('action', action);
    const token = Auth.token();
    if (token) url.searchParams.set('token', token);
    if (data) url.searchParams.set('data', JSON.stringify(data));
    url.searchParams.set('_', Date.now());
    let res;
    try {
      res = await fetch(url.toString(), { method: 'GET', redirect: 'follow' });
    } catch (e) {
      throw new ApiError('Tidak dapat terhubung ke server. Periksa koneksi internet.');
    }
    let j;
    try { j = await res.json(); } catch (e) { throw new ApiError('Respon server tidak valid.'); }
    if (!j.ok) {
      if (j.code === 'AUTH' && action !== 'login') {
        Auth.clear();
        setTimeout(() => location.reload(), 1200);
      }
      throw new ApiError(j.error || 'Terjadi kesalahan.', j.errors, j.code);
    }
    return j.data;
  }
};

async function sha256(s) {
  const b = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s));
  return Array.from(new Uint8Array(b)).map(x => x.toString(16).padStart(2, '0')).join('');
}
