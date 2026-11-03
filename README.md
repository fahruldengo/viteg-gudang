# VITEG Gudang

Sistem barang masuk/keluar modem FWA MyRepublic, dicatat per IMEI, SN, dan MAC address.
Frontend di GitHub Pages, backend Google Apps Script (GET saja), database Google Sheets.

## Struktur

```
github/              → upload ke repository GitHub
  index.html
  css/style.css
  js/config.js       → isi API_URL sekali, JANGAN ditimpa saat update
  js/api.js
  js/app.js
apps-script/
  Code.gs            → tempel ke editor Apps Script (jangan diupload ke GitHub)
```

## Instalasi

### 1. Google Sheets + Apps Script
1. Buat Google Spreadsheet baru, misalnya "VITEG Gudang DB".
2. Buka menu **Ekstensi → Apps Script**, hapus isi bawaan, lalu tempel seluruh isi `Code.gs`. Simpan.
3. Pilih fungsi `setup` di toolbar, lalu klik **Jalankan** dan izinkan akses.
   Sheet `Users`, `Cabang`, `Unit`, `Transaksi`, dan `Sesi` akan dibuat, beserta akun awal **admin / admin123**.
4. Klik **Terapkan → Deployment baru → Aplikasi web**:
   - Jalankan sebagai: **Saya**
   - Yang memiliki akses: **Siapa saja**
5. Salin URL Web App (berakhiran `/exec`).

### 2. GitHub Pages
1. Isi `API_URL` di `js/config.js` dengan URL tadi.
2. Upload isi folder `github/` ke repository lewat web GitHub.
3. Buka **Settings → Pages → Deploy from branch → main / (root)**.
4. Buka situsnya, login sebagai `admin` / `admin123`, lalu **segera ganti password**.

### 3. Pengaturan awal
1. Menu **Cabang**: tambahkan semua cabang.
2. Menu **Pengguna**: buat akun admin cabang dan pilih cabangnya.

## Hak akses

| Fitur | Admin Utama | Admin Cabang |
|---|---|---|
| Barang masuk | Ke cabang mana saja | Hanya ke cabangnya |
| Barang keluar | Semua cabang | Hanya stok cabangnya |
| Mutasi antar cabang | ✓ | – |
| Stok, riwayat, dashboard | Semua cabang | Hanya cabangnya |
| Lacak unit | ✓ | ✓ |
| Hapus unit salah input | ✓ | – |
| Kelola cabang & pengguna | ✓ | – |

## Aturan data
- **IMEI** wajib 15 digit. **SN** dan **MAC** wajib diisi. MAC boleh ditulis `aabbcc112233`, `AA-BB-CC-11-22-33`, atau `AA:BB:CC:11:22:33`; sistem menyimpannya sebagai `AA:BB:CC:11:22:33`.
- IMEI, SN, dan MAC tidak boleh dipakai dua unit berbeda.
- Unit berstatus KELUAR yang di-input masuk lagi dengan IMEI, SN, dan MAC yang sama otomatis dicatat sebagai **RETUR**.
- Penyimpanan bersifat semua-atau-tidak: kalau ada satu unit bermasalah, tidak ada yang disimpan dan daftar errornya ditampilkan.
- Data dikirim dalam paket 25 unit (masuk) dan 50 unit (keluar/mutasi) supaya URL GET tidak terlalu panjang.
- Semua kolom sheet berformat teks, jadi IMEI tidak berubah menjadi notasi ilmiah.
- Unit bisa dicari dengan IMEI, SN, atau MAC di halaman keluar, mutasi, dan lacak.

## Update aplikasi
- Perubahan frontend: timpa file di GitHub **kecuali `js/config.js`**.
- Perubahan `Code.gs`: tempel kode baru, lalu **Terapkan → Kelola deployment → Edit → Versi: Baru**. Dengan cara ini URL tetap sama.
