# HR Master — Excel Add-in (POC)

Tab **HR Master** di ribbon Excel + panel HTML (task pane) untuk Onboarding, Change Organization,
Change Direct Report, dan Offboarding pada workbook **SSOT Employee Master.xlsx**.

## Cara kerjanya

```
manifest.xml  ──(di-upload ke Excel)──►  Excel tahu: nama add-in, tombol ribbon, alamat halaman
                                              │
                                              ▼
GitHub Pages (hosting gratis)  ──►  taskpane.html + core.js + app.js  (tampil di panel kanan Excel)
                                              │  Office.js
                                              ▼
                     Workbook: tblMaster · tblDept · tblJob · tblLog
```

| File | Isi |
|---|---|
| `manifest.xml` | "KTP" add-in: nama, ikon, tab ribbon, URL halaman. **Satu-satunya file yang di-upload ke Excel.** |
| `taskpane.html` | Tampilan panel (HTML + CSS). |
| `core.js` | Aturan bisnis (validasi, cascade, cek bawahan, cek rantai atasan). Tidak tergantung Excel. |
| `app.js` | Form, preview live, dan baca/tulis ke Excel lewat Office.js. |
| `demo-data.js` | Data **fiktif** untuk Demo Mode (halaman dibuka di browser biasa). |
| `commands.html` | File wajib manifest (kosong). |
| `assets/` | Ikon. |

## Setup (gratis)

### 1. Hosting di GitHub Pages
1. Daftar / login di https://github.com.
2. **New repository** → nama: `hr-master-addin` → **Public** → **Create repository**.
3. Klik **uploading an existing file** → drag **semua isi** folder ini (termasuk folder `assets`) → **Commit changes**.
4. **Settings → Pages** → Source: **Deploy from a branch** → Branch: `main` / `(root)` → **Save**.
5. Tunggu 1–2 menit, buka `https://USERNAME.github.io/hr-master-addin/taskpane.html`.
   Harus muncul panel dengan badge kuning **Demo mode**.

### 2. Sesuaikan manifest
Buka `manifest.xml` di Notepad → **Ctrl+H** → ganti `GITHUB_USERNAME` dengan username GitHub Anda
(huruf kecil semua) → **Replace All** → Save. Simpan file ini di laptop.

### 3. Pasang di Excel (sideload)
1. Upload **SSOT Employee Master.xlsx** ke OneDrive → buka di Excel web.
2. **Home → Add-ins → More Add-ins** (atau **Insert → Add-ins**) → tab **My Add-ins** →
   **Upload My Add-in** (kanan atas) → pilih `manifest.xml` → **Upload**.
3. Tab **HR Master** muncul di ribbon. Badge di panel berubah menjadi **Excel**.

> Sideload hanya berlaku untuk Anda dan bisa hilang jika cache browser dibersihkan — cukup upload ulang.
> Agar muncul untuk seluruh tim, admin M365 men-deploy manifest yang sama lewat **Admin center → Integrated apps**.

## Kalau mentok
| Gejala | Solusi |
|---|---|
| Tidak ada **Upload My Add-in** / diblok tenant | Pakai akun Microsoft pribadi (outlook.com, gratis) di office.com, upload workbook ke OneDrive pribadi, sideload di sana. |
| Panel kosong / "add-in error" | Cek URL di langkah 1.5 bisa dibuka, dan semua `GITHUB_USERNAME` sudah diganti. |
| "Gagal membaca workbook" | File yang terbuka harus SSOT Employee Master.xlsx (punya tabel tblMaster, tblDept, tblJob, tblLog). |
| Perubahan kode tidak muncul | Tunggu GitHub Pages selesai deploy (tab **Actions**), lalu tutup-buka panel atau refresh halaman Excel. |

## Catatan POC
- Password proteksi sheet (`SSOT2026`) ada di `app.js` dan repo ini publik — hanya untuk demo.
- Jangan upload data karyawan asli ke repo. Data asli tetap di workbook OneDrive.
- Office.js tidak memberi identitas user tanpa SSO, jadi **Submitted By** tetap diisi manual (diingat di browser).
