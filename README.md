# Kas Kamu — Catatan Keuangan

Web app pencatatan keuangan yang tersambung ke Google Sheets.

## Struktur
```
public/index.html      ← aplikasi (frontend)
apps-script/Code.gs    ← backend (Google Apps Script)
vercel.json            ← konfigurasi Vercel
```

## 1. Backend (Google Apps Script)
1. Buka spreadsheet template → Extensions → Apps Script.
2. Hapus isi Code.gs, tempel isi `apps-script/Code.gs` → Save.
3. Deploy → New deployment → Web app. Execute as: **Me**, Who has access: **Anyone** → Deploy.
4. Salin URL yang berakhiran `/exec`.

## 2. Upload ke GitHub
1. Buat repo baru (mis. `kas-kamu`).
2. Upload seluruh isi folder ini.

## 3. Deploy ke Vercel
1. vercel.com → Add New → Project → Import repo `kas-kamu`.
2. Framework Preset: **Other**, Build Command: kosongkan, Output Directory: `public`.
3. Deploy → buka URL `*.vercel.app`. Setiap push ke GitHub otomatis ter-deploy ulang.

## 4. Sambungkan
Buka app → titik tiga kanan atas → Pengaturan → tempel URL `/exec` → Sambungkan.

## Pasang di Home Screen
- iPhone (Safari): Share → Add to Home Screen.
- Android (Chrome): ⋮ → Add to Home screen.

## Update
- Frontend: ganti `public/index.html`, commit & push.
- Backend: tempel ulang Code.gs → Deploy → Manage deployments → Edit → Version: New version.
