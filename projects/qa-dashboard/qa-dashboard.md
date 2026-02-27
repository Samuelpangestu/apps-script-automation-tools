# QA Dashboard

> Apps Script aggregator untuk monitoring kualitas **lintas modul/proyek** secara real-time.  
> Auto-refresh setiap 1 jam. Mendukung 11–20 modul aktif secara bersamaan.  
> Dikelola oleh **QA Team INA Digital**.

---

## Daftar Isi

1. [Overview](#overview)
2. [Arsitektur](#arsitektur)
3. [Cara Setup](#cara-setup)
4. [Struktur Tab](#struktur-tab)
5. [Mengisi Config](#mengisi-config)
6. [Menjalankan Refresh](#menjalankan-refresh)
7. [Auto-Refresh Trigger](#auto-refresh-trigger)
8. [Membaca Overview](#membaca-overview)
9. [Charts & Visualisasi](#charts--visualisasi)
10. [Referensi Fungsi](#referensi-fungsi)
11. [Troubleshooting](#troubleshooting)

---

## Overview

QA Dashboard adalah **Google Sheets terpisah** yang secara otomatis mengambil data dari semua modul sheet (template QA Test Management) dan menampilkannya dalam satu tampilan terintegrasi untuk stakeholder.

### Yang bisa dipantau

| Fitur | Keterangan |
|-------|-----------|
| **RAG Status** | Merah/Kuning/Hijau per modul berdasarkan Pass Rate |
| **Blocker Alert** | TC Critical & High yang FAILED/BLOCKED + bug priority Medium→Critical yang masih Open |
| **Bug Tracking** | Total bug, open, blocker, critical per modul |
| **Coverage** | Breakdown per SubModul: total TC, passed, failed, auto% |
| **Charts** | Bar, pie, stacked column, trend line untuk stakeholder |
| **Trend History** | Snapshot pass rate + bug count setiap refresh |
| **Perf Result** | Status PASS/FAIL dari PerfTest tiap modul |

---

## Arsitektur

```
┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐
│  QA - MOD-AUTH  │  │  QA - MOD-USER  │  │  QA - MOD-PAY   │
│  (Sheet modul)  │  │  (Sheet modul)  │  │  (Sheet modul)  │
└────────┬────────┘  └────────┬────────┘  └────────┬────────┘
         │                   │                     │
         └───────────────────┼─────────────────────┘
                             │  Apps Script openById()
                             ▼
              ┌──────────────────────────┐
              │       QA DASHBOARD       │
              │  ┌──────────────────┐    │
              │  │ Config           │    │
              │  │ Overview (RAG)   │    │
              │  │ Blockers (alert) │    │
              │  │ Coverage         │    │
              │  │ History (trend)  │    │
              │  │ _Raw (cache)     │    │
              │  └──────────────────┘    │
              │  Auto-refresh / 1 jam    │
              └──────────────────────────┘
```

**Prasyarat:** Setiap file modul menggunakan template **QA Test Management v38+**.

---

## Cara Setup

### Prasyarat

- Akun Google dengan akses ke semua file modul
- Google Apps Script (V8 runtime)
- File modul sudah dibuat menggunakan `QA_Sheet_v38.js`

### Langkah Instalasi

1. **Buat Google Sheets baru** — beri nama: `QA Dashboard`
2. Buka **Extensions > Apps Script**
3. Hapus semua kode default
4. Paste seluruh isi file `QA_Portfolio_Dashboard.js`
5. Klik **Save** (Ctrl+S)
6. Pilih function **`createDashboard`** → klik **Run**
7. Izinkan permission yang diminta
8. Tunggu dialog konfirmasi muncul

### Permission yang Dibutuhkan

| Permission | Alasan |
|-----------|--------|
| `spreadsheets` | Baca/tulis dashboard sheet |
| `script.external_request` | Akses modul sheet via `openById()` |
| `script.scriptapp` | Setup trigger otomatis |

> Pilih **Advanced > Go to [project name]** jika muncul warning "app not verified" — ini normal untuk script internal.

---

## Struktur Tab

| Tab | Warna | Fungsi |
|-----|-------|--------|
| **Config** | Abu-abu gelap | Daftar modul dan Spreadsheet ID |
| **Overview** | Biru | Ringkasan semua modul dengan RAG status + charts |
| **Blockers** | Merah | Alert TC Critical/High yang FAIL/BLOCKED |
| **Coverage** | Hijau | Detail coverage per SubModul semua modul |
| **History** | Ungu | Log pass rate + bug count setiap refresh + trend chart |
| **_Raw** | Abu-abu | Cache data internal (jangan diedit manual) |

---

## Mengisi Config

Buka tab **Config** dan isi satu baris per modul:

| Kolom | Keterangan | Contoh |
|-------|-----------|--------|
| **Active (Y/N)** | Y = aktif di-pull, N = skip | `Y` |
| **Modul Name** | Nama modul/proyek | `MOD-AUTH` |
| **PIC / Team / Squad** | Otomatis diisi dari Summary modul setelah refresh | `Team Platform` |
| **Project / Sprint** | Otomatis diisi dari Summary modul setelah refresh | `Sprint 12` |
| **Spreadsheet ID** | ID dari URL file modul | `1BxiMVs0XRA5nF...` |
| **Link** | Otomatis terisi (jangan diedit) | *(auto)* |
| **Notes** | Keterangan tambahan | `Modul Authentication` |

### Auto-Populate dari Summary

Setelah `refreshDashboard()` berjalan, kolom **PIC / Team / Squad** dan **Project / Sprint** di Config akan **otomatis diisi/diupdate** dari tab Summary tiap modul:

| Kolom Config | Diambil dari Summary | Sel |
|---|---|---|
| PIC / Team / Squad | `PIC QA:` | B5 |
| Project / Sprint | `Project / Sprint:` | B2 |

Tidak perlu mengisi kolom ini secara manual — cukup pastikan Summary modul sudah terisi.

### Cara Mendapatkan Spreadsheet ID

Buka file modul di Google Sheets, ambil ID dari URL:

```
https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit
                                        ^^^^^^^^^^^^^^^^
                                        Ini yang dicopy
```

### Tips Config

- Set **Active = N** untuk modul yang tidak dalam sprint aktif — tidak di-pull, performa lebih cepat
- Modul dengan ID `PASTE_SPREADSHEET_ID_HERE` otomatis dilewati
- Bisa menampung hingga **20 modul aktif** secara bersamaan

---

## Menjalankan Refresh

### Manual

1. Buka Apps Script editor
2. Pilih function **`refreshDashboard`**
3. Klik **Run**

### Apa yang terjadi saat refresh

1. Baca daftar modul aktif dari Config
2. Loop setiap modul: buka file via `openById()`, baca TC_Master, TC_Execution, API_Master, API_Execution, PerfTest, BugReport, Summary
3. Hitung statistik: total, passed, failed, pass rate, exec rate, auto rate
4. Hitung bug stats: total, open, blocker (Medium→Critical + status Open/InProgress/Reopen), critical
5. Identifikasi blocker TC (Critical/High yang FAIL atau BLOCKED)
6. Tulis ke Overview, Blockers, Coverage
7. **Append** satu baris ke History (data historis tidak dihapus)
8. **Generate charts** otomatis di Overview dan History
9. Auto-update Config dengan PIC dan Sprint dari Summary tiap modul
10. Update timestamp "Last refreshed" di baris 2 Overview

> **Estimasi waktu:** ~2-5 detik per modul. Untuk 15 modul aktif: ~45-75 detik.

---

## Auto-Refresh Trigger

### Aktivasi (sekali saja)

1. Di Apps Script, pilih function **`setupTrigger`**
2. Klik **Run**
3. Trigger terdaftar — dashboard auto-refresh **setiap 1 jam**

### Cek Status Trigger

Apps Script editor → ikon jam 🕐 di sidebar kiri (*Triggers*) → lihat daftar trigger aktif.

### Ubah Jadwal

```javascript
// Ganti di fungsi setupTrigger() sebelum run:

// Setiap 1 jam (default)
.timeBased().everyHours(1).create();

// Setiap 6 jam
.timeBased().everyHours(6).create();

// Setiap hari jam 08:00
.timeBased().atHour(8).everyDays(1).create();

// Setiap Senin jam 07:00
.timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
```

---

## Membaca Overview

### Layout Overview

- **Baris 1**: Judul dashboard
- **Baris 2**: Timestamp *Last refreshed* (kanan atas)
- **Baris 3–4**: Group header dan kolom header
- **Baris 5+**: Data satu baris per modul
- **Baris terakhir**: TOTAL / AVERAGE
- **Di bawah tabel**: Charts otomatis

### Kolom Overview

| Kolom | Keterangan |
|-------|-----------|
| Modul | Nama modul |
| PIC / Team / Squad | PIC dan tim — otomatis dari Summary |
| Project / Sprint | Project & sprint — otomatis dari Summary |
| Web Total/Passed/Failed/Block | Statistik TC Web/Mobile |
| **Web Pass%** | Pass rate — RAG: hijau >=80%, kuning 50-79%, merah <50% |
| API Total/Passed/Failed/Block | Statistik TC API |
| **API Pass%** | Pass rate API — RAG sama |
| **Perf** | PASS / FAIL / -- dari PerfTest |
| **Bugs** | Total bug di BugReport |
| **Blocker** | Bug Open/InProgress/Reopen dengan priority Medium, High, atau Critical |
| **Critical** | Bug dengan priority Critical |
| Notes | Error info |

### Warna Otomatis

| Kondisi | Warna |
|---------|-------|
| Pass Rate >= 80% | Hijau |
| Pass Rate 50–79% | Kuning |
| Pass Rate < 50% | Merah |
| Failed > 0 | Merah (kolom Failed) |
| Blocked > 0 | Oranye (kolom Blocked) |
| Blocker bugs > 0 | Merah (kolom Blocker) |
| Critical bugs > 0 | Merah (kolom Critical) |
| Perf = PASS | Hijau |
| Perf = FAIL | Merah |

---

## Charts & Visualisasi

Charts otomatis digenerate setiap `refreshDashboard()` di bawah tabel Overview dan di tab History.

### Overview Charts

| Chart | Jenis | Isi |
|-------|-------|-----|
| Pass Rate per Module | Horizontal Bar | Web Pass% vs API Pass% per modul — biru tua |
| Bug Distribution | Donut Pie | Proporsi total bug per modul |
| TC Status per Module | Stacked Column | Passed (hijau) / Failed (merah) / Blocked (oranye) |
| Open Blockers & Critical | Horizontal Bar | Blocker count dan Critical bug count per modul |

### History Chart

| Chart | Jenis | Isi |
|-------|-------|-----|
| Pass Rate Trend | Line Chart | Web Pass% dan API Pass% over time — berguna untuk laporan sprint |

> Charts ditempatkan otomatis di bawah tabel data. Jika data belum ada, charts akan kosong dan terisi setelah refresh pertama.

---

## Referensi Fungsi

| Fungsi | Kapan dijalankan | Keterangan |
|--------|-----------------|-----------|
| `createDashboard()` | **Sekali** saat instalasi | Buat semua tab dan struktur |
| `refreshDashboard()` | Setiap refresh (manual/auto) | Pull data, hitung stats, generate charts |
| `setupTrigger()` | **Sekali** setelah createDashboard | Aktifkan auto-refresh per jam |

---

## Troubleshooting

### Error: "You do not have permission to access the requested document"

**Penyebab:** Akun yang menjalankan script tidak punya akses ke file modul.  
**Solusi:** Share file modul ke akun Google yang dipakai untuk run script (minimal Viewer).

### Modul tidak muncul di Overview setelah refresh

**Cek:**
1. Pastikan kolom Active = `Y` di Config
2. Pastikan Spreadsheet ID benar (tidak ada spasi, bukan URL penuh)
3. Cek Apps Script Logs: **View > Logs** — cari baris `ERROR [nama modul]`

### Bug stats semua 0 padahal ada data di BugReport

**Penyebab:** File modul tidak menggunakan template v38+ atau tab BugReport tidak ada.  
**Cek:** Buka file modul, pastikan tab `BugReport` ada dan data dimulai dari **baris 5**.

### Charts tidak muncul setelah refresh

**Penyebab:** Data belum ada saat chart dibuat (modul baru).  
**Solusi:** Run `refreshDashboard()` sekali lagi setelah data terisi.

### History terus bertambah sangat panjang

Ini by design — History menyimpan semua snapshot historis untuk trend chart.  
Jika terlalu panjang (>10.000 baris), hapus baris lama secara manual atau kurangi frekuensi refresh.

### Apps Script timeout (>6 menit)

**Penyebab:** Terlalu banyak modul aktif atau file modul sangat besar.  
**Solusi:** Kurangi jumlah modul aktif (set N untuk modul non-aktif di sprint ini).

---

## Kontak

**QA Team INA Digital**  
Email: departemen.qa@inadigital.co.id

Untuk pertanyaan teknis terkait Apps Script: buka issue di repository ini.

---

*Dokumentasi ini berlaku untuk QA_Portfolio_Dashboard v1.0+*  
*Membutuhkan QA Test Management template v38+ di setiap modul sheet*