# QA Test Management — Template

> Google Sheets template untuk manajemen test case, eksekusi, dan monitoring kualitas per modul/proyek.  
> Dikelola oleh **QA Team INA Digital**.

---

## Daftar Isi

1. [Overview](#overview)
2. [Cara Setup](#cara-setup)
3. [Struktur Tab](#struktur-tab)
4. [Panduan Pengisian](#panduan-pengisian)
5. [Konvensi TC_ID](#konvensi-tc_id)
6. [Prioritas & Blocker](#prioritas--blocker)
7. [RBAC & Role Testing](#rbac--role-testing)
8. [Bug Report](#bug-report)
9. [Performance Test](#performance-test)
10. [Summary & Dashboard](#summary--dashboard)
11. [Integrasi Portfolio Dashboard](#integrasi-portfolio-dashboard)

---

## Overview

Template ini mencakup satu modul/proyek secara lengkap:

- **TC_Master** — repository semua test case Web/Mobile
- **API_Master** — repository semua test case API
- **TC_Execution** — log eksekusi per tanggal run (Web/Mobile)
- **API_Execution** — log eksekusi per tanggal run (API)
- **BugReport** — log bug Web, Mobile, dan API dalam satu tab
- **Summary** — KPI overview, komposisi status, trend, coverage, bug summary
- **PerfTest** — rekam dan evaluasi hasil performance test (K6/JMeter)
- **Appendix** — definisi, panduan, dan referensi

---

## Cara Setup

### Prasyarat
- Akun Google dengan akses Google Sheets
- Google Apps Script (via Extensions > Apps Script)

### Langkah Instalasi

1. Buat **Google Sheets baru** (beri nama sesuai modul, misal: `QA - MOD-AUTH`)
2. Buka **Extensions > Apps Script**
3. Hapus semua kode default di editor
4. Paste seluruh isi file `QA_Sheet_v38.js`
5. Klik **Save** (Ctrl+S)
6. Pilih function **`createQASheet`** di dropdown
7. Klik **Run** — izinkan permission bila diminta
8. Tunggu hingga muncul dialog konfirmasi berhasil (~30-60 detik)

> **Catatan:** Proses create memerlukan waktu karena Apps Script membuat 8 sheet sekaligus dengan conditional formatting dan formula. Jangan tutup browser.

---

## Struktur Tab

| Tab | Warna | Fungsi |
|-----|-------|--------|
| **Summary** | Biru tua | Dashboard utama: KPI, pie chart, trend, coverage, bug summary |
| **Appendix** | Abu-abu | Panduan & referensi |
| **BugReport** | Biru | Log bug Web, Mobile, dan API |
| **TC_Execution** | Biru | Log eksekusi test case Web/Mobile per run |
| **TC_Master** | Biru | Repository TC Web/Mobile |
| **API_Execution** | Ungu | Log eksekusi test case API per run |
| **API_Master** | Ungu | Repository TC API |
| **PerfTest** | Biru tua | Rekam hasil performance test |

---

## Panduan Pengisian

### TC_Master

| Kolom | Keterangan |
|-------|-----------|
| No | Nomor urut |
| SubModul | Kode sub-modul, misal: `1.1`, `1.2` |
| TC_ID | ID unik — lihat [Konvensi TC_ID](#konvensi-tc_id) |
| Feature | Nama fitur/halaman spesifik |
| Priority | Critical / High / Medium / Low / Lowest |
| Platform | Web / Mobile / Web & Mobile |
| Test Type | Positive / Negative / Edge Case |
| Automated | Automated / Manual / To Do / Cannot be Automated |
| Version | Versi fitur yang diuji |
| Role | Role RBAC yang menjalankan skenario |
| Scenario | Deskripsi skenario pengujian |
| Steps | Langkah-langkah test |
| Expected | Hasil yang diharapkan |
| Test Level | Smoke / Regression |

### TC_Execution

- **Kolom A–G**: Auto-sync dari TC_Master (jangan diedit)
- **Kolom H dst**: Satu kolom per tanggal run — isi dengan `PASSED`, `FAILED`, `BLOCKED`, atau `IN PROGRESS`
- **Kolom Z (LATEST STATUS)**: Otomatis dihitung dari run terakhir
- **Kolom Screenshot (AA dst)**: **10 kolom screenshot dinamis** — satu kolom per tanggal run, label otomatis mengambil tanggal dari baris 2

#### Screenshot Dinamis

Setiap run memiliki kolom screenshot-nya sendiri. Header kolom screenshot otomatis menyesuaikan tanggal run di baris 2, sehingga:

```
Run: 2025-01-10  → Screenshot: "2025-01-10 Shot"
Run: 2025-02-05  → Screenshot: "2025-02-05 Shot"
```

Paste link Google Drive, Jira attachment, atau URL gambar di kolom yang sesuai dengan run-nya.

#### Menambah Run Baru

1. Tambahkan kolom baru di sebelah kanan run terakhir (sebelum kolom Z)
2. Isi baris 2 dengan tanggal (format `YYYY-MM-DD`)
3. Isi baris 3 dengan nama run (misal: `Sprint 12 - Reg`)
4. Isi status di setiap baris data
5. Kolom screenshot untuk run baru otomatis terbentuk di kolom AA+

---

## Konvensi TC_ID

### Web/Mobile

```
Format : [APP].[FEAT].[000]

[APP]  = Kode aplikasi  → WEB, MOB, ADM, USR, SHP
[FEAT] = Kode fitur     → LOG, DASH, PRF, CHK, RPT
[000]  = Nomor urut 3 digit, mulai 001

Contoh:
  WEB.LOG.001   = Web, Login, TC ke-1
  WEB.LOG.002   = Web, Login, TC ke-2 (negative)
  MOB.DASH.001  = Mobile, Dashboard, TC ke-1
  ADM.USR.015   = Admin, User Mgmt, TC ke-15
```

### API

```
Format : API.[SVC].[FEAT].[000]

[SVC]  = Kode service   → AUTH, USER, ORD, PAY, INV
[FEAT] = Kode endpoint  → LOG, LIST, CRT, UPD, DEL
[000]  = Nomor urut 3 digit, mulai 001

Contoh:
  API.AUTH.LOG.001  = Auth, Login endpoint, TC ke-1
  API.USER.CRT.002  = User, Create, TC ke-2 (negative)
  API.PAY.CHK.005   = Payment, Checkout, TC ke-5
```

#### Aturan Penting

- TC_ID **wajib unik** — jangan pernah reuse ID yang sudah ada
- **Jangan ubah** TC_ID yang sudah punya history di Execution
- Urutkan: **Positive dulu** (001) → Negative → Edge Case

---

## Prioritas & Blocker

| Priority | Dampak | Konsekuensi jika FAIL |
|----------|--------|----------------------|
| **Critical** | Blocker utama | Release **DITAHAN** — wajib PASS |
| **High** | Blocker | Harus PASS di sprint yang sama — perlu approval PM |
| **Medium** | Potential blocker | Fix sebelum UAT — flagged ke Tech Lead |
| **Low** | Non-blocker | Fix di sprint berikutnya |
| **Lowest** | Nice to have | Opsional |

> **Smoke Test** = TC dengan priority Critical + High + Medium yang wajib dijalankan setiap run.  
> **Regression** = TC dengan priority Low + Lowest untuk full cycle testing.

---

## RBAC & Role Testing

Kolom **Role** di TC_Master/API_Master merujuk pada **peran pengguna (RBAC)** yang menjalankan skenario.

**Contoh role:** Admin, User, Viewer, Operator, Supervisor, Guest, Super Admin

### Pola Skenario RBAC

Untuk setiap endpoint/fitur sensitif, buat minimal 3 TC:

1. **Role yang BERHAK** → expected: `200/201`
2. **Role yang TIDAK berhak** → expected: `403 Forbidden`
3. **Tanpa token/login** → expected: `401 Unauthorized`

---

## Bug Report

Tab **BugReport** menampung semua bug dari Web, Mobile, dan API dalam satu tempat. Kolom **Type** digunakan sebagai pembeda.

### Kolom BugReport

| Group | Kolom | Keterangan |
|-------|-------|-----------|
| Identification | Bug ID | Format: `BUG-WEB-001` / `BUG-MOB-001` / `BUG-API-001` |
| | Type | Web / Mobile / API |
| | Priority | Critical / High / Medium / Low |
| | Status | Open / In Progress / Fixed / Verified / Closed / Won't Fix / Reopen |
| Classification | Feature | Nama fitur yang terdampak |
| | SubModul | Kode sub-modul |
| Detail | Title / Summary | Deskripsi singkat bug |
| | Environment | Dev / Staging UAT / Production / All |
| | Steps to Reproduce | Langkah-langkah untuk mereproduksi |
| | Expected Result | Hasil yang seharusnya |
| | Actual Result | Hasil aktual yang terjadi |
| Ownership | Related TC_ID | TC_ID yang berkaitan (jika ada) |
| | Reported By | Nama QA yang menemukan |
| | Assigned To | Nama developer yang handle |
| Timeline | Date Found | Tanggal bug ditemukan |
| | Date Fixed | Tanggal bug diperbaiki |
| Reference | Sprint | Sprint tempat bug ditemukan |
| | Jira / Link | Link tiket Jira atau issue tracker |
| | Notes | Catatan tambahan (browser, device, dll) |
| | Screenshot / Evidence | Link screenshot atau video evidence |

### Bug ID Convention

```
Format: BUG-[TYPE]-[000]

BUG-WEB-001  = Web/UI bug ke-1
BUG-MOB-001  = Mobile bug ke-1
BUG-API-001  = API bug ke-1

Nomor urut 3 digit, mulai 001. Jangan reuse ID.
```

### Priority sebagai Indikator Blocker

Berbeda dari TC Priority, Bug Priority menentukan apakah bug tersebut dihitung sebagai **blocker** di dashboard:

| Priority | Warna | Dihitung Blocker? |
|----------|-------|-------------------|
| **Critical** | Merah | Ya — showstopper, release DITAHAN |
| **High** | Oranye | Ya — fitur utama tidak bisa digunakan |
| **Medium** | Biru muda | **Ya** — ada workaround tapi experience buruk |
| **Low** | Hijau muda | Tidak — minor/kosmetik |

> Bug dengan status Open / In Progress / Reopen dan priority Medium ke atas semua dihitung sebagai **Blocker** di Portfolio Dashboard.

### Status Flow

```
Open → In Progress → Fixed → Verified → Closed
                   ↘ Won't Fix
         Fixed ← Reopen (jika masih reproducible)
```

### Screenshot / Evidence

Kolom terakhir (Screenshot / Evidence) — paste link:
- Google Drive (pastikan akses sudah di-share ke reviewer)
- Jira attachment URL
- Direct image URL

---

## Performance Test

Tab **PerfTest** digunakan untuk mencatat dan mengevaluasi hasil performance test (K6, JMeter, Locust, dll).

### Kolom Input

| Kolom | Keterangan |
|-------|-----------|
| Scenario | Nama skenario test |
| RPS Actual | Request per second yang dicapai |
| Error % | Persentase error |
| P90 / P95 / P99 | Response time percentile (ms) |
| VU | Jumlah Virtual User |
| CPU / Memory | Resource usage (%) |

### Threshold (Baris 11)

Isi threshold sesuai SLA yang disepakati. STATUS otomatis PASS/FAIL berdasarkan perbandingan actual vs threshold:

- **RPS**: PASS jika actual **>=** threshold
- **Error%, P90, P95, P99, CPU, Memory**: PASS jika actual **<=** threshold

---

## Summary & Dashboard

Tab **Summary** berisi:

| Section | Konten |
|---------|--------|
| **TEST DESCRIPTION** | Info sesi: project, sprint, QA lead, PIC QA, environment, status |
| **A. Status Overview** | KPI cards: Total, Passed, Failed, Blocked, Pass Rate, Auto Rate, Exec Rate |
| **B. Komposisi Status** | Pie chart dengan persentase per status (Web+Mobile & API terpisah) |
| **C. Trend Eksekusi** | Line chart pass rate per tanggal run — otomatis dari Execution |
| **D. Coverage per SubModul** | Tabel coverage: total, smoke, regression, auto%, pass% |
| **E. Bug Summary** | KPI cards bug: Total, Open, In Progress, Fixed, Verified, Critical, High, Medium (Blocker) |

### Section E: Bug Summary

Section E otomatis menghitung dari tab BugReport mulai **baris 5** (baris pertama data). Data dibagi menjadi dua panel:
- Kiri: **Web + Mobile** (Type = Web atau Mobile)
- Kanan: **API** (Type = API)

### Interpretasi Warna

| Warna | Arti |
|-------|------|
| Hijau (>=80%) | Target tercapai |
| Kuning (50-79%) | Perlu perhatian |
| Merah (<50%) | Critical — perlu tindakan segera |

---

## Integrasi Portfolio Dashboard

Setelah template ini terisi, Spreadsheet ID file ini bisa didaftarkan ke **QA Dashboard** (Portfolio) untuk monitoring lintas modul.

Cara mendapatkan Spreadsheet ID:
```
URL: https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit
```

Data berikut dari tab Summary akan **otomatis ditarik** ke Portfolio Dashboard saat refresh:

| Field Summary | Sel | Tampil di Portfolio |
|---|---|---|
| `Project / Sprint:` | B2 | Kolom Project / Sprint |
| `PIC QA:` | B5 | Kolom PIC / Team / Squad |

Pastikan kedua field ini sudah terisi di Summary agar Portfolio Dashboard menampilkan data yang akurat.

Lihat dokumentasi: [README QA Dashboard](./README_QA_PortfolioDashboard.md)

---

## Kontak

**QA Team INA Digital**  
Email: departemen.qa@inadigital.co.id

---

*Dokumentasi ini berlaku untuk QA_Sheet v38+*