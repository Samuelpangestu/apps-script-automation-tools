# QA Test Management -- Standard Template

Script otomasi untuk membangun ekosistem Test Management di Google Sheets secara instan. Dirancang khusus untuk tim QA INA Digital (Peruri).

## 🚀 Fitur Utama
* **Instant Setup**: Membuat 7 tab standar QA (TC Master, Execution, API, PerfTest, Dashboard) hanya dengan satu klik.
* **Auto-Sync**: Data di tab Execution tersinkron otomatis dengan Master.
* **Smart Dashboard**: Grafik distribusi status (Pie Chart) dan Trend Pass Rate (Line Chart) yang update secara real-time.
* **Performance Monitoring**: Modul khusus untuk mencatat hasil K6/JMeter dengan perhitungan PASS/FAIL otomatis terhadap SLA.
* **RBAC Focus**: Kolom khusus untuk pengujian akses kontrol (Role Based Access Control).

## 🛠 Cara Penggunaan
1. Buka [Google Sheets](https://sheets.new).
2. Pergi ke menu **Extensions** > **Apps Script**.
3. Copy semua kode dari folder `src/` di repo ini ke editor Apps Script.
4. Simpan project, lalu jalankan fungsi `createQASheet()`.
5. Berikan izin (Authorize) saat diminta.

## 📊 Struktur Template
- **TC_Master**: Input test case Web/Mobile.
- **TC_Execution**: Log eksekusi per tanggal/sprint.
- **API_Master/Execution**: Manajemen testing endpoint API.
- **Summary**: Dashboard eksekutif dan coverage per modul.
- **PerfTest**: Evaluasi hasil Performance Test (Latency, RPS, Error Rate).
