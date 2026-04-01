# WhatsApp Announcement - VAPT Integration

## Message untuk Tim Security (Copy-paste ready):

---

🔒 *VAPT BLOCKER TRACKING - NEW FEATURE*

Halo tim Security! 👋

QA Dashboard sekarang sudah terintegrasi dengan VAPT Findings kalian!

*📊 Apa yang baru?*

✅ *Automatic Pull Data*
- Data dari "Ad Hoc VAPT" + "Regular VAPT" otomatis ditarik setiap refresh
- Real-time tracking semua aplikasi

✅ *VAPT Blocker Monitoring*
- *Blocker* = Medium + High + Critical findings yang masih *OPEN*
- Sorted berdasarkan jumlah blocker (terbanyak di atas)
- Color coding: 🔴 Red (ada blocker) | 🟢 Green (0 blocker)

✅ *Summary Metrics*
- Total aplikasi yang sudah di-VAPT
- Total blocker di semua aplikasi
- Breakdown by severity (Critical, High, Medium)

✅ *Historical Tracking*
- Daily snapshots untuk monitor progress
- Bisa lihat trend penurunan blocker

---

*📍 Cara Akses:*

1. Buka QA Dashboard Testing:
   https://docs.google.com/spreadsheets/d/[TESTING_DASHBOARD_ID]/edit

2. Lihat tab *"VAPT"* (warna orange)

3. Summary metrics di atas, detail table di bawah

---

*🎯 Benefit untuk Security Team:*

• Monitor blocker per aplikasi dengan mudah
• Prioritize fixing based on blocker count
• Track progress closure findings
• Data selalu up-to-date (auto-refresh)

---

*🔄 Update Schedule:*

Data di-refresh otomatis setiap kali QA Dashboard refresh (biasanya daily).

Manual refresh bisa dilakukan via Apps Script:
Extensions > Apps Script > Run `refreshDashboard()`

---

*📝 Notes:*

- Data source: VAPT Spreadsheet kalian (Ad Hoc + Regular VAPT)
- Target: *0 blocker* di semua aplikasi
- Fokus: Medium-Critical Open findings

---

*👥 Need Help?*

Contact: [Your Name/Team QA]

---

**Mari kita target 0 VAPT blocker! 🎯**

