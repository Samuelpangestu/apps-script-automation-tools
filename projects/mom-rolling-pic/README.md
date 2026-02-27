# 🎲 MOM Rolling & PIC Reminder System

**Automated rotation system for meeting note-takers with integrated notifications**

> Fair PIC assignment for daily standup meetings with automatic Google Doc creation and multi-channel reminders.

---

## 📋 **Overview**

This Google Apps Script automates the assignment of "Person In Charge" (PIC) for recording Minutes of Meeting (MOM) during daily standup meetings. The system ensures fair rotation, sends automated reminders, and creates MOM documents in Google Drive.

---

## ✨ **Features**

### 🎲 **Smart Rotation**
- Weighted random assignment based on history
- Fair distribution across all team members
- Configurable schedule (Mon, Wed, Fri by default)
- Manual override option

### 📝 **Auto Documentation**
- Automatic Google Doc creation for each meeting
- Standardized MOM template
- Organized in dedicated Drive folder
- Title format: `[MOM] YYYY-MM-DD - Standup`

### 📧 **Multi-Channel Notifications**
- **Email:** Personalized reminder to assigned PIC
- **Google Chat:** Team announcement via webhook
- Customizable message templates
- Scheduled delivery before meeting time

### 📊 **History Tracking**
- Complete audit trail of all assignments
- Date, assigned person, attendance log
- Easy review and reporting
- Automatic logging

---

## 🛠️ **Tech Stack**

- **Google Apps Script** (V8 Runtime)
- **SpreadsheetApp** - Data storage & UI
- **DocumentApp** - MOM template generation
- **DriveApp** - File organization
- **MailApp** - Email notifications
- **UrlFetchApp** - Google Chat webhook
- **Time-based Triggers** - Automated scheduling

---

## 📦 **Components**

### **Sheet: Tim** (Team Roster)
Stores team member information:
- Name
- Email address

### **Sheet: Jadwal** (Schedule History)
Tracks all assignments:
- Date
- Assigned PIC
- Auto-populated

### **Sheet: Config** (Configuration)
Settings:
- Active days (Mon/Wed/Fri)
- Google Chat Webhook URL
- Drive Folder ID for MOM docs

---

## 🚀 **Setup Guide**

### **Step 1: Create Google Sheet**

1. Create new Google Sheet
2. Open **Extensions** → **Apps Script**
3. Copy code from `src/Code.js`
4. Save project

### **Step 2: Initial Configuration**

Run from custom menu:
```
Standup Roller → 1. Setup Awal (Jalankan Pertama Kali)
```

This creates:
- ✅ Sheet "Tim" with headers
- ✅ Sheet "Jadwal" with headers
- ✅ Sheet "Config" with default settings

### **Step 3: Configure Settings**

1. **Add Team Members** in sheet "Tim":
   ```
   | Nama          | Email                |
   |---------------|----------------------|
   | John Doe      | john@example.com     |
   | Jane Smith    | jane@example.com     |
   ```

2. **Set Drive Folder** (optional):
   - Create folder for MOM docs in Google Drive
   - Get Folder ID from URL: `https://drive.google.com/drive/folders/FOLDER_ID_HERE`
   - Paste in Config sheet

3. **Set Google Chat Webhook** (optional):
   - Create webhook in Google Chat Space
   - Paste URL in Config sheet

### **Step 4: Activate Automation**

Run from custom menu:
```
Standup Roller → 2. Aktifkan Trigger Otomatis
```

This creates time-based trigger to run every Mon/Wed/Fri at 8:00 AM.

---

## 💻 **Usage**

### **Automated (Recommended)**

Once triggers are set up:
1. Script runs automatically on scheduled days
2. PIC is randomly selected
3. Google Doc is created
4. Email & Chat notifications sent
5. Assignment logged to "Jadwal" sheet

### **Manual**

From custom menu:
- **Test Rolling Sekarang** - Trigger assignment manually
- **Test Buat MOM Doc Saja** - Create doc without rolling
- **Lihat History Jadwal** - Open schedule sheet

---

## 📧 **Notification Examples**

### **Email**
```
Subject: [Reminder] Kamu jadi PIC MOM hari ini

Hi John,

Reminder: Kamu ditugaskan sebagai PIC untuk mencatat MOM standup hari ini.

Tanggal: 27 Februari 2026
Doc MOM: [Link ke Google Doc]

Terima kasih!
```

### **Google Chat**
```
🎲 **Standup MOM Roller**

📅 Tanggal: 27 Februari 2026
👤 PIC: John Doe
📝 Doc MOM: [Link]

@John jangan lupa catat ya! 😊
```

---

## 🔧 **Configuration**

### **Change Schedule Days**

Edit in `Code.js`:
```javascript
var DAYS_ACTIVE = [1, 3, 5]; // Mon=1, Wed=3, Fri=5
// Change to [1,2,3,4,5] for Mon-Fri
```

### **Change Trigger Time**

Default: 8:00 AM
To change, modify in `setupTrigger()`:
```javascript
.atHour(8)  // Change to desired hour (0-23)
```

### **Customize Email Template**

Edit in `sendEmailToSelectedPerson()`:
```javascript
var subject = "Your custom subject";
var body = "Your custom message...";
```

---

## 📊 **How It Works**

### **Rolling Algorithm**

1. **Get team members** from "Tim" sheet
2. **Get assignment history** from "Jadwal" sheet
3. **Calculate weights**:
   - Recent assignments get lower weight
   - People not assigned recently get higher weight
4. **Weighted random selection**
5. **Log to history**

### **Document Creation**

1. Check if Drive Folder ID is configured
2. Create new Google Doc with template
3. Set title: `[MOM] YYYY-MM-DD - Standup`
4. Add to specified folder (or root if not configured)
5. Get shareable link

### **Notification Flow**

```
Trigger (Time-based)
   ↓
Roll Assignment
   ↓
Create MOM Doc
   ↓
Send Email → Assigned PIC
   ↓
Send Chat → Team Channel
   ↓
Log to Jadwal Sheet
```

---

## 🐛 **Troubleshooting**

### **Trigger Not Running**

1. Check if trigger exists:
   - **Extensions** → **Apps Script** → **Triggers** (clock icon)
2. Verify active days match current day
3. Check execution logs for errors

### **Email Not Sent**

1. Verify email addresses in "Tim" sheet
2. Check Gmail sending limits (500/day for personal accounts)
3. Review execution transcript for errors

### **Google Chat Not Working**

1. Verify webhook URL is correct
2. Test webhook manually with curl:
   ```bash
   curl -X POST -H 'Content-Type: application/json' \
     -d '{"text":"Test message"}' \
     YOUR_WEBHOOK_URL
   ```
3. Ensure webhook has not been revoked

### **Document Not Created**

1. Check Drive Folder ID is valid
2. Verify script has permission to access Drive
3. Check if folder is in Trash

---

## 📈 **Analytics**

Query assignment history:
```javascript
function getAssignmentStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var jadwal = ss.getSheetByName("Jadwal");
  var data = jadwal.getDataRange().getValues();

  // Count assignments per person
  var counts = {};
  data.slice(1).forEach(row => {
    var name = row[1];
    counts[name] = (counts[name] || 0) + 1;
  });

  Logger.log(counts);
}
```

---

## 🔒 **Permissions Required**

- ✅ **Spreadsheets** - Read/write sheet data
- ✅ **Gmail** - Send emails as user
- ✅ **Drive** - Create files in Drive
- ✅ **External requests** - Call Google Chat webhook
- ✅ **Triggers** - Run time-based automation

---

## 🎯 **Use Cases**

- Daily standup meeting management
- Weekly sync meeting assignments
- Sprint planning note-taker rotation
- Any recurring meeting with rotating responsibilities

---

## 📝 **Customization Ideas**

### **Add Backup PIC**
Select 2 people, primary + backup

### **Holiday Skip**
Check calendar API to skip company holidays

### **Reminder Follow-up**
Send reminder 1 hour before meeting

### **Integration with Calendar**
Auto-create calendar event with assigned PIC

### **Slack Integration**
Replace Google Chat with Slack webhook

---

## 📞 **Support**

For issues or questions:
- Check [Troubleshooting](#troubleshooting) section
- Review execution logs in Apps Script Editor
- Contact: departemen.qa@inadigital.co.id

---

## 📄 **License**

© INA Digital Team - Internal Tool

---

## 🔗 **Related Projects**

- [QA Dashboard](../qa-dashboard/) - Test management dashboard
- [QA Test Management](../qa-test-management/) - Test case template

---

<div align="center">

**Part of [Google Apps Script Portfolio](../../README.md)**

Made with ❤️ for better team collaboration

</div>
