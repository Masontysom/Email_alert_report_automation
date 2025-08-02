# 📧 Automated Email Alert Report Generator

An end-to-end automated solution built at **Company Pvt. Ltd.** to parse Gmail alerts and generate visually formatted Excel reports for daily monitoring—replacing a previously manual process.

## 🧩 Project Background

I was initially assigned the responsibility of checking critical backup alerts and daily status emails manually each day at 11 AM. This involved:

* Logging into Gmail,
* Filtering and reading emails from multiple client groups,
* Extracting details like alert time, message, device, and severity,
* Copy-pasting them into a formatted Excel sheet.

This repetitive and time-sensitive task demanded precision, speed, and consistency. Recognizing the inefficiency and error-prone nature of the manual process, I proposed and developed a fully automated solution.

## 🚀 Solution Overview

The **Email Alert Report Generator** is a Python-based utility (compiled into a standalone `.exe`) that:

* Connects securely to a Gmail inbox via **IMAP**.
* Searches for emails from **today and yesterday** matching critical alert criteria.
* Extracts:

  * Customer group,
  * Alert time,
  * Device name,
  * Alert message,
  * Email subject and embedded HTML content.
* Filters and formats the alerts intelligently.
* Saves the result into a **color-coded Excel report**, grouped by customer, with proper alignment and spacing.

This executable is then scheduled to run **automatically every day at 11:00 AM** using **Windows Task Scheduler** via PowerShell.

---

## 📂 Features

✅ Fetch Gmail alerts using IMAP
✅ Parse both plain-text and HTML content using **BeautifulSoup**
✅ Intelligent detection of group names, devices, and alert times
✅ Alert filtering (e.g., only “Critical” alerts)
✅ Highlighting for common issues:

* 🔴 Backup Failed (Red)
* 🟡 Quota Exceeded (Light Red)
* 🟢 Incident Detected (Green)

✅ Final report saved in Excel with:

* Merged customer rows
* Wrapped message fields
* Spacing between different client sections
* Summary sheet with stats

---

## ⚙️ Technologies Used

* Python 3.11
* PyInstaller (`--onefile` build)
* `imaplib`, `email`, `dotenv`, `re`, `bs4`, `pandas`, `xlsxwriter`
* Windows Task Scheduler
* PowerShell scripting
* Acronis Cyber Protect script management

---

## 🔁 Automation Setup

### ✅ Task Scheduler

* Triggers the executable every day at **11:00 AM**.
* Ensures no manual intervention is needed.
* `.env` file is stored alongside `.exe` in the same `dist/` directory for secure credential management.

### ✅ Acronis Panel Execution

* Uses PowerShell to `Set-Location` to the directory and execute `.\Email_Report_Generator.exe`
* Useful for IT admins to trigger or monitor the job remotely via the Acronis Console.

---

## 📅 Workflow Summary

```plaintext
Gmail Inbox
   ↓
Filter: "DAILY STATUS REPORT" with Critical Alerts
   ↓
Extract Device, Time, Group, Message
   ↓
Filter Only Today's and Yesterday's Alerts
   ↓
Formatted Excel Report with Conditional Coloring and Summary
   ↓
Saved to Shared Network Folder
   ↓
(Automatically triggered every day @ 11AM)
```

---

## 📁 Folder Structure

```bash
📁 Report_Generator_Application/
├── dist/
│   ├── Email_Report_Generator.exe
│   ├── .env                 # Contains EMAIL_ID and APP_PASSWORD
│   ├── launch_script.ps1    # PowerShell launcher
├── Email_Report_Generator.py  # Main Python source code
```

---

## 📌 Result & Appreciation

The solution completely eliminated the manual work involved in generating daily reports. It runs **unattended**, delivers reports in consistent quality and format.

