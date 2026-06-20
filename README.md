# 🖥️ Production Efficiency Dashboard

A real-time web dashboard built during my internship at **CommScope** to monitor 
production efficiency across assembly lines and shifts.

---

## 📌 About the Project

This dashboard was developed to help the production team track actual vs target 
output across 4 assembly lines (Line 1–4) and 3 shifts (A/B/C) in real time.
It auto-detects changes in Excel data files and updates the dashboard without 
any manual refresh.

---

## ⚙️ Features

- 📊 Live Actual vs Target efficiency tracking
- 🏭 Supports 4 Lines × 3 Shifts (A/B/C)
- 📦 Products tracked: OPTITAP, PRODIGY, BT-COF, JUMPER, 1x4, 1x8
- 🔄 Auto file watcher — detects & consolidates new Excel files automatically
- 🔍 Filter by Product, Date, and Shift
- 📤 Auto-pushes efficiency data back to Excel for reporting
- 🌐 Clean web UI built with HTML/JS frontend

---

## 🛠️ Tech Stack

- **Backend:** Python, Flask
- **Excel Handling:** openpyxl
- **File Watching:** watchdog
- **Frontend:** HTML, CSS, JavaScript

---

## 🚀 Setup & Installation

### 1. Clone the repo
git clone https://github.com/sarthak-cloud/production-dashboard.git
cd production-dashboard

### 2. Install dependencies
pip install flask openpyxl watchdog

### 3. Configure your database path
Open `processor.py` and update:
DATABASE_ROOT = r"your\database\folder\path"

Open `watcher.py` and update:
SOURCE_FOLDER = r"your\source\folder\path"

### 4. Set up Data_sheet.xlsx
Create a `Data_sheet.xlsx` in the root folder with two sheets:
- `Targets` — columns: Product Name | Target Value
- `Efficiency` — used for auto-push output (can be blank initially)

### 5. Excel DB file column order
Year | Month | Product | Length | Product Type | Link | Sheet | 
MID | Order No | Team Leader | Date | Shift | Clock ID | 
Product Name | Line | Produced | OK | NOK

### 6. Run the app
python app.py

Open your browser at: http://localhost:1111

---

## 📁 Project Structure

production-dashboard/
│
├── app.py              # Flask app & API routes
├── processor.py        # Data processing & dashboard logic
├── consolidator.py     # Excel file consolidation
├── watcher.py          # Auto file watcher (watchdog)
├── templates/
│   └── dashboard.html  # Frontend UI
└── README.md

---

## ⚠️ Note

The actual production database files are confidential and not included 
in this repository. Follow the setup steps above to configure your own 
data source.

---

## 👨‍💻 Author

**Sarthak** — Electrical & Computer Science Engineering Student  
Built during internship at CommScope (May–June 2026)
