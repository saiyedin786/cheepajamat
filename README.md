# Chheepa Samaj Khanda Falsa, Jodhpur — Census & Directory Portal

A modern Census & Directory web application with an administrative panel for **Chheepa Samaj Khanda Falsa, Jodhpur**.

---

## 🌟 Key Features

1. **Public Member & Family Directory**:
   - Multi-criteria instant search (by Name, Hindi Name, Father/Spouse, Mobile, Mohalla, Gotra, Education, Profession).
   - Family grouping by `Family ID` linking members with the Head of Household.
   - Demographics & Analytics (gender ratio, literacy rates, youth percentage, employment breakdown).
   - Printable ID cards and profile sheets.

2. **🔐 Password-Protected Admin Panel**:
   - Secure session-based authentication backed by hashed passwords in SQLite (`admin_users` table).
   - Only authorized administrators can access `/admin`.
   - Access control checks on all sensitive import, edit, delete, and maintenance endpoints.
   - Password change tool inside the admin security settings.

3. **📤 Multi-Format Bulk Upload (CSV, Excel, JSON)**:
   - **Excel Files (`.xlsx`, `.xls`)**: Ingest spreadsheets with automatic data type handling and column normalization.
   - **CSV Files (`.csv`)**: Full support for comma/tab/semicolon delimiters with UTF-8 BOM encoding for Hindi (Devanagari) text.
   - **JSON Files (`.json`)**: Import structured JSON arrays or backup payloads.
   - **🔍 Interactive Data Preview**: Preview mapped columns, total rows detected, and the first 5 records in a live table before saving to the database.
   - **Ingestion Strategies**:
     - **Append Mode**: Adds new records while generating unique member codes for conflicts.
     - **Update & Merge Mode**: Matches existing records by Member Code or Name + Primary Phone and updates their details.
     - **Replace All Mode**: Wipes the database and loads the uploaded file fresh (with explicit confirmation).
   - **Downloadable Sample Templates**: Pre-formatted templates for CSV, Excel (.xlsx), and JSON ready for data entry.

4. **📥 Instant Data Export**:
   - Download complete database as a styled Excel workbook (`.xlsx`).
   - Download as CSV (UTF-8 BOM).
   - Download as full JSON Backup.

---

## 🔑 Default Administrator Login

| Field | Default Value | Notes |
|---|---|---|
| **Login URL** | `http://127.0.0.1:5000/admin/login` | Also accessible via the **🔐 Admin Panel** button |
| **Admin Username** | `admin` | Stored in `admin_users` table |
| **Admin Password** | `admin123` | Can be changed in Admin Panel -> Security Tab |

---

## 🚀 How to Run

### Method 1: Using the Batch Launcher (Recommended)
Double-click `run_portal.bat`. It will:
1. Verify Python & required packages.
2. Install any missing dependencies from `requirements.txt`.
3. Launch the local web server at `http://127.0.0.1:5000`.
4. Open the portal in your default web browser.

### Method 2: Manual Terminal Command
```bash
# Install dependencies
pip install -r requirements.txt

# Start the Flask portal
python app.py
```

Open your browser at:
- **Public Directory Portal**: [http://127.0.0.1:5000](http://127.0.0.1:5000)
- **Admin Control Panel**: [http://127.0.0.1:5000/admin](http://127.0.0.1:5000/admin)
- **Admin Login**: [http://127.0.0.1:5000/admin/login](http://127.0.0.1:5000/admin/login)

---

## 🛠 Tech Stack

- **Backend**: Python 3, Flask 3.1, Werkzeug (Security & Password Hashing), SQLite 3
- **Spreadsheets**: openpyxl 3.1
- **Frontend**: HTML5, Tailwind CSS, Vanilla JavaScript (Fast, zero-build, responsive)
