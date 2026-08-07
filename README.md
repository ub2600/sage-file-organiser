# Sage File Organiser

A standalone Windows desktop application designed to reorganize, clean, transform, and segment **Sage CSV exports** (Customers, Sales Orders/Invoices, Products, Transactions).

Built on the architecture of `csv-organizer`, with native support for Sage data formats.

---

## Features

### Tab 1: Sage File Organizer
- **Auto-Detection Engine**: Automatically detects whether an uploaded CSV is a Sage **Customer**, **Sales Order**, **Product**, or **Generic** export based on header signatures.
- **Customer CSV Processing**: Standardizes field names (`account_ref` → `Customer_ID`, `name` → `Customer_Name`, `e_mail` → `Email`), formats phone numbers, re-orders columns, and attaches custom segment tags.
- **Sales Order CSV Processing**: Standardizes order fields, cleans dates into normalized `YYYY-MM-DD` format, and structures pricing/VAT totals.
- **Product CSV Processing**: Normalizes stock codes, pricing, stock levels, and category metadata.
- **Encoding Protection**: Auto-detects file encodings (`UTF-8`, `UTF-8 with BOM`, `CP1252`, `Latin-1`) and displays byte-level diagnostics on corrupt input files.

### Tab 2: Sage Data Splitter & Joiner
- **Multi-File Processing**: Takes Sage **Customer CSV** and **Sales Order CSV**.
- **Customer Segmentation**:
  - **Active Customers**: Automatically filters and exports customers who have recent orders.
  - **Inactive Customers**: Segments accounts with no recent order history into a separate clean CSV.
- **Clean Phone & Address Normalization**: Strips invalid characters, fixes trailing `.0` values from numerical exports, and ensures proper leading zeros on UK numbers.

---

## How to Build the Standalone Windows `.exe` (No Python Required)

1. **Create a GitHub Repository**:
   - Go to GitHub and create a repository named `sage-file-organiser`.
2. **Push the Code**:
   ```bash
   git init
   git add .
   git commit -m "Initial commit of Sage File Organiser"
   git branch -M main
   git remote add origin https://github.com/YOUR_USERNAME/sage-file-organiser.git
   git push -u origin main
   ```
3. **Download executable**:
   - Navigate to the **Actions** tab on your GitHub repository.
   - Click the latest workflow run for `Build Windows EXE`.
   - Download the generated `SageFileOrganiser-Windows` zip under **Artifacts** to retrieve your standalone `SageFileOrganiser.exe`.

---

## Local Development & Usage

### Running Locally with Python
```bash
python sage_file_organiser.py
```

### Packaging Locally with PyInstaller
```bash
pip install -r requirements.txt
pyinstaller --onefile --windowed --name "SageFileOrganiser" sage_file_organiser.py
```
The compiled executable will be placed in the `dist/` directory.

---

## Technical Stack
- **GUI Framework**: Python `tkinter` + `ttk` with custom dark theme design system.
- **CSV Processing**: Standard `csv` library, regex cleaning, and `datetime` reformatting.
- **Build Engine**: PyInstaller + GitHub Actions workflow runner (`windows-latest`).
