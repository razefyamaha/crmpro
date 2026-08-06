# CRM Pro

A complete Customer Relationship Management system built with **Bootstrap 5**,
**Flask**, and **Microsoft Excel as the database**.

All your data lives in `data/CRM_Database.xlsx`. The app reads it, shows it in a
dashboard, and writes your changes straight back into the same Excel file. You
can also edit that Excel file directly in Microsoft Excel — the app picks up your
changes next time it loads a page.

---

## Features

- **Dashboard** — KPI cards (total quotations, value, pipeline, won/lost), plus charts:
  status distribution, quotations per month, **top companies by completed sales**, and
  a **Potential Customers** panel (leads by stage with pipeline value).
- **Monthly Target tracker (CRM card)** — target set to **RM 60,000 / month** (editable
  in the Excel `Settings` sheet). Shows closed-sale value for the **last 3 months**
  (June, July, August) with progress bars vs the 60k target and a combined total.
- **Quotations** — searchable, filterable, sortable, paginated table with full
  **Add / Edit / Delete**. New quotation numbers are auto-generated (`QT#####`).
  The company field is a **dropdown of your existing customers** — pick one or type
  a new customer.
- **Customers (Potential Customers)** — dedicated page listing active leads (not yet
  Won/Lost) with a filterable/sortable table, lead source, pipeline value, quick
  stage-dropdown, Edit/Delete, and a **View** button that opens each customer's detail page.
- **Customer Detail page** — click the **👁 View** button on a customer to open their
  detail page: full contact info, their quotations, and a **document library** of
  uploaded quotation files. Uploaded PDFs appear as a **thumbnail grid** — click any
  thumbnail to open the full quotation in a Bootstrap modal, or download the original.
  Each uploaded file appears in a **list showing its file name and upload date**
  (with a small thumbnail preview that falls back to a PDF icon if it can't render).
  Clicking a row opens the quotation in a full-screen Bootstrap modal. PDF pages are
  **rendered as images server-side**, so the quotation always shows inline in the
  page — it never depends on the browser's PDF/download settings. The viewer has **zoom controls**
  (in/out/fit-width/reset, Ctrl+wheel), a **page count + page jump** (prev/next and a
  "go to page" box) for multi-page PDFs, and fit-to-width by default.
  The same inline viewer is on the **Quotations list** — any quotation with an
  uploaded PDF shows a **View** button that opens it from the list.
  On the Quotations page you can also **search by PDF filename** and filter to
  **"With PDF file" / "Without file"**.
- **Upload quotation PDFs** — from the customer detail page, the **New Lead** sidebar,
  the **New Quotation** form, or the **Save as Quotation** form. Files are stored in a
  **per-customer folder** (`data/uploads/<Customer Name>/`). If you upload a quotation
  PDF, CRM Pro **auto-reads it and fills in the customer's details** (contact person,
  phone, address) automatically.
- **Products** — product list (from the Excel `Products` sheet) with **Cost Price,
  SRP, Margin %, Selling Price and Profit**. When you enter a **margin %, it
  automatically shows the selling price** (Selling = Cost ÷ (1 − Margin%)). Full
  Add / Edit / Delete.
- **Quotation Calculator** — build a quotation from your products: pick
  products + quantities, see cost, selling price, profit and overall margin, then
  save it directly as a quotation. Also a quick selling-price calculator.
- **Forecast** — a page for customer sales projections: **select an existing customer
  or add a new one**, add **products/items** with quantities, and see totals live.
  Each line item has **editable Cost, SRP and Margin %** — the selling price is
  recalculated from cost + margin, and changes **update the linked product** in the
  Products sheet. **8% SST is added automatically for Service or Software** category
  items (others are excluded). Forecasts are saved to Excel (new `Forecasts` sheet),
  listed with status tracking (Open/Won/Lost), and you can **view the detail** (line
  items, cost/SRP/margin/SST/profit breakdown, totals) or **edit** any forecast. The
  dashboard has a **Forecast card** showing total forecast value, potential profit,
  count, and open pipeline.
- **Export / Backup / Shutdown** — top-bar buttons to download the live Excel
  database (backup) or a CSV of all quotations, plus a **Shutdown** button that
  stops the server right from inside the app.

---

## Requirements

- Python 3.9+
- Internet connection on first load (Bootstrap / Chart.js are loaded from CDN)

## How to run

### Windows (easiest — one click)
Just double-click **`run.bat`**. It automatically:
1. Checks Python is installed.
2. Installs the required libraries (flask, openpyxl) on first run.
3. Starts the server and **opens your browser** at http://127.0.0.1:5000 automatically.

Close the black window (or press Ctrl+C) to stop the server.

> If Python isn't installed, install it from https://www.python.org/downloads/
> and tick **"Add Python to PATH"** during installation, then run `run.bat` again.

### Mac / Linux (or Windows command line)
```bash
cd crmpro

# first time only:
pip install -r requirements.txt

# run:
python app.py
```

Then open your browser at: **http://127.0.0.1:5000**

Press `Ctrl+C` in the terminal to stop the server.

---

## The Excel database

File: **`data/CRM_Database.xlsx`** (generated from your Quotation Report PDF —
126 quotations, 83 companies).

| Sheet | Columns |
|-------|---------|
| **Quotations** | Quotation No, Date Created, Time, Company, Status, Amount (RM), Created By |
| **Companies**  | Company, Industry, Contact Person, Email, Phone, Lead Source, Stage, Notes, Address |
| **Products**   | Product ID, Product Name, Category, Cost Price, SRP, Margin %, Selling Price, Profit, Notes |
| **Settings**   | Key / Value — contains `monthly_target` (set to 60000). Change it in Excel to update the target. |

Uploaded quotation PDFs are kept in **`data/uploads/<Customer Name>/`** (a separate
folder per customer), outside the Excel file but alongside it in the `data` folder.

**Important:** the app writes back to this file. Close the Excel file in MS Excel
while the web app is running, or you may get a "file locked" error.

To start fresh or reset: use the **Backup .xlsx** button on the top bar to
download the current file before making big changes.

**Products sheet** — 3 sample products are pre-loaded so you can try the margin →
selling-price calculator. Delete them from the app (or Excel) once you add your own.

**Settings sheet** — the row `monthly_target = 60000` drives the dashboard target.
Change the value here in Excel and refresh the dashboard to use a different target.

---

## Project structure

```
crmpro/
├── app.py                  # Flask app (all routes)
├── db.py                   # Excel read/write layer (openpyxl)
├── requirements.txt
├── run.bat / run.sh        # launchers
├── data/
│   └── CRM_Database.xlsx   # <-- your Excel database
├── templates/              # Bootstrap 5 HTML pages
│   ├── base.html
│   ├── dashboard.html
│   ├── quotations.html
│   └── companies.html
└── static/css/style.css
```

## Build a standalone .exe (Windows)

If you want a single `CRM Pro.exe` you can run anywhere on Windows (no Python needed):

1. Copy the whole `crmpro` folder to your **Windows** PC.
2. Double-click **`build_exe.bat`**.
3. When it finishes, your app is at **`dist\CRM Pro.exe`**, with the `data` folder
   (your Excel database) sitting right next to it.
4. Double-click `CRM Pro.exe` to run — the browser opens automatically.

> The `.exe` must be built **on Windows** (PyInstaller only makes Windows `.exe`
> when run on Windows). The build script handles installing PyInstaller for you.
> `app.py` and `db.py` are already written to work both from source and as a
> frozen `.exe` (the data folder stays next to the `.exe` so your database
> persists and is easy to back up).

## Editing the data in Excel

Open `data/CRM_Database.xlsx` in Microsoft Excel. The **Quotations** sheet maps
1-to-1 to what the app shows. Add/remove rows and save — the next page load in the
app will reflect your changes. Keep the header row as-is.
