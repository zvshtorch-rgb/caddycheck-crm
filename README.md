# CaddyCheck CRM

A desktop GUI application for managing project revenue, debt tracking, and monthly invoice generation for CaddyCheck/iRetailCheck maintenance contracts.

---

## Features

- **Dashboard** – Summary cards (total income, debt, paid, monthly/yearly income, active projects, cameras), filter by year/month/country/paid status, export to Excel.
- **Projects** – Full project list from "Projects overview", search/filter, revenue breakdown per maintenance year.
- **Invoice Details** – All invoices from "Invoice details", filter by year/project/maintenance year/paid status, full debt summary per project.
- **Monthly Invoice Generation** – Select month + year, preview all projects included, generate Excel invoice from template, open file, send by email.
- **Settings** – SMTP email configuration, default recipients/subject/body, SMTP connection test.

---

## Setup

### 1. Requirements

- Python 3.10+
- Install dependencies:

```bash
pip install -r requirements.txt
```

### 2. Data Files

Place your Excel files in the `data/` directory:

```
data/
  CaddyCheckProjectsInfo.xlsx     ← Projects and invoice history
  CC_M_inv_8669_Dec_2025.xlsx     ← Invoice template
```

### 3. Run the Application

```bash
python app.py
```

---

## Project Structure

```
e:/CRM_caddycheck/
├── app.py                        ← Entry point
├── requirements.txt
├── README.md
├── data/                         ← Excel source files (not committed)
├── output/                       ← Generated invoices and exports
├── config/
│   ├── settings.py               ← Constants, paths, email config loader
│   └── email_config.json         ← Saved email settings (auto-created)
├── models/
│   ├── project.py                ← Project dataclass
│   └── invoice.py                ← Invoice and DebtSummary dataclasses
├── services/
│   ├── excel_service.py          ← Read/parse Excel files, compute debt
│   ├── invoice_service.py        ← Generate monthly invoice Excel files
│   ├── email_service.py          ← Send emails via SMTP
│   └── report_service.py         ← Export dashboard to Excel
└── ui/
    ├── main_window.py            ← Main window, sidebar navigation, data loader
    ├── dashboard_page.py         ← Dashboard with cards and table
    ├── projects_page.py          ← Projects list and revenue calculator
    ├── invoices_page.py          ← Invoice list and debt summary
    ├── monthly_invoice_page.py   ← Monthly invoice generator and email sender
    └── settings_page.py          ← SMTP and email settings
```

---

## Business Rules

### Revenue Calculation

| Maintenance Year | Rate per Camera |
|-----------------|-----------------|
| Y1 (first year) | €778            |
| Y2+ (subsequent)| €228            |

Maintenance year is determined by: `invoice_year - installation_year + 1`

The calculation is isolated in `services/invoice_service.py → _determine_maintenance_year()` for easy adjustment.

### Invoice Generation

Generated invoices reuse the template structure from `CC_M_inv_8669_Dec_2025.xlsx`, preserving headers, bank details, and formatting. Projects are included based on matching `payment month` from the Projects overview.

### Debt Calculation

Expected = sum of `num_cams × rate` for each year from `installation_year` to current year.
Debt = Expected − Paid invoices.

---

## Email Configuration

1. Go to **Settings** page.
2. Enter your SMTP host, port, credentials, and default recipients.
3. Click **Test Connection** to verify.
4. Click **Save Settings**.

For Gmail, use:
- SMTP Host: `smtp.gmail.com`
- Port: `587`
- TLS: enabled
- Use an [App Password](https://support.google.com/accounts/answer/185833) instead of your account password.

---

## Generated Files

All generated files are saved to the `output/` directory:
- Monthly invoices: `CC_M_inv_<number>_<Mon>_<Year>.xlsx`
- Dashboard exports: `CaddyCheck_Dashboard_<timestamp>.xlsx`

---

## Scheduled Monthly Sending

Use `monthly_auto_send.py` to generate and email the monthly invoice automatically on the 1st of the month.

Default behavior:
- If you run it on `2026-05-01`, it prepares and sends the invoice for `April 2026`.
- It uses the configured default recipients and CC addresses.
- It saves invoice rows to the ledger before sending unless you disable that.
- It refuses to resend the same month/year if that month already exists in the sent invoice log unless you pass `--force`.

For shared/cloud use:
- Run `migrations/create_sent_invoices.sql` in the Supabase SQL editor so sent invoices persist across app restarts.

Manual test:

```bash
python monthly_auto_send.py --dry-run
```

Live send:

```bash
python monthly_auto_send.py
```

Useful overrides:

```bash
python monthly_auto_send.py --month April --year 2026
python monthly_auto_send.py --to finance@example.com --cc ops@example.com
python monthly_auto_send.py --source excel
python monthly_auto_send.py --force
```

Recommended Windows Task Scheduler setup:

1. Create a Basic Task named `CaddyCheck Monthly Invoice`.
2. Trigger: `Monthly`, every `1` month, on day `1`.
3. Start time: choose the local time you want the email to go out.
4. Action: `Start a program`.
5. Program/script: path to your Python executable.
6. Add arguments: `monthly_auto_send.py`.
7. Start in: the project folder.

Example values:

```text
Program/script: C:\Python311\python.exe
Add arguments: monthly_auto_send.py
Start in: F:\caddycheck-crm
```

Recommended first run in Task Scheduler:
- Use `monthly_auto_send.py --dry-run` first.
- Confirm the target month, invoice number, recipients, and total in the task history/logs.
- Then switch the task arguments back to `monthly_auto_send.py`.

---

## Orders Workflow

The CRM now supports project orders as a separate workflow from Projects.

Before using it in the shared cloud app:
- Run `migrations/create_orders.sql` in the Supabase SQL editor.

How to use the Orders page:
- Open the `📦 Orders` page in the Streamlit app.
- Use `Import Orders` to upload a customer order file (`PDF`, `XLSX`, or `CSV`).
- Or use `New Order` to create an order row manually.

Tracked order fields:
- Order reference
- Project name
- Country
- Ordered cameras
- Payment month
- Install year
- Requested activation date
- Status
- Notes

Supported order statuses:
- `New`
- `Ordered`
- `In Progress`
- `Installed`
- `Active`
- `Cancelled`

Creating CRM projects from orders:
- Use `Create Missing Projects From Selected Orders` on the Orders page.
- This creates projects in the Projects list from order rows that do not yet exist as projects.

Storage behavior:
- If the `orders` table exists in Supabase, orders are stored centrally.
- If the table is missing, the app falls back to local JSON storage and shows a warning in the UI.

---

## Slide Content: Supabase, Streamlit, and GitHub

If you want a single slide that explains how the app is connected, you can use this:

**Title:** How CaddyCheck CRM connects Streamlit, Supabase, and GitHub

**Slide text:**
- **Streamlit** is the user interface. It shows the dashboard, orders, projects, invoices, and settings.
- **Supabase** is the live database. The app reads and writes projects, invoices, orders, tickets, and subscriptions there.
- **GitHub** stores the source code and updates. When the app changes, the code is committed and deployed from GitHub.
- Optional local files act as a fallback when Supabase data is not available.

**Simple flow diagram:**

```text
User
  ↓
Streamlit app
  ↓
Supabase database
  ↓
GitHub repository
```

**Short speaker note:**
The Streamlit app is the front end that the user interacts with. Supabase is the backend database where the app saves live data. GitHub keeps the codebase versioned and is the place from which the app is deployed and maintained.
