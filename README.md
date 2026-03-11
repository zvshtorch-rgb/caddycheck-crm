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
