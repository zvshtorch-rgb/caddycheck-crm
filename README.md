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

## License Expiry Email Alerts

Use `license_expiry_alert.py` to email the configured recipients when project
licenses are about to expire. By default it alerts for licenses expiring exactly
3 days from today.

Run this migration once in the Supabase SQL editor so repeated scheduled runs do
not resend the same alert:

```sql
create table if not exists license_expiry_alert_log (
  id            bigint generated by default as identity primary key,
  project_name  text not null,
  license_eop   date not null,
  days_before   integer not null default 3,
  sent_to       text,
  sent_cc       text,
  sent_at       timestamptz not null default now(),
  unique (project_name, license_eop, days_before)
);

create index if not exists license_expiry_alert_log_sent_at_idx
  on license_expiry_alert_log(sent_at desc);

create index if not exists license_expiry_alert_log_license_eop_idx
  on license_expiry_alert_log(license_eop);

notify pgrst, 'reload schema';
```

Manual test:

```bash
python license_expiry_alert.py --dry-run
```

Live send:

```bash
python license_expiry_alert.py
```

Useful overrides:

```bash
python license_expiry_alert.py --days-before 3
python license_expiry_alert.py --to admin@example.com --cc ops@example.com
python license_expiry_alert.py --force
```

A scheduled GitHub Actions workflow is available at
[.github/workflows/license_expiry_alert.yml](.github/workflows/license_expiry_alert.yml).
It runs daily at 06:00 UTC and uses GitHub repository secrets:

- Supabase: `SUPABASE_URL`, `SUPABASE_SERVICE_ROLE_KEY`
- SMTP: `SMTP_HOST`, `SMTP_PORT`, `SMTP_USE_TLS`, `SMTP_USERNAME`,
  `SMTP_PASSWORD`, `SMTP_SENDER_EMAIL`, `SMTP_SENDER_NAME`
- Recipients: `EMAIL_DEFAULT_RECIPIENTS`, `EMAIL_DEFAULT_CC`

The job runs as a standalone Python process and does not depend on Streamlit.

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

## Purchase-Order Approval Workflow (MVP)

Customer purchase orders that arrive **by email as a PDF** can be automatically
ingested and routed to the CEO for approval via a secure email link.

> Note: this uses a dedicated `purchase_orders` table (separate from the internal
> camera-ordering `orders` table) so it cannot affect existing CRM data.

### One-time setup

1. Run `migrations/create_purchase_order_approvals.sql` in the Supabase SQL editor.
   This creates `incoming_order_emails`, `purchase_orders`, and `order_approvals`,
   and the private `purchase-order-pdfs` storage bucket is created automatically on
   first upload.
2. Add the configuration below as environment variables (for the poller / cron) or
   in Streamlit secrets (for the app).

### How it works

1. `order_intake_poll.py` reads the configured mailbox for unread emails with a PDF.
2. Each PDF is stored privately in Supabase; a `purchase_orders` row is created.
3. A secure, single-use approval token (valid 7 days) is generated — **only its
   SHA-256 hash is stored** — and the CEO is emailed a link
   `https://<app>/?approval=<token>`.
4. The CEO opens the link (no login required), reviews the order and a short-lived
   signed PDF URL, then **Approves / Rejects / Requests correction**.
5. Admins can review all orders and statuses on the `✅ Order Approvals` page.

### Running the poller

```bash
py order_intake_poll.py            # process new emails
py order_intake_poll.py --dry-run  # inspect without writing anything
```

`order_intake_poll.py` is a **standalone Python process** — it does *not* need the
Streamlit runtime. The Streamlit Cloud app cannot run the poller reliably (it only
executes while a browser session is open), so schedule the poller externally using
one of the options below.

#### Option 1 — Windows Task Scheduler

Create a scheduled task that runs the poller every few minutes.

**Program/script:**

```text
py
```

**Add arguments:**

```text
order_intake_poll.py
```

**Start in (working directory):**

```text
F:\caddycheck-crm
```

To capture logs to a file, point the task at a small wrapper instead. Create
`run_order_intake_poll.bat` in the repo root:

```bat
@echo off
cd /d F:\caddycheck-crm
py order_intake_poll.py >> logs\order_intake_poll.log 2>&1
```

Then set the task **Program/script** to the `.bat` file. The required environment
variables (see the table below) must be set as **system or user environment
variables**, or exported at the top of the `.bat` file with `set NAME=value`.
Recommended trigger: *Daily*, repeat task every *5 or 10 minutes* for a duration of
*1 day*, *indefinitely*.

#### Option 2 — GitHub Actions scheduled workflow

A ready-made workflow lives at
[.github/workflows/order_intake_poll.yml](.github/workflows/order_intake_poll.yml).
It runs every 10 minutes (change the `cron` to `*/5 * * * *` for every 5 minutes),
installs `requirements.txt`, and runs `python order_intake_poll.py`.

Add the configuration as **GitHub repository secrets**
(*Settings → Secrets and variables → Actions → New repository secret*):

- Supabase: `SUPABASE_URL`, `SUPABASE_SERVICE_ROLE_KEY`
- Approval links / recipients: `APP_BASE_URL`, `ORDER_APPROVAL_CEO_EMAIL`
- Mailbox (IMAP): `ORDER_INTAKE_PROVIDER`, `ORDER_INTAKE_IMAP_HOST`,
  `ORDER_INTAKE_IMAP_PORT`, `ORDER_INTAKE_IMAP_USERNAME`,
  `ORDER_INTAKE_IMAP_PASSWORD`, `ORDER_INTAKE_IMAP_FOLDER`
- Mailbox (Graph, only if `ORDER_INTAKE_PROVIDER=graph`):
  `ORDER_INTAKE_GRAPH_TENANT_ID`, `ORDER_INTAKE_GRAPH_CLIENT_ID`,
  `ORDER_INTAKE_GRAPH_CLIENT_SECRET`, `ORDER_INTAKE_GRAPH_MAILBOX`,
  `ORDER_INTAKE_GRAPH_FOLDER`
- SMTP: `SMTP_HOST`, `SMTP_PORT`, `SMTP_USE_TLS`, `SMTP_USERNAME`,
  `SMTP_PASSWORD`, `SMTP_SENDER_EMAIL`, `SMTP_SENDER_NAME`

You can also trigger a run manually from the **Actions** tab (`workflow_dispatch`).

### Configuration

Shared:

| Setting | Description |
| --- | --- |
| `SUPABASE_URL` | Supabase project URL (poller only; app uses secrets). |
| `SUPABASE_SERVICE_ROLE_KEY` | Service role key (poller only). |
| `APP_BASE_URL` | Public base URL of the Streamlit app, e.g. `https://your-app.streamlit.app`. |
| `ORDER_APPROVAL_CEO_EMAIL` | CEO recipient(s), comma-separated. |
| `ORDER_INTAKE_PROVIDER` | `imap` (default, also Gmail/O365 IMAP) or `graph`. |

IMAP provider (`ORDER_INTAKE_PROVIDER=imap`):

| Setting | Description |
| --- | --- |
| `ORDER_INTAKE_IMAP_HOST` | e.g. `imap.gmail.com` or `outlook.office365.com`. |
| `ORDER_INTAKE_IMAP_PORT` | Default `993`. |
| `ORDER_INTAKE_IMAP_USERNAME` | Mailbox username. |
| `ORDER_INTAKE_IMAP_PASSWORD` | Mailbox password / app password. |
| `ORDER_INTAKE_IMAP_FOLDER` | Default `INBOX`. |

Microsoft Graph provider (`ORDER_INTAKE_PROVIDER=graph`, preferred for Office 365):

| Setting | Description |
| --- | --- |
| `ORDER_INTAKE_GRAPH_TENANT_ID` | Azure AD tenant id. |
| `ORDER_INTAKE_GRAPH_CLIENT_ID` | App registration client id. |
| `ORDER_INTAKE_GRAPH_CLIENT_SECRET` | App registration secret. |
| `ORDER_INTAKE_GRAPH_MAILBOX` | Mailbox UPN/object id to read. |
| `ORDER_INTAKE_GRAPH_FOLDER` | Default `Inbox`. |

Email sending reuses the existing SMTP settings. When the poller runs standalone
(Task Scheduler / GitHub Actions) it has no Streamlit secrets, so configure SMTP
via these environment variables (they override the app config only when set):

| Setting | Description |
| --- | --- |
| `SMTP_HOST` | SMTP server host, e.g. `smtp.gmail.com`. |
| `SMTP_PORT` | Default `587`. |
| `SMTP_USE_TLS` | `true` (STARTTLS) or `false` (SSL). Default `true`. |
| `SMTP_USERNAME` | SMTP username / login. |
| `SMTP_PASSWORD` | SMTP password / app password. |
| `SMTP_SENDER_EMAIL` | From address (defaults to `SMTP_USERNAME`). |
| `SMTP_SENDER_NAME` | Display name, e.g. `CaddyCheck CRM`. |

### Security

- Approval links use a cryptographically random token; only the SHA-256 hash is
  stored in the database.
- Tokens expire after 7 days and become single-use once a decision is recorded.
- Purchase-order PDFs live in a private bucket and are only ever exposed through
  short-lived signed URLs.

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
