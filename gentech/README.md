# Gentec Billing System

Flask-based billing web app for generating invoices with JSON master data, preview, and download in PDF + Excel.

## Features
- Customer master from JSON with `Add Customer` screen.
- Item description to HSN/SAC coupling from JSON.
- Line items start with one default row and support adding/removing rows up to configured single-page limit.
- Description can be selected from master list or entered manually.
- HSN/SAC and Unit Price auto-fill from known descriptions and remain editable.
- Reference supports dropdown values, `None`, and manual entry.
- Editable Qty and Unit Price with automatic line amount and GST totals.
- Transport charge (default `0`) added after GST as non-taxable add-on.
- Per-invoice `Include Letterhead/Logo` checkbox.
- Fixed top header-band spacing in both PDF and Excel:
  - Checked: shows company letterhead details.
  - Unchecked: keeps the same top space blank for pre-printed letterhead paper.
- If `/Users/karthra3/Documents/gentech/gentech/letterhead.*` or `/Users/karthra3/Documents/gentech/gentech/letterpad.*` exists, that image is used in preview, PDF, and Excel.
- Invoice history with re-download links.

## Run Locally
1. Create venv and install dependencies:
   ```bash
   python -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```
2. Start server:
   ```bash
   python app.py
   ```
3. Open in browser:
   - `http://localhost:5000`

## Run on Windows LAN
1. Open PowerShell as Administrator.
2. If script execution is blocked, run:
   ```powershell
   Set-ExecutionPolicy -Scope Process Bypass
   ```
3. Deploy and start as Windows Service (installs Python if missing, creates `.venv`, installs deps, starts service):
   ```powershell
   .\tools\deploy_windows.ps1 -Port 5000
   ```
4. Access from LAN: `http://<host-ip>:5000`

### Windows Service Controls
- Start service:
  ```powershell
  .\tools\start_service.ps1
  ```
- Stop service:
  ```powershell
  .\tools\stop_service.ps1
  ```

Service logs are written to:
- `data\service_logs\service.out.log`
- `data\service_logs\service.err.log`

## Data Files
Auto-created in `/data` on first run:
- `company_settings.json`
- `customers.json`
- `items.json`
- `sequence.json`
- `invoices/*.json`
- `generated/pdf/*.pdf`
- `generated/xlsx/*.xlsx`
