# TOABH Contracts Dashboard

Very simple signed-contract dashboard backed by Google Sheets + Google Apps Script.

## What this includes
- React dashboard with only 3 tabs:
  - All Signed Contracts
  - Contracts Expiring
  - Contract Renewal
- Signed-only filtering
- One-time AI PDF enrichment flow driven by Apps Script
- Renewals sheet creation + version bumping
- Frontend actions for rescan, renewal, and notes

## Run locally
```bash
cd toabh-contracts-dashboard
npm install
cp .env.example .env
# add your Apps Script web app URL
npm run dev
```

## Build
```bash
npm run build
```

## Apps Script setup
1. Open your Google Sheet.
2. Open Extensions → Apps Script.
3. Copy `apps-script/Code.gs` into the Apps Script project.
4. Copy `apps-script/appsscript.json` into the manifest.
5. In Script Properties, set:
   - `MAIN_SHEET_NAME` (default: `Contracts`)
   - `RENEWALS_SHEET_NAME` (default: `Renewals`)
   - `GEMINI_API_KEY`
   - `CONTRACT_GENERATOR_WEBHOOK_URL` (optional, if your existing contract generator is exposed by webhook)
6. Run `setupSheetsAndTriggers()` once.
7. Deploy the script as a Web App with access for your dashboard.
8. Paste that URL into `VITE_APPS_SCRIPT_URL`.

## Notes
- AI scans only when `Signed PDF URL` exists and `AI Scan Status` is blank or `Pending`.
- Scanned contracts are read from sheet values after first scan.
- `Rescan with AI` explicitly resets a row back to `Pending`.
- Renewal row creation is done in the `Renewals` sheet.
- The sample includes a webhook hook for your existing contract-generation flow.
