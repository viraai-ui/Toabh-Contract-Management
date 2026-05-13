# TOABH Contracts Dashboard

Very simple signed-contract dashboard backed by the existing Google Sheet + bound Google Apps Script + Zoho Sign flow.

## What this includes
- React dashboard with only 3 tabs:
  - All Signed Contracts
  - Contracts Expiring
  - Contract Renewal
- Signed-only filtering using `Signed PDF URL` as the hard gate
- One-time AI PDF enrichment flow driven by Apps Script
- Backward-compatible main sheet handling for the existing `Contract Links` tab
- Renewal sheet creation + version bumping
- Frontend actions for rescan, renewal, Zoho refresh, and notes

## Run locally
```bash
cd toabh-contracts-dashboard
npm install
cp .env.example .env
# current Apps Script web app URL is prefilled in .env.example
npm run dev
```

## Build
```bash
npm run build
```

## Apps Script setup
1. Open the existing Google Sheet:
   - Spreadsheet ID: `18Bv8MGdxsNS5MQNVhHMjdjVT-hhCFQsMcM2jAw2jNRI`
   - Main tab: `Contract Links`
2. Open Extensions → Apps Script.
3. Copy `apps-script/Code.gs` into the Apps Script project.
4. Copy `apps-script/appsscript.json` into the manifest.
5. In Script Properties, set:
   - `MAIN_SHEET_NAME` (default: `Contract Links`)
   - `RENEWALS_SHEET_NAME` (default: `Renewals`)
   - `GEMINI_API_KEY`
   - `CONTRACT_GENERATOR_WEBHOOK_URL` (your existing contract generation webhook)
   - `ZOHO_SIGN_WEBHOOK_URL` (your existing Zoho send webhook)
   - `ZOHO_STATUS_WEBHOOK_URL` (your existing Zoho status/signed-PDF sync webhook)
6. Run `setupSheetsAndTriggers()` once.
7. Update the Web App deployment so the same Apps Script URL serves the new dashboard actions.
8. Keep using `VITE_APPS_SCRIPT_URL` in the frontend.

## Notes
- AI scans only when `Signed PDF URL` exists and `AI Scan Status` is blank, `Pending`, `Failed`, or `Rescan Requested`.
- Scanned contracts are read from sheet values after first scan.
- `Rescan with AI` explicitly marks the row `Rescan Requested` and rescans immediately.
- Renewal row creation is done in the `Renewals` sheet.
- The main dashboard list only includes rows where `Signed PDF URL` is not blank.
- Existing unsigned / draft / pending rows remain in the original system and are not shown in the dashboard.
