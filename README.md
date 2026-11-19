# INVO BUILD – SRB / PRA / GST Sales Tax Invoices

Static web app (no build step, no server cost). Data is stored in your Google Sheet via Apps Script.
Exports Excel (same A4 layout, live formulas) and PDF (A4, via Print → Save as PDF).

## Run locally (VS Code)
Open the folder in VS Code → install **Live Server** → right-click `index.html` → *Open with Live Server*.
With no backend URL configured the app stores data in the browser (good for testing).

## Go live on GitHub (free)
1. `git init && git add . && git commit -m "first version"`
2. Create a repo on GitHub, `git remote add origin <url>`, `git push -u origin main`
3. Repo → Settings → Pages → Branch `main` / root → Save. Your link: `https://<user>.github.io/<repo>/`

## Backend (two separate Google Sheets)
- **CONTROL sheet** (the sheet that holds the script): only the `Users` tab. This is where you approve / block clients.
- **DATA sheet** (`INVO BUILD - Data`, created automatically by `setup`): `Invoices` and `Settings` tabs. Each client's invoices and business details (name, address, NTN, default rates) are saved here, under their email, so they follow the account to any device.
Follow the steps at the top of `Code.gs`, then paste the Web App URL into `const API_URL=''` near the top of the script in `index.html` and push. While `API_URL` is empty the app runs in local test mode (data stays in the browser, no login).

## Login / signup and blocking
- Users sign up inside the app (email + password). The account starts **inactive**; you get an email.
- Approve: in the `Users` sheet set `Active = TRUE` (optionally an `Expiry` date). Your own account: sign up once, then set your own row to TRUE.
- Block: set `Active = FALSE` or let `Expiry` pass. It applies on the very next request (the app locks, nothing saves/prints/exports).
- Each user only sees their own invoices (stored under their email in the `Invoices` sheet). Passwords are stored salted + hashed, never in plain text.
- Basic security only: no email password-reset (reset by deleting the user's row so they sign up again, or ask them to), no login attempt limit.
- Enforcement lives in the backend, not the website code. Keep the Google Sheet / Apps Script owned by you only.

## Rates
Each line item has its own rate (any number; suggestions 5, 8, 10, 12, 13, 14, 15, 16, 17, 18, 20).
Defaults per type are set in Settings.

## Fixes versus the old Excel file
- GST sheet total row showed 15% while items used 18% → total rate now shows the real rate (blank if mixed).
- Tax was hardcoded (`E9*15/100`) → now always `Value × Rate` from the rate cell.

## Lines
No limit on item lines. Up to 20 lines fit on one A4 page; more lines continue on the next page (header repeats, total on the last page). Excel export grows its rows and `SUM` ranges automatically.