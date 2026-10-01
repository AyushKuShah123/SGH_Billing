# Shivam Gift House Billing

Responsive billing web app designed for Vercel. It works on mobile and desktop and includes:
- INR and NPR
- Admin-renamable app/shop name
- GST/PAN/custom document label
- Admin tax label and tax rate
- Amount/percentage discount
- Cash / UPI / Balance
- Auto bill number format `SGH-YYDDM0001`
- Amount in words
- Bill history and outstanding view
- Dark/light theme and admin accent color
- Browser printing
- Local browser persistence
- Optional Google Sheets sync through Vercel + Google Apps Script

## Deploy the website
1. Create a GitHub repository and upload this entire folder.
2. Go to https://vercel.com, sign in, choose **Add New → Project**, import the GitHub repository, and deploy.
3. No build command is required. Vercel serves `index.html` and the `/api` function.

## Google Sheets setup
1. Create a Google Sheet.
2. Open **Extensions → Apps Script**.
3. Copy `google-apps-script/Code.gs` into the Apps Script editor.
4. Replace `CHANGE_ME` for `SECRET` with a long random secret and replace `SHEET_ID` with the spreadsheet ID.
5. Deploy the Apps Script as a Web App, execute as the owner, and allow access to anyone with the link.
6. In Vercel Project Settings → Environment Variables add:
   - `GOOGLE_APPS_SCRIPT_URL` = your Apps Script web-app URL
   - `GOOGLE_SHEETS_SECRET` = exactly the same secret used in Code.gs
7. Redeploy Vercel.

The Apps Script automatically creates one tab for days 1–15 and another for days 16–end of the month.

## Important production note
The current project is a deployable front-end/MVP. Authentication and role enforcement should be added before using it for real business data. Do not put admin passwords or Google credentials in `index.html`.
