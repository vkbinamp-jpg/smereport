# Company Wise Health Checkup MIS (Static HTML)

A browser-only MIS dashboard built with HTML, CSS and vanilla JavaScript. It reads `.csv`, `.xlsx`, and `.xls` files from one fixed Google Drive folder using Google OAuth and the Drive API. It does not write to Drive and does not send patient rows to an application backend.

## Files

- `index.html` – page layout
- `styles.css` – responsive styling
- `app.js` – Google Drive read-only integration, CSV / Excel parsing, filters, MIS calculations
- `config.js.example` – configuration template; copy to `config.js`

## Important security limitations

1. The requested `admin/admin` login is a client-side page gate. Anyone who can inspect the published JavaScript can discover/bypass it. Do not use this as real access control for patient information.
2. GitHub Pages is static hosting. It cannot securely keep a password, OAuth client secret, service-account key, or private API credential. Never add a client secret or private key to this repository.
3. Google OAuth is the actual Google Drive authorization. Configure the OAuth consent screen for the intended organization/test users and restrict the app's allowed Google accounts where possible. The optional `ALLOWED_GOOGLE_EMAIL` field is displayed as configuration guidance only; this static implementation does not independently verify an email identity.
4. The OAuth access token and parsed patient rows are held in JavaScript memory only. The token is not written to localStorage/sessionStorage. Reloading the page clears them. The browser still temporarily receives file bytes to parse the workbook.
5. This app requests the `drive.readonly` scope. It does not call Drive create, update, rename, or delete endpoints.
6. For production use with real patient data, use a properly authenticated and access-controlled hosting/backend arrangement, even if the reporting UI remains static.

## Google Cloud setup

1. Create/select a Google Cloud project.
2. Enable **Google Drive API**.
3. Configure the OAuth consent screen. Add only the authorized users / organization as appropriate.
4. Create an OAuth Client ID of type **Web application**.
5. Add your exact GitHub Pages origin to **Authorized JavaScript origins**, e.g. `https://YOUR-USERNAME.github.io`. If using a project site, the origin is still `https://YOUR-USERNAME.github.io` (not the repository path).
6. Create an API key and restrict it to the Google Drive API and your GitHub Pages HTTP referrer/origin. The API key is not a substitute for OAuth.
7. Copy `config.js.example` to `config.js` and fill in:
   - `GOOGLE_CLIENT_ID`
   - `GOOGLE_API_KEY`
   - `DRIVE_FOLDER_ID` (the folder ID from the Google Drive folder URL)
8. Add `config.js` to the website root alongside `index.html`.

## Publish with GitHub Pages

1. Create a GitHub repository, e.g. `company-healthcheckup-mis`.
2. Upload `index.html`, `styles.css`, `app.js`, and your configured `config.js`.
3. Open **Settings → Pages**.
4. Select **Deploy from a branch**, choose `main` and `/ (root)`, then Save.
5. Wait for GitHub Pages to publish the site.
6. Add the published origin to the OAuth client's Authorized JavaScript origins.

Do not commit `config.js` to a public repository if you prefer not to expose your folder ID or API key. Note: a browser app necessarily exposes its OAuth client ID and any API key it uses. Restrict the API key by API and referrer. For a public repository, use a private repository or inject the public config at deploy time; GitHub Pages itself does not provide secret server-side runtime variables.

## How to use

1. Open the published site.
2. Login using the basic page gate `admin` / `admin` (change in `app.js` if you still want this convenience gate; it is not security).
3. Click **Connect Google Drive** and authorize the Google account that can read the folder.
4. The app reads direct child files with `.csv`, `.xlsx`, or `.xls` extensions from the configured folder.
5. Select a date range and optional company, UHID or file filters, then click **Apply Filters**.
6. Click a company name to view patient-level UHID details.
7. Click **Refresh Data** to re-read files added to the folder.

## Data and calculation notes

- Company grouping uses `Ref` (or recognized aliases such as Company Name).
- Patient identity uses `UHID` (or recognized aliases such as Hospital ID).
- A patient is counted once per normalized company within the filtered date range.
- The same UHID under two different companies counts once in each company.
- Company normalization trims and collapses whitespace and ignores case for grouping. It does not fuzzy-merge different names.
- Missing UHID and missing company rows are excluded from company unique-patient totals and shown in Data Quality.
- Duplicate rows are flagged by company + UHID and do not increase unique-patient count.
- Date filtering is applied before distinct UHID counting.
- The parser attempts to find the header row within the first 20 rows. If the required columns are not detected, the worksheet is reported as requiring manual mapping. Manual column mapping UI is not included in this first implementation and must be added after reviewing a real sample workbook.
- Current implementation lists files directly inside the configured folder (not nested subfolders).
- It processes all sheets in each workbook.
- The browser temporarily loads the workbook bytes and parsed rows in memory. Closing/reloading the page clears them.
- No patient-level data is written to application logs. The optional CSV export is initiated only when the user clicks Export CSV.

## Before using real data

Test with a copy containing dummy/anonymized patient data. Confirm actual column headers, date formats, header-row placement, and whether company names in `Ref` are consistent. The the billing summary CSV previously shared does not contain patient-level UHID/Ref records and cannot validate the patient parser.
