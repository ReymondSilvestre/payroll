# ARK Payroll - GitHub frontend + Google Apps Script backend

## 1. Backend (Google Sheets database)
1. script.google.com -> New project. Paste `apps-script/Code.gs` (there is no HTML file in Apps Script now).
2. Project Settings -> show `appsscript.json` and replace it with:
```json
{"timeZone":"Asia/Manila","exceptionLogging":"STACKDRIVER","runtimeVersion":"V8",
 "webapp":{"executeAs":"USER_DEPLOYING","access":"ANYONE_ANONYMOUS"},
 "oauthScopes":["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive","https://www.googleapis.com/auth/script.scriptapp"]}
```
3. Run `setup` once, authorize, and copy the admin password from the Execution log.
4. Deploy -> New deployment -> Web app (Execute as: Me, Access: Anyone). Copy the URL ending in `/exec`.
   After any code change: Deploy -> Manage deployments -> Edit -> New version.

## 2. Frontend (GitHub)
1. Open `frontend/app.js` and replace `PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE` with your `/exec` URL.
2. Upload `index.html`, `style.css`, `app.js` to a GitHub repository (root, or keep them in `/frontend`).
3. Settings -> Pages -> deploy from branch `main` (folder `/` or `/frontend`). Open the Pages URL and sign in as `admin`.

## Security
The `/exec` URL is public but harmless: every action needs a login token and is authorized by role on the server. The spreadsheet ID and all data stay inside Apps Script; never share the spreadsheet. Do not put passwords or the spreadsheet ID in the frontend files.
Daily backups go to Drive folder `ARK_Payroll_Backups`. Recovery: set Script Property `DB_ID` to a backup file's ID.
Verify the SSS, PhilHealth, Pag-IBIG and BIR tables against current official circulars before live use.
