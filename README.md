# Drive Image Exporter

A browser-based React app that scans Google Drive folders and exports image metadata and links to an `.xlsx` workbook.

## Features

- Paste a Google Drive folder URL or folder ID.
- Google OAuth 2.0 sign-in using read-only `drive.readonly` scope.
- Lists image files in the selected folder.
- Optional recursive scan of nested folders.
- Exports Image Name, Format, Copy Link URL, Direct Image URL, and Folder Path.
- Excel generation happens in the browser. No app backend or database is required.
- Supports pagination and Shared Drive listing parameters.

## Requirements

- Node.js 18+
- A Google Cloud project with the Google Drive API enabled.
- OAuth 2.0 Client ID (Web application).

## Run locally

```bash
npm install
npm run dev
```

Open the local URL printed by Vite, normally `http://localhost:5173`.

## Google Cloud configuration

1. Open [Google Cloud Console](https://console.cloud.google.com/).
2. Create or select a project.
3. Enable **Google Drive API**.
4. Configure Google Auth Platform / OAuth consent screen.
   - Choose External if this is for consumer Google accounts.
   - While the app is in Testing, add the Google accounts that will use it as test users.
   - For a public production app, complete Google's verification requirements if prompted.
5. Create an OAuth Client ID. Choose **Web application**.
6. Add Authorized JavaScript origins:
   - `http://localhost:5173` for local development.
   - `https://YOUR-VERCEL-DOMAIN.vercel.app` after deployment.
   - Add any custom domain origin if you use one.
7. No redirect URI is needed for this Google Identity Services token flow.
8. Copy the Client ID ending in `.apps.googleusercontent.com`.

In the app, open **Settings**, paste the OAuth Client ID, save, then click **Continue with Google**.

## Deploy to Vercel

1. Push this project to a GitHub repository.
2. In Vercel, choose **Add New → Project** and import the repository.
3. Keep the framework preset as **Vite**. Build command: `npm run build`; output directory: `dist`.
4. Deploy.
5. Add the exact production origin (for example `https://your-project.vercel.app`) to the OAuth client's **Authorized JavaScript origins** in Google Cloud Console.
6. Open the deployed app, enter the OAuth Client ID in Settings, and sign in.

No server-side environment variable is required because the OAuth Client ID is a public browser identifier, not a client secret. Never put an OAuth client secret in this app.

## Important notes

- The app only requests `https://www.googleapis.com/auth/drive.readonly`.
- The signed-in Google account must have access to the target folder. A shared folder can be accessed if that account has permission.
- Public-link access without signing in is intentionally not implemented. The app uses the authenticated user's Drive permissions.
- The `Direct Image URL` column uses the same `https://lh3.googleusercontent.com/d/FILE_ID` pattern as the Apps Script version. Google does not guarantee this URL format as a permanent public delivery API; access may depend on file sharing and can change. The Drive `Copy Link URL` is the more reliable link to retain.
- Files are listed from folders the signed-in account can access. Shortcuts are not automatically followed.
- OAuth access tokens are held in page memory and cleared when the page is refreshed or disconnected. The Client ID is saved in browser local storage.
- If the OAuth consent screen is in Testing, only configured test users can sign in. Google may also expire test-mode refresh/access grants according to its testing rules.

## Project structure

```text
drive-image-exporter/
├── index.html
├── package.json
├── README.md
└── src/
    ├── App.jsx
    ├── main.jsx
    └── styles.css
```
