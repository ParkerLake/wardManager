# Ward Manager — Setup Guide

A step-by-step guide to get Ward Manager running on GitHub Pages.

---

## Prerequisites
- A Google account (the one you'll use as the "admin")
- Your ward's Google Sheet already created
- Git installed, and this repo cloned to your computer

---

## Step 1 — Prepare your Google Sheet

Your sheet needs these tabs (exact names, case-sensitive):

| Tab name      | Columns (Row 1 = header) |
|--------------|--------------------------|
| Appointments | Name · Status · Owner · Purpose · Appt Date · Notes |
| Callings     | Calling · Name · Stage · Notes |
| Releasings   | Calling · Name · Stage · Notes |
| Meetings     | Meeting Type · Date · Agenda Item · Owner · Notes · Done |
| Members      | Name · Calling · Phone · Email · Notes |

Copy the Spreadsheet ID from the URL:
`https://docs.google.com/spreadsheets/d/**SPREADSHEET_ID**/edit`

---

## Step 2 — Create a Google Cloud project & OAuth client

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a new project (or use an existing one)
3. Go to **APIs & Services → Library** and enable:
   - **Google Sheets API**
   - **Google Drive API**
4. Go to **APIs & Services → OAuth consent screen**
   - User type: **External**
   - Fill in App name (e.g. "Ward Manager"), support email, developer email
   - Scopes: add `openid`, `email`, `profile`, `https://www.googleapis.com/auth/spreadsheets`
   - Add test users: every email that needs access
5. Go to **APIs & Services → Credentials → Create Credentials → OAuth 2.0 Client ID**
   - Application type: **Web application**
   - Authorised JavaScript origins: `https://YOUR-USERNAME.github.io`
   - Click Create and copy the **Client ID**

---

## Step 3 — Configure the app

Edit `src/config.js`:

```js
const config = {
  GOOGLE_CLIENT_ID: "YOUR_CLIENT_ID.apps.googleusercontent.com",
  SPREADSHEET_ID:   "YOUR_SPREADSHEET_ID",
  ALLOWED_EMAILS: [
    "bishop@example.com",
    "firstcounsellor@example.com",
    "secondcounsellor@example.com",
    "clerk@example.com",
  ],
  WARD_NAME: "Your Ward Name",
};
```

---

## Step 4 — Enable GitHub Pages

1. Push your code to the `main` branch
2. Go to your repo on GitHub → **Settings → Pages**
3. Source: **GitHub Actions**
4. The first deploy will run automatically on push

Your app will be live at: `https://YOUR-USERNAME.github.io/wardManager`

---

## Step 5 — First login

1. Open the app URL
2. Click **Sign in with Google**
3. Click **Test Connection** to verify Sheets access
4. Click **↓ Load Data** to pull your data

---

## Adding more users

1. Add their email to `ALLOWED_EMAILS` in `config.js`
2. Add them as a **test user** in Google Cloud OAuth consent screen
3. Push to `main` — the app redeploys automatically in ~2 minutes

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| "Not authorised" on login | Email not in `ALLOWED_EMAILS` |
| Blank page after deploy | Check Actions tab for build errors |
| "Cannot reach spreadsheet" | Check `SPREADSHEET_ID` in config.js; ensure Sheets API is enabled |
| "400: redirect_uri_mismatch" | Add your GitHub Pages URL to Authorised JS Origins in Google Cloud Console |
| Missing tabs warning | Create the required tabs in your Google Sheet (see Step 1) |

---

## Local development

```bash
npm install
npm start   # runs at http://localhost:3000
```

For local dev, add `http://localhost:3000` to Authorised JavaScript Origins in Google Cloud Console.
