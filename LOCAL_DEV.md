# Local Development Guide

Run Ward Manager locally so you can test changes instantly without deploying.

---

## One-time setup

### 1 — Add localhost to Google Cloud Console

This is the only extra step needed for local dev.

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. APIs & Services → Credentials → click your OAuth 2.0 Client ID
3. Under **Authorised JavaScript origins**, add:
   ```
   http://localhost:3000
   ```
4. Save — takes effect immediately (no waiting)

You only need to do this once. Your GitHub Pages origin stays there too — both work at the same time.

---

### 2 — Install dependencies (only once)

```bash
cd wardManager   # or wherever your repo is
npm install
```

---

## Running locally

```bash
npm start
```

The app opens at **http://localhost:3000** automatically.

- Hot reload is on — save any file and the browser updates instantly
- No push/deploy needed
- Uses your real Google Sheet (same data as production)
- Sign in with Google works exactly the same as on GitHub Pages

---

## Typical workflow

1. `npm start` → app opens in browser
2. Make a change in `src/App.js` (or any file)
3. Save → browser auto-refreshes within 1-2 seconds
4. When happy → `git add -A && git commit -m "..." && git push`
5. GitHub Actions deploys to production in ~2 minutes

---

## Common issues

| Problem | Fix |
|---------|-----|
| "redirect_uri_mismatch" popup error | Add `http://localhost:3000` to Authorised JS Origins (Step 1 above) |
| Blank page / console error | Run `npm install` first |
| Port 3000 already in use | Another app is on 3000. Kill it or run `PORT=3001 npm start` then update your Google Cloud origin to match |
| Changes not showing | Make sure you saved the file — CRA hot reload requires a save |

---

## Notes

- Local and production share the **same Google Sheet** — any saves you make locally write to the real sheet
- If you want a separate test sheet, set up a second config (swap `SPREADSHEET_ID` in `config.js` before testing)
- Your login session persists in `sessionStorage` — you won't need to sign in again unless you close the tab
