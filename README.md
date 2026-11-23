# Promoter Share Portfolio

**App:** https://grgsuman23.github.io/Promoter-Porfolio/  ·  **Price feed:** https://raw.githubusercontent.com/grgsuman23/Promoter-Porfolio/main/prices.json

A mobile-installable app (PWA) hosted on **GitHub Pages**. Your data stays in **your Google Sheet**, which the app reaches through a small **Apps Script** API protected by your **app PIN**. A **GitHub Action** fetches nepalstock.com prices every 15 minutes during market hours.

```
Phone / laptop app (GitHub Pages)  ──PIN──▶  Apps Script API  ──▶  Google Sheet (your data)
                                                   ▲
GitHub Action ── nepalstock.com ──▶ prices.json ───┘  (read on each LTP refresh)
```

> This repo must be **public** for free GitHub Pages. It contains only app code and public market prices: no holdings and no PIN. Never write your PIN or Sheet data into any file here.

---

## 1. Backend: Google Sheet + Apps Script (one time)
1. Create a Google Sheet → **Extensions → Apps Script**.
2. Paste `apps-script/Code.gs` into `Code.gs`. Add an HTML file named **Index** and paste `apps-script/Index.html`. In ⚙ Project Settings tick *Show appsscript.json* and paste `apps-script/appsscript.json`. Save.
3. Run **setup** (approve permissions: Advanced → Go to project).
4. Reload the Google Sheet. In the new menu choose **Promoter Portfolio → Set app PIN / password**, enter your PIN (at least 4 characters) twice. It is stored only as a salted hash in the script's settings.
5. **Deploy → New deployment → Web app**: Execute as **Me**, Who has access **Anyone** → Deploy → copy the **Web app URL** (`…/exec`).
   *"Anyone" is needed so the phone app can reach it without a Google login. Every request still needs a sign-in made with your PIN, and the Apps Script-hosted page refuses anyone except you.*

Already using the Apps Script version? Replace `Code.gs` and `Index.html`, save, reload the Sheet, then do steps 4–5. The old "Only myself" deployment keeps working.

## 2. Host the app on GitHub Pages
1. ✅ Done: this repo holds the app.
2. **Settings → Pages → Build and deployment**: Source **Deploy from a branch**, Branch **main**, folder **/ (root)** → Save.
3. After about a minute the app is live at **https://grgsuman23.github.io/Promoter-Porfolio/**
4. Open it, paste the **Web app URL**, enter your **PIN**, and click **Connect**.

## 3. Install it on your phone
- **Android (Chrome)**: open the link → tap **Install app** (top bar) or ⋮ → **Install app / Add to Home screen**.
- **iPhone (Safari)**: open the link → Share ⬆ → **Add to Home Screen**.

It opens full-screen with its own icon. The last loaded portfolio is kept on the device, so it opens offline and refreshes when you're back online.

**PIN & sign-in**
- Each device signs in with the PIN and stays signed in for **30 days**. The PIN itself is never stored on the phone; only a sign-in token is.
- 5 wrong PINs lock sign-in for **1 hour**.
- **Settings → Sign out** signs this device out. **Disconnect this device** also forgets the web app link and the offline copy.
- Lost a phone? In the Sheet use **Promoter Portfolio → Sign out all devices**, or set a new PIN (this also signs out every device).
- A 4-digit PIN is easy to remember but also easier to guess. Keep the web app URL private too.

**Want an actual .apk file?** Go to [pwabuilder.com](https://www.pwabuilder.com), enter your GitHub Pages link → **Package for stores → Android** → download, then install the `.apk` on your phone (allow "install unknown apps"). This is optional; the installed PWA behaves the same.

## 4. Automatic nepalstock.com prices
1. In the repo: **Actions** → enable workflows → **NEPSE price feed → Run workflow**. About a minute later `prices.json` appears in the repo.
2. In the app: **Prices & Settings → GitHub price feed** → paste `https://raw.githubusercontent.com/grgsuman23/Promoter-Porfolio/main/prices.json` → **Save** → **Refresh LTP**.
3. Turn on **Auto-refresh during market hours** so the Sheet's values update by themselves.

The workflow runs every 15 min, Sun–Thu, 10:45–15:30 Nepal time. GitHub may delay scheduled runs a few minutes, and each update also republishes the Pages site, which is harmless. If a run fails, open **Actions → the failed run** to see why; the app falls back to Merolagani or your manual rates.

## Updating the app later
Replace `index.html` in the repo (and `apps-script/*` in Apps Script, then **Deploy → Manage deployments → ✏ → New version**). In `sw.js` bump `VERSION` (e.g. `psp-v2` → `psp-v3`) so installed phones pick up the new version.

## Files
| Path | What it is |
|---|---|
| `index.html` | The app (same file as `apps-script/Index.html`) |
| `manifest.webmanifest`, `sw.js`, `icons/` | Makes it installable and work offline |
| `apps-script/` | Backend to paste into Google Apps Script |
| `fetch_prices.py`, `.github/workflows/nepse-prices.yml` | nepalstock.com price feed → `prices.json` |
