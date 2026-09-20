# ZALVRA Store — Google Sheets Database Setup

## Quick setup (auto-creates sheets)

You only need a blank spreadsheet. The script creates **Products** and **Orders** sheets and headers for you.

### Step 1: Create a spreadsheet

1. Go to [Google Sheets](https://sheets.google.com) → **Blank** spreadsheet
2. Name it e.g. `ZALVRA Store`

### Step 2: Add the Apps Script

1. **Extensions → Apps Script**
2. Delete any default code
3. Paste everything from `GoogleAppsScript.js`
4. **Save** and name the project `ZALVRA API`

### Step 3: Deploy as Web App

1. **Deploy → New deployment**
2. Gear icon → **Web app**
3. Settings:
   - **Description:** ZALVRA Store API
   - **Execute as:** Me
   - **Who has access:** Anyone
4. **Deploy** → authorize when asked
5. **Copy the Web App URL**
   Example: `https://script.google.com/macros/s/AKfycb.../exec`

### Step 4: Connect Admin

1. Open `admin.html` → login (`zalvra@admin.store` / `zalvra`)
2. Paste the URL under **Google Sheets Database**
3. **Save URL** → **Test Connection**
4. You should see **Connected to Google Sheets**

Sheets are created automatically on first connection (no manual headers needed).

---

## What syncs

| Action | Direction |
|--------|-----------|
| Place order on store | → Orders sheet |
| Add / Edit / Delete product in admin | → Products sheet |
| Update order status / Delete order | → Orders sheet |
| Store homepage / category products | ← Products sheet |
| Admin orders list | ← Orders sheet |

---

## Speed features

- **Read cache** (~45 seconds) so the website loads products/orders faster
- **Writes** only update the sheet (no heavy re-read in the same request)
- **Auto-setup** once per spreadsheet
- **Efficient row lookups** by id column only when updating/deleting

After changing the script code: **Deploy → Manage deployments → Edit → New version**.

---

## Without Google Sheets

Leave the Script URL empty. Everything works in **local mode** (browser storage only) on that device.

---

## Product images column

`images` = pipe-separated ImgBB URLs:

`https://i.ibb.co/a.jpg|https://i.ibb.co/b.jpg`

Upload images in **Admin → Products**.
