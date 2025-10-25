# Voice Over Organizer (Google Sheets + Apps Script)

A Google Sheets sidebar that builds two views from your Google Drive folders:

- **Voice (final roles/releases)** with “Current / Past” status and optional Final Link
- **Auditions** with quick filter chips, instant search, and **auto-sync** of statuses from Voice

This repo ships ready-to-deploy with **clasp** so nobody needs to copy/paste code (no indentation/syntax headaches).

---

## 🔧 Prerequisites

- A Google account
- **Google Drive (web)** and (recommended) **Google Drive for desktop**
  - Install Drive for desktop and sign in
  - In Drive (web) create or choose two folders:
    - `Voice Projects` (final roles / releases)
    - `Auditions` (all audition takes)
  - Anything you place in the local synced folders will auto-appear in Drive

> ### 📌 Where do I find the folder ID?
> Open your folder in Drive and copy the URL. The **ID is the long string after `/folders/`**.
>
> Example:  
> `https://drive.google.com/drive/folders/1AbCDefGhijkLMNOPq` → **ID =** `1AbCDefGhijkLMNOPq`
>
> You can paste **either the whole URL _or_ just the ID** (the script can handle both if you wire that in).  
> In this version, you set the constants directly in `Code.gs` (see below).

---

## 🚀 Installation (No Copy/Paste) — Using `clasp`

> Requires Node.js (LTS).

1) Install and log in:
```bash
npm i -g @google/clasp
clasp login
```

2) Create a **new Google Sheet** (name it e.g. “Voice Tracker”). Copy the Spreadsheet ID from its URL:  
`https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit#gid=0`

3) Clone this repo and bind it to your Sheet:
```bash
git clone https://github.com/Squ1zM0/Voice-Over-Organizer.git
cd Voice-Over-Organizer

# Create the bound Apps Script project (ties it to your Sheet)
clasp create --parentId <SPREADSHEET_ID>

# Ensure .clasp.json has: { "scriptId": "...", "rootDir": "src" }
# If rootDir is missing, copy the example:
cp .clasp.json.example .clasp.json
# Then paste the scriptId from the file clasp just created, and keep "rootDir": "src"
```

4) Push the source:
```bash
clasp push
```

5) Open the Sheet → **Extensions → Apps Script** → run any function (e.g. `onOpen`) once and **Authorize**.

6) Edit IDs at the top of `src/Code.gs`:
```js
var VOICE_FOLDER_ID   = 'PASTE_VOICE_FOLDER_ID_HERE';
var AUDITION_FOLDER_ID = 'PASTE_AUDITION_FOLDER_ID_HERE';
```
> These are the Drive folder IDs from “Where do I find the folder ID?” above.

7) Back in the Sheet, use **Voice Tracker** menu:
- **Refresh Sheet** → builds the Voice sheet
- **Refresh Auditions** → builds the Auditions sheet
- **Show Tracker Panel** → open the sidebar (Dashboard, Updates, Search, Insights, Auditions, Actions)

---

## 🧰 Manual Install (fallback)

If a user can’t use `clasp`:

1) Create a new Google Sheet → **Extensions → Apps Script**
2) Create two files in the editor:
   - `Code.gs` → copy contents from `src/Code.gs`
   - `Sidebar.html` → copy contents from `src/Sidebar.html`
3) Click **Save**, then run any function (e.g., `onOpen`) once → **Authorize**
4) Edit the two constants at the top of `Code.gs`:
   ```js
   var VOICE_FOLDER_ID   = 'PASTE_VOICE_FOLDER_ID_HERE';
   var AUDITION_FOLDER_ID = 'PASTE_AUDITION_FOLDER_ID_HERE';
   ```
5) Return to the Sheet → **Voice Tracker** menu:
   - **Refresh Sheet** (Voice)
   - **Refresh Auditions**
   - **Show Tracker Panel**

> Manual paste can be error-prone due to hidden formatting—prefer the `clasp` path when possible.

---

## 🧭 Daily Use

- **Dashboard**: one-click filters for **Current / Past / All** (Voice sheet)
- **Updates**: update a single row or apply bulk **Final Link** + **Status**
- **Search**: filter Voice by **Character / Folder / File name** + **Status**
- **Insights**: KPIs + recent additions across Voice & Auditions
- **Auditions**: status chips (**All / Pending / Submitted / Booked / Passed**) + instant search
- **Actions**: export visible rows to **CSV** and a summary **PDF**

### Auto-Sync Rules
- Voice **Current** → Auditions **Booked**  
- Voice **Past** → Auditions **Submitted**  
- Never overwrite **Passed**

### Optional Automation
**Voice Tracker → Install Auto-Refresh (hourly)** to keep Voice in sync with Drive automatically.

---

## 🔍 Troubleshooting

- **“Invalid folder ID”**: Copy the folder URL directly from Drive and paste the long ID.
- **Nothing loads**: Open **Extensions → Apps Script**, run `onOpen` once, and authorize.
- **Chips/search not filtering (Auditions)**: Click **Refresh** in the Auditions tab, then try again.
- **Bulk update changed 0 rows**: Ensure selected **Character(s)** exist in Voice; if filtering by Files, the URLs must match exactly.

---

## 🤝 Contributing

PRs welcome. Please keep `src/Code.gs` and `src/Sidebar.html` in sync with the README and avoid introducing copy/paste setup.

License: **MIT**
