VOICE PROJECT TRACKER — README
================================

A simple, script-powered dashboard that scans your Google Drive voice-over folders, lists every file in a Google Sheet, lets you add “Final Production” links, track Current/Past status, and sync statuses with your separate Auditions tab. It also includes a small HTML UI (buttons, tabs, KPIs) you can open in the Sheet.

--------------------------------------------------------------------
1) WHAT THIS TOOL DOES
--------------------------------------------------------------------
- Recursively scans two Google Drive folders:
  • Voice Projects (finished or in-progress recordings)
  • Auditions (raw/alt takes, submissions)

- Writes rows into your Google Sheet with:
  • Folder path
  • File name
  • Character (auto-parsed from file name)
  • Date added
  • MIME type
  • File link
  • Final Production link
  • Status

- Syncs Voice status → Auditions status:
  • Voice "Current" → Auditions "Booked"
  • Voice "Past"    → Auditions "Submitted"
  • Never overwrites "Passed"

- Adds a sidebar UI: filters, quick search, bulk updates, CSV/PDF export, Insights, and a Donate link/button.


--------------------------------------------------------------------
2) PREREQUISITES
--------------------------------------------------------------------
- A Google account
- Google Sheets
- Access to Google Drive (web)
- (Optional, recommended) Google Drive for desktop installed and signed in


--------------------------------------------------------------------
3) INSTALL GOOGLE DRIVE FOR DESKTOP (TO KEEP LOCAL FILES SYNCED)
--------------------------------------------------------------------
Windows
  1. Search the web for “Google Drive for desktop download” (official Google page).
  2. Download and run the installer.
  3. Sign in with your Google account.
  4. In Drive settings:
     - Choose Stream files (recommended) or Mirror files.
     - Confirm your local Drive letter (Windows often maps it to G:).
  5. You will see Google Drive in File Explorer. Anything in My Drive or Shared drives will sync.

macOS
  1. Download Google Drive for desktop from Google’s site.
  2. Open the .dmg, drag the app into Applications.
  3. Sign in with your Google account.
  4. In Finder, you will see Google Drive mounted in the sidebar.

Tip: Work from your desktop DAW while Drive syncs to the cloud. The script reads the cloud folders—sync keeps the cloud in step with your desktop.


--------------------------------------------------------------------
4) CREATE THE GOOGLE SHEET + APPS SCRIPT
--------------------------------------------------------------------
1. Create a new Google Sheet (e.g., “Voice Project Tracker”).
2. Click Extensions → Apps Script.
3. In the Script Editor:
   - Create a file named Code.gs and paste the entire server script
     (everything from “// ========== MAIN SCRIPT ==========” down to the last function).
   - Create a file named Sidebar.html and paste the entire HTML block
     (everything from “<!DOCTYPE html>” to “</html>”). You may keep the Donate
     link as an <a> with target="_blank".
4. Click Save.


--------------------------------------------------------------------
5) CONFIGURE YOUR FOLDER IDS
--------------------------------------------------------------------
At the top of Code.gs, set these constants:

  var VOICE_FOLDER_ID    = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX';
  var AUDITION_FOLDER_ID = 'YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY';

How to find a folder ID:
- In Drive (web), open the folder.
- The URL looks like: https://drive.google.com/drive/folders/1AbCDeFGhijkLMNOPqR...
- Copy the long string after “/folders/” — that is the folder ID.


--------------------------------------------------------------------
6) FILE NAMING FORMAT (CHARACTER AUTO-DETECTION)
--------------------------------------------------------------------
The script extracts the Character from each file’s name using this pattern:

  Fred French - <Character Name> - <anything else>.wav

In collectFiles(), the line is:

  var match = fileName.match(/Fred French - (.+?) -/);

It looks for “Fred French - ” followed by the character name, then “ -”.

Examples that parse correctly:
  - Fred French - Captain Orion - Take 3.wav    → Character = Captain Orion
  - Fred French - Shop Announcer - v2.mp3       → Character = Shop Announcer

If your display name is not “Fred French”, choose one of the following:

Option A — Replace the literal name
  Change the regex to match your name:
    var match = fileName.match(/Your Name - (.+?) -/);

Option B — Make it more generic
  If you ensure the file names follow:
    <Performer Name> - <Character Name> - <rest>
  you can generalize the regex and ignore the performer name entirely:
    var match = fileName.match(/^[^-]+-\s*(.+?)\s*-/);
  This captures the text between the first and second hyphen blocks as Character.

Recommendation:
  Adopt a consistent naming standard, for example:
    Your Name - Character - Project - 2025-11-10 - take01.wav

If a file does not match the pattern, the script sets Character to “Unknown”.


--------------------------------------------------------------------
7) USING THE UI & WHAT EACH BUTTON DOES
--------------------------------------------------------------------
After saving scripts:
1) Return to the Sheet and refresh.
2) Run any function when prompted and accept permissions.

You will see a “Voice Tracker” menu in the Sheet. Menu items:
- Refresh Auditions: scans the Auditions folder and rewrites the Auditions tab.
- Sync Auditions from Voice: applies the logic (Current → Booked, Past → Submitted).
- Show Tracker Panel: opens the sidebar UI (tabs, buttons).
- Refresh Sheet: scans the Voice Projects folder and rewrites the main tab.
- Export Summary PDF: saves a PDF snapshot of the whole Sheet into your Voice Projects folder.
- Install Auto-Refresh (hourly): adds a time trigger for hourly refresh.
- Remove Auto-Refresh: deletes that time trigger.
- Export Visible Auditions CSV / Voice CSV: exports only currently visible (filtered) rows to CSV.

In the Sidebar (HTML UI):
- Dashboard: quick filter buttons for Voice (Current/Past/All).
- Updates:
    Single update: choose a Character, optional specific File URL, add a Final Production link and/or set Status.
    Bulk update: choose multiple Characters (and optional specific File(s)), then apply a single Final link or Status to all.
- Search: free-text filters for Folder/File/Character plus a Status filter.
- Insights: KPIs and a Recent Activity table compiled from both sheets.
- Auditions: status chips (All/Pending/Submitted/Booked/Passed), quick search, Sync Status button.
- Actions: quick links to refresh or export PDF.


--------------------------------------------------------------------
8) FIRST-RUN CHECKLIST
--------------------------------------------------------------------
- Confirm VOICE_FOLDER_ID and AUDITION_FOLDER_ID.
- Ensure file names follow the pattern (section 6).
- Open Voice Tracker → Refresh Sheet.
- Open Voice Tracker → Refresh Auditions.
- Open Voice Tracker → Show Tracker Panel, then try:
    Dashboard: Current / Past / All
    Updates: add a Final Production link to one file, set Status
    Auditions tab: click Sync Status, verify changes
- Optional: Install Auto-Refresh (hourly).


--------------------------------------------------------------------
9) TROUBLESHOOTING
--------------------------------------------------------------------
Authorization prompts
  - Run any menu action and approve OAuth permissions.

No rows appear after Refresh
  - Check folder IDs.
  - Confirm the Drive folders actually contain files.
  - Confirm your account has access (especially for Shared drives).

Character shows “Unknown”
  - File name did not match the expected pattern.
  - Fix the file name or switch to the generic regex.

Donate link does not open
  - Use an anchor tag (<a href="...">) with target="_blank" instead of a button with window.open().

Auditions status not updating
  - Use Voice Tracker → Sync Auditions from Voice (or “Sync Status” in the sidebar).
  - Confirm file names in Auditions match those in Voice (same Character + same File Name).

Performance is slow with very large folders
  - Organize into subfolders to keep lists manageable.


--------------------------------------------------------------------
10) CUSTOMIZATIONS (OPTIONAL)
--------------------------------------------------------------------
- Stage name:
  Change the regex to your name or use the generic version.

- Status words:
  Auditions: ["Pending", "Submitted", "Booked", "Passed"]
  Voice:     ["Current", "Past"]
  If you add new statuses, update conditional formatting and sync logic in syncAuditionsWithProjects().

- Auto-refresh frequency:
  In installAutoRefresh(), change .everyHours(1) to .everyMinutes(15) if desired.


--------------------------------------------------------------------
11) RECOMMENDED FILE NAME TEMPLATE
--------------------------------------------------------------------
Your Name - Character Name - Project Name - 2025-11-10 - take01.wav


--------------------------------------------------------------------
12) QUICK START (TL;DR)
--------------------------------------------------------------------
1. Install Google Drive for desktop and sign in; let it sync your voice and audition folders.
2. Create a Google Sheet → Extensions → Apps Script.
3. Paste Code.gs and Sidebar.html.
4. Set VOICE_FOLDER_ID and AUDITION_FOLDER_ID in Code.gs.
5. In the Sheet, open Voice Tracker menu:
   - Refresh Sheet
   - Refresh Auditions
   - Show Tracker Panel
6. Use file names like “Your Name - Character - …” for auto Character detection.
7. Use Updates to add Final Production links and set statuses.
8. Click Sync to propagate Voice statuses into Auditions.

You’re ready to roll.
