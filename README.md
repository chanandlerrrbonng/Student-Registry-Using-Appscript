# NoQs Registry — Student Management System

A Google Apps Script web app that runs entirely inside Google Sheets. It gives you a dark-mode student records dashboard with AI-powered bulk import (Gemini), role-based access control, audit logging, analytics charts, and at-risk email alerts — no server or database required.

---

## Table of Contents

1. [Prerequisites](#1-prerequisites)
2. [Create the Google Sheet](#2-create-the-google-sheet)
3. [Open the Apps Script Editor](#3-open-the-apps-script-editor)
4. [Copy the Code Files](#4-copy-the-code-files)
5. [Run the One-Time Setup](#5-run-the-one-time-setup)
6. [Get a Gemini API Key](#6-get-a-gemini-api-key)
7. [Add the API Key to Script Properties](#7-add-the-api-key-to-script-properties)
8. [Add Yourself as an Editor](#8-add-yourself-as-an-editor)
9. [Deploy as a Web App](#9-deploy-as-a-web-app)
10. [Open the App](#10-open-the-app)
11. [Verify Everything Works](#11-verify-everything-works)
12. [Project File Reference](#12-project-file-reference)
13. [Troubleshooting](#13-troubleshooting)

---

## 1. Prerequisites

Before you start, make sure you have:

- A **Google account** (personal Gmail or Google Workspace).
- Access to **Google Drive** and **Google Sheets** (free).
- Access to **Google AI Studio** to create a free Gemini API key (free tier is sufficient).

No installs, no Node.js, no CLI tools — everything runs in the browser.

---

## 2. Create the Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and click **+ Blank** to create a new spreadsheet.
2. Give it any temporary name — the setup script will rename it to **"NoQs Registry — Student Management System"** automatically.
3. Keep this tab open. You'll come back to it after pasting the code.

> **Important:** The spreadsheet must be open (not just Drive) when you run the setup. The script is "container-bound" — it lives inside this specific file.

---

## 3. Open the Apps Script Editor

From inside your new Google Sheet:

1. Click **Extensions** in the top menu bar.
2. Click **Apps Script**.

A new browser tab opens showing the Apps Script editor. By default it has one file called `Code.gs` with an empty `myFunction`. You'll replace everything in this editor with the project files.

---

## 4. Copy the Code Files

The project has **7 files** — 4 script files (`.gs`) and 3 HTML files. You need to create each one in the editor and paste the matching content from the GitHub repo.

### Step-by-step for each file

#### A — Rename the default file to `setup`

The editor starts with a file called `Code.gs`. Rename it first:

1. In the left panel, hover over `Code.gs` and click the **⋮ (three dots)** menu.
2. Click **Rename** and type `setup` (the editor adds `.gs` automatically).
3. Select all the placeholder code inside it and **delete it**.
4. Paste the contents of **`setup.gs`** from the GitHub repo.

#### B — Add the remaining `.gs` files

For each of the files below, click the **+** button next to "Files" in the left panel, choose **Script**, name it exactly as shown, then paste the matching file content.

| File name to create | Source file in GitHub repo |
|---|---|
| `Service` | `Service.gs` |
| `Repository` | `Repository.gs` |
| `AI` | `AI.gs` |
| `Auth` | `Auth.gs` |

#### C — Add the HTML files

For each HTML file, click **+** → choose **HTML**, name it exactly as shown, then paste the content.

| File name to create | Source file in GitHub repo |
|---|---|
| `Index` | `Index.html` |
| `Script` | `Script.html` |
| `Styles` | `Styles.html` |

> **Names are case-sensitive.** `Index` ≠ `index`. Match them exactly or `doGet()` and `include()` will break.

#### D — Save everything

Press **Ctrl + S** (or **Cmd + S** on Mac) after pasting each file, or click the floppy-disk **Save** icon. Make sure none of the file tabs show an unsaved indicator (a dot next to the name).

---

## 5. Run the One-Time Setup

This step builds all three sheets (Students, Config, Audit\_Log), applies formatting, data validation, named ranges, and seeds 15 sample student records.

1. In the Apps Script editor, open **`setup.gs`** by clicking it in the left panel.
2. In the toolbar, click the **function dropdown** (it may say "Select function") and choose **`buildEntireRegistry`**.
3. Click the **▶ Run** button.

**First run — authorization prompt:**

Google will show a dialog: *"Authorization required."*

1. Click **Review permissions**.
2. Choose your Google account.
3. You may see a **"Google hasn't verified this app"** warning — this is expected for personal scripts.
4. Click **Advanced** → **Go to NoQs Registry (unsafe)**.
5. Click **Allow**.

The script will run. Watch the **Execution log** at the bottom of the editor — it should print a series of `✓` lines and end with:

```
=== NoQs Registry Setup: Complete ===
```

Switch back to your Google Sheet tab. You should now see a **Students** tab with 15 sample rows, a dark teal tab colour, and the spreadsheet renamed to "NoQs Registry — Student Management System".

> **If you need to re-run setup** (e.g. after making changes), it is safe to run `buildEntireRegistry()` again on an **empty** Students sheet. On a sheet that already has data it skips the seed step to protect your records.

---

## 6. Get a Gemini API Key

The AI bulk import feature uses Google's Gemini API. You need a free API key from Google AI Studio.

1. Go to [aistudio.google.com](https://aistudio.google.com).
2. Sign in with the **same Google account** you used for the Sheet (not strictly required, but keeps things tidy).
3. In the left sidebar, click **Get API key**.
4. Click **Create API key**.
5. In the dropdown, select **"Create API key in new project"** (or pick an existing Google Cloud project if you have one).
6. Click **Create API key**.
7. A key will be generated — it looks like `AIzaSy...`. Click the **copy icon** to copy it.

> **Keep this key private.** Do not paste it into your code files or commit it to GitHub. You will add it through Script Properties in the next step, which keeps it out of version control.

The free tier of Gemini API is sufficient for normal use of this app. Rate limits apply but the code already handles retries automatically.

---

## 7. Add the API Key to Script Properties

Script Properties are a secure, server-side key-value store — they are never visible in your code or to the browser.

1. Go back to the **Apps Script editor** tab.
2. Click the **⚙ Project Settings** icon in the left sidebar (looks like a gear).
3. Scroll down to the **Script properties** section.
4. Click **Add script property**.
5. In the **Property** field, type exactly:
   ```
   GEMINI_API_KEY
   ```
6. In the **Value** field, paste the API key you copied from AI Studio.
7. Click **Save script properties**.

The AI parsing feature will now work. The key is stored securely and never leaves Google's servers.

---

## 8. Add Yourself as an Editor

The app uses role-based access control (RBAC). Only emails listed in the Config sheet's **Editor Allowlist** can add, edit, delete, or import students. Viewers can only browse and export.

The setup script automatically adds the email of whoever ran `buildEntireRegistry()` as the first editor. To verify or add more editors:

1. In your **Google Sheet**, click **Extensions → Apps Script**.
2. In the editor, click **View → Logs** (or check the Execution Log from a prior run) — the setup log shows which email was added.

To add more editors manually:

1. In the Google Sheet, click **View → Hidden sheets** if you don't see the Config tab, then click **Config** to unhide it temporarily.
2. In **column J** (labelled `EDITOR EMAILS`), add one email address per row starting from **J2**.
3. Re-hide the sheet by right-clicking its tab → **Hide sheet**.

> Email matching is case-insensitive. The default role for anyone *not* in the list is `viewer` (set in cell **K2** of the Config sheet).

---

## 9. Deploy as a Web App

This step generates the public URL that anyone can visit to use the app.

1. In the Apps Script editor, click **Deploy** (top-right) → **New deployment**.
2. Click the **⚙ gear icon** next to "Select type" and choose **Web app**.
3. Fill in the settings:

   | Setting | Value |
   |---|---|
   | **Description** | `NoQs Registry v1` (or anything you like) |
   | **Execute as** | **Me** (`your@email.com`) |
   | **Who has access** | **Anyone** (or "Anyone with Google account" if you want login-gated access) |

4. Click **Deploy**.
5. You may be asked to authorize again — click **Authorize access** and repeat the same steps as in Step 5.
6. After deployment, a dialog shows your **Web app URL**. It looks like:
   ```
   https://script.google.com/macros/s/AKfycb.../exec
   ```
7. Copy this URL and click **Done**.

> **Every time you edit the code**, you must create a **New deployment** (or click **Deploy → Manage deployments → Edit → New version**) for changes to take effect on the live URL. The URL itself stays the same.

---

## 10. Open the App

Paste the Web app URL into any browser tab and press Enter.

You should see the **NoQs Registry** dashboard with:
- A dark sidebar with Records, Analytics, and (for editors) Bulk Import tabs.
- The 15 sample student records in the table.
- Your role shown by which UI elements are visible — editors see Add Student and Bulk Import; viewers only see the table and analytics.

---

## 11. Verify Everything Works

Run the built-in health check to confirm all sheets and settings are correct:

1. In the Apps Script editor, open **`setup.gs`**.
2. In the function dropdown, select **`runHealthCheck`**.
3. Click **▶ Run**.
4. Open the **Execution log** — you should see:

```
=== Health Check ===
Students  : EXISTS
Config    : EXISTS
Audit_Log : EXISTS
  Data rows          : 15
  Columns            : 10
  Frozen rows        : 1
  Bandings           : 1
  Cond. format rules : 7
  Subject validation : SET
  Config hidden      : true
  ID Counter         : 15
  Subjects           : [Mathematics, Science, English, ...]
  Audit_Log hidden   : true
  Audit entries      : 15
  Named ranges       : [STUDENT_DATA, SUBJECT_LIST, ID_COUNTER]
=== Health Check Complete ===
```

Then test the AI import:

1. Open the web app URL.
2. Click **Bulk Import** in the sidebar (only visible if you're an editor).
3. Paste a few lines like:
   ```
   Aryan Kapoor 84 Mathematics 12/03/2005
   Priya Mehta 45 Bio 22.09.2006
   Rohan Verma 92 CS 15th July 2003
   ```
4. Click **Parse with Gemini**.
5. A preview table should appear. Click **Import Selected Rows**.
6. The dashboard will refresh with the new students added.

---

## 12. Project File Reference

| File | Purpose |
|---|---|
| `setup.gs` | One-time sheet builder. Run `buildEntireRegistry()` once. Also contains `runHealthCheck()` and `getSubjectList()`. |
| `Service.gs` | Business logic layer. All functions exposed to `google.script.run` live here. Validates input, never touches the sheet directly. |
| `Repository.gs` | Data access layer. The **only** file allowed to call `SpreadsheetApp`. Handles locking, caching, ID generation, and audit writes. |
| `AI.gs` | Gemini API integration. Proposes rows only — never writes. Includes retry logic, model fallback, and a local regex parser as offline fallback. |
| `Auth.gs` | Role-based access control. Reads the editor allowlist from the Config sheet and resolves each visitor's role (`editor` or `viewer`). |
| `Index.html` | Main HTML template served by `doGet()`. Pulls in Styles and Script via `include()`. |
| `Script.html` | All frontend JavaScript — state management, table rendering, form handling, AI import flow, analytics charts. |
| `Styles.html` | All CSS — dark mode design system, responsive layout, component styles. |

---

## 13. Troubleshooting

**"Authorization required" keeps appearing**
Re-run `buildEntireRegistry()` and go through the authorization flow again. This happens once per Google account.

**"You have view-only access" when you should be an editor**
Your email is not in the Config sheet's Editor Allowlist (column J). Follow Step 8 to add it, then reload the web app.

**"GEMINI\_API\_KEY missing" error in Bulk Import**
The Script Property was not saved correctly. Repeat Step 7. Make sure the property name is spelled exactly `GEMINI_API_KEY` with no spaces.

**AI parsing returns "HTTP 503"**
This is a transient Gemini outage. The code retries automatically up to 3 times. If it still fails, the local regex parser kicks in as a fallback — your data will still be parsed without Gemini.

**Changes to the code aren't showing in the web app**
You must create a **new deployment version** after every code change. Go to **Deploy → Manage deployments**, click the pencil ✏ icon, change the version to **"New version"**, and click **Deploy**.

**The Students sheet is empty after running setup**
This can happen if the script timed out. Run `buildEntireRegistry()` again — it is safe to re-run on an empty sheet.

**"Students sheet not found. Run buildEntireRegistry() first."**
The setup hasn't been run yet, or the sheet was renamed/deleted. Run `buildEntireRegistry()` from the `setup.gs` file.

**The web app URL returns a blank page**
Make sure `Index.html`, `Script.html`, and `Styles.html` are all present in the editor with those exact names (capital I, capital S). Then create a new deployment.

---

## Architecture Overview

```
Browser (Web App URL)
        │
        ▼
   Index.html          ← served by doGet() in setup.gs
   Script.html         ← frontend JS, state, rendering
   Styles.html         ← CSS, dark mode
        │
        │  google.script.run (RPC)
        ▼
   Service.gs          ← validates input, builds responses
        │
        ├──▶ AI.gs          ← Gemini API calls + local fallback
        ├──▶ Auth.gs         ← RBAC: editor / viewer resolution
        └──▶ Repository.gs   ← reads/writes Google Sheets
                  │
                  ▼
          Google Sheets (Students, Config, Audit_Log)
```

---

*Built with Google Apps Script, Tailwind CSS, Chart.js, and the Gemini API.*
