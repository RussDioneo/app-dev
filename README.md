# app-dev
My First Repository

Writes rows from the Apploi userscript straight into your SharePoint `.xlsx`, via Microsoft Graph.

**Architecture**

```
Apploi page  ──►  userscript panel  ──►  http://127.0.0.1:8765  ──►  Graph API  ──►  SharePoint .xlsx
                 (parses, you review)     (Python, holds the token)
```

Only the local service holds credentials. The browser never sees a token, and nothing is exposed off your machine (bound to `127.0.0.1`).

The **Copy** button keeps working with no service running — this is purely additive.

---

## Step 1 — Make the sheet a Table

`rows/add` only works on a real Excel Table, not a plain range.

1. Open the workbook in **Excel Online** (or desktop).
2. Select the range **including the header row**.
3. **Insert → Table** → tick "My table has headers".
4. **Table Design → Table Name** → set it to `Candidates` (or note whatever you pick — it goes in the config).

### Your layout (14 columns)

The writer **locates columns by header name**, not by position, so extra columns are fine.

| Col | Header | Written by the tool? |
|---|---|---|
| A | Shift Date | ❌ left alone |
| B | FACILITY | ✅ |
| C | Date Applied | ✅ |
| D | Applicant Name | ✅ |
| E | TermList Result | ✅ (formula) |
| F | TermList Last 4 Digits | ✅ (formula) |
| G | Position | ✅ |
| H | Date of Interview | ✅ |
| I | Phone number | ✅ |
| J | Action Taken (Call) | ❌ left alone |
| K | Text Sent (Y/N) | ❌ left alone |
| L | Email sent (Y/N) | ❌ left alone |
| M | Status | ❌ left alone |
| N | Remarks | ❌ left alone |

Two consequences worth knowing:

- **On a re-run, only B–I are touched.** Your manually-maintained cells (Shift Date, Action Taken, Text Sent, Email Sent, Status, Remarks) are never overwritten.
- **The VLOOKUP still works.** `INDIRECT("D"&ROW())` points at Applicant Name, which is column D in this layout — so the formulas are already correct here.

Header matching is case-insensitive and tolerant of the messy multi-line headers, e.g. `Date of Interview '03/05/25, 10:30 AM' [copy from the interview dialogue box]` still resolves to column H. If a header is renamed beyond recognition, override it in `config.json`:

```json
"columns": {
  "facility": "FACILITY",
  "dateApplied": "Date Applied",
  "name": "Applicant Name",
  "termListResult": "TermList Result",
  "termListLast4": "TermList Last 4 Digits",
  "position": "Position",
  "interview": "Date of Interview",
  "phone": "Phone number"
}
```

> ⚠️ **About the Copy button:** it still produces the 8 original fields (`FACILITY` … `Phone number`). With `Shift Date` in column A, paste it starting at **column B**, not A. The *Send to SharePoint* path handles all 14 columns automatically.

---

## Step 2 — Register the app in Microsoft Entra ID

You only do this once. (Needs permission to register apps in your tenant; if blocked, ask your M365 admin.)

1. Go to **portal.azure.com** → **Microsoft Entra ID** → **App registrations** → **New registration**
2. Name: `CCH Sheet Writer`
3. Supported account types: **Accounts in this organizational directory only (single tenant)**
4. Redirect URI: leave **blank** (device-code flow doesn't need one)
5. **Register**
6. On the app page, copy and keep:
   - **Application (client) ID** → `clientId`
   - **Directory (tenant) ID** → `tenantId`
7. **Authentication** → scroll to *Advanced settings* → **Allow public client flows** = **Yes** → **Save**
   *(Required for device-code sign-in. Without it you'll get `AADSTS7000218`.)*
8. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** → search **`Files.ReadWrite.All`** → **Add permissions**
9. Click **Grant admin consent** if you can. If not, the first sign-in will prompt you personally (unless your tenant disables user consent — then you need an admin).

No client secret is needed. This is a **public client**; the token cache on disk is the credential.

---

## Step 3 — Configure

First run creates the template:

```powershell
cd "C:\Users\okik\Documents\Obsidian Vault\CCH"
python sharepoint_writer.py
```

Edit `config.json`:

```json
{
  "tenantId": "00000000-0000-0000-0000-000000000000",
  "clientId": "11111111-1111-1111-1111-111111111111",
  "workbookUrl": "https://yourtenant.sharepoint.com/sites/Recruiting/Shared Documents/Candidate Tracker.xlsx",
  "tableName": "Candidates",
  "port": 8765
}
```

`workbookUrl`: open the file in SharePoint, copy the address bar URL, and delete everything from `?` onward.

---

## Step 4 — Verify (read-only, writes nothing)

```powershell
python sharepoint_writer.py --check
```

It prints a code and a URL — open the URL, enter the code, sign in. Then you should see:

```
[graph] site: Recruiting (contoso.sharepoint.com,xxx,yyy)
[graph] drive: Documents
[graph] file:  Candidate Tracker.xlsx (48,221 bytes)
[graph] tables: Candidates
[graph] table range: Sheet1!A1:H420

--- CHECK OK -------------------------------------------------
  read + discover work. Nothing was written.
```

If that succeeds, the hard part is done.

---

## Step 5 — Run the service

```powershell
python sharepoint_writer.py
```

Leave the window open while you work. You'll see `[ready] listening on http://127.0.0.1:8765`.

---

## Step 6 — Update the userscript

Re-paste [Apploi Copy for Excel.user.js](Apploi%20Copy%20for%20Excel.user.js) into Tampermonkey (the header now has `@grant GM_xmlhttpRequest` and `@connect 127.0.0.1`).

The panel gains a **Send to SharePoint** button:

- `Row added to sheet` — new row appended
- `Updated row 37` — same facility + name already existed, so it was overwritten in place (this is what makes reschedules safe)
- `SharePoint: writer not running…` — step 5 isn't running

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| `AADSTS7000218` | "Allow public client flows" isn't set to Yes (Step 2.7) |
| `AADSTS65001` / insufficient privileges | Consent missing — add `Files.ReadWrite.All` and grant consent |
| `No Excel tables found` | The range isn't a Table (Step 1) |
| `Table 'Candidates' not found` | Name mismatch — use Table Design → Table Name |
| `writer not running on 127.0.0.1:8765` | Start Step 5 |
| Row lands but formulas are text | The sheet is in a weird state; re-open in Excel Online — Graph writes `=` strings as live formulas |
| Sign-in loop after a password change | Delete `token_cache.json`, run again |
| Conflicts / lost edits | Someone has the file open in **desktop** Excel with unsaved changes. Co-author in Excel Online, or ask them to close it |

---

## Security notes

- `token_cache.json` holds a **refresh token** — treat it like a password. Don't sync it, don't share the folder.
- The service binds to `127.0.0.1` only; nothing on your network can reach it.
- `Files.ReadWrite.All` is broad (all files *you* can access). If your admin objects, a narrower `Sites.Selected` grant scoped to one site is possible but requires admin setup.
- Removing access later: delete the app registration, or revoke consent at **myapps.microsoft.com**.

---

## What it does on a re-run

Before appending, the service reads the table and looks for a row where the **FACILITY** column and the **Applicant Name** column both match (located by header, so whichever letter they sit in). If found, it overwrites **only the fields it owns** (FACILITY, Date Applied, Applicant Name, both TermList formulas, Position, Date of Interview, Phone number) and leaves Action Taken, Text Sent, Email sent, Status and Remarks exactly as you left them — so a rescheduled interview updates the line without wiping your notes.

That match is intentionally loose (ignores date/position), which corrects the common "I already entered this person" case. If you ever legitimately need two rows for the same person at the same facility, append manually.
