# Apps Script (the server side of LPP Log)

The app's data lives in a Google Sheet, and the code that reads and writes
that sheet lives in Google Apps Script — **not** in this repo automatically.
Google does not version it the way GitHub does, so copies are kept here by hand.

| File | What it is |
|---|---|
| `lpp-log-backend-v10.gs` | The backend as of 2026-10-03 (v9 + getLog returns Job Start / Job End) |
| `nightly-backup.gs` | Makes a dated copy of the sheet every night |

## Whenever the Apps Script changes

Paste the new version over the `-v9` file (rename to the new version number)
and commit. That way there is always a copy that isn't inside Google.

## Services tab layout

The service list is JSON in column B of the `DATA` rows. Since 2026-10-11
(LPP Log script v22, FieldPro script v15) it continues in B3, B4 … once it
passes 40,000 characters, because Google Sheets refuses cells over 50,000.
Pieces after the first start with `~`; to read by hand, drop each `~` and join
in order. While it fits in one cell, B2 holds plain JSON exactly as before.
The FieldPro backend reads and writes this tab the same way.

## Sheet

`1ZLuqBz61IXa5CEQieNQbhTLxH6EJYGmxJiXjeyaswa0`
