# Apps Script (the server side of LPP Log)

The app's data lives in a Google Sheet, and the code that reads and writes
that sheet lives in Google Apps Script — **not** in this repo automatically.
Google does not version it the way GitHub does, so copies are kept here by hand.

| File | What it is |
|---|---|
| `lpp-log-backend-v9.gs` | The live backend as of 2026-08-27 |
| `nightly-backup.gs` | Makes a dated copy of the sheet every night |

## Whenever the Apps Script changes

Paste the new version over the `-v9` file (rename to the new version number)
and commit. That way there is always a copy that isn't inside Google.

## Sheet

`1ZLuqBz61IXa5CEQieNQbhTLxH6EJYGmxJiXjeyaswa0`
