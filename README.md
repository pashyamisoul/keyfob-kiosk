# Keyfob Kiosk

A touch-friendly kiosk web app for borrowing and returning office keyfobs,
with a Google Sheet as the database. Runs on any iPad or Android tablet:
no app store, no server, no API keys. Everything lives inside your
Google Workspace.

## How it works

```
Tablet (kiosk browser, full screen)
        |  opens the web app URL
        v
Google Apps Script web app  (Code.gs + Index.html)
        |  reads / writes
        v
Google Sheet
  +-- Fob Log   one row per borrow; the return updates the same row
  +-- Teams     team names shown in the kiosk dropdown
```

There is no fixed fob inventory. The receptionist hands over a fob and the
borrower types in the number printed on it. Which fobs are currently out is
derived from the *Fob Log*.

## Features

- **Check-out**: the borrower picks **Employee** or **External visitor**.
  Employees give their company email (domain enforced), team, and reporting
  manager. Externals give any valid email, their company, and the contact
  person they are visiting. A fob number that is already out is rejected
  with the current holder's name, so typos cannot double-book a fob.
- **Check-in**: pick the fob (shown with the borrower's name); name and
  email prefill from the borrower on record but stay editable. The original
  log row is marked `RETURNED` with a timestamp, keeping one clean row per
  borrow cycle.
- **Emails**: confirmation on check-out and check-in, plus an automatic
  reminder once a fob has been out 24 hours, repeated every further
  24 hours until it is returned.
- **GDPR-aware**: external visitors see a privacy notice with a required
  consent checkbox (recorded in the sheet), and their personal data is
  redacted automatically 30 days after the fob is returned.
- **Kiosk behavior**: large touch targets, inline validation, returns to
  the start screen after 90 s idle, race-safe via script locks.
- **Demo mode**: open `Index.html` directly in any browser and it runs on
  in-memory sample data, no deployment needed.

## Setup (one time, ~10 minutes)

1. Create a Google Sheet (e.g. "Keyfob Tracking").
2. Open **Extensions → Apps Script** from that sheet.
3. Replace the default `Code.gs` content with [Code.gs](Code.gs).
4. **File → New → HTML file**, name it exactly `Index`, paste in
   [Index.html](Index.html).
5. (Optional) In **Project Settings → Show manifest**, replace
   `appsscript.json` with the one here, and set your timezone.
6. Edit the `CFG` object at the top of `Code.gs`: set `EMAIL_DOMAIN` to
   your company domain. Also replace `example.com` in the two email
   validation spots in `Index.html`.
7. Run `setup` once from the editor toolbar and approve the permission
   prompt. This creates the *Fob Log* and *Teams* tabs (edit the team list
   directly in the sheet).
8. Run `installTriggers` once (approves a second prompt). This installs the
   hourly reminder job and the daily GDPR redaction job.
9. **Deploy → New deployment → Web app**: execute as **Me**, access
   **Anyone within your domain** (or **Anyone** if the kiosk tablet and
   phones are not signed into Workspace accounts).
10. Open the deployment URL on the tablet. On iPad: Add to Home Screen +
    Guided Access gives a locked, app-like kiosk.

## Configuration

All knobs are in the `CFG` object at the top of `Code.gs`:

| Setting | Default | Meaning |
| --- | --- | --- |
| `EMAIL_DOMAIN` | `example.com` | Domain enforced for employees |
| `SEND_CONFIRMATION_EMAIL` | `true` | Email the borrower on check-out and check-in |
| `REMINDER_EVERY_HOURS` | `24` | First reminder after 24 h out, then every further 24 h |
| `EXTERNAL_RETENTION_DAYS` | `30` | Days after return before external visitors' data is redacted |

## License

MIT
