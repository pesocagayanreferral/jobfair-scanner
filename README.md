# Job Fair QR Check-In & Interview Tracking System

A zero-cost check-in, attendance, and interview-outcome tracking system built for PESO job fair events. Handles pre-registration via Google Forms, QR-code e-tickets, live multi-device scanning, and DOLE-format interview result logging.

## Architecture

```
Google Form (pre-registration)
        │
        ▼
Google Sheet ("Form Responses")  ◄──── source of truth for all data
        │
        ▼
Google Apps Script (Code.gs)
  • Generates unique ID per registrant; serves the ticket/QR lookup API
  • Serves a JSON/JSONP API (doGet) for the scanner front end
  • Manages check-in, interview logging, stats, search
        │
        ▼
Scanner Web App (index.html, hosted free on GitHub Pages)
  • Camera-based QR scanning (check-in + interview recording)
  • Manual search/check-in fallback
  • Live dashboard counters
```

Two extra sheets in the same workbook:
- **Interview Log** — one row per interview outcome (a jobseeker can appear multiple times)
- **Vacancy List** — company/position reference data, powers autocomplete

## Features

- **Pre-registration → self-serve QR pass** — after submitting the form, registrants open a landing page, enter their email, and screenshot their QR code. No email is sent, so no Gmail spam-flagging or quota risk.
- **QR check-in scanning** — camera-based, works on any phone browser, with haptic feedback and on-screen confirmation.
- **Interview result recording** — scan a checked-in jobseeker to log an interview outcome (Qualified / Not Qualified / Hired On The Spot / Near Hires) against a specific Company and Position, with autocomplete sourced from the Vacancy List. Supports multiple interview records per jobseeker, with an explicit staff confirmation step before adding a repeat entry.
- **Manual fallback** — search any registrant by name (checked-in or not) and check them in or record an interview manually, for when a QR code won't scan.
- **Live dashboard** — running totals for Checked In and Hired On The Spot counts, visible across every staff device simultaneously.
- **Multi-device safe** — concurrent scanning from several staff phones is protected against double check-ins via server-side locking.

## Files

| File | Purpose |
|---|---|
| `Code.gs` | Apps Script backend — bound to the Form Responses sheet |
| `index.html` | Scanner front end — deployed on GitHub Pages |
| `ticket.html` | Registrant QR pass page — deployed on GitHub Pages, linked from the Form confirmation |

## Setup

1. **Google Form** — collects registrant info; response destination is the main "Form Responses" sheet.
2. **Apps Script** (`Extensions → Apps Script` from the Sheet):
   - Paste in `Code.gs`.
   - Edit the header constants at the top of the file to match your Form's exact question text (see Configuration below).
   - Set a trigger: `onFormSubmitHandler`, event source *From spreadsheet*, event type *On form submit*.
   - `Deploy → New deployment → Web app` — **Execute as: Me**, **Who has access: Anyone** (required for the scanner page to reach it without a login wall).
3. **GitHub Pages**:
   - Create a public repo, upload `index.html` and `ticket.html`.
   - Paste your Apps Script Web App `/exec` URL into the `APPS_SCRIPT_URL` constant near the top of the `<script>` block in both files.
   - Enable Pages under repo Settings → Pages → Deploy from branch → `main` / root.
   - Bookmark the resulting URL on staff phones.
4. **Google Form confirmation** — in Settings → Presentation, tell registrants to open your `ticket.html` URL, enter the email they registered with, and screenshot their QR pass. No e-ticket email is sent anymore.

## Configuration

Edit these constants in `Code.gs` to match your Form's actual column headers (Google Sheets copies Form question text verbatim, including punctuation):

```javascript
EMAIL_COLUMN_HEADER, LAST_NAME_HEADER, FIRST_NAME_HEADER, MIDDLE_NAME_HEADER,
GENDER_HEADER, BIRTHDATE_HEADER, BARANGAY_HEADER, MUNICIPALITY_HEADER, PROVINCE_HEADER,
VACANCY_COMPANY_HEADER, VACANCY_POSITION_HEADER
```

A **"Vacancy List"** sheet tab is expected in the same workbook, with columns matching `VACANCY_COMPANY_HEADER` / `VACANCY_POSITION_HEADER`.

## Known limitations

- **Google Forms can be auto-flagged.** Google's automated moderation has, on at least one occasion, incorrectly flagged the registration form as inappropriate and blocked sharing. If this happens: duplicate the form (*File → Make a copy*) and, in the new copy's Responses tab, link it to the **existing** response Sheet rather than creating a new one — this keeps the entire backend working unchanged, since everything is bound to the Sheet, not the Form.
- **GitHub Pages requires a public repo** on the free tier — making the repo private would break hosting unless upgraded to a paid GitHub plan. Repo visibility and edit access are separate settings: the repo can stay public while only the owner retains push access (no collaborators added).
- **Apps Script quotas**: 30 simultaneous executions per user, single-threaded execution model. Comfortably sufficient for a handful of concurrent scanner devices; not designed for large-scale simultaneous traffic.
- **CacheService** (used to cache dashboard reads) is not guaranteed-persistent storage — it's used only for read-side performance, never as the authoritative record of a check-in or interview.

## Possible future paths (not yet implemented)

If the event scale outgrows Google Sheets, options discussed include a custom HTML registration form (removing Google Forms from the pipeline entirely) or a managed free-tier database (e.g. Supabase). Any such migration should be planned and tested well ahead of an event date, not attempted as an emergency fix — the current Form/Sheet/Apps Script pipeline should remain the primary system unless a full migration has been deliberately built and tested in advance.
