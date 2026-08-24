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
  • Generates unique ID + QR code per registrant, emails ticket
  • Serves a JSON/JSONP API (doGet) for the scanner front end
  • Manages check-in, interview logging, stats, search
        │
        ▼
Scanner Web App (scanner.html, hosted free on GitHub Pages)
  • Camera-based QR scanning (check-in + interview recording)
  • Manual search/check-in fallback
  • Live dashboard counters
```

Two extra sheets in the same workbook:
- **Interview Log** — one row per interview outcome (a jobseeker can appear multiple times)
- **Vacancy List** — company/position reference data, powers autocomplete

## Features

- **Pre-registration → automatic e-ticket email** with embedded QR code (unique ID per registrant), with automatic retries on both QR generation and email delivery.
- **QR check-in scanning** — camera-based, works on any phone browser, with haptic feedback and on-screen confirmation.
- **Interview result recording** — scan a checked-in jobseeker to log an interview outcome (Qualified / Not Qualified / Hired On The Spot / Near Hires) against a specific Company and Position, with autocomplete sourced from the Vacancy List. Supports multiple interview records per jobseeker, with an explicit staff confirmation step before adding a repeat entry.
- **Manual fallback** — search any registrant by name (checked-in or not) and check them in or record an interview manually, for when a QR code won't scan.
- **Live dashboard** — running totals for Checked In and Hired On The Spot counts, visible across every staff device simultaneously.
- **Multi-device safe** — concurrent scanning from several staff phones is protected against double check-ins via server-side locking.

## Files

| File | Purpose |
|---|---|
| `Code.gs` | Apps Script backend — bound to the Form Responses sheet |
| `index.html` | Applicant registration form (multi-step, mobile-friendly) — GitHub Pages default entry point |
| `scanner.html` | Scanner front end — deployed on GitHub Pages |
| `api.js` | API abstraction for the registration form — holds the Apps Script endpoint |

## Applicant Registration Form

Replaces the Google Form pipeline with a self-hosted form. Submissions go straight into the **first sheet of the same workbook**, so every registrant is immediately visible to the existing scanner/check-in system.

```
Applicant Browser → index.html → api.js (submitApplicant)
                  → Apps Script doPost (Code.gs)
                  → Google Sheet + Google Drive (uploads) + Gmail ticket
```

### Setup

1. Update `Code.gs` in your Apps Script project, then redeploy the web app:
   **Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy**.
   Reusing the same deployment keeps the existing `/exec` URL, so `index.html`, `scanner.html`, and `ticket.html` keep working unchanged.
2. In `api.js`, set `APPS_SCRIPT_URL` to that web app `/exec` URL (it plays the role of an environment variable; GitHub Pages has none). No secrets go in this file.
3. Host `index.html` (registration) + `api.js` alongside `scanner.html` on GitHub Pages and share the page link with applicants.

### Behavior

- **System-generated columns** (`Timestamp`, `Unique ID`, `Status`, `Check-in Time`, `Email Status`, `Interview Result`) are written server-side at submission time; applicants never see or edit them.
- **Headers**: any missing canonical headers are appended automatically without moving existing columns. For a brand-new spreadsheet you can run `setupRegistrationSheet()` once from the editor to write the exact header order.
- **Reference numbers**: each submission gets a sequential ID like `JF26-00001` plus a UUID; both appear on the success screen.
- **Duplicate protection**: a submission with the *same email + first name + last name* as an existing row is rejected server-side; the applicant is shown their original reference number.
- **Uploads**: resume (PDF/DOC/DOCX) and PWD ID (JPG/PNG/PDF, ≤5 MB enforced client-side, 8 MB server-side) are stored in a Drive folder named **Job Fair Applicant Uploads**; the file link is written into the sheet column.
- **Ticket email**: sent immediately after registration when possible. The success screen only claims an email was sent if the backend confirms it.

## Security & Reliability Remediations

Applied following a static production-readiness audit. All protections are server-side (`Code.gs`); client checks remain as UX only.

- **Spreadsheet formula injection** — registration rows are written via `writeRegistrationRow_()`: every column except `Timestamp`/`Birthdate` is forced to plain-text (`'@'`) number format before `setValues`, so values beginning with `=`, `+`, `-`, `@`, or TAB are stored literally and can never execute as formulas. *(Note: the pre-existing Interview Log writer used by scanner action `hots` is out of scope here — see Remaining Risks.)*
- **Email HTML injection** — applicant-controlled names are HTML-escaped (`escapeHtml_`) inside `sendTicketEmail()` before interpolation into the message body.
- **Abuse damping** — layered, anonymous-friendly: (1) hidden honeypot field rejected server-side; (2) per-email cooldown of 3 accepted registrations per hour (`CacheService` counter, fails open); (3) request caps — ≤15 MB body, ≤40 fields, ≤1000 chars/field; (4) uploads validated server-side (extension whitelist, MIME whitelist, ≤8 MB decoded) and rolled back (trashed) if a later upload or the Sheet write fails; (5) error codes distinguish malformed / validation / duplicate / rate-limited / busy / upload / storage / internal failures, mapped to friendly messages in `api.js`.
- **Duplicate response** returns only the reference number — never the internal UUID or any storage identifier.
- **Ticket lookup** (`action:'ticket'`) is now implemented: email lookup, rate-limited to 5 lookups/min/email, returning only the name and ticket UUID (the same QR payload the confirmation email already delivers). Pending-email-status triggers the page's built-in retry.
- **Schema hazard** — the live workbook may still be attached to a legacy Google Form whose question texts differ from the canonical 36 headers. If both sources write to sheet[0], semantic duplicate columns can appear and the scanner may read legacy columns for new rows. Before production: either detach the Form deliberately, point the form responses to their own spreadsheet, or start registration on a fresh workbook initialized once with `setupRegistrationSheet()`. This cannot be verified from source code.

### Known remaining risks (deferred by scope)

The pre-existing public scanner API actions (`stats`, `list`, `all`, `hots_info`, `hots`) remain unauthenticated and unthrottled, and `hots` writes staff-supplied strings into the Interview Log without formula neutralization. Hardening these is a separate task because they belong to the legacy scanner architecture.

## Form Choices Management

Selectable options in the registration form are configuration-driven via a **"Form Choices"** sheet tab (auto-created and seeded with the original choices on first use). Schema: `Choice ID · Field Key · Choice Value · Display Label · Active · Sort Order · Created At · Updated At`.

- **Managed fields**: Interview Location, Gender, Civil Status, Education Attainment, Employment Preference, PESO Assistance Programs. Yes/No questions (First-time Jobseeker, Returning OFW?, Returning Worker, Skills Training, Disability) are **locked** — their values drive conditional logic. PESO "**None**" is a protected value (cannot be deactivated/renamed; exclusivity logic is value-based).
- **Deactivation over deletion**: inactive choices disappear from new forms; historical submissions keep their stored values untouched.
- **Public read**: `GET action=form_choices` returns active choices only (cached 5 min, invalidated on every mutation). `index.html` falls back to its built-in defaults if the endpoint fails — the form never becomes unusable.
- **Server-side authority**: registration validation checks submitted values against ACTIVE configured choices; fabricated/inactive values are rejected regardless of what the browser shows.

### Admin access provisioning

1. In the Apps Script editor run **`generateFormChoicesAdminKey()`** once. It stores a key in Script Properties and logs it once (execution logs are owner-only).
2. Share the key with authorized staff. They open `admin/form-choices.html`, paste it once per browser session, and manage choices.
3. To rotate: re-run `generateFormChoicesAdminKey()` (old key stops working immediately).

**Authorization model & limitations:** mutations (`add/update/toggle/reorder/list`) are server-authorized against the Script Properties key, sent only in POST bodies (never URLs), with failed attempts rate-limited (10 failures → 15-minute lockout) and audited to a "Choices Audit Log" tab. This is a **shared-secret** model: there are no individual staff identities under the current anonymous Apps Script architecture, so audit entries record actions but not who performed them. Do not treat this as multi-user access control.

## Setup

1. **Google Form** — collects registrant info; response destination is the main "Form Responses" sheet.
2. **Apps Script** (`Extensions → Apps Script` from the Sheet):
   - Paste in `Code.gs`.
   - Edit the header constants at the top of the file to match your Form's exact question text (see Configuration below).
   - Set a trigger: `onFormSubmitHandler`, event source *From spreadsheet*, event type *On form submit*.
   - Optional: a time-driven trigger for `retryPendingEmails` (daily) as a safety net for any emails that hit Gmail's sending quota.
   - `Deploy → New deployment → Web app` — **Execute as: Me**, **Who has access: Anyone** (required for the scanner page to reach it without a login wall).
3. **GitHub Pages**:
   - Create a public repo, upload the site files (`index.html` is the public landing page; the scanner lives at `scanner.html`).
   - Paste your Apps Script Web App `/exec` URL into the `APPS_SCRIPT_URL` constant near the top of the `<script>` block.
   - Enable Pages under repo Settings → Pages → Deploy from branch → `main` / root.
   - Bookmark the resulting URL on staff phones.

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
