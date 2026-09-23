# GitHub (Azure AD SSO) login checker

Playwright automation that reads users from an Excel file and, for each one,
drives a real browser through the **GitHub → Azure AD (Entra) SSO** login flow
using a **userPrincipalName + Temporary Access Pass (TAP)**, then writes back
whether that user could sign in.

Intended for an admin verifying access for accounts in **their own**
organization.

## What it does

For every row in `State of Maryland.xlsx` (sheet `Users`) it:

1. Opens `https://github.com/login` (configurable) in a real Chromium browser.
2. Types the `GitHub_Username` and clicks **"Sign in with your identity
   provider"** to hand off to Azure AD.
3. On Azure AD, enters the `userPrincipalName` and the `tap` (Temporary Access
   Pass) in place of a password. (If a row has no TAP it falls back to
   `password`; if the org shows GitHub's classic form instead, it uses
   `GitHub_Username` + `github_password`.)
4. Determines the outcome and writes it back to the same row:

| status         | meaning                                                        |
|----------------|----------------------------------------------------------------|
| `SUCCESS`      | Ended up signed in to GitHub as the user.                      |
| `FAILED`       | Credentials were rejected by Azure AD or GitHub.               |
| `MFA_REQUIRED` | Sign-in needs an MFA / 2FA step that can't run unattended.     |
| `ERROR`        | Couldn't determine the outcome (timeout, bot challenge, etc.). |
| `SKIPPED`      | Row was missing `userPrincipalName` or TAP/`password`.         |

It also fills `detail` (explanation) and `checked_at` (timestamp), and **saves
after every user**, so progress is never lost if a run is interrupted.

## Setup (one time)

```bash
bash setup.sh          # creates .venv, installs deps, downloads Chromium
```

## Prepare your data

```bash
./.venv/bin/python make_template.py   # optional: creates a blank workbook with the right columns
```

Use the shared `State of Maryland.xlsx` (sheet `Users`), or run
`make_template.py` for a blank one. Columns:

- `userPrincipalName` – the Azure AD UPN used to sign in
- `password` – the Azure AD password (*optional*; only used when a row has no TAP)
- `tap` – the Temporary Access Pass, entered in place of the password
- `GitHub_Username` – the GitHub (EMU) username typed on the GitHub page
- `github_password` – *optional*; only used if the org shows GitHub's classic
  login form. Leave blank to fall back to `password`.

> TAPs are single-use / time-limited. Generate fresh passes before a run, and
> expect `FAILED` on rows whose TAP has already expired or been consumed.

(`status`, `detail`, `checked_at` are written by the tool — leave them blank.)

## Run

```bash
./.venv/bin/python check_logins.py                 # test everyone not yet checked
./.venv/bin/python check_logins.py --recheck       # re-test everyone
./.venv/bin/python check_logins.py --only bob_org  # test specific users
./.venv/bin/python check_logins.py --headless      # no visible window
```

By default rows that already have a `status` are skipped; use `--recheck` to redo them.

## Configuration — `config.json`

| key                           | purpose                                                       |
|-------------------------------|---------------------------------------------------------------|
| `excel_file`                  | path to the workbook                                           |
| `sheet_name`                  | worksheet name (`null` = first sheet)                         |
| `start_url`                   | where login begins (e.g. an org SSO URL)                     |
| `columns`                     | rename any expected column to match your sheet               |
| `headless`                    | `false` (recommended) shows the browser                      |
| `clear_session_between_users` | `true` uses a fresh, cookie-free session per user            |
| `*_timeout_ms`                | navigation / per-step timeouts                               |

## Important notes

- **Run with a visible browser (`headless: false`, the default).** GitHub and
  Azure AD frequently block headless browsers, which shows up as `ERROR`.
- **MFA:** accounts that require an interactive MFA prompt can't be completed
  unattended — they're reported as `MFA_REQUIRED`. Running headed lets you
  complete the prompt by hand within the step timeout if you want a real result.
- **The exact login page varies by org** (EMU vs. SAML SSO, custom branding).
  The flow auto-detects the common Microsoft and GitHub screens; if your org
  uses different fields, adjust the selectors in `attempt_login()` and the
  `start_url` in `config.json`.
- **Credentials are sensitive.** The workbook holds plaintext passwords/TAPs and is
  git-ignored. Keep it protected and delete it when you're done. Only run this
  against accounts you're authorized to test.
- Repeated automated sign-ins can trigger account lockouts or security alerts;
  the `between_users_delay_ms` setting spaces attempts out.

## Files

- `check_logins.py` – the automation
- `make_template.py` – creates a starter workbook
- `config.json` – settings
- `setup.sh` – installs everything
- `requirements.txt` – Python dependencies
```
