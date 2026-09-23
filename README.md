# GitHub (Azure AD SSO) login checker

Playwright automation that reads users from an Excel file and, for each one,
drives a real browser through the **GitHub → Azure AD (Entra) SSO** login flow,
then writes back whether that user could sign in.

Intended for an admin verifying access for accounts in **their own**
organization.

## What it does

For every row in `users.xlsx` it:

1. Opens `https://github.com/login` (configurable) in a real Chromium browser.
2. Follows the redirect chain to Azure AD, entering the `azure_login` and
   `azure_password`. (If the org shows GitHub's classic username/password form
   instead, it uses `github_username` + `github_password`, falling back to
   `azure_password`.)
3. Determines the outcome and writes it back to the same row:

| status         | meaning                                                        |
|----------------|----------------------------------------------------------------|
| `SUCCESS`      | Ended up signed in to GitHub as the user.                      |
| `FAILED`       | Credentials were rejected by Azure AD or GitHub.               |
| `MFA_REQUIRED` | Sign-in needs an MFA / 2FA step that can't run unattended.     |
| `ERROR`        | Couldn't determine the outcome (timeout, bot challenge, etc.). |
| `SKIPPED`      | Row was missing `azure_login` / `azure_password`.              |

It also fills `detail` (explanation) and `checked_at` (timestamp), and **saves
after every user**, so progress is never lost if a run is interrupted.

## Setup (one time)

```bash
bash setup.sh          # creates .venv, installs deps, downloads Chromium
```

## Prepare your data

```bash
./.venv/bin/python make_template.py   # creates users.xlsx with the right columns
```

Open `users.xlsx` and fill in real rows. Columns:

- `github_username` – identifies the user / GitHub account
- `azure_login` – the Azure AD email/UPN used to sign in
- `azure_password` – the Azure AD password
- `github_password` – *optional*; only used if the org shows GitHub's classic
  login form. Leave blank to fall back to `azure_password`.

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
- **Credentials are sensitive.** `users.xlsx` holds plaintext passwords and is
  git-ignored. Keep it protected and delete it when you're done. Only run this
  against accounts you're authorized to test.
- Repeated automated sign-ins can trigger account lockouts or security alerts;
  the `between_users_delay_ms` setting spaces attempts out.

## Files

- `check_logins.py` – the automation
- `make_template.py` – creates a starter `users.xlsx`
- `config.json` – settings
- `setup.sh` – installs everything
- `requirements.txt` – Python dependencies
```
