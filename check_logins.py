#!/usr/bin/env python3
"""
GitHub (via Azure AD / Entra SSO) login verifier.

Reads users from an Excel file, and for each row drives a real browser through
the GitHub login -> Azure AD SSO flow, then writes back whether the user was
able to sign in successfully.

Excel columns (names configurable in config.json):
  github_username  - identifies the user / GitHub account
  azure_login      - the Azure AD (Entra) email/UPN used to sign in
  azure_password   - the Azure AD password
  github_password  - OPTIONAL. Used if the org shows GitHub's classic
                     username/password form instead of redirecting to Azure.
                     Falls back to azure_password if the column is absent/blank.

Written back per row:
  status      - SUCCESS | FAILED | MFA_REQUIRED | ERROR | SKIPPED
  detail      - human-readable explanation
  checked_at  - ISO timestamp

Usage:
  python check_logins.py                 # uses config.json
  python check_logins.py --config x.json
  python check_logins.py --headless
  python check_logins.py --only alice_org,bob_org   # only these github_usernames
"""

import argparse
import json
import sys
import time
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout


# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #
def load_config(path: str) -> dict:
    defaults = {
        "excel_file": "users.xlsx",
        "sheet_name": None,
        "start_url": "https://github.com/login",
        "columns": {
            "github_username": "github_username",
            "azure_login": "azure_login",
            "azure_password": "azure_password",
            "github_password": "github_password",
            "status": "status",
            "detail": "detail",
            "checked_at": "checked_at",
        },
        "headless": False,
        "slow_mo_ms": 0,
        "nav_timeout_ms": 30000,
        "step_timeout_ms": 20000,
        "between_users_delay_ms": 1500,
        "success_url_substring": "github.com",
        "clear_session_between_users": True,
    }
    cfg_path = Path(path)
    if cfg_path.exists():
        user_cfg = json.loads(cfg_path.read_text())
        # shallow merge, with a nested merge for "columns"
        cols = {**defaults["columns"], **(user_cfg.get("columns") or {})}
        defaults.update(user_cfg)
        defaults["columns"] = cols
    return defaults


# --------------------------------------------------------------------------- #
# Excel helpers
# --------------------------------------------------------------------------- #
class Sheet:
    """Thin wrapper around an openpyxl worksheet with header-name access."""

    def __init__(self, wb, ws, col_cfg):
        self.wb = wb
        self.ws = ws
        self.col_cfg = col_cfg
        self.header_row = 1
        self.headers = {}  # normalized header name -> column index (1-based)
        self._read_headers()
        self._ensure_output_columns()

    def _read_headers(self):
        for idx, cell in enumerate(self.ws[self.header_row], start=1):
            if cell.value is not None:
                self.headers[str(cell.value).strip()] = idx

    def _ensure_output_columns(self):
        """Make sure status/detail/checked_at columns exist; create if missing."""
        for key in ("status", "detail", "checked_at"):
            name = self.col_cfg[key]
            if name not in self.headers:
                new_idx = (max(self.headers.values()) if self.headers else 0) + 1
                self.ws.cell(row=self.header_row, column=new_idx, value=name)
                self.headers[name] = new_idx

    def col_index(self, key, required=True):
        name = self.col_cfg.get(key, key)
        idx = self.headers.get(name)
        if idx is None and required:
            raise KeyError(
                f"Required column '{name}' not found. Found columns: "
                f"{sorted(self.headers)}"
            )
        return idx

    def iter_rows(self):
        for row in range(self.header_row + 1, self.ws.max_row + 1):
            yield row

    def get(self, row, key, required=True):
        idx = self.col_index(key, required=required)
        if idx is None:
            return None
        val = self.ws.cell(row=row, column=idx).value
        return None if val is None else str(val).strip()

    def set(self, row, key, value):
        idx = self.col_index(key)
        self.ws.cell(row=row, column=idx, value=value)

    def save(self, path):
        self.wb.save(path)


# --------------------------------------------------------------------------- #
# Login flow (state machine)
# --------------------------------------------------------------------------- #
def _visible(page, selector, timeout=1200):
    """Return True if selector is present and visible within a short timeout."""
    try:
        loc = page.locator(selector).first
        loc.wait_for(state="visible", timeout=timeout)
        return True
    except PWTimeout:
        return False
    except Exception:
        return False


def _safe_fill(page, selector, value, timeout=6000):
    """Fill a field, tolerating brief 'not enabled' states. Returns True on success."""
    try:
        loc = page.locator(selector).first
        loc.wait_for(state="visible", timeout=timeout)
        loc.fill(value, timeout=timeout)
        return True
    except Exception:
        return False


def _text_present(page, needles, timeout=800):
    try:
        body = page.locator("body").inner_text(timeout=timeout).lower()
    except Exception:
        return False
    return any(n.lower() in body for n in needles)


def attempt_login(page, cfg, github_username, azure_login, azure_password,
                  github_password):
    """
    Drive the browser through the login chain. Returns (status, detail).
    status in {SUCCESS, FAILED, MFA_REQUIRED, ERROR}.
    """
    step_to = cfg["step_timeout_ms"]
    success_sub = cfg["success_url_substring"]

    page.set_default_timeout(step_to)
    try:
        page.goto(cfg["start_url"], wait_until="domcontentloaded",
                  timeout=cfg["nav_timeout_ms"])
    except PWTimeout:
        return "ERROR", f"Timed out loading start_url {cfg['start_url']}"

    filled_ms_user = False
    filled_ms_pass = False
    filled_gh = False

    # The login can bounce GitHub <-> Microsoft several times; loop until we
    # reach a terminal state or run out of iterations.
    for _ in range(24):
        page.wait_for_timeout(600)  # let redirects settle
        url = page.url.lower()

        # ---- Terminal: signed in to GitHub ----------------------------------
        if success_sub in url and _logged_into_github(page, github_username):
            return "SUCCESS", f"Signed in to GitHub as expected (url={page.url})"

        # ---- Microsoft: MFA / additional verification -----------------------
        if _is_mfa_challenge(page):
            return ("MFA_REQUIRED",
                    "Azure AD requested MFA / additional verification "
                    "(cannot be completed unattended).")

        # ---- Microsoft: explicit credential error ---------------------------
        ms_err = _microsoft_error(page)
        if ms_err:
            return "FAILED", f"Azure AD rejected sign-in: {ms_err}"

        # ---- Microsoft: email/username step ---------------------------------
        if _visible(page, 'input[name="loginfmt"]'):
            if filled_ms_user:
                # Field reappeared -> usually an unrecognized account.
                return "FAILED", "Azure AD did not accept the account (loginfmt)."
            if not _safe_fill(page, 'input[name="loginfmt"]', azure_login):
                return "ERROR", "Azure AD username field was not editable (possible bot challenge)."
            _click_first(page, ['#idSIButton9', 'input[type="submit"]',
                                'button[type="submit"]'])
            filled_ms_user = True
            continue

        # ---- Microsoft: password step ---------------------------------------
        if _visible(page, 'input[name="passwd"]'):
            if not _safe_fill(page, 'input[name="passwd"]', azure_password):
                return "ERROR", "Azure AD password field was not editable (possible bot challenge)."
            _click_first(page, ['#idSIButton9', 'input[type="submit"]',
                                'button[type="submit"]'])
            filled_ms_pass = True
            continue

        # ---- Microsoft: "Stay signed in?" -----------------------------------
        if _text_present(page, ["stay signed in"]) or _visible(page, "#idBtn_Back"):
            # Click "No" to keep sessions clean.
            if not _click_first(page, ["#idBtn_Back"]):
                _click_first(page, ["#idSIButton9", 'input[type="submit"]'])
            continue

        # ---- GitHub: classic username/password form -------------------------
        if _visible(page, "#login_field") and _visible(page, "#password"):
            gh_pw = github_password or azure_password
            _safe_fill(page, "#login_field", github_username)
            if not _safe_fill(page, "#password", gh_pw):
                return "ERROR", ("GitHub password field was not editable "
                                 "(GitHub often blocks headless browsers — "
                                 "try headless=false).")
            _click_first(page, ['input[name="commit"]', 'button[type="submit"]',
                                'input[type="submit"]'])
            filled_gh = True
            continue

        # ---- GitHub: 2FA -----------------------------------------------------
        if _github_2fa(page):
            return ("MFA_REQUIRED",
                    "GitHub requested a 2FA code (cannot be completed unattended).")

        # ---- GitHub: sign-in error ------------------------------------------
        gh_err = _github_error(page)
        if gh_err:
            return "FAILED", f"GitHub rejected sign-in: {gh_err}"

        # ---- GitHub: "Continue"/SSO button ----------------------------------
        if _click_first(page, ['button:has-text("Continue")',
                               'a:has-text("Continue")',
                               'button:has-text("Sign in with")',
                               'a:has-text("Single sign-on")']):
            continue

        # Nothing actionable recognized this round; loop again to allow redirects.

    # Fell out of the loop without a terminal state.
    if success_sub in page.url.lower() and _logged_into_github(page, github_username):
        return "SUCCESS", f"Signed in to GitHub (url={page.url})"
    return ("ERROR",
            f"Could not determine outcome. Final url={page.url}. "
            f"(ms_user={filled_ms_user}, ms_pass={filled_ms_pass}, gh={filled_gh})")


def _click_first(page, selectors):
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            if loc.count() > 0 and loc.is_visible():
                loc.click(timeout=4000)
                return True
        except Exception:
            continue
    return False


def _logged_into_github(page, github_username):
    # Most reliable signal: <meta name="user-login" content="...">
    try:
        content = page.locator('meta[name="user-login"]').first.get_attribute(
            "content", timeout=1500)
        if content and content.strip():
            return True
    except Exception:
        pass
    # Fallback: the top-right user menu present on authenticated pages.
    for sel in ['[aria-label="View profile and more"]',
                'button[aria-label*="View profile"]',
                'summary.Header-link img.avatar']:
        if _visible(page, sel, timeout=800):
            return True
    return False


def _is_mfa_challenge(page):
    if _visible(page, "#idTxtBx_SAOTCC_OTC", timeout=600):  # enter code
        return True
    if _visible(page, "#idDiv_SAOTCAS_Description", timeout=600):  # approve request
        return True
    needles = [
        "approve sign in request", "approve the request", "enter code",
        "we texted your phone", "verify your identity", "more information required",
        "open your authenticator app", "check your microsoft authenticator",
        "additional security verification", "verify it's you",
    ]
    return _text_present(page, needles)


def _microsoft_error(page):
    for sel in ["#usernameError", "#passwordError", "#error", "#errorText"]:
        try:
            loc = page.locator(sel).first
            if loc.count() > 0 and loc.is_visible():
                txt = (loc.inner_text(timeout=800) or "").strip()
                if txt:
                    return txt
        except Exception:
            continue
    try:
        alert = page.locator('[role="alert"]').first
        if alert.count() > 0 and alert.is_visible():
            txt = (alert.inner_text(timeout=800) or "").strip()
            # Ignore benign informational alerts.
            low = txt.lower()
            if txt and any(k in low for k in
                           ["incorrect", "isn't", "does not", "doesn't",
                            "can't find", "cannot find", "wrong", "invalid",
                            "account or password", "didn't work", "locked",
                            "blocked", "disabled"]):
                return txt
    except Exception:
        pass
    return None


def _github_2fa(page):
    for sel in ['#app_totp', 'input[name="otp"]', '#otp']:
        if _visible(page, sel, timeout=600):
            return True
    return _text_present(page, ["two-factor authentication",
                                "verification code", "authenticator app"])


def _github_error(page):
    for sel in [".flash-error", "#js-flash-container .flash-error"]:
        try:
            loc = page.locator(sel).first
            if loc.count() > 0 and loc.is_visible():
                txt = (loc.inner_text(timeout=800) or "").strip()
                if txt:
                    return txt
        except Exception:
            continue
    if _text_present(page, ["incorrect username or password"]):
        return "Incorrect username or password."
    return None


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main():
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--headless", action="store_true",
                        help="override config to run without a visible browser")
    parser.add_argument("--only", default=None,
                        help="comma-separated github_usernames to test (others skipped)")
    parser.add_argument("--recheck", action="store_true",
                        help="re-test rows that already have a status")
    args = parser.parse_args()

    cfg = load_config(args.config)
    if args.headless:
        cfg["headless"] = True

    only = None
    if args.only:
        only = {u.strip() for u in args.only.split(",") if u.strip()}

    excel_path = Path(cfg["excel_file"])
    if not excel_path.exists():
        print(f"ERROR: Excel file not found: {excel_path.resolve()}", file=sys.stderr)
        print("Run `python make_template.py` to create a starter file.", file=sys.stderr)
        sys.exit(2)

    wb = load_workbook(excel_path)
    ws = wb[cfg["sheet_name"]] if cfg["sheet_name"] else wb.active
    sheet = Sheet(wb, ws, cfg["columns"])

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=cfg["headless"],
                                    slow_mo=cfg["slow_mo_ms"])
        context = None
        if not cfg["clear_session_between_users"]:
            context = browser.new_context()

        total = ok = failed = other = 0
        for row in sheet.iter_rows():
            gh_user = sheet.get(row, "github_username", required=False)
            az_login = sheet.get(row, "azure_login", required=False)
            az_pass = sheet.get(row, "azure_password", required=False)
            gh_pass = sheet.get(row, "github_password", required=False)

            if not gh_user and not az_login:
                continue  # blank row

            if only is not None and gh_user not in only:
                continue

            existing_status = sheet.get(row, "status", required=False)
            if existing_status and not args.recheck:
                print(f"[skip] {gh_user or az_login}: already '{existing_status}' "
                      f"(use --recheck to redo)")
                continue

            if not az_login or not az_pass:
                sheet.set(row, "status", "SKIPPED")
                sheet.set(row, "detail", "Missing azure_login or azure_password")
                sheet.set(row, "checked_at", datetime.now().isoformat(timespec="seconds"))
                sheet.save(excel_path)
                print(f"[skip] row {row}: missing credentials")
                continue

            total += 1
            label = gh_user or az_login
            print(f"[test] {label} ...", flush=True)

            per_user_ctx = context or browser.new_context()
            page = per_user_ctx.new_page()
            page.set_default_navigation_timeout(cfg["nav_timeout_ms"])
            try:
                status, detail = attempt_login(
                    page, cfg, gh_user or "", az_login, az_pass, gh_pass)
            except Exception as e:  # noqa: BLE001
                status, detail = "ERROR", f"Unhandled exception: {e!r}"
            finally:
                page.close()
                if context is None:
                    per_user_ctx.close()

            sheet.set(row, "status", status)
            sheet.set(row, "detail", detail)
            sheet.set(row, "checked_at", datetime.now().isoformat(timespec="seconds"))
            sheet.save(excel_path)  # save after each user so progress is never lost

            if status == "SUCCESS":
                ok += 1
            elif status == "FAILED":
                failed += 1
            else:
                other += 1
            print(f"       -> {status}: {detail}")

            time.sleep(cfg["between_users_delay_ms"] / 1000.0)

        if context is not None:
            context.close()
        browser.close()

    print("\n==== Summary ====")
    print(f"Tested : {total}")
    print(f"SUCCESS: {ok}")
    print(f"FAILED : {failed}")
    print(f"Other  : {other} (MFA_REQUIRED / ERROR / SKIPPED)")
    print(f"Results written to: {excel_path.resolve()}")


if __name__ == "__main__":
    main()
