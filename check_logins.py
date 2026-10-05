#!/usr/bin/env python3
"""
GitHub (via Azure AD / Entra SSO) login verifier.

Reads users from an Excel file, and for each row drives a real browser through
the GitHub login -> Azure AD SSO flow, then writes back whether the user was
able to sign in successfully.

Login flow (matches the "State of Maryland" workbook):
  1. Open https://github.com/login and type the GitHub username.
  2. Click "Sign in with your identity provider" to hand off to Azure AD.
  3. On Azure AD, enter the userPrincipalName (UPN) and the Temporary Access
     Pass (TAP) instead of a password.
  4. Write back whether the user reached an authenticated GitHub session.

Excel columns (names configurable in config.json):
  github_username  - the GitHub (EMU) username entered on the GitHub page
  azure_login      - the Azure AD (Entra) userPrincipalName used to sign in
  azure_tap        - the Temporary Access Pass (used in place of the password)
  azure_password   - OPTIONAL fallback secret, used only when a row has no TAP
  github_password  - OPTIONAL. Used only if the org shows GitHub's classic
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
        "excel_file": "State of Maryland.xlsx",
        "sheet_name": "Users",
        "start_url": "https://github.com/login",
        "columns": {
            "github_username": "GitHub_Username",
            "azure_login": "userPrincipalName",
            "azure_tap": "tap",
            "azure_password": "password",
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

        # Which checks to run per user.
        "check_github": True,
        "check_windows365": False,
        "check_copilot": False,
        "check_powerautomate": False,
        "check_m365": False,
        "check_cowork": False,

        # Windows 365 portal check settings.
        "windows365_url": "https://windows365.microsoft.com/",
        "windows365_portal_host": "windows.cloud.microsoft",

        # Copilot Studio portal check settings.
        "copilot_url": "https://copilotstudio.microsoft.com/",
        "copilot_portal_host": "copilotstudio.microsoft.com",

        # Power Automate portal check settings.
        "powerautomate_url": "https://make.powerautomate.com/",
        "powerautomate_portal_host": "make.powerautomate.com",

        # Microsoft 365 portal check settings.
        "m365_url": "https://m365.cloud.microsoft/?auth=2",
        "m365_portal_host": "m365.cloud.microsoft",

        # Cowork (Microsoft 365) access check settings.
        "cowork_url": "https://m365.cloud.microsoft/cowork",
    }
    # Default names for the per-check output columns (created only when the
    # relevant check is enabled).
    defaults["columns"]["w365_status"] = "w365_status"
    defaults["columns"]["w365_detail"] = "w365_detail"
    defaults["columns"]["copilot_status"] = "copilot_status"
    defaults["columns"]["copilot_detail"] = "copilot_detail"
    defaults["columns"]["pa_status"] = "pa_status"
    defaults["columns"]["pa_detail"] = "pa_detail"
    defaults["columns"]["m365_status"] = "m365_status"
    defaults["columns"]["m365_detail"] = "m365_detail"
    defaults["columns"]["cowork_status"] = "cowork_status"
    defaults["columns"]["cowork_detail"] = "cowork_detail"
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

    def __init__(self, wb, ws, col_cfg, output_keys=None):
        self.wb = wb
        self.ws = ws
        self.col_cfg = col_cfg
        # Which output columns to create if missing. Defaults to the GitHub
        # check's columns for backwards compatibility.
        self.output_keys = output_keys or ("status", "detail", "checked_at")
        self.header_row = 1
        self.headers = {}  # normalized header name -> column index (1-based)
        self._read_headers()
        self._ensure_output_columns()

    def _read_headers(self):
        for idx, cell in enumerate(self.ws[self.header_row], start=1):
            if cell.value is not None:
                self.headers[str(cell.value).strip()] = idx

    def _ensure_output_columns(self):
        """Create any configured output columns that are missing. Only the keys
        in self.output_keys are created, so a sheet that runs one check doesn't
        gain columns for the other."""
        for key in self.output_keys:
            name = self.col_cfg.get(key)
            if not name:
                continue
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


def _field_value(page, selector):
    """Return the current value of an input, or '' if unavailable."""
    try:
        return page.locator(selector).first.input_value(timeout=800) or ""
    except Exception:
        return ""


def _fill_and_verify(page, selector, value, timeout=6000):
    """Fill a field and confirm the value stuck (retry once). Returns bool."""
    for _ in range(2):
        if not _safe_fill(page, selector, value, timeout=timeout):
            continue
        if _field_value(page, selector) == value:
            return True
        # Value didn't stick (typeahead/overlay) — clear and retry.
        try:
            page.locator(selector).first.fill("", timeout=1500)
        except Exception:
            pass
    return _field_value(page, selector) == value


def _text_present(page, needles, timeout=800):
    try:
        body = page.locator("body").inner_text(timeout=timeout).lower()
    except Exception:
        return False
    return any(n.lower() in body for n in needles)


def attempt_login(page, cfg, github_username, azure_login, azure_tap,
                  azure_password, github_password):
    """
    Drive the browser through the login chain. Returns (status, detail).
    status in {SUCCESS, FAILED, MFA_REQUIRED, ERROR}.

    Flow: GitHub username -> "Sign in with your identity provider" -> Azure AD
    (UPN + Temporary Access Pass).
    """
    step_to = cfg["step_timeout_ms"]
    success_sub = cfg["success_url_substring"]

    # The Azure secret is the TAP; fall back to a password only if no TAP given.
    azure_secret = azure_tap or azure_password
    secret_kind = "TAP" if azure_tap else "password"
    if not azure_secret:
        return "ERROR", "Row has neither a Temporary Access Pass nor a password."

    page.set_default_timeout(step_to)
    try:
        page.goto(cfg["start_url"], wait_until="domcontentloaded",
                  timeout=cfg["nav_timeout_ms"])
    except PWTimeout:
        return "ERROR", f"Timed out loading start_url {cfg['start_url']}"

    filled_gh_user = False
    clicked_idp = False
    filled_ms_user = False
    filled_ms_secret = False
    secret_submits = 0

    # The login can bounce GitHub <-> Microsoft several times; loop until we
    # reach a terminal state or run out of iterations.
    for _ in range(28):
        page.wait_for_timeout(600)  # let redirects settle
        url = page.url.lower()

        # ---- Terminal: signed in to GitHub ----------------------------------
        if success_sub in url and _logged_into_github(page, github_username):
            return "SUCCESS", f"Signed in to GitHub as expected (url={page.url})"

        # ---- Microsoft: explicit credential error ---------------------------
        ms_err = _microsoft_error(page)
        if ms_err:
            return "FAILED", f"Azure AD rejected sign-in: {ms_err}"

        # ---- Microsoft: username (UPN) step ---------------------------------
        # The first Azure screen asks for the userPrincipalName. Fill ONLY the
        # username here (some tenants also render a hidden/adjacent password box,
        # but the flow is username-first).
        loginfmt_vis = _visible(page, 'input[name="loginfmt"]', timeout=700)
        if loginfmt_vis and not filled_ms_user:
            if not _fill_and_verify(page, 'input[name="loginfmt"]', azure_login):
                return "ERROR", ("Azure AD username field was not editable "
                                 "(possible bot challenge).")
            _click_first(page, ['#idSIButton9', 'input[type="submit"]',
                                'button[type="submit"]'])
            filled_ms_user = True
            continue

        # ---- Microsoft: Temporary Access Pass / password step ---------------
        # Only after the UPN has been submitted do we enter the secret, so we
        # never type it into the username screen's adjacent password box.
        if filled_ms_user:
            secret_sel = _azure_secret_field(page)
            if secret_sel:
                if not _fill_and_verify(page, secret_sel, azure_secret):
                    return "ERROR", (f"Azure AD {secret_kind} field was not "
                                     "editable (possible bot challenge).")
                _click_first(page, ['#idSIButton9', '#idA_SAOTCC_Continue',
                                    'input[type="submit"]', 'button[type="submit"]'])
                filled_ms_secret = True
                secret_submits += 1
                if secret_submits > 3:
                    # Field keeps coming back -> the secret isn't being accepted.
                    return "FAILED", f"Azure AD did not accept the {secret_kind}."
                continue
            # UPN screen reappeared after we submitted it -> account not accepted.
            if loginfmt_vis:
                return "FAILED", "Azure AD did not accept the account (loginfmt)."

        # ---- Microsoft: offer to switch to the Temporary Access Pass --------
        if azure_tap and _switch_to_tap(page):
            continue

        # ---- Microsoft: MFA / additional verification -----------------------
        if _is_mfa_challenge(page):
            return ("MFA_REQUIRED",
                    "Azure AD requested MFA / additional verification "
                    "(cannot be completed unattended).")

        # ---- Microsoft: "Stay signed in?" -----------------------------------
        if _text_present(page, ["stay signed in"]) or _visible(page, "#idBtn_Back"):
            # Click "No" to keep sessions clean.
            if not _click_first(page, ["#idBtn_Back"]):
                _click_first(page, ["#idSIButton9", 'input[type="submit"]'])
            continue

        # ---- GitHub: username + "Sign in with your identity provider" -------
        if _visible(page, "#login_field"):
            if not filled_gh_user:
                if not _safe_fill(page, "#login_field", github_username):
                    return "ERROR", ("GitHub username field was not editable "
                                     "(GitHub often blocks headless browsers — "
                                     "try headless=false).")
                filled_gh_user = True
                # Typing an EMU username flips the submit button to
                # "Sign in with your identity provider"; give the UI a beat and
                # re-evaluate on the next loop iteration.
                page.wait_for_timeout(900)
                continue
            if _click_identity_provider(page):
                clicked_idp = True
                continue
            # No IdP button on this page: fall back to the classic form if the
            # password box is present.
            if _visible(page, "#password"):
                gh_pw = github_password or azure_password
                if gh_pw and _safe_fill(page, "#password", gh_pw):
                    _click_first(page, ['input[name="commit"]',
                                        'button[type="submit"]',
                                        'input[type="submit"]'])
                    continue
                return "ERROR", ("GitHub showed a password form but no identity "
                                 "provider button and no usable github_password.")
            # IdP button may not be rendered yet; loop to let it appear.
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
            f"(gh_user={filled_gh_user}, idp={clicked_idp}, "
            f"ms_user={filled_ms_user}, ms_secret={filled_ms_secret})")


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


def _click_identity_provider(page):
    """Click GitHub's "Sign in with your identity provider" button/link.

    On github.com/login the submit control is an <input type="submit"> whose
    *value* becomes "Sign in with your identity provider" once an EMU username
    has been typed, so match on the value attribute as well as link/button text.
    """
    selectors = [
        'input[type="submit"][value*="identity provider" i]',
        'input[value*="identity provider" i]',
        'a:has-text("Sign in with your identity provider")',
        'button:has-text("Sign in with your identity provider")',
        'a:has-text("identity provider")',
        'button:has-text("identity provider")',
        'a:has-text("Single sign-on")',
        'a[href*="/sso"]',
    ]
    return _click_first(page, selectors)


# The Temporary Access Pass is entered on Microsoft's own screen. Depending on
# the tenant it appears in a dedicated "accesspass" box, in the generic OTC box,
# or simply in the standard password box.
_TAP_FIELD_SELECTORS = [
    'input[name="accesspass"]',
    'input[placeholder*="Temporary Access Pass" i]',
    'input[placeholder*="access pass" i]',
]


def _azure_secret_field(page):
    """
    Return the selector of the field where the TAP / password should be typed,
    or None if no such field is currently visible.
    """
    for sel in _TAP_FIELD_SELECTORS:
        if _visible(page, sel, timeout=600):
            return sel
    # If the page is clearly a TAP screen, the OTC box is the place to type it.
    if _text_present(page, ["temporary access pass"]) and \
            _visible(page, "#idTxtBx_SAOTCC_OTC", timeout=600):
        return "#idTxtBx_SAOTCC_OTC"
    # Standard "Enter password" box (also accepts a TAP in many tenants).
    if _visible(page, 'input[name="passwd"]', timeout=600):
        return 'input[name="passwd"]'
    return None


def _switch_to_tap(page):
    """Click a "Use your Temporary Access Pass instead" style link, if present."""
    selectors = [
        'a:has-text("Temporary Access Pass")',
        'button:has-text("Temporary Access Pass")',
        'a:has-text("Use a Temporary Access Pass")',
        '#idA_PWD_SwitchToCredPicker',
    ]
    return _click_first(page, selectors)


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
# Windows 365 portal check
# --------------------------------------------------------------------------- #
# The portal at windows365.microsoft.com redirects (after Azure AD sign-in) to
# the Windows App web client at windows.cloud.microsoft, where the user's Cloud
# PCs are listed under "Devices". Each Cloud PC renders a card whose test-id
# starts with "cloudPC-card-" and a connect button test-id "cloudpc-trigger-
# connect"; its aria-label reads "Connect to <name>. Press Enter to connect".
_W365_CARD_SELECTOR = ('[data-testid^="cloudPC-card-"], '
                       '[data-testid="cloudpc-trigger-connect"]')


def _w365_cloud_pcs(page):
    """Return a list of Cloud PC display names visible on the Devices page."""
    try:
        return page.evaluate(
            """() => {
                const names = new Set();
                document.querySelectorAll(
                    '[data-testid="cloudpc-trigger-connect"],'
                    + '[data-testid^="cloudPC-card-"]'
                ).forEach(e => {
                    let n = e.getAttribute('aria-label') || e.innerText || '';
                    n = n.replace(/^Connect to\\s*/i, '')
                         .replace(/\\.\\s*Press Enter.*$/i, '')
                         .replace(/\\s+/g, ' ').trim();
                    if (n) names.add(n);
                });
                return Array.from(names);
            }"""
        ) or []
    except Exception:
        return []


def _w365_power_state(page):
    """Best-effort read of a Cloud PC's power/status hint, or '' if none."""
    for sel in ['[data-testid="icon-status-indicator-button"]',
                '[data-testid^="cpc-status-indicator-"]']:
        try:
            loc = page.locator(sel).first
            if loc.count() > 0:
                label = (loc.get_attribute("aria-label") or "").strip()
                if label:
                    return label
        except Exception:
            continue
    return ""


def check_windows365(page, cfg, azure_login, azure_tap, azure_password):
    """
    Sign in to the Windows 365 portal with the UPN + Temporary Access Pass and
    report whether a Cloud PC exists for the user. Returns (status, detail).

    status in {SUCCESS, NO_CLOUDPC, FAILED, MFA_REQUIRED, ERROR}.
      SUCCESS    - at least one Windows 365 Cloud PC is present for the user.
      NO_CLOUDPC - signed in, but the user has no Cloud PC assigned.
      FAILED     - Azure AD rejected the sign-in.
      MFA_REQUIRED / ERROR - as for the GitHub check.
    """
    step_to = cfg["step_timeout_ms"]
    portal_host = cfg.get("windows365_portal_host", "windows.cloud.microsoft")

    azure_secret = azure_tap or azure_password
    secret_kind = "TAP" if azure_tap else "password"
    if not azure_secret:
        return "ERROR", "Row has neither a Temporary Access Pass nor a password."

    page.set_default_timeout(step_to)
    try:
        page.goto(cfg.get("windows365_url", "https://windows365.microsoft.com/"),
                  wait_until="domcontentloaded", timeout=cfg["nav_timeout_ms"])
    except PWTimeout:
        return "ERROR", "Timed out loading the Windows 365 portal."

    filled_ms_user = False
    filled_ms_secret = False
    secret_submits = 0

    # Phase 1: drive the Azure AD sign-in until we land on the portal host.
    for _ in range(32):
        page.wait_for_timeout(600)
        url = page.url.lower()

        if portal_host in url:
            break

        ms_err = _microsoft_error(page)
        if ms_err:
            return "FAILED", f"Azure AD rejected sign-in: {ms_err}"

        if _visible(page, 'input[name="loginfmt"]', timeout=700) and not filled_ms_user:
            if not _fill_and_verify(page, 'input[name="loginfmt"]', azure_login):
                return "ERROR", ("Azure AD username field was not editable "
                                 "(possible bot challenge).")
            _click_first(page, ['#idSIButton9', 'input[type="submit"]',
                                'button[type="submit"]'])
            filled_ms_user = True
            continue

        if filled_ms_user:
            secret_sel = _azure_secret_field(page)
            if secret_sel:
                if not _fill_and_verify(page, secret_sel, azure_secret):
                    return "ERROR", (f"Azure AD {secret_kind} field was not "
                                     "editable (possible bot challenge).")
                _click_first(page, ['#idSIButton9', '#idA_SAOTCC_Continue',
                                    'input[type="submit"]', 'button[type="submit"]'])
                filled_ms_secret = True
                secret_submits += 1
                if secret_submits > 3:
                    return "FAILED", f"Azure AD did not accept the {secret_kind}."
                continue

        if azure_tap and _switch_to_tap(page):
            continue

        if _is_mfa_challenge(page):
            return ("MFA_REQUIRED",
                    "Azure AD requested MFA / additional verification "
                    "(cannot be completed unattended).")

        # "Stay signed in?" — only after the secret, and detected by its own
        # text (never by the generic submit button, which also appears on the
        # username screen). Click Yes to carry the session into the portal.
        if filled_ms_secret and _text_present(page, ["stay signed in"]):
            if not _click_first(page, ["#idSIButton9"]):
                _click_first(page, ["#idBtn_Back"])
            continue
    else:
        return ("ERROR",
                f"Did not reach the Windows 365 portal. Final url={page.url} "
                f"(ms_user={filled_ms_user}, ms_secret={filled_ms_secret}).")

    # Phase 2: on the portal — dismiss the first-run tour, open Devices,
    # and look for a Cloud PC.
    try:
        page.wait_for_selector('[data-testid="nav-devices"]',
                               timeout=cfg["nav_timeout_ms"])
    except PWTimeout:
        return "ERROR", "Windows 365 portal did not finish loading (no navigation)."
    page.wait_for_timeout(1500)

    # The guided-tour overlay (Next / Not now) can sit over the nav.
    for _ in range(3):
        if _click_first(page, ['[data-testid="not-now-button"]',
                               'button:has-text("Not now")']):
            page.wait_for_timeout(1200)
            break
        if _click_first(page, ['[data-testid="next-button"]']):
            page.wait_for_timeout(900)
            continue
        break

    _click_first(page, ['[data-testid="nav-devices"]',
                        'button:has-text("Go to devices")'])

    # Give the device list time to load, then poll for a Cloud PC card.
    names = []
    for _ in range(12):
        page.wait_for_timeout(1000)
        names = _w365_cloud_pcs(page)
        if names:
            break

    if names:
        state = _w365_power_state(page)
        detail = f"Cloud PC present: {', '.join(names)}"
        if state:
            detail += f" ({state})"
        return "SUCCESS", detail

    # No card found. Distinguish "no Cloud PC" from a page that never rendered.
    if _text_present(page, ["no cloud pc", "don't have", "do not have",
                            "no devices", "nothing here", "no resources"]):
        return "NO_CLOUDPC", "Signed in, but no Windows 365 Cloud PC is assigned."
    if _visible(page, _W365_CARD_SELECTOR, timeout=800):
        return "SUCCESS", "Cloud PC present."
    return "NO_CLOUDPC", ("Signed in to the portal, but no Cloud PC was found on "
                          "the Devices page.")


# --------------------------------------------------------------------------- #
# Copilot Studio portal check
# --------------------------------------------------------------------------- #
# copilotstudio.microsoft.com signs in via Azure AD and then the SPA itself
# performs a *second* token acquisition, so the login screens (UPN + TAP) can
# appear twice before the app loads. Success is the authenticated app route,
# e.g. .../environments/~personal/home.
_AZURE_FATAL_ERR = ("incorrect", "isn't recogniz", "does not exist",
                    "couldn't find", "can't find", "cannot find",
                    "account or password", "didn't work", "locked",
                    "blocked", "disabled", "expired", "invalid")


def _azure_error_is_fatal(text):
    low = (text or "").lower()
    return any(k in low for k in _AZURE_FATAL_ERR)


# Signals that a Microsoft web app is authenticated. Classic Office portals show
# the "me control" / account manager; the newer *.cloud.microsoft Copilot apps
# (m365 / cowork) instead show a waffle app-launcher, a nav footer and a
# "<name>, Work account" control.
_AUTHED_SELECTORS = ('[aria-label^="Account manager"]',
                     'button[aria-label*="Account manager" i]',
                     '[data-automationid="meControl"]',
                     '#meControlButton',
                     '[data-testid="app-launcher-waffle-button"]',
                     '[aria-label*="Work account" i]',
                     '[data-testid="nav-footer"]')


def _portal_authed(page):
    for sel in _AUTHED_SELECTORS:
        try:
            if page.locator(sel).count() > 0:
                return True
        except Exception:
            continue
    return False


def _check_portal_login(page, cfg, azure_login, azure_tap, azure_password,
                        *, url_start, host, label, require_path=(),
                        require_authed=False):
    """
    Generic "can this user sign in to a Microsoft web portal" check. Drives the
    Azure AD UPN + Temporary Access Pass flow (handling the repeat auth round
    that SPAs like Copilot Studio / Power Automate perform) until the
    authenticated app loads on `host`. Returns (status, detail).

    require_path: optional substrings; if given, success also requires the URL
    path to contain one of them (guards against a brief pre-auth landing on the
    same host before the app bounces back to the login page).
    require_authed: if True, success also requires an authenticated UI signal
    (the account-manager control). Needed for portals like m365.cloud.microsoft
    that render an anonymous marketing page on the same host.

    status in {SUCCESS, FAILED, MFA_REQUIRED, ERROR}.
    """
    azure_secret = azure_tap or azure_password
    secret_kind = "TAP" if azure_tap else "password"
    if not azure_secret:
        return "ERROR", "Row has neither a Temporary Access Pass nor a password."

    page.set_default_timeout(cfg["step_timeout_ms"])
    try:
        page.goto(url_start, wait_until="domcontentloaded",
                  timeout=cfg["nav_timeout_ms"])
    except PWTimeout:
        return "ERROR", f"Timed out loading the {label} portal."

    secret_submits = 0
    login_rounds = 0
    stable = 0

    def _path_ok(url):
        return (not require_path) or any(s in url for s in require_path)

    for _ in range(90):
        page.wait_for_timeout(700)
        url = page.url.lower()
        on_login = ("microsoftonline" in url) or ("login.live" in url)

        # Only a clear credential-rejection message is fatal; transient
        # "enter a valid email" flashes between rounds are ignored.
        ms_err = _microsoft_error(page)
        if ms_err and _azure_error_is_fatal(ms_err):
            return "FAILED", f"Azure AD rejected sign-in: {ms_err}"

        # Success: on the app host, not on a login page, no sign-in field, and
        # (if required) on an authenticated route / showing an authed signal.
        on_host = (host in url and not on_login and _path_ok(url)
                   and not _visible(page, 'input[name="loginfmt"]', 300))
        if on_host and (not require_authed or _portal_authed(page)):
            stable += 1
            if stable >= 4:
                return "SUCCESS", f"Signed in to {label} (url={page.url})"
            continue
        stable = 0

        # Anonymous marketing landing on the app host (require_authed portals):
        # click "Sign in" to kick off the org auth.
        if on_host and require_authed and not _portal_authed(page):
            if _click_first(page, ['a[href*="/login"]', 'a:has-text("Sign in")',
                                   'button:has-text("Sign in")']):
                continue

        # Azure: username (may appear more than once).
        if _visible(page, 'input[name="loginfmt"]', 500):
            if _field_value(page, 'input[name="loginfmt"]') != azure_login:
                if not _fill_and_verify(page, 'input[name="loginfmt"]', azure_login):
                    return "ERROR", ("Azure AD username field was not editable "
                                     "(possible bot challenge).")
            _click_first(page, ['#idSIButton9', 'input[type="submit"]',
                                'button[type="submit"]'])
            login_rounds += 1
            if login_rounds > 4:
                return "ERROR", "Azure AD kept asking for the username."
            continue

        # Azure: Temporary Access Pass / password.
        secret_sel = _azure_secret_field(page) if on_login else None
        if secret_sel:
            if not _fill_and_verify(page, secret_sel, azure_secret):
                return "ERROR", (f"Azure AD {secret_kind} field was not editable "
                                 "(possible bot challenge).")
            _click_first(page, ['#idSIButton9', '#idA_SAOTCC_Continue',
                                'input[type="submit"]', 'button[type="submit"]'])
            secret_submits += 1
            if secret_submits > 5:
                return "FAILED", f"Azure AD did not accept the {secret_kind}."
            continue

        if azure_tap and _switch_to_tap(page):
            continue

        if _is_mfa_challenge(page):
            return ("MFA_REQUIRED",
                    "Azure AD requested MFA / additional verification "
                    "(cannot be completed unattended).")

        # "Stay signed in?" — Yes, so the session carries into the app's second
        # token request (avoids an endless re-prompt loop).
        if on_login and _text_present(page, ["stay signed in"]):
            if not _click_first(page, ["#idSIButton9"]):
                _click_first(page, ["#idBtn_Back"])
            continue

        # OAuth consent / permissions prompt.
        if _click_first(page, ['input[value="Accept"]',
                               'button:has-text("Accept")',
                               'button:has-text("Allow")',
                               'button:has-text("Yes")']):
            continue

    # Ran out of iterations.
    final = page.url.lower()
    if (host in final and not (("microsoftonline" in final) or ("login.live" in final))
            and _path_ok(final) and (not require_authed or _portal_authed(page))):
        return "SUCCESS", f"Signed in to {label} (url={page.url})"
    return ("ERROR",
            f"Did not reach the {label} app. Final url={page.url} "
            f"(login_rounds={login_rounds}, secret_submits={secret_submits}).")


def check_copilot_studio(page, cfg, azure_login, azure_tap, azure_password):
    """Sign-in check for Copilot Studio (reaches .../environments/.../home)."""
    return _check_portal_login(
        page, cfg, azure_login, azure_tap, azure_password,
        url_start=cfg.get("copilot_url", "https://copilotstudio.microsoft.com/"),
        host=cfg.get("copilot_portal_host", "copilotstudio.microsoft.com"),
        label="Copilot Studio",
        require_path=("/environments/", "/home"))


def check_power_automate(page, cfg, azure_login, azure_tap, azure_password):
    """Sign-in check for Power Automate (make.powerautomate.com)."""
    return _check_portal_login(
        page, cfg, azure_login, azure_tap, azure_password,
        url_start=cfg.get("powerautomate_url", "https://make.powerautomate.com/"),
        host=cfg.get("powerautomate_portal_host", "make.powerautomate.com"),
        label="Power Automate",
        require_path=())


def check_m365(page, cfg, azure_login, azure_tap, azure_password):
    """Sign-in check for the Microsoft 365 portal (m365.cloud.microsoft).

    Note: m365.cloud.microsoft shows an anonymous marketing page by default, so
    we deep-link to a protected route (?auth=2) to force the org sign-in."""
    return _check_portal_login(
        page, cfg, azure_login, azure_tap, azure_password,
        url_start=cfg.get("m365_url", "https://m365.cloud.microsoft/?auth=2"),
        host=cfg.get("m365_portal_host", "m365.cloud.microsoft"),
        label="Microsoft 365",
        require_authed=True)


def check_cowork(page, cfg, azure_login, azure_tap, azure_password):
    """
    Check whether the user has access to Cowork on the Microsoft 365 portal.
    Signs in (UPN + TAP), opens m365.cloud.microsoft/cowork, and looks for the
    access banner ("You have access to Cowork"). Returns (status, detail).

    status in {SUCCESS, NO_ACCESS, FAILED, MFA_REQUIRED, ERROR}.
      SUCCESS   - the user has access to Cowork.
      NO_ACCESS - signed in, but Cowork is not available to the user.
    """
    cowork_url = cfg.get("cowork_url", "https://m365.cloud.microsoft/cowork")
    host = cfg.get("m365_portal_host", "m365.cloud.microsoft")

    # Reuse the shared Azure login to reach an authenticated m365 session.
    status, detail = _check_portal_login(
        page, cfg, azure_login, azure_tap, azure_password,
        url_start=cowork_url, host=host, label="Cowork (Microsoft 365)",
        require_authed=True)
    if status != "SUCCESS":
        return status, detail

    # Make sure we're on the Cowork route, then read the access banner.
    try:
        if "cowork" not in page.url.lower():
            page.goto(cowork_url, wait_until="domcontentloaded",
                      timeout=cfg["nav_timeout_ms"])
    except PWTimeout:
        return "ERROR", "Timed out loading the Cowork page."

    has_text = ("you have access to cowork", "great news")
    no_text = ("don't have access", "do not have access", "no access to cowork",
               "not licensed", "isn't available", "is not available",
               "request access", "doesn't have access")
    # Poll briefly while the Cowork page renders its banner.
    for _ in range(12):
        page.wait_for_timeout(1000)
        if _text_present(page, has_text):
            return "SUCCESS", "User has access to Cowork."
        if _text_present(page, no_text):
            return "NO_ACCESS", "Signed in, but no access to Cowork."

    # Banner never appeared — report what we ended on rather than guessing.
    return "NO_ACCESS", ("Signed in, but the Cowork access banner was not found "
                         f"(url={page.url}).")


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
                        help="comma-separated usernames (GitHub username or UPN) "
                             "to test; others skipped")
    parser.add_argument("--recheck", action="store_true",
                        help="re-test rows that already have a status")
    args = parser.parse_args()

    cfg = load_config(args.config)
    if args.headless:
        cfg["headless"] = True

    do_github = bool(cfg.get("check_github", True))
    do_w365 = bool(cfg.get("check_windows365", False))
    do_copilot = bool(cfg.get("check_copilot", False))
    do_pa = bool(cfg.get("check_powerautomate", False))
    do_m365 = bool(cfg.get("check_m365", False))
    do_cowork = bool(cfg.get("check_cowork", False))
    if not (do_github or do_w365 or do_copilot or do_pa or do_m365 or do_cowork):
        print("ERROR: all checks are disabled in config.", file=sys.stderr)
        sys.exit(2)

    only = None
    if args.only:
        only = {u.strip() for u in args.only.split(",") if u.strip()}

    excel_path = Path(cfg["excel_file"])
    if not excel_path.exists():
        print(f"ERROR: Excel file not found: {excel_path.resolve()}", file=sys.stderr)
        print("Run `python make_template.py` to create a starter file.", file=sys.stderr)
        sys.exit(2)

    # Which sheets to process: a name, a list of names, or None (active sheet).
    wb = load_workbook(excel_path)
    sheet_cfg = cfg.get("sheet_name")
    if isinstance(sheet_cfg, list):
        sheet_names = sheet_cfg
    elif sheet_cfg:
        sheet_names = [sheet_cfg]
    else:
        sheet_names = [wb.active.title]

    output_keys = ["checked_at"]
    if do_github:
        output_keys = ["status", "detail"] + output_keys
    if do_w365:
        output_keys = ["w365_status", "w365_detail"] + output_keys
    if do_copilot:
        output_keys = ["copilot_status", "copilot_detail"] + output_keys
    if do_pa:
        output_keys = ["pa_status", "pa_detail"] + output_keys
    if do_m365:
        output_keys = ["m365_status", "m365_detail"] + output_keys
    if do_cowork:
        output_keys = ["cowork_status", "cowork_detail"] + output_keys

    totals = {"github": {"SUCCESS": 0, "FAILED": 0, "other": 0},
              "w365": {"SUCCESS": 0, "FAILED": 0, "other": 0},
              "copilot": {"SUCCESS": 0, "FAILED": 0, "other": 0},
              "pa": {"SUCCESS": 0, "FAILED": 0, "other": 0},
              "m365": {"SUCCESS": 0, "FAILED": 0, "other": 0},
              "cowork": {"SUCCESS": 0, "FAILED": 0, "other": 0}}
    tested = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=cfg["headless"],
                                    slow_mo=cfg["slow_mo_ms"])
        shared_context = None
        if not cfg["clear_session_between_users"]:
            shared_context = browser.new_context()

        for sheet_name in sheet_names:
            if sheet_name not in wb.sheetnames:
                print(f"[warn] sheet '{sheet_name}' not found; skipping.")
                continue
            sheet = Sheet(wb, wb[sheet_name], cfg["columns"], output_keys)
            print(f"\n==== Sheet: {sheet_name} ====")

            for row in sheet.iter_rows():
                gh_user = sheet.get(row, "github_username", required=False) if do_github else None
                az_login = sheet.get(row, "azure_login", required=False)
                az_tap = sheet.get(row, "azure_tap", required=False)
                az_pass = sheet.get(row, "azure_password", required=False)
                gh_pass = sheet.get(row, "github_password", required=False) if do_github else None

                if not gh_user and not az_login:
                    continue  # blank row

                if only is not None and gh_user not in only and az_login not in only:
                    continue

                # Per-check skip: only run a check that hasn't been done (unless
                # --recheck). A row is skipped only when nothing is left to do.
                gh_done = sheet.get(row, "status", required=False) if do_github else None
                w_done = sheet.get(row, "w365_status", required=False) if do_w365 else None
                cp_done = sheet.get(row, "copilot_status", required=False) if do_copilot else None
                pa_done = sheet.get(row, "pa_status", required=False) if do_pa else None
                m_done = sheet.get(row, "m365_status", required=False) if do_m365 else None
                cw_done = sheet.get(row, "cowork_status", required=False) if do_cowork else None
                need_github = do_github and (args.recheck or not gh_done)
                need_w365 = do_w365 and (args.recheck or not w_done)
                need_copilot = do_copilot and (args.recheck or not cp_done)
                need_pa = do_pa and (args.recheck or not pa_done)
                need_m365 = do_m365 and (args.recheck or not m_done)
                need_cowork = do_cowork and (args.recheck or not cw_done)
                if not (need_github or need_w365 or need_copilot or need_pa
                        or need_m365 or need_cowork):
                    print(f"[skip] {gh_user or az_login}: already done "
                          f"(use --recheck to redo)")
                    continue

                if not az_login or not (az_tap or az_pass):
                    miss = "Missing userPrincipalName or TAP/password"
                    if need_github:
                        sheet.set(row, "status", "SKIPPED")
                        sheet.set(row, "detail", miss)
                    if need_w365:
                        sheet.set(row, "w365_status", "SKIPPED")
                        sheet.set(row, "w365_detail", miss)
                    if need_copilot:
                        sheet.set(row, "copilot_status", "SKIPPED")
                        sheet.set(row, "copilot_detail", miss)
                    if need_pa:
                        sheet.set(row, "pa_status", "SKIPPED")
                        sheet.set(row, "pa_detail", miss)
                    if need_m365:
                        sheet.set(row, "m365_status", "SKIPPED")
                        sheet.set(row, "m365_detail", miss)
                    if need_cowork:
                        sheet.set(row, "cowork_status", "SKIPPED")
                        sheet.set(row, "cowork_detail", miss)
                    sheet.set(row, "checked_at", datetime.now().isoformat(timespec="seconds"))
                    sheet.save(excel_path)
                    print(f"[skip] row {row}: missing credentials")
                    continue

                tested += 1
                label = gh_user or az_login
                print(f"[test] {label} ...", flush=True)

                per_user_ctx = shared_context or browser.new_context()
                try:
                    if need_github:
                        page = per_user_ctx.new_page()
                        page.set_default_navigation_timeout(cfg["nav_timeout_ms"])
                        try:
                            status, detail = attempt_login(
                                page, cfg, gh_user or "", az_login, az_tap, az_pass, gh_pass)
                        except Exception as e:  # noqa: BLE001
                            status, detail = "ERROR", f"Unhandled exception: {e!r}"
                        finally:
                            page.close()
                        sheet.set(row, "status", status)
                        sheet.set(row, "detail", detail)
                        sheet.set(row, "checked_at", datetime.now().isoformat(timespec="seconds"))
                        sheet.save(excel_path)
                        _tally(totals["github"], status)
                        print(f"       github -> {status}: {detail}")

                    if need_w365:
                        page = per_user_ctx.new_page()
                        page.set_default_navigation_timeout(cfg["nav_timeout_ms"])
                        try:
                            status, detail = check_windows365(
                                page, cfg, az_login, az_tap, az_pass)
                        except Exception as e:  # noqa: BLE001
                            status, detail = "ERROR", f"Unhandled exception: {e!r}"
                        finally:
                            page.close()
                        sheet.set(row, "w365_status", status)
                        sheet.set(row, "w365_detail", detail)
                        sheet.set(row, "checked_at", datetime.now().isoformat(timespec="seconds"))
                        sheet.save(excel_path)
                        _tally(totals["w365"], status)
                        print(f"       win365 -> {status}: {detail}")

                    if need_copilot:
                        page = per_user_ctx.new_page()
                        page.set_default_navigation_timeout(cfg["nav_timeout_ms"])
                        try:
                            status, detail = check_copilot_studio(
                                page, cfg, az_login, az_tap, az_pass)
                        except Exception as e:  # noqa: BLE001
                            status, detail = "ERROR", f"Unhandled exception: {e!r}"
                        finally:
                            page.close()
                        sheet.set(row, "copilot_status", status)
                        sheet.set(row, "copilot_detail", detail)
                        sheet.set(row, "checked_at", datetime.now().isoformat(timespec="seconds"))
                        sheet.save(excel_path)
                        _tally(totals["copilot"], status)
                        print(f"       copilot -> {status}: {detail}")

                    if need_pa:
                        page = per_user_ctx.new_page()
                        page.set_default_navigation_timeout(cfg["nav_timeout_ms"])
                        try:
                            status, detail = check_power_automate(
                                page, cfg, az_login, az_tap, az_pass)
                        except Exception as e:  # noqa: BLE001
                            status, detail = "ERROR", f"Unhandled exception: {e!r}"
                        finally:
                            page.close()
                        sheet.set(row, "pa_status", status)
                        sheet.set(row, "pa_detail", detail)
                        sheet.set(row, "checked_at", datetime.now().isoformat(timespec="seconds"))
                        sheet.save(excel_path)
                        _tally(totals["pa"], status)
                        print(f"       powerautomate -> {status}: {detail}")

                    if need_m365:
                        page = per_user_ctx.new_page()
                        page.set_default_navigation_timeout(cfg["nav_timeout_ms"])
                        try:
                            status, detail = check_m365(
                                page, cfg, az_login, az_tap, az_pass)
                        except Exception as e:  # noqa: BLE001
                            status, detail = "ERROR", f"Unhandled exception: {e!r}"
                        finally:
                            page.close()
                        sheet.set(row, "m365_status", status)
                        sheet.set(row, "m365_detail", detail)
                        sheet.set(row, "checked_at", datetime.now().isoformat(timespec="seconds"))
                        sheet.save(excel_path)
                        _tally(totals["m365"], status)
                        print(f"       m365 -> {status}: {detail}")

                    if need_cowork:
                        page = per_user_ctx.new_page()
                        page.set_default_navigation_timeout(cfg["nav_timeout_ms"])
                        try:
                            status, detail = check_cowork(
                                page, cfg, az_login, az_tap, az_pass)
                        except Exception as e:  # noqa: BLE001
                            status, detail = "ERROR", f"Unhandled exception: {e!r}"
                        finally:
                            page.close()
                        sheet.set(row, "cowork_status", status)
                        sheet.set(row, "cowork_detail", detail)
                        sheet.set(row, "checked_at", datetime.now().isoformat(timespec="seconds"))
                        sheet.save(excel_path)
                        _tally(totals["cowork"], status)
                        print(f"       cowork -> {status}: {detail}")
                finally:
                    if shared_context is None:
                        per_user_ctx.close()

                time.sleep(cfg["between_users_delay_ms"] / 1000.0)

        if shared_context is not None:
            shared_context.close()
        browser.close()

    print("\n==== Summary ====")
    print(f"Users tested: {tested}")
    if do_github:
        g = totals["github"]
        print(f"GitHub   -> SUCCESS: {g['SUCCESS']}  FAILED: {g['FAILED']}  "
              f"Other: {g['other']}")
    if do_w365:
        w = totals["w365"]
        print(f"Win365   -> SUCCESS: {w['SUCCESS']}  FAILED: {w['FAILED']}  "
              f"Other: {w['other']} (NO_CLOUDPC / MFA_REQUIRED / ERROR / SKIPPED)")
    if do_copilot:
        cp = totals["copilot"]
        print(f"Copilot  -> SUCCESS: {cp['SUCCESS']}  FAILED: {cp['FAILED']}  "
              f"Other: {cp['other']} (MFA_REQUIRED / ERROR / SKIPPED)")
    if do_pa:
        pa = totals["pa"]
        print(f"PowerAut -> SUCCESS: {pa['SUCCESS']}  FAILED: {pa['FAILED']}  "
              f"Other: {pa['other']} (MFA_REQUIRED / ERROR / SKIPPED)")
    if do_m365:
        m = totals["m365"]
        print(f"M365     -> SUCCESS: {m['SUCCESS']}  FAILED: {m['FAILED']}  "
              f"Other: {m['other']} (MFA_REQUIRED / ERROR / SKIPPED)")
    if do_cowork:
        cw = totals["cowork"]
        print(f"Cowork   -> SUCCESS: {cw['SUCCESS']}  FAILED: {cw['FAILED']}  "
              f"Other: {cw['other']} (NO_ACCESS / MFA_REQUIRED / ERROR / SKIPPED)")
    print(f"Results written to: {excel_path.resolve()}")


def _tally(counter, status):
    if status == "SUCCESS":
        counter["SUCCESS"] += 1
    elif status == "FAILED":
        counter["FAILED"] += 1
    else:
        counter["other"] += 1


if __name__ == "__main__":
    main()
