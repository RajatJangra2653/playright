#!/usr/bin/env bash
# One-time setup: create a virtualenv, install deps, download the Chromium browser.
set -euo pipefail
cd "$(dirname "$0")"

echo "==> Creating virtual environment (.venv)"
python3 -m venv .venv

echo "==> Installing Python dependencies"
./.venv/bin/python -m pip install --upgrade pip
./.venv/bin/python -m pip install -r requirements.txt

echo "==> Downloading Chromium for Playwright"
./.venv/bin/python -m playwright install chromium

echo
echo "Setup complete. Next:"
echo "  ./.venv/bin/python make_template.py     # create users.xlsx (once)"
echo "  # edit users.xlsx with real accounts, then:"
echo "  ./.venv/bin/python check_logins.py"
