#!/usr/bin/env python3
"""Create a starter users.xlsx with the expected columns and two sample rows."""
import json
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font

cfg = json.loads(Path("config.json").read_text())
cols = cfg["columns"]
out = Path(cfg["excel_file"])

headers = [
    cols["github_username"],
    cols["azure_login"],
    cols["azure_password"],
    "github_password",     # optional; leave blank to fall back to azure_password
    cols["status"],
    cols["detail"],
    cols["checked_at"],
]

wb = Workbook()
ws = wb.active
ws.title = "users"
ws.append(headers)
for cell in ws[1]:
    cell.font = Font(bold=True)

# Example rows — replace with real data.
ws.append(["alice_org", "alice@contoso.com", "REPLACE_ME", "", "", "", ""])
ws.append(["bob_org", "bob@contoso.com", "REPLACE_ME", "", "", "", ""])

# Widen columns a little for readability.
widths = [20, 28, 18, 18, 14, 60, 22]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[chr(64 + i)].width = w

if out.exists():
    print(f"Refusing to overwrite existing {out}. Delete it first if you want a fresh template.")
else:
    wb.save(out)
    print(f"Created {out.resolve()} — fill in real users and passwords, then run check_logins.py")
