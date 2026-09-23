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
    cols["azure_login"],      # userPrincipalName
    cols["azure_password"],   # password (optional fallback)
    cols["azure_tap"],        # Temporary Access Pass (preferred secret)
    cols["github_username"],  # GitHub_Username
    cols["status"],
    cols["detail"],
    cols["checked_at"],
]

wb = Workbook()
ws = wb.active
ws.title = cfg.get("sheet_name") or "Users"
ws.append(headers)
for cell in ws[1]:
    cell.font = Font(bold=True)

# Example rows — replace with real data.
ws.append(["user01@example.com", "REPLACE_ME", "REPLACE_TAP", "user01_org", "", "", ""])
ws.append(["user02@example.com", "REPLACE_ME", "REPLACE_TAP", "user02_org", "", "", ""])

# Widen columns a little for readability.
widths = [32, 18, 18, 22, 14, 60, 22]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[chr(64 + i)].width = w

if out.exists():
    print(f"Refusing to overwrite existing {out}. Delete it first if you want a fresh template.")
else:
    wb.save(out)
    print(f"Created {out.resolve()} — fill in real users and passwords, then run check_logins.py")
