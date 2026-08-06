"""Generate CRM_Database.xlsx from parsed quotation data."""
import json, os, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC = "/home/user/crmpro/scripts/quotes.json"
OUT = "/home/user/crmpro/data/CRM_Database.xlsx"

with open(SRC) as f:
    quotes = json.load(f)

# ---- Date parsing (dd/mm/yy) -> python date ----
def to_date(s):
    d, m, y = s.split("/")
    return f"20{y}-{int(m):02d}-{int(d):02d}"

HEAD_FILL = PatternFill("solid", fgColor="1B2A4A")
HEAD_FONT = Font(color="FFFFFF", bold=True, size=11)
thin = Side(style="thin", color="D0D5E0")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
BAND = PatternFill("solid", fgColor="F4F6FB")
ALT = PatternFill("solid", fgColor="FFFFFF")

def style_sheet(ws, widths, n_cols, n_rows):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for c in ws.iter_cols(1, n_cols):
        ws.cell(row=1, column=c[0].column).fill = HEAD_FILL
        ws.cell(row=1, column=c[0].column).font = HEAD_FONT
        ws.cell(row=1, column=c[0].column).alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"
    for row in ws.iter_rows(min_row=2, max_row=n_rows, max_col=n_cols):
        for cell in row:
            cell.border = BORDER
            if cell.row % 2 == 0:
                cell.fill = BAND
        ws.row_dimensions[row[0].row].height = 20

wb = Workbook()

# ============ Sheet 1: Quotations ============
ws = wb.active
ws.title = "Quotations"
ws.append(["Quotation No", "Date Created", "Time", "Company", "Status", "Amount (RM)", "Created By"])
for q in quotes:
    ws.append([q["qno"], to_date(q["date"]), q["time"], q["company"], q["status"], q["amount"], q["created_by"]])

style_sheet(ws, [14, 14, 10, 38, 22, 14, 12], 7, ws.max_row)
# number format for amount
for r in range(2, ws.max_row + 1):
    ws.cell(row=r, column=6).number_format = '#,##0.00'

# ============ Sheet 2: Companies (leads) ============
ws2 = wb.create_sheet("Companies")
ws2.append(["Company", "Industry", "Contact Person", "Email", "Phone", "Lead Source", "Stage", "Notes"])
# unique companies, sorted
companies = sorted({q["company"] for q in quotes})
for c in companies:
    ws2.append([c, "", "", "", "", "Quotation", "New", ""])
style_sheet(ws2, [40, 22, 22, 28, 18, 16, 14, 30], 8, ws2.max_row)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
wb.save(OUT)
print("Saved:", OUT, "| quotations:", ws.max_row - 1, "| companies:", ws2.max_row - 1)
