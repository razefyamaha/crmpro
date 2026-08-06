"""Parse the Quotation Report PDF into a structured list of quotations."""
import pdfplumber, re, json

PDF = "/home/user/uploads/Quotation Report.pdf"

dates = []      # (date, time, company)
statuses = []   # (status, qno, amount)

DATE_RE = re.compile(r"^(\d{1,2}/\d{1,2}/\d{2})\s+(\d{1,2}:\d{2}\s[AP]M)\s+(.+)$")
STAT_RE = re.compile(r"^(Quotation In Progress|Deal Lost|Completed|Deal Won)\s+(QT\d+)\s+([\d.,]+)$")

with pdfplumber.open(PDF) as pdf:
    pages = pdf.pages
    # Pages 0-2 -> date/company columns
    for p in pages[0:3]:
        for line in (p.extract_text() or "").split("\n"):
            m = DATE_RE.match(line)
            if m:
                dates.append({"date": m.group(1), "time": m.group(2), "company": m.group(3)})
    # Pages 3-5 -> status/qno/amount columns
    for p in pages[3:6]:
        for line in (p.extract_text() or "").split("\n"):
            m = STAT_RE.match(line)
            if m:
                statuses.append({"status": m.group(1), "qno": m.group(2), "amount": m.group(3)})
    # Pages 6-8 -> created by column
    created = []
    for p in pages[6:9]:
        for line in (p.extract_text() or "").split("\n"):
            line = line.strip()
            if line in ("Razef", "Jimmy Law"):
                created.append(line)

print("dates:", len(dates), "statuses:", len(statuses), "created:", len(created))

rows = []
n = min(len(dates), len(statuses), len(created))
for i in range(n):
    d = dates[i]
    s = statuses[i]
    rows.append({
        "qno": s["qno"],
        "date": d["date"],
        "time": d["time"],
        "company": d["company"],
        "status": s["status"],
        "amount": float(s["amount"].replace(",", "")),
        "created_by": created[i],
    })

print("combined rows:", len(rows))
print(json.dumps(rows, indent=2)[:2000])
with open("/home/user/crmpro/scripts/quotes.json", "w") as f:
    json.dump(rows, f, indent=2)
