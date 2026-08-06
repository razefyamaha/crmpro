"""
pdf_parse.py - extract customer + quotation details from an uploaded
quotation PDF so the customer detail view can be auto-populated.

The parser targets the common quotation layout seen in CRM Pro (ARK IT GLOBAL)
PDFs, e.g.:
    Company Code :300-S0004 Date :01-07-2026
    Company Name :SEAMEC MALAYSIA SDN BHD Payment Term :COD
    PIC Name :MS TEOH Remarks :-
    Customer Phone :+60163378743
    Number
    Customer Address :-
"""
import re


def extract_text(pdf_path):
    import pdfplumber
    chunks = []
    with pdfplumber.open(pdf_path) as pdf:
        for p in pdf.pages:
            t = p.extract_text() or ""
            if t:
                chunks.append(t)
    return "\n".join(chunks)


def _clean(v):
    if v is None:
        return ""
    v = v.strip().strip(":").strip()
    v = v.replace("Number", "").strip()
    if v in ("-", ":"):
        return ""
    return v


_KNOWN = (r"Company\s+Code|Company\s+Name|PIC\s+Name|Customer\s+Phone|"
          r"Customer\s+Address|Payment\s+Term|Remarks|Date|Quotation\s+No|"
          r"Created\s+By|Total|Subtotal|Tax|Powered")


def parse_quotation_pdf(pdf_path):
    """Return a dict of extracted fields (empty string when not found)."""
    text = extract_text(pdf_path)
    if not text:
        return {}

    def grab(pattern):
        # Capture value up to the next known field label or end of line, so a
        # line like "Company Name :X Payment Term :COD" yields only "X".
        m = re.search(pattern + rf"(?=\s+{_KNOWN}\s*:|\n|$)", text)
        return _clean(m.group(1)) if m else ""

    def grab_simple(pattern):
        m = re.search(pattern, text)
        return _clean(m.group(1)) if m else ""

    # Phone: the field is sometimes split as "Customer Phone :+6016..." then a
    # stray "Number" word on the next line. Capture up to the next known field.
    phone = ""
    m = re.search(r"Customer\s+Phone\s*:\s*([^\n]+)" + rf"(?=\s+{_KNOWN}\s*:|\n|$)", text)
    if m:
        phone = _clean(m.group(1))

    total = ""
    for pat in (r"Total\s+Incl\.?\s*Tax\s*\(?\s*MYR\s*\)?\s*:\s*([\d.,]+)",
                r"Total\s*\(?\s*MYR\s*\)?\s*:\s*([\d.,]+)"):
        m = re.search(pat, text, re.I)
        if m:
            total = m.group(1).replace(",", "").replace(" ", "")
            break

    return {
        "company": grab(r"Company\s+Name\s*:\s*([A-Z0-9][^:]*?)"),
        "company_code": grab_simple(r"Company\s+Code\s*:\s*([A-Za-z0-9\-]+)"),
        "pic": grab(r"PIC\s+Name\s*:\s*([A-Z0-9][^:]*?)"),
        "phone": phone,
        "address": grab_simple(r"Customer\s+Address\s*:\s*(.+?)(?=\s*Customer\s+|\n|$)"),
        "payment_term": grab_simple(r"Payment\s+Term\s*:\s*(.+?)(?=\s*Remarks|\n|$)"),
        "remarks": grab(r"Remarks\s*:\s*([^\n]+)"),
        "quotation_no": grab_simple(r"Quotation\s+No\s*:\s*(QT\d+)"),
        "date": grab_simple(r"\bDate\s*:\s*(\d{1,2}-\d{1,2}-\d{4})"),
        "total": total,
    }


def _num(s):
    """Parse a string that may contain commas/currency into a float, or 0."""
    if s is None:
        return 0.0
    s = re.sub(r"[^\d.,]", "", str(s))
    s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_quotation_lines(pdf_path):
    """Extract the line-items table from a quotation PDF.

    Returns a list of dicts:
        {no, name, description, qty, unit, unit_price, selling, cost, margin}
    Selling = the unit price read from the PDF (cost/margin start blank so the
    user can fill them in while editing).
    """
    text = extract_text(pdf_path)
    if not text:
        return []
    lines = text.splitlines()
    items = []
    # A data row looks like: "1. Software 2 SEAT 721.65 1,443.30 115.46 1,558.76"
    # followed by optional "Product Name:" and "Description :" lines.
    row_re = re.compile(
        r"^\s*(\d+)[.)]\s+([^\d]+?)\s+(\d+)\s+([A-Za-z]{2,5})\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s*$"
    )
    pending_name = ""
    pending_desc = ""
    pending = None

    def flush():
        nonlocal pending, pending_name, pending_desc
        if pending is None:
            return
        name = (pending_name or pending[1]).strip()
        items.append({
            "no": pending[0],
            "name": name,
            "description": pending_desc.strip(),
            "qty": pending[2],
            "unit": pending[3],
            "unit_price": pending[4],
            "selling": pending[4],
            "cost": "",
            "margin": "",
            "profit": "",
        })
        pending = None
        pending_name = ""
        pending_desc = ""

    for raw in lines:
        ln = raw.strip()
        m = row_re.match(ln)
        if m:
            flush()
            pending = (int(m.group(1)), m.group(2).strip(), int(m.group(3)),
                       m.group(4).strip(), _num(m.group(5)))
            continue
        pm = re.match(r"^Product\s+Name\s*:\s*(.+)$", ln, re.I)
        if pm:
            pending_name = pm.group(1).strip()
            continue
        dm = re.match(r"^Description\s*:\s*(.+)$", ln, re.I)
        if dm:
            pending_desc = dm.group(1).strip()
            continue
    flush()
    return items

