"""
CRM Pro - Customer Relationship Management
Bootstrap 5 front-end + Microsoft Excel as the database (openpyxl)

Run:
    python app.py
Then open http://127.0.0.1:5000
"""
import os
import io
import csv
import json
import re
import sys
import time
import threading
import webbrowser
from datetime import datetime
from collections import Counter

from flask import (Flask, render_template, request, redirect, url_for,
                   flash, Response, jsonify, send_file, abort, session)
from functools import wraps
from werkzeug.utils import secure_filename
from werkzeug.security import check_password_hash

import db
import pdf_parse
from openpyxl import load_workbook

# Resolve the base/data folders so the app works both from source and when
# bundled into a standalone .exe (PyInstaller).
if getattr(sys, "frozen", False):
    BASE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.executable)))
    DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(sys.executable)), "data")
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    DATA_DIR = os.path.join(BASE_DIR, "data")

os.makedirs(DATA_DIR, exist_ok=True)

app = Flask(__name__,
            template_folder=os.path.join(BASE_DIR, "templates"),
            static_folder=os.path.join(BASE_DIR, "static"))
app.secret_key = "crmpro-secret-key"
app.config["CRM_VERSION"] = "2.0 (with Login & User Management)"
app.config["JSON_SORT_KEYS"] = False


# ---------------- Authentication ----------------
def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "user" not in session:
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapped


def role_required(*roles):
    def decorator(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            if "user" not in session:
                return redirect(url_for("login", next=request.path))
            if session.get("role") not in roles:
                flash("You do not have permission to access that page.", "danger")
                return redirect(url_for("dashboard"))
            return view(*args, **kwargs)
        return wrapped
    return decorator


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = db.find_user(username)
        if user and user.get("active") and check_password_hash(user["password"], password):
            session["user"] = user["username"]
            session["name"] = user["name"]
            session["role"] = user["role"]
            flash(f"Welcome back, {user['name'] or user['username']}!", "success")
            nxt = request.args.get("next")
            if nxt and nxt.startswith("/"):
                return redirect(nxt)
            return redirect(url_for("dashboard"))
        flash("Invalid username or password.", "danger")
        return render_template("login.html")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("You have been logged out.", "info")
    return redirect(url_for("login"))


@app.before_request
def require_login():
    """Require login for all pages except the login/logout/static endpoints."""
    allowed = ("login", "logout", "static")
    if request.endpoint and request.endpoint in allowed:
        return None
    if "user" not in session:
        return redirect(url_for("login", next=request.path))
    return None


@app.context_processor
def inject_globals():
    """Expose the list of existing customer names to all templates so company
    fields can offer a dropdown (datalist) of existing customers while still
    allowing a brand-new customer to be typed in."""
    try:
        names = set()
        for c in db.load_companies():
            if c.get("company"):
                names.add(c["company"])
        for q in db.load_quotations():
            if q.get("company"):
                names.add(q["company"])
        company_names = sorted(names)
    except Exception:
        company_names = []
    return {
        "company_names": company_names,
        "current_theme": db.current_theme(),
        "current_user": session.get("user"),
        "current_user_name": session.get("name") or session.get("user"),
        "current_role": session.get("role"),
        "crm_version": app.config.get("CRM_VERSION"),
    }


# ---------------- Uploads / PDF auto-fill ----------------
def _save_quote_file(company, upload):
    """Save an uploaded quotation file into the customer's folder and return
    (saved_filename, parsed_details). saved_filename is the actual name on disk
    (may differ from the original if there was a name collision)."""
    if not upload or not upload.filename:
        return None, {}
    fname = secure_filename(upload.filename)
    if not fname:
        return None, {}
    folder = db.customer_folder(company)
    ext = os.path.splitext(fname)[1].lower()
    dest = os.path.join(folder, fname)
    base, e = os.path.splitext(fname)
    n = 1
    while os.path.exists(dest):
        dest = os.path.join(folder, f"{base}_{n}{e}")
        n += 1
    upload.save(dest)
    saved_name = os.path.basename(dest)
    parsed = {}
    if ext == ".pdf":
        try:
            parsed = pdf_parse.parse_quotation_pdf(dest)
        except Exception:
            parsed = {}
    return saved_name, parsed



def _find_quote_file(company, qno):
    """Find an uploaded PDF for a quotation by matching its QT number in the
    customer's folder (e.g. 'QT01250_..._Quotation.pdf')."""
    if not qno:
        return None
    qno = str(qno).upper()
    try:
        folder = db.customer_folder(company)
        if not os.path.isdir(folder):
            return None
        for name in os.listdir(folder):
            if name.lower().endswith(".pdf") and qno in name.upper():
                return name
    except Exception:
        return None
    return None


_QNO_RE = re.compile(r"(QT\d{4,})", re.IGNORECASE)


def _extract_qno(text):
    """Pull the first quotation number (e.g. 'QT01250') out of a filename."""
    if not text:
        return None
    m = _QNO_RE.search(text)
    return m.group(1).upper() if m else None


# ---------------- helpers ----------------
def month_key(date_str):
    try:
        d = datetime.strptime(str(date_str)[:10], "%Y-%m-%d")
        return d.strftime("%Y-%m")
    except (ValueError, TypeError):
        return ""


def month_label(mk):
    try:
        return datetime.strptime(mk, "%Y-%m").strftime("%b %y")
    except ValueError:
        return mk


def _month_key_from_ym(y, m):
    return f"{y:04d}-{m:02d}"


def closed_sale_amounts(quotes):
    """Sum of 'closed' sales (Completed + Deal Won) grouped by month key."""
    months = {}
    for q in quotes:
        if q["status"] not in ("Completed", "Deal Won"):
            continue
        mk = month_key(q["date"])
        if mk:
            months[mk] = months.get(mk, 0.0) + q["amount"]
    return months


def monthly_target_data(quotes, n=3):
    """Last n months (ending at current month) closed-sale values vs target."""
    today = datetime.now()
    y, m = today.year, today.month
    months = []
    for _ in range(n):
        months.append((y, m))
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    months.reverse()

    closed = closed_sale_amounts(quotes)
    target = db.monthly_target()
    rows = []
    cumulative = 0.0
    for y, m in months:
        mk = _month_key_from_ym(y, m)
        val = round(closed.get(mk, 0.0), 2)
        cumulative += val
        rows.append({
            "label": datetime(y, m, 1).strftime("%b %y"),
            "month_key": mk,
            "value": val,
            "target": target,
            "progress": min(100.0, val / target * 100) if target else 0,
        })
    return rows, cumulative, target


# ---------------- Dashboard ----------------
@app.route("/")
def dashboard():
    quotes = db.load_quotations()
    companies = db.load_companies()

    total_count = len(quotes)
    total_amount = sum(q["amount"] for q in quotes)
    by_status = Counter(q["status"] for q in quotes)
    pipeline = sum(q["amount"] for q in quotes
                   if q["status"] in ("Quotation In Progress", "Negotiating"))
    won_amount = sum(q["amount"] for q in quotes
                     if q["status"] in ("Deal Won", "Completed"))

    # status cards
    cards = [
        {"label": "Total Quotations", "value": f"{total_count:,}",
         "sub": f"RM {total_amount:,.2f} total value", "icon": "bi-file-earmark-text", "color": "primary"},
        {"label": "In Progress", "value": f"{by_status.get('Quotation In Progress', 0):,}",
         "sub": f"Pipeline RM {pipeline:,.2f}", "icon": "bi-hourglass-split", "color": "info"},
        {"label": "Completed + Won", "value": f"{by_status.get('Completed', 0) + by_status.get('Deal Won', 0):,}",
         "sub": f"Won RM {won_amount:,.2f}", "icon": "bi-check2-circle", "color": "success"},
        {"label": "Deal Lost", "value": f"{by_status.get('Deal Lost', 0):,}",
         "sub": f"Companies {len(companies):,}", "icon": "bi-x-octagon", "color": "danger"},
    ]

    # status distribution chart
    status_labels = ["Quotation In Progress", "Completed", "Deal Won", "Deal Lost"]
    status_values = [by_status.get(s, 0) for s in status_labels]

    # monthly trend
    months = {}
    for q in quotes:
        mk = month_key(q["date"])
        if mk:
            months.setdefault(mk, {"count": 0, "amount": 0.0})
            months[mk]["count"] += 1
            months[mk]["amount"] += q["amount"]
    month_keys = sorted(months.keys())
    trend_labels = [month_label(m) for m in month_keys]
    trend_counts = [months[m]["count"] for m in month_keys]
    trend_amounts = [round(months[m]["amount"], 2) for m in month_keys]

    # top companies by completed sales (Completed + Deal Won), with amount
    comp_closed = {}
    for q in quotes:
        if q["status"] not in ("Completed", "Deal Won"):
            continue
        comp_closed.setdefault(q["company"], [0, 0.0])
        comp_closed[q["company"]][0] += 1
        comp_closed[q["company"]][1] += q["amount"]
    top_comp = sorted(comp_closed.items(), key=lambda x: x[1][0], reverse=True)[:5]
    top_comp_labels = [c for c, _ in top_comp]
    top_comp_counts = [v[0] for _, v in top_comp]
    top_comp_amounts = [round(v[1], 2) for _, v in top_comp]

    # monthly target (60k) - last 3 months closed sales
    target_rows, cumulative_target, target = monthly_target_data(quotes, 3)
    target_labels = json.dumps([r["label"] for r in target_rows])
    target_values = json.dumps([r["value"] for r in target_rows])
    target_progress = json.dumps([round(r["progress"], 1) for r in target_rows])
    target_total = round(sum(r["value"] for r in target_rows), 2)

    # recent quotations
    recent = sorted(quotes, key=lambda q: (str(q["date"]), str(q["time"])), reverse=True)[:6]

    # sales-by-month selector: user picks a month to view completed/won sales
    sel_month = request.args.get("smonth", "").strip()  # e.g. "2026-06"
    # all months present in data, for the dropdown
    months_avail = sorted({month_key(q["date"]) for q in quotes if month_key(q["date"])}, reverse=True)
    if sel_month not in months_avail and months_avail:
        sel_month = months_avail[0]  # default to most recent month with data

    # closed-sales chart: Jan -> current month for a selected year, with a
    # per-month count + value of Completed / Deal Won sales.
    chart_year = request.args.get("cyear", "").strip() or str(datetime.now().year)
    closed_by_month = {}  # "2026-01" -> {"count":n,"amount":x}
    for q in quotes:
        mk = month_key(q["date"])
        if not mk or mk[:4] != str(chart_year):
            continue
        if q["status"] in ("Completed", "Deal Won"):
            closed_by_month.setdefault(mk, {"count": 0, "amount": 0.0})
            closed_by_month[mk]["count"] += 1
            closed_by_month[mk]["amount"] += q["amount"]
    # build Jan..current-month labels
    cur_month = datetime.now().month if str(datetime.now().year) == str(chart_year) else 12
    closed_labels = []
    closed_counts = []
    closed_amounts = []
    for m in range(1, cur_month + 1):
        key = f"{chart_year}-{m:02d}"
        d = closed_by_month.get(key, {"count": 0, "amount": 0.0})
        closed_labels.append(datetime(int(chart_year), m, 1).strftime("%b"))
        closed_counts.append(d["count"])
        closed_amounts.append(round(d["amount"], 2))
    chart_years = sorted({month_key(q["date"])[:4] for q in quotes if month_key(q["date"])}, reverse=True)
    if chart_year not in chart_years and chart_years:
        chart_year = chart_years[0]

    month_sales = []
    month_total = 0.0
    month_won_count = 0
    if sel_month:
        for q in quotes:
            if month_key(q["date"]) == sel_month:
                if q["status"] in ("Completed", "Deal Won"):
                    month_sales.append(q)
                    month_total += q["amount"]
                    month_won_count += 1
        month_sales.sort(key=lambda q: str(q["date"]), reverse=True)
    month_label_str = ""
    if sel_month:
        try:
            month_label_str = datetime.strptime(sel_month, "%Y-%m").strftime("%B %Y")
        except ValueError:
            month_label_str = sel_month

    # forecast stats for the dashboard card
    forecasts = db.load_forecasts()
    fc_total_value = round(sum(f["value"] for f in forecasts), 2)
    fc_total_profit = round(sum(f["profit"] for f in forecasts), 2)
    fc_count = len(forecasts)
    fc_open = sum(1 for f in forecasts if f["status"] == "Open")

    return render_template(
        "dashboard.html",
        cards=cards,
        fc_total_value=fc_total_value,
        fc_total_profit=fc_total_profit,
        fc_count=fc_count,
        fc_open=fc_open,
        months_avail=months_avail,
        sel_month=sel_month,
        month_label_str=month_label_str,
        month_sales=month_sales,
        month_total=round(month_total, 2),
        month_won_count=month_won_count,
        chart_year=chart_year,
        chart_years=chart_years,
        closed_labels=json.dumps(closed_labels),
        closed_counts=json.dumps(closed_counts),
        closed_amounts=json.dumps(closed_amounts),
        target_rows=target_rows,
        target=target,
        target_total=target_total,
        target_labels=target_labels,
        target_values=target_values,
        target_progress=target_progress,
        cumulative_target=round(cumulative_target, 2),
        total_amount=total_amount,
        status_labels=json.dumps(status_labels),
        status_values=json.dumps(status_values),
        trend_labels=json.dumps(trend_labels),
        trend_counts=json.dumps(trend_counts),
        trend_amounts=json.dumps(trend_amounts),
        top_comp_labels=json.dumps(top_comp_labels),
        top_comp_counts=json.dumps(top_comp_counts),
        top_comp_amounts=json.dumps(top_comp_amounts),
        recent=recent,
        active="dashboard",
    )


# ---------------- Quotations ----------------
PAYMENT_TERMS = ["COD", "7 Days", "14 Days", "30 Days", "60 Days", "On Delivery"]


@app.route("/quotations")
def quotations():
    quotes = db.load_quotations()
    search = request.args.get("search", "").strip().lower()
    status = request.args.get("status", "").strip()
    files = request.args.get("files", "").strip()
    date_from = request.args.get("date_from", "").strip()
    date_to = request.args.get("date_to", "").strip()
    month = request.args.get("month", "").strip()   # e.g. "2026-06"
    this_year = request.args.get("year", "").strip()  # "current"
    sort = request.args.get("sort", "date")
    order = request.args.get("order", "desc")

    # filter
    filtered = quotes
    if search:
        filtered = [q for q in filtered
                    if search in str(q["company"]).lower()
                    or search in str(q["qno"]).lower()
                    or search in str(q["created_by"]).lower()
                    or search in (_find_quote_file(q["company"], q["qno"]) or "").lower()]
    if status:
        filtered = [q for q in filtered if q["status"] == status]
    if files == "yes":
        filtered = [q for q in filtered if _find_quote_file(q["company"], q["qno"])]
    elif files == "no":
        filtered = [q for q in filtered if not _find_quote_file(q["company"], q["qno"])]
    if month:
        filtered = [q for q in filtered if str(q["date"])[:7] == month]
    if this_year == "current":
        cy = str(datetime.now().year)
        filtered = [q for q in filtered if str(q["date"])[:4] == cy]
    if date_from:
        filtered = [q for q in filtered if str(q["date"]) >= date_from]
    if date_to:
        filtered = [q for q in filtered if str(q["date"]) <= date_to]

    # sort
    rev = (order == "desc")
    if sort == "amount":
        filtered.sort(key=lambda q: q["amount"], reverse=rev)
    elif sort == "company":
        filtered.sort(key=lambda q: str(q["company"]).lower(), reverse=rev)
    elif sort == "qno":
        filtered.sort(key=lambda q: str(q["qno"]), reverse=rev)
    elif sort == "status":
        filtered.sort(key=lambda q: str(q["status"]), reverse=rev)
    else:
        filtered.sort(key=lambda q: (str(q["date"]), str(q["time"])), reverse=rev)

    # pagination
    page = request.args.get("page", 1, type=int)
    per_page = 15
    total = len(filtered)
    pages = max(1, (total + per_page - 1) // per_page)
    page = min(page, pages)
    start = (page - 1) * per_page
    paged = filtered[start:start + per_page]
    # flag which quotations have an uploaded PDF + attach company contact info
    companies = db.load_companies()
    cmap = {str(c["company"]).lower(): c for c in companies}
    paged = []
    for q in filtered[start:start + per_page]:
        c = cmap.get(str(q["company"]).lower(), {})
        paged.append(dict(q,
                          has_file=bool(_find_quote_file(q["company"], q["qno"])),
                          company_cid=c.get("id"),
                          company_contact=c.get("contact", ""),
                          company_phone=c.get("phone", ""),
                          company_email=c.get("email", ""),
                          company_address=c.get("address", "")))

    return render_template(
        "quotations.html",
        quotes=paged,
        total=total,
        page=page,
        pages=pages,
        search=search,
        status=status,
        files=files,
        date_from=date_from,
        date_to=date_to,
        month=month,
        this_year=this_year,
        cur_year=datetime.now().year,
        sort=sort,
        order=order,
        statuses=db.STATUSES,
        reps=sorted({q["created_by"] or "" for q in quotes}),
        active="quotations",
    )


@app.route("/quotations/parse-preview", methods=["POST"])
def quotation_parse_preview():
    """Read an uploaded quotation PDF and return the extracted customer /
    quotation details as JSON so the New Quotation form can auto-fill."""
    upload = request.files.get("quote_file")
    if not upload or not upload.filename:
        return jsonify({"ok": False, "error": "No file"}), 400
    fname = secure_filename(upload.filename)
    if not fname.lower().endswith(".pdf"):
        return jsonify({"ok": False, "error": "Not a PDF"}), 400
    import tempfile
    tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    tmp_path = tmp.name
    try:
        upload.save(tmp_path)
        try:
            parsed = pdf_parse.parse_quotation_pdf(tmp_path)
        except Exception:
            parsed = {}
    finally:
        tmp.close()
        try:
            os.remove(tmp_path)
        except Exception:
            pass
    return jsonify({"ok": True, "data": parsed or {}})


@app.route("/quotations/add", methods=["POST"])
def quotation_add():
    quotes = db.load_quotations()
    qno = request.form.get("qno") or db.next_qno(quotes)
    date = request.form.get("date") or db.today()
    time = request.form.get("time") or db.now_time()
    company = request.form.get("company", "").strip()
    if not company:
        flash("Company is required.", "danger")
        return redirect(url_for("quotations"))
    try:
        amount = float(request.form.get("amount", 0))
    except ValueError:
        amount = 0.0
    status = request.form.get("status") or "Quotation In Progress"
    created_by = request.form.get("created_by", "").strip()

    quotes.append({
        "id": len(quotes),
        "qno": qno.upper(),
        "date": date,
        "time": time,
        "company": company,
        "status": status,
        "amount": amount,
        "created_by": created_by,
    })
    db.save_quotations(quotes)
    # upload quotation file (if any) - just store it in the customer's folder
    quote_file = request.files.get("quote_file")
    if quote_file and quote_file.filename:
        _save_quote_file(company, quote_file)
    companies = db.load_companies()
    existing = next((c for c in companies if c["company"].lower() == company.lower()), None)
    if not existing:
        companies.append({"company": company, "industry": "", "contact": "",
                          "email": "", "phone": "", "source": "Quotation",
                          "stage": "Quoted", "notes": "", "address": "", "pipeline": {}})
        db.save_companies(companies)
    flash(f"Quotation {qno.upper()} added.", "success")
    # If the add came from a customer detail page, return there so the new
    # quotation is visible right away in context.
    if request.form.get("back") == "customer":
        try:
            back_cid = int(request.form.get("back_cid") or "")
        except (TypeError, ValueError):
            back_cid = None
        if back_cid is not None:
            return redirect(url_for("customer_detail", cid=back_cid))
    return redirect(url_for("quotations"))


@app.route("/quotations/edit", methods=["POST"])
def quotation_edit():
    qid = request.form.get("id", type=int)
    quotes = db.load_quotations()
    if qid is None or not (0 <= qid < len(quotes)):
        flash("Quotation not found.", "danger")
        return redirect(url_for("quotations"))
    q = quotes[qid]
    q["qno"] = (request.form.get("qno") or q["qno"]).upper()
    q["date"] = request.form.get("date") or q["date"]
    q["time"] = request.form.get("time") or q["time"]
    q["company"] = request.form.get("company", "").strip() or q["company"]
    q["status"] = request.form.get("status") or q["status"]
    q["payment_term"] = request.form.get("payment_term", "").strip() or q.get("payment_term", "")
    q["remarks"] = request.form.get("remarks", "").strip() or q.get("remarks", "")
    q["created_by"] = request.form.get("created_by", "").strip() or q["created_by"]

    # full edit: rebuild line items
    pids = request.form.getlist("product_id")
    qtys = request.form.getlist("qty")
    descs = request.form.getlist("desc")
    specs = request.form.getlist("spec")
    catalogues = request.form.getlist("catalogue")
    costs = request.form.getlist("cost")
    margins = request.form.getlist("margin")
    srps = request.form.getlist("srp")
    ssts = request.form.getlist("sst")
    products = db.load_products()
    prod_map = {str(p["id"]): p for p in products}

    # Determine the number of line rows from the longest submitted field list.
    n_rows = max(len(qtys), len(descs), len(costs), len(margins), len(pids))

    lines = []
    total_selling = 0.0
    for i in range(n_rows):
        pid = pids[i] if i < len(pids) else ""
        p = prod_map.get(pid) if pid else None
        desc = descs[i].strip() if i < len(descs) else ""
        if not p and not desc:
            continue
        try:
            qty = int(float(qtys[i] or 1)) if i < len(qtys) else 1
        except (ValueError, TypeError):
            qty = 1
        if qty <= 0:
            qty = 1
        def _num(v, default):
            try:
                return float(v) if v not in (None, "") else default
            except (TypeError, ValueError):
                return default
        name = p["name"] if p else (descs[i].strip() if i < len(descs) else "Item")
        cost = _num(costs[i] if i < len(costs) else None, p["cost"] if p else 0)
        margin = _num(margins[i] if i < len(margins) else None, p["margin"] if p else 0)
        srp = _num(srps[i] if i < len(srps) else None, p.get("srp", 0) if p else 0)
        selling = cost / (1 - margin / 100) if (0 <= margin < 100) else cost
        profit = selling - cost
        apply_sst = (i < len(ssts)) and (ssts[i] == "1")
        subtotal = selling * qty
        tax = round(subtotal * SST_RATE, 2) if apply_sst else 0.0
        total_selling += subtotal
        lines.append({
            "no": len(lines) + 1, "name": name, "qty": qty,
            "spec": (specs[i] if i < len(specs) else "").strip(),
            "catalogue": (catalogues[i] if i < len(catalogues) else "").strip(),
            "cost": round(cost, 2), "margin": round(margin, 2),
            "srp": round(srp, 2), "selling": round(selling, 2),
            "profit": round(profit, 2), "sst": apply_sst,
            "subtotal": round(subtotal, 2), "tax": tax,
            "line_total": round(subtotal + tax, 2),
        })

    if lines:
        q["lines"] = lines
        # amount = total selling price only
        q["amount"] = round(total_selling, 2)
    else:
        try:
            q["amount"] = float(request.form.get("amount"))
        except (ValueError, TypeError):
            pass

    db.save_quotations(quotes)

    # update customer (company) contact fields if provided
    companies = db.load_companies()
    comp = next((c for c in companies if str(c["company"]).lower() == str(q["company"]).lower()), None)
    if comp:
        for fld in ("contact", "phone", "email", "address"):
            v = request.form.get(fld, "").strip()
            if v:
                comp[fld] = v
        db.save_companies(companies)

    flash(f"Quotation {q['qno']} updated.", "success")
    return redirect(url_for("quotations"), code=303)


@app.route("/quotations/delete", methods=["POST"])
def quotation_delete():
    qid = request.form.get("id", type=int)
    quotes = db.load_quotations()
    if qid is not None and 0 <= qid < len(quotes):
        qno = quotes[qid]["qno"]
        quotes.pop(qid)
        db.save_quotations(quotes)
        flash(f"Quotation {qno} deleted.", "success")
    return redirect(url_for("quotations"))


# ---------------- Customers (leads) data routes ----------------
# The Companies page was removed (the Customers page covers it), but these
# routes power the Customers page's add / edit / delete / stage actions and
# the "New Lead" sidebar. They read/write the same "Companies" Excel sheet.
@app.route("/companies/add", methods=["POST"])
def company_add():
    companies = db.load_companies()
    company = request.form.get("company", "").strip()
    if not company:
        flash("Company name is required.", "danger")
        return redirect(url_for("quotations"))

    quote_file = request.files.get("quote_file")
    if quote_file and quote_file.filename:
        _save_quote_file(company, quote_file)

    existing = next((c for c in companies if c["company"].lower() == company.lower()), None)
    if existing:
        # "use existing customer" - update supplied fields, keep the rest
        for fld in ("industry", "contact", "email", "phone", "source", "notes", "address"):
            v = request.form.get(fld, "").strip()
            if v:
                existing[fld] = v
        st = request.form.get("stage", "").strip()
        if st:
            existing["stage"] = st
        db.save_companies(companies)
        msg = f"'{company}' already exists — details updated."
        msg += " Quotation file uploaded." if quote_file and quote_file.filename else ""
        flash(msg, "success")
        return redirect(url_for("quotations"))

    new = {
        "company": company,
        "industry": request.form.get("industry", "").strip(),
        "contact": request.form.get("contact", "").strip(),
        "email": request.form.get("email", "").strip(),
        "phone": request.form.get("phone", "").strip(),
        "source": request.form.get("source", "").strip() or "Manual",
        "stage": request.form.get("stage") or "New",
        "notes": request.form.get("notes", "").strip(),
        "address": request.form.get("address", "").strip(),
        "pipeline": {},
    }
    companies.append(new)
    db.save_companies(companies)
    flash(f"{company} added.", "success")
    return redirect(url_for("quotations"))


@app.route("/companies/edit", methods=["POST"])
def company_edit():
    cid = request.form.get("id", type=int)
    companies = db.load_companies()
    if cid is None or not (0 <= cid < len(companies)):
        flash("Company not found.", "danger")
        return redirect(url_for("quotations"))
    c = companies[cid]
    c["company"] = request.form.get("company", "").strip() or c["company"]
    c["industry"] = request.form.get("industry", "").strip()
    c["contact"] = request.form.get("contact", "").strip()
    c["email"] = request.form.get("email", "").strip()
    c["phone"] = request.form.get("phone", "").strip()
    c["source"] = request.form.get("source", "").strip()
    c["stage"] = request.form.get("stage") or c["stage"]
    c["notes"] = request.form.get("notes", "").strip()
    c["address"] = request.form.get("address", "").strip()
    db.save_companies(companies)
    flash(f"{c['company']} updated.", "success")
    return redirect(url_for("customer_detail", cid=cid))


@app.route("/companies/delete", methods=["POST"])
def company_delete():
    cid = request.form.get("id", type=int)
    companies = db.load_companies()
    if cid is not None and 0 <= cid < len(companies):
        name = companies[cid]["company"]
        companies.pop(cid)
        db.save_companies(companies)
        flash(f"{name} deleted.", "success")
    return redirect(url_for("quotations"))


@app.route("/companies/stage", methods=["POST"])
def company_stage():
    """Quickly update only a lead's stage (keeps all other fields intact)."""
    cid = request.form.get("id", type=int)
    stage = request.form.get("stage", "").strip()
    companies = db.load_companies()
    if cid is not None and 0 <= cid < len(companies):
        c = companies[cid]
        c["stage"] = stage if stage else c["stage"]
        db.save_companies(companies)
        flash(f"'{c['company']}' stage updated to {c['stage']}.", "success")
    return redirect(url_for("quotations"))


# ---------------- Potential Customers ----------------
ACTIVE_STAGES = ["New", "Contacted", "Quoted", "Negotiating"]


@app.route("/customer/<int:cid>")
def customer_detail(cid):
    companies = db.load_companies()
    if not (0 <= cid < len(companies)):
        flash("Customer not found.", "danger")
        return redirect(url_for("quotations"))
    customer = companies[cid]
    quotes = db.load_quotations()
    c_quotes = [q for q in quotes
                if str(q["company"]).lower() == str(customer["company"]).lower()]
    c_quotes.sort(key=lambda q: (str(q["date"]), str(q["time"])), reverse=True)
    files = db.list_uploads(customer["company"])
    # Annotate each file with the quotation number found in its name and that
    # quotation's total amount, so the list can show a price badge. This only
    # reads/quotes existing data - it never auto-fills or modifies anything.
    for f in files:
        qno = _extract_qno(f["name"])
        f["linked_qno"] = qno
        f["linked_amount"] = ""
        f["linked_date"] = ""
        if qno:
            m = next((q for q in quotes
                      if str(q["qno"]).upper() == qno
                      and str(q["company"]).lower() == str(customer["company"]).lower()), None)
            if m:
                f["linked_amount"] = m.get("amount")
                f["linked_date"] = m.get("date")
                # Auto-extract line items from the PDF so they can be edited.
                if not m.get("lines"):
                    try:
                        pdf_lines = pdf_parse.parse_quotation_lines(
                            db.customer_file_path(customer["company"], f["name"]))
                        if pdf_lines:
                            m["lines"] = pdf_lines
                    except Exception:
                        pass
    return render_template(
        "customer_detail.html",
        customer=customer,
        quotes=c_quotes,
        files=files,
        products=db.load_products(),
        statuses=db.STATUSES,
        active="board",
    )


@app.route("/customer/<int:cid>/upload", methods=["POST"])
def customer_upload(cid):
    companies = db.load_companies()
    if not (0 <= cid < len(companies)):
        flash("Customer not found.", "danger")
        return redirect(url_for("quotations"))
    customer = companies[cid]
    upload = request.files.get("quote_file")
    saved_fname, _parsed = _save_quote_file(customer["company"], upload)
    if saved_fname:
        flash(f"File uploaded to {customer['company']}'s folder.", "success")
    else:
        flash("No file selected.", "warning")
    return redirect(url_for("customer_detail", cid=cid))


@app.route("/customer/<int:cid>/file/<path:fname>")
def customer_file(cid, fname):
    companies = db.load_companies()
    if not (0 <= cid < len(companies)):
        abort(404)
    customer = companies[cid]
    folder = db.customer_folder(customer["company"])
    safe = os.path.basename(fname)  # prevent path traversal
    path = os.path.join(folder, safe)
    if not os.path.isfile(path):
        abort(404)
    return send_file(path)


@app.route("/customer/<int:cid>/file/<path:fname>/thumb")
def customer_file_thumb(cid, fname):
    """Serve (and lazily generate) a cached thumbnail of a PDF's first page."""
    placeholder = os.path.join(BASE_DIR, "static", "img", "pdf_placeholder.png")
    companies = db.load_companies()
    if not (0 <= cid < len(companies)):
        return send_file(placeholder, mimetype="image/png")
    customer = companies[cid]
    safe = os.path.basename(fname)
    folder = db.customer_folder(customer["company"])
    src = os.path.join(folder, safe)
    if not os.path.isfile(src):
        return send_file(placeholder, mimetype="image/png")
    if not safe.lower().endswith(".pdf"):
        return send_file(placeholder, mimetype="image/png")
    thumb = db.thumbnail_path(customer["company"], safe)
    # regenerate if missing or older than the source file
    try:
        if not thumb or not os.path.isfile(thumb) or os.path.getmtime(thumb) < os.path.getmtime(src):
            if not db.render_thumbnail(src, thumb):
                return send_file(placeholder, mimetype="image/png")
        return send_file(thumb, mimetype="image/jpeg", max_age=3600)
    except Exception:
        # Can't render this PDF (e.g. encrypted/corrupt) - show the placeholder.
        return send_file(placeholder, mimetype="image/png")


@app.route("/customer/<int:cid>/file/<path:fname>/pages")
def customer_file_pages(cid, fname):
    """Return the page count of an uploaded PDF as JSON (for page navigation)."""
    companies = db.load_companies()
    if not (0 <= cid < len(companies)):
        abort(404)
    customer = companies[cid]
    safe = os.path.basename(fname)
    path = os.path.join(db.customer_folder(customer["company"]), safe)
    if not os.path.isfile(path) or not safe.lower().endswith(".pdf"):
        return jsonify({"pages": 1})
    return jsonify({"pages": db.pdf_page_count(path)})





@app.route("/quotations/<qno>/file")
def quotation_file(qno):
    """Serve the uploaded PDF linked to a quotation (matched by QT number)."""
    quotes = db.load_quotations()
    q = next((x for x in quotes if str(x["qno"]).upper() == qno.upper()), None)
    if not q:
        abort(404)
    fname = _find_quote_file(q["company"], q["qno"])
    if not fname:
        abort(404)
    return send_file(os.path.join(db.customer_folder(q["company"]), fname))


@app.route("/quotations/<qno>/file/pages")
def quotation_file_pages(qno):
    quotes = db.load_quotations()
    q = next((x for x in quotes if str(x["qno"]).upper() == qno.upper()), None)
    if not q:
        return jsonify({"pages": 1})
    fname = _find_quote_file(q["company"], q["qno"])
    if not fname:
        return jsonify({"pages": 1})
    path = os.path.join(db.customer_folder(q["company"]), fname)
    return jsonify({"pages": db.pdf_page_count(path)})


@app.route("/quotations/<qno>/file/page/<int:page>")
def quotation_file_page(qno, page):
    """Render one page of a quotation's PDF as an image (for inline viewing)."""
    quotes = db.load_quotations()
    q = next((x for x in quotes if str(x["qno"]).upper() == qno.upper()), None)
    if not q:
        abort(404)
    fname = _find_quote_file(q["company"], q["qno"])
    if not fname:
        abort(404)
    src = os.path.join(db.customer_folder(q["company"]), fname)
    if not os.path.isfile(src):
        abort(404)
    # URL page numbers are 1-based; render_page expects 0-based.
    idx = page - 1
    img = db.page_image_path(q["company"], fname, idx)
    try:
        if not os.path.isfile(img) or os.path.getmtime(img) < os.path.getmtime(src):
            if not db.render_page(src, idx, img):
                abort(404)
        return send_file(img, mimetype="image/jpeg", max_age=3600)
    except Exception:
        abort(404)


@app.route("/customer/<int:cid>/file/<path:fname>/page/<int:page>")
def customer_file_page(cid, fname, page):
    """Render one page of a customer's PDF as an image (for inline viewing)."""
    companies = db.load_companies()
    if not (0 <= cid < len(companies)):
        abort(404)
    customer = companies[cid]
    safe = os.path.basename(fname)
    src = os.path.join(db.customer_folder(customer["company"]), safe)
    if not os.path.isfile(src):
        abort(404)
    if not safe.lower().endswith(".pdf"):
        return send_file(src)  # not a PDF, serve as-is
    # URL page numbers are 1-based; render_page expects 0-based.
    idx = page - 1
    img = db.page_image_path(customer["company"], safe, idx)
    try:
        if not os.path.isfile(img) or os.path.getmtime(img) < os.path.getmtime(src):
            if not db.render_page(src, idx, img):
                abort(404)
        return send_file(img, mimetype="image/jpeg", max_age=3600)
    except Exception:
        abort(404)


@app.route("/customer/<int:cid>/file/delete", methods=["POST"])
def customer_file_delete(cid):
    companies = db.load_companies()
    if not (0 <= cid < len(companies)):
        flash("Customer not found.", "danger")
        return redirect(url_for("quotations"))
    customer = companies[cid]
    fname = os.path.basename(request.form.get("name", ""))  # prevent path traversal
    if fname:
        folder = db.customer_folder(customer["company"])
        path = os.path.join(folder, fname)
        if os.path.isfile(path):
            os.remove(path)
            flash(f"'{fname}' deleted.", "success")
    return redirect(url_for("customer_detail", cid=cid))


# ---------------- Forecasts ----------------
SST_RATE = 0.08


def _is_sst_category(cat):
    """8% SST applies to Service or Software product categories."""
    c = (cat or "").lower()
    return "service" in c or "software" in c


def _build_forecast_lines(pids, qtys, costs, srps, margins, ssts, products):
    """Build line items from the submitted form arrays. Editable cost/SRP/margin
    are stored ONLY in the forecast (they are NOT written back to the Products
    sheet). Each line has a manual 8% SST toggle (ssts list: '1'=apply, else no).
    Returns (lines, total_value, total_cost, total_profit)."""
    prod_map = {str(p["id"]): p for p in products}
    lines = []

    def _num(v, default):
        try:
            return float(v) if v not in (None, "") else default
        except (TypeError, ValueError):
            return default

    for i, pid in enumerate(pids):
        p = prod_map.get(pid)
        if not p:
            continue
        qty = int(float(qtys[i] or 1)) if i < len(qtys) else 1
        if qty <= 0:
            qty = 1
        cost = _num(costs[i] if i < len(costs) else None, p["cost"])
        srp = _num(srps[i] if i < len(srps) else None, p.get("srp", 0))
        margin = _num(margins[i] if i < len(margins) else None, p.get("margin", 0))
        # selling price from cost + margin
        selling = cost / (1 - margin / 100) if (0 <= margin < 100) else cost
        profit = selling - cost
        # manual SST toggle: '1' = apply 8%, otherwise no SST
        apply_sst = (i < len(ssts)) and (ssts[i] == "1")
        lines.append({
            "name": p["name"], "qty": qty,
            "cost": round(cost, 2), "selling": round(selling, 2),
            "profit": round(profit, 2), "srp": round(srp, 2),
            "margin": round(margin, 2), "category": p.get("category", ""),
            "sst": apply_sst, "sst_amount": round(selling * qty * SST_RATE, 2) if apply_sst else 0.0,
        })

    total_value = 0.0
    total_cost = 0.0
    total_profit = 0.0
    for ln in lines:
        line_value = ln["selling"] * ln["qty"]
        line_sst = ln["sst_amount"]
        line_cost = ln["cost"] * ln["qty"]
        line_profit = line_value + line_sst - line_cost
        total_value += line_value + line_sst
        total_cost += line_cost
        total_profit += line_profit

    return lines, round(total_value, 2), round(total_cost, 2), round(total_profit, 2)


@app.route("/forecast")
def forecast_page():
    products = db.load_products()
    forecasts = db.load_forecasts()
    forecasts.sort(key=lambda f: (str(f["date"]), f["fc_no"]), reverse=True)

    # search + filter
    search = request.args.get("search", "").strip().lower()
    date_from = request.args.get("date_from", "").strip()
    date_to = request.args.get("date_to", "").strip()
    status_f = request.args.get("status", "").strip()
    month_f = request.args.get("month", "").strip()   # e.g. "06" or "6"
    year_f = request.args.get("year", "").strip()      # e.g. "2026"

    filtered = forecasts
    if search:
        filtered = [f for f in filtered
                    if search in str(f["customer"]).lower()
                    or search in f["fc_no"].lower()
                    or search in str(f["items"]).lower()]
    if status_f:
        filtered = [f for f in filtered if f["status"] == status_f]
    if month_f:
        mm = int(month_f)
        filtered = [f for f in filtered if str(f["date"])[:7].endswith(f"-{mm:02d}")]
    if year_f:
        yy = str(year_f)
        filtered = [f for f in filtered if str(f["date"])[:4] == yy]
    if date_from:
        filtered = [f for f in filtered if str(f["date"]) >= date_from]
    if date_to:
        filtered = [f for f in filtered if str(f["date"]) <= date_to]

    # stats (over filtered view)
    total_value = round(sum(f["value"] for f in filtered), 2)
    total_profit = round(sum(f["profit"] for f in filtered), 2)
    open_value = round(sum(f["value"] for f in filtered if f["status"] == "Open"), 2)
    # years available in the data (for the year dropdown)
    years_avail = sorted({str(f["date"])[:4] for f in forecasts if f.get("date")}, reverse=True)
    return render_template(
        "forecast.html",
        products=products,
        forecasts=filtered,
        total_value=total_value,
        total_profit=total_profit,
        open_value=open_value,
        fc_count=len(filtered),
        fc_open=sum(1 for f in filtered if f["status"] == "Open"),
        search=search,
        date_from=date_from,
        date_to=date_to,
        status_f=status_f,
        month_f=month_f,
        year_f=year_f,
        years_avail=years_avail,
        cur_year=datetime.now().year,
        active="forecast",
    )


@app.route("/forecasts/add", methods=["POST"])
def forecast_add():
    forecasts = db.load_forecasts()
    customer = request.form.get("customer", "").strip()
    date = request.form.get("date") or db.today()
    notes = request.form.get("notes", "").strip()
    status = request.form.get("status") or "Open"

    # collect line items: product_id[], qty[], cost[], srp[], margin[]
    pids = request.form.getlist("product_id")
    qtys = request.form.getlist("qty")
    costs = request.form.getlist("cost")
    srps = request.form.getlist("srp")
    margins = request.form.getlist("margin")
    ssts = request.form.getlist("sst")
    products = db.load_products()

    lines, total_value, total_cost, total_profit = _build_forecast_lines(
        pids, qtys, costs, srps, margins, ssts, products)

    if not customer:
        flash("Customer is required.", "danger")
        return redirect(url_for("forecast_page"), code=303)
    if not lines:
        flash("Add at least one product/item.", "danger")
        return redirect(url_for("forecast_page"), code=303)

    item_summary = "; ".join(f"{l['qty']}x {l['name']}" for l in lines)

    forecasts.append({
        "id": len(forecasts),
        "fc_no": db.next_forecast_no(forecasts),
        "date": date,
        "customer": customer,
        "value": total_value,
        "cost": total_cost,
        "profit": total_profit,
        "items": item_summary,
        "status": status,
        "notes": notes,
        "lines": lines,
    })
    db.save_forecasts(forecasts)

    # if customer is new, add to companies
    companies = db.load_companies()
    if not any(c["company"].lower() == customer.lower() for c in companies):
        companies.append({"company": customer, "industry": "", "contact": "",
                          "email": "", "phone": "", "source": "Forecast",
                          "stage": "New", "notes": "", "address": "", "pipeline": {}})
        db.save_companies(companies)
        flash(f"Forecast saved & new customer '{customer}' added.", "success")
    else:
        flash(f"Forecast {forecasts[-1]['fc_no']} saved.", "success")
    return redirect(url_for("forecast_page"), code=303)


@app.route("/forecasts/delete", methods=["POST"])
def forecast_delete():
    fid = request.form.get("id", type=int)
    forecasts = db.load_forecasts()
    if fid is not None and 0 <= fid < len(forecasts):
        fc_no = forecasts[fid]["fc_no"]
        forecasts.pop(fid)
        db.save_forecasts(forecasts)
        flash(f"Forecast {fc_no} deleted.", "success")
    return redirect(url_for("forecast_page"), code=303)


@app.route("/forecasts/status", methods=["POST"])
def forecast_status():
    fid = request.form.get("id", type=int)
    status = request.form.get("status", "").strip()
    forecasts = db.load_forecasts()
    if fid is not None and 0 <= fid < len(forecasts):
        forecasts[fid]["status"] = status if status else forecasts[fid]["status"]
        db.save_forecasts(forecasts)
        flash(f"Forecast {forecasts[fid]['fc_no']} status updated.", "success")
    return redirect(url_for("forecast_page"), code=303)


@app.route("/forecasts/edit", methods=["POST"])
def forecast_edit():
    fid = request.form.get("id", type=int)
    forecasts = db.load_forecasts()
    if fid is None or not (0 <= fid < len(forecasts)):
        flash("Forecast not found.", "danger")
        return redirect(url_for("forecast_page"), code=303)
    f = forecasts[fid]
    f["date"] = request.form.get("date") or f["date"]
    f["customer"] = request.form.get("customer", "").strip() or f["customer"]
    f["status"] = request.form.get("status") or f["status"]
    f["notes"] = request.form.get("notes", "").strip()

    # rebuild lines (editable cost/srp/margin update the product)
    pids = request.form.getlist("product_id")
    qtys = request.form.getlist("qty")
    costs = request.form.getlist("cost")
    srps = request.form.getlist("srp")
    margins = request.form.getlist("margin")
    ssts = request.form.getlist("sst")
    products = db.load_products()

    lines, total_value, total_cost, total_profit = _build_forecast_lines(
        pids, qtys, costs, srps, margins, ssts, products)
    if lines:
        f["lines"] = lines
        f["value"] = total_value
        f["cost"] = total_cost
        f["profit"] = total_profit
        f["items"] = "; ".join(f"{l['qty']}x {l['name']}" for l in lines)
    db.save_forecasts(forecasts)
    flash(f"Forecast {f['fc_no']} updated.", "success")
    return redirect(url_for("forecast_page"), code=303)


# ---------------- Products ----------------
@app.route("/products")
def products():
    products = db.load_products()
    search = request.args.get("search", "").strip().lower()
    cat = request.args.get("category", "").strip()
    if search:
        products = [p for p in products if search in p["name"].lower()
                    or search in p["pid"].lower()
                    or search in p["category"].lower()]
    if cat:
        products = [p for p in products if p["category"] == cat]
    categories = sorted({p["category"] for p in products if p["category"]})
    return render_template(
        "products.html",
        products=products,
        search=search,
        category=cat,
        categories=categories,
        active="products",
    )


@app.route("/products/add", methods=["POST"])
def product_add():
    products = db.load_products()
    name = request.form.get("name", "").strip()
    if not name:
        flash("Product name is required.", "danger")
        return redirect(url_for("products"))
    try:
        cost = float(request.form.get("cost", 0) or 0)
    except ValueError:
        cost = 0.0
    try:
        margin = float(request.form.get("margin", 0) or 0)
    except ValueError:
        margin = 0.0
    try:
        srp = float(request.form.get("srp", 0) or 0)
    except ValueError:
        srp = 0.0
    # selling price from cost + margin (margin % on selling price)
    selling = cost / (1 - margin / 100) if (0 <= margin < 100) else cost
    profit = selling - cost
    products.append({
        "pid": db.next_product_id(products),
        "name": name,
        "category": request.form.get("category", "").strip(),
        "cost": cost,
        "srp": srp,
        "margin": margin,
        "selling": round(selling, 2),
        "profit": round(profit, 2),
        "notes": request.form.get("notes", "").strip(),
    })
    db.save_products(products)
    flash(f"Product '{name}' added.", "success")
    return redirect(url_for("products"))


@app.route("/products/edit", methods=["POST"])
def product_edit():
    pid = request.form.get("id", type=int)
    products = db.load_products()
    if pid is None or not (0 <= pid < len(products)):
        flash("Product not found.", "danger")
        return redirect(url_for("products"))
    p = products[pid]
    p["name"] = request.form.get("name", "").strip() or p["name"]
    p["category"] = request.form.get("category", "").strip()
    try:
        p["cost"] = float(request.form.get("cost") or 0)
        p["margin"] = float(request.form.get("margin") or 0)
        p["srp"] = float(request.form.get("srp") or 0)
    except ValueError:
        pass
    selling = p["cost"] / (1 - p["margin"] / 100) if (0 <= p["margin"] < 100) else p["cost"]
    p["selling"] = round(selling, 2)
    p["profit"] = round(selling - p["cost"], 2)
    p["notes"] = request.form.get("notes", "").strip()
    db.save_products(products)
    flash(f"Product '{p['name']}' updated.", "success")
    return redirect(url_for("products"))


@app.route("/products/delete", methods=["POST"])
def product_delete():
    pid = request.form.get("id", type=int)
    products = db.load_products()
    if pid is not None and 0 <= pid < len(products):
        name = products[pid]["name"]
        products.pop(pid)
        db.save_products(products)
        flash(f"Product '{name}' deleted.", "success")
    return redirect(url_for("products"))


@app.route("/products/json")
def products_json():
    return jsonify(db.load_products())


# ---------------- Quotation Calculator / Forecast ----------------
@app.route("/calculator")
def calculator():
    products = db.load_products()
    return render_template("calculator.html", products=products, active="calculator")


@app.route("/api/calc")
def api_calc():
    try:
        cost = float(request.args.get("cost", 0) or 0)
        margin = float(request.args.get("margin", 0) or 0)
        qty = int(request.args.get("qty", 1) or 1)
    except ValueError:
        return jsonify({"error": "invalid input"}), 400
    if not (0 <= margin < 100):
        return jsonify({"error": "margin must be 0-99"}), 400
    selling = cost / (1 - margin / 100)
    profit = selling - cost
    return jsonify({
        "unit_selling": round(selling, 2),
        "unit_profit": round(profit, 2),
        "total_selling": round(selling * qty, 2),
        "total_profit": round(profit * qty, 2),
        "total_cost": round(cost * qty, 2),
        "margin": margin,
    })


# ---------------- Reports / Export ----------------
@app.route("/export/quotations.csv")
def export_csv():
    quotes = db.load_quotations()
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Quotation No", "Date", "Time", "Company", "Status", "Amount (RM)", "Created By"])
    for q in quotes:
        w.writerow([q["qno"], q["date"], q["time"], q["company"], q["status"], q["amount"], q["created_by"]])
    data = buf.getvalue()
    return Response(data, mimetype="text/csv",
                    headers={"Content-Disposition": "attachment; filename=quotations.csv"})


@app.route("/export/database")
def export_database():
    path = os.path.join(DATA_DIR, "CRM_Database.xlsx")
    return send_file(path, as_attachment=True,
                     download_name="CRM_Database.xlsx")


@app.route("/api/summary")
def api_summary():
    quotes = db.load_quotations()
    by_status = Counter(q["status"] for q in quotes)
    return jsonify({
        "total": len(quotes),
        "total_value": round(sum(q["amount"] for q in quotes), 2),
        "by_status": by_status,
    })


@app.route("/shutdown", methods=["POST"])
@role_required("superadmin", "admin")
def shutdown():
    """Gracefully stop the CRM Pro server from inside the app."""
    def _stop():
        # give the response a moment to reach the browser, then exit
        time.sleep(1.0)
        os._exit(0)
    threading.Thread(target=_stop, daemon=True).start()
    html = """<html><head><meta charset="utf-8"><title>CRM Pro</title></head>
    <body style="margin:0;font-family:Segoe UI,Arial,sans-serif;background:#0b1f3a;color:#fff;
      display:flex;align-items:center;justify-content:center;height:100vh">
      <div style="text-align:center">
        <h2 style="margin-bottom:.4rem">CRM Pro has shut down.</h2>
        <p style="opacity:.85">You can now close this tab and the server window.</p>
      </div>
    </body></html>"""
    return html


# ---------------- Leads (contacts) ----------------
@app.route("/leads")
def leads_page():
    all_leads = db.load_leads()
    search = request.args.get("search", "").strip().lower()
    status = request.args.get("status", "").strip()
    source = request.args.get("source", "").strip()
    month = request.args.get("month", "").strip()
    year = request.args.get("year", "").strip()

    leads = all_leads
    if search:
        leads = [l for l in leads if search in l["name"].lower()
                 or search in l["company"].lower()
                 or search in l["email"].lower()
                 or search in l["phone"].lower()]
    if status:
        leads = [l for l in leads if l["status"] == status]
    if source:
        leads = [l for l in leads if l["source"] == source]
    if month:
        mm = int(month)
        leads = [l for l in leads if str(l["date"])[:7].endswith(f"-{mm:02d}")]
    if year:
        leads = [l for l in leads if str(l["date"])[:4] == str(year)]

    # stats (over filtered view)
    total_leads = len(leads)
    active = sum(1 for l in leads if l["status"] in ("New", "Contacted"))
    quoted = sum(1 for l in leads if l["status"] == "Quotation In Progress")
    won = sum(1 for l in leads if l["status"] == "Deal Won")

    years_avail = sorted({str(l["date"])[:4] for l in all_leads if l.get("date")}, reverse=True)
    sources_avail = sorted({l["source"] for l in all_leads if l.get("source")})
    return render_template("leads.html", leads=leads, statuses=db.LEAD_STATUSES,
                           search=search, status=status, source=source,
                           month=month, year=year,
                           years_avail=years_avail, sources_avail=sources_avail,
                           total_leads=total_leads, active_count=active,
                           quoted_count=quoted, won_count=won,
                           active="leads")


@app.route("/leads/add", methods=["POST"])
def lead_add():
    leads = db.load_leads()
    name = request.form.get("name", "").strip()
    if not name:
        flash("Lead name is required.", "danger")
        return redirect(url_for("leads_page"), code=303)
    status = request.form.get("status") or "New"
    date = request.form.get("date") or db.today()
    leads.append({
        "id": len(leads),
        "name": name,
        "company": request.form.get("company", "").strip(),
        "phone": request.form.get("phone", "").strip(),
        "email": request.form.get("email", "").strip(),
        "source": request.form.get("source", "").strip(),
        "status": status,
        "date": date,
        "notes": request.form.get("notes", "").strip(),
    })
    db.save_leads(leads)
    flash(f"Lead '{name}' added.", "success")
    return redirect(url_for("leads_page"), code=303)


@app.route("/leads/edit", methods=["POST"])
def lead_edit():
    lid = request.form.get("id", type=int)
    leads = db.load_leads()
    if lid is None or not (0 <= lid < len(leads)):
        flash("Lead not found.", "danger")
        return redirect(url_for("leads_page"), code=303)
    l = leads[lid]
    l["name"] = request.form.get("name", "").strip() or l["name"]
    l["company"] = request.form.get("company", "").strip()
    l["phone"] = request.form.get("phone", "").strip()
    l["email"] = request.form.get("email", "").strip()
    l["source"] = request.form.get("source", "").strip()
    l["status"] = request.form.get("status") or l["status"]
    l["date"] = request.form.get("date") or l["date"]
    l["notes"] = request.form.get("notes", "").strip()
    db.save_leads(leads)
    flash(f"Lead '{l['name']}' updated.", "success")
    return redirect(url_for("leads_page"), code=303)


@app.route("/leads/status", methods=["POST"])
def lead_status():
    lid = request.form.get("id", type=int)
    leads = db.load_leads()
    if lid is None or not (0 <= lid < len(leads)):
        return redirect(url_for("leads_page"), code=303)
    l = leads[lid]
    new_status = request.form.get("status", "").strip()
    l["status"] = new_status or l["status"]
    db.save_leads(leads)
    flash(f"Lead '{l['name']}' status updated.", "success")
    return redirect(url_for("leads_page"), code=303)


@app.route("/leads/delete", methods=["POST"])
def lead_delete():
    lid = request.form.get("id", type=int)
    leads = db.load_leads()
    if lid is not None and 0 <= lid < len(leads):
        name = leads[lid]["name"]
        leads.pop(lid)
        db.save_leads(leads)
        flash(f"Lead '{name}' deleted.", "success")
    return redirect(url_for("leads_page"), code=303)


# ---------------- Customer Board (Kanban) ----------------
BOARD_STAGES = [
    ("New", "New"), ("Contacted", "Contacted"), ("Quoted", "Quoted"),
    ("Negotiating", "Negotiating"), ("Won", "Won"), ("Lost", "Deal Lost"),
]


@app.route("/board")
def customer_board():
    companies = db.load_companies()
    quotes = db.load_quotations()
    search = request.args.get("search", "").strip().lower()
    # per-company pipeline value from in-progress quotes
    inprogress = {}
    for q in quotes:
        if q["status"] == "Quotation In Progress":
            inprogress[str(q["company"]).lower()] = inprogress.get(str(q["company"]).lower(), 0.0) + q["amount"]
    cards = []
    for c in companies:
        stage = c.get("stage") or "New"
        if stage not in [s[0] for s in BOARD_STAGES]:
            stage = "New"
        # filter by search
        if search and not (search in str(c["company"]).lower()
                           or search in str(c.get("contact", "")).lower()
                           or search in str(c.get("email", "")).lower()
                           or search in str(c.get("source", "")).lower()):
            continue
        cards.append({
            "id": c["id"],
            "company": c["company"],
            "contact": c.get("contact", ""),
            "phone": c.get("phone", ""),
            "email": c.get("email", ""),
            "source": c.get("source", ""),
            "stage": stage,
            "value": round(inprogress.get(str(c["company"]).lower(), 0.0), 2),
        })
    return render_template("board.html", cards=cards, stages=BOARD_STAGES,
                           search=search, active="board")


@app.route("/board/move", methods=["POST"])
def board_move():
    """AJAX: move a card to a new stage."""
    cid = request.form.get("id", type=int)
    stage = request.form.get("stage", "").strip()
    companies = db.load_companies()
    valid_stages = [s[0] for s in BOARD_STAGES]
    if cid is not None and 0 <= cid < len(companies) and stage in valid_stages:
        companies[cid]["stage"] = stage
        db.save_companies(companies)
        return jsonify({"ok": True})
    return jsonify({"ok": False}), 400


# ---------------- Settings / UI Themes ----------------
THEMES = [
    {"id": "blue", "name": "Modern Blue", "desc": "Classic navy & blue — clean and professional.",
     "nav": "#0b1f3a", "nav2": "#1b3a6b", "accent": "#4f7cff"},
    {"id": "emerald", "name": "Emerald", "desc": "Fresh teal & green — calm and modern.",
     "nav": "#064e3b", "nav2": "#0f766e", "accent": "#10b981"},
    {"id": "sunset", "name": "Sunset Amber", "desc": "Warm amber & orange — bold and vibrant.",
     "nav": "#431407", "nav2": "#b45309", "accent": "#f59e0b"},
]


@app.route("/settings")
@role_required("superadmin", "admin")
def settings_page():
    theme = db.current_theme()
    users = db.load_users()
    return render_template("settings.html", themes=THEMES, current_theme=theme,
                           monthly_target=db.monthly_target(), users=users,
                           roles=db.ROLES, active="settings")


@app.route("/settings/theme", methods=["POST"])
@role_required("superadmin", "admin")
def settings_theme():
    theme = request.form.get("theme", "").strip()
    if theme in ("blue", "emerald", "sunset"):
        db.set_setting("ui_theme", theme)
        flash("UI theme updated.", "success")
    return redirect(url_for("settings_page"), code=303)


@app.route("/settings/target", methods=["POST"])
@role_required("superadmin", "admin")
def settings_target():
    try:
        val = float(request.form.get("monthly_target"))
    except (TypeError, ValueError):
        val = db.monthly_target()
    db.set_setting("monthly_target", val)
    flash("Monthly target updated.", "success")
    return redirect(url_for("settings_page"), code=303)


@app.route("/settings/users/add", methods=["POST"])
@role_required("superadmin", "admin")
def user_add():
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")
    name = request.form.get("name", "").strip()
    role = request.form.get("role", "user").strip()
    if role not in db.ROLES:
        role = "user"
    if not username or not password:
        flash("Username and password are required.", "danger")
        return redirect(url_for("settings_page"), code=303)
    if db.find_user(username):
        flash(f"Username '{username}' already exists.", "danger")
        return redirect(url_for("settings_page"), code=303)
    from werkzeug.security import generate_password_hash
    users = db.load_users()
    users.append({
        "username": username, "password": generate_password_hash(password),
        "name": name, "role": role, "active": 1,
    })
    db.save_users(users)
    flash(f"User '{username}' created.", "success")
    return redirect(url_for("settings_page"), code=303)


@app.route("/settings/users/edit", methods=["POST"])
@role_required("superadmin", "admin")
def user_edit():
    orig = request.form.get("orig_username", "").strip()
    username = request.form.get("username", "").strip()
    name = request.form.get("name", "").strip()
    role = request.form.get("role", "user").strip()
    active = 1 if request.form.get("active") in ("1", "on", "yes") else 0
    if role not in db.ROLES:
        role = "user"
    users = db.load_users()
    target = next((u for u in users if u["username"] == orig), None)
    if not target:
        flash("User not found.", "danger")
        return redirect(url_for("settings_page"), code=303)
    # Prevent removing your own superadmin privileges accidentally.
    if orig == session.get("user") and target["role"] == "superadmin" and role != "superadmin":
        flash("You cannot change your own superadmin role.", "danger")
        return redirect(url_for("settings_page"), code=303)
    # Prevent deactivating your own account.
    if orig == session.get("user") and not active:
        flash("You cannot deactivate your own account.", "danger")
        return redirect(url_for("settings_page"), code=303)
    new_password = request.form.get("password", "")
    from werkzeug.security import generate_password_hash
    if new_password:
        target["password"] = generate_password_hash(new_password)
    target["username"] = username
    target["name"] = name
    target["role"] = role
    target["active"] = active
    # If editing the logged-in user, update the session.
    if orig == session.get("user"):
        session["name"] = name
        session["role"] = role
    db.save_users(users)
    flash(f"User '{username}' updated.", "success")
    return redirect(url_for("settings_page"), code=303)


@app.route("/settings/users/delete", methods=["POST"])
@role_required("superadmin", "admin")
def user_delete():
    username = request.form.get("username", "").strip()
    if username == session.get("user"):
        flash("You cannot delete your own account.", "danger")
        return redirect(url_for("settings_page"), code=303)
    users = db.load_users()
    users = [u for u in users if u["username"] != username]
    # Keep at least one superadmin.
    if not any(u["role"] == "superadmin" and u["active"] for u in users):
        flash("Cannot delete: at least one active superadmin is required.", "danger")
        return redirect(url_for("settings_page"), code=303)
    db.save_users(users)
    flash(f"User '{username}' deleted.", "success")
    return redirect(url_for("settings_page"), code=303)


@app.errorhandler(500)
def handle_500(e):
    """Show a friendly error page instead of a bare Internal Server Error."""
    return render_template("error.html", code=500,
                           message="Something went wrong. Please go back and try again."), 500


@app.errorhandler(404)
def handle_404(e):
    return render_template("error.html", code=404,
                           message="The page you're looking for was not found."), 404


def open_browser():
    # Give the server a moment to start, then open the app in the default browser
    try:
        threading.Timer(1.5, lambda: webbrowser.open("http://127.0.0.1:5000")).start()
    except Exception:
        pass


if __name__ == "__main__":
    print("=" * 55)
    print("  CRM Pro  |  http://127.0.0.1:5000")
    print("  Excel DB  :", db.DB_PATH)
    print("  Opening your browser automatically...")
    print("  Press Ctrl+C to stop.")
    print("=" * 55)
    open_browser()
    app.run(host="127.0.0.1", port=5000, debug=False)
