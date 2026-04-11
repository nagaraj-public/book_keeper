import os
from datetime import date, datetime
from flask import (
    Blueprint,
    jsonify,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    current_app,
    send_file,
)
from werkzeug.utils import secure_filename
from models import db, Client, Income, Expense, HourEntry, MonthlyBilling, FeeDefaults
from excel_import import import_clients_from_excel

bp = Blueprint("main", __name__)


def register_routes(app):
    app.register_blueprint(bp)


# --------------- Pages ---------------

@bp.route("/")
def dashboard():
    year = request.args.get("year", date.today().year, type=int)
    month = request.args.get("month", date.today().month, type=int)

    incomes = Income.query.filter(
        db.extract("year", Income.date) == year,
        db.extract("month", Income.date) == month,
    ).all()
    expenses = Expense.query.filter(
        db.extract("year", Expense.date) == year,
        db.extract("month", Expense.date) == month,
    ).all()
    hours = HourEntry.query.filter(
        db.extract("year", HourEntry.date) == year,
        db.extract("month", HourEntry.date) == month,
    ).all()

    total_income = sum(i.total for i in incomes)
    total_expenses = sum(e.total for e in expenses)
    total_hours = sum(h.hours for h in hours)
    total_btw_income = sum(i.btw_amount for i in incomes)
    total_btw_expenses = sum(e.btw_amount for e in expenses)
    profit = total_income - total_expenses

    return render_template(
        "dashboard.html",
        year=year,
        month=month,
        total_income=total_income,
        total_expenses=total_expenses,
        total_hours=total_hours,
        total_btw_income=total_btw_income,
        total_btw_expenses=total_btw_expenses,
        profit=profit,
        incomes=incomes,
        expenses=expenses,
        hours=hours,
    )


@bp.route("/clients")
def clients_page():
    tab = request.args.get("tab", "all")
    all_clients = Client.query.order_by(Client.name).all()

    def _filter(stype=None, status=None, mode=None):
        q = Client.query.order_by(Client.name)
        if stype:
            q = q.filter(Client.student_type == stype)
        if status:
            q = q.filter(Client.status == status)
        if mode:
            q = q.filter(Client.delivery_mode == mode)
        return q.all()

    clients_by_tab = {
        "all":          all_clients,
        "adults":       _filter(stype="adult"),
        "kids":         _filter(stype="child"),
        "individual":   _filter(stype="individual"),
        "online":       _filter(stype="online"),
        "active":       _filter(status="active"),
        "discontinued": _filter(status="discontinued"),
        "pending":      _filter(status="pending"),
        "mode_offline": _filter(mode="offline"),
        "mode_online":  _filter(mode="online"),
    }
    counts = {k: len(v) for k, v in clients_by_tab.items()}

    return render_template(
        "clients.html",
        clients=clients_by_tab.get(tab, all_clients),
        tab=tab,
        counts=counts,
    )


@bp.route("/income")
def income_page():
    incomes = Income.query.order_by(Income.date.desc()).all()
    clients = Client.query.order_by(Client.name).all()
    return render_template("income.html", incomes=incomes, clients=clients)


@bp.route("/expenses")
def expenses_page():
    expenses = Expense.query.order_by(Expense.date.desc()).all()
    return render_template("expenses.html", expenses=expenses)


@bp.route("/hours")
def hours_page():
    hours = HourEntry.query.order_by(HourEntry.date.desc()).all()
    clients = Client.query.order_by(Client.name).all()
    return render_template("hours.html", hours=hours, clients=clients)


@bp.route("/report")
def report_page():
    year = request.args.get("year", date.today().year, type=int)
    return render_template("report.html", year=year, report=generate_year_report(year))


# --------------- Client CRUD ---------------

@bp.route("/clients/add", methods=["POST"])
def add_client():
    client = Client(
        name=request.form.get("name", "").strip(),
        company_name=request.form.get("company_name", "").strip(),
        kvk_number=request.form.get("kvk_number", "").strip(),
        btw_number=request.form.get("btw_number", "").strip(),
        email=request.form.get("email", "").strip(),
        phone=request.form.get("phone", "").strip(),
        address=request.form.get("address", "").strip(),
        notes=request.form.get("notes", "").strip(),
        student_type=request.form.get("student_type", "adult"),
        status=request.form.get("status", "active"),
        delivery_mode=request.form.get("delivery_mode", "offline"),
        monthly_fee=float(request.form.get("monthly_fee")) if request.form.get("monthly_fee") else None,
        default_btw_rate=float(request.form.get("default_btw_rate", 21)),
    )
    if not client.name:
        flash("Client name is required.", "danger")
        return redirect(url_for("main.clients_page"))
    db.session.add(client)
    db.session.commit()
    flash(f"Client '{client.name}' added.", "success")
    return redirect(url_for("main.clients_page"))


@bp.route("/clients/edit/<int:client_id>", methods=["POST"])
def edit_client(client_id):
    client = Client.query.get_or_404(client_id)
    client.name = request.form.get("name", client.name).strip()
    client.company_name = request.form.get("company_name", "").strip()
    client.kvk_number = request.form.get("kvk_number", "").strip()
    client.btw_number = request.form.get("btw_number", "").strip()
    client.email = request.form.get("email", "").strip()
    client.phone = request.form.get("phone", "").strip()
    client.address = request.form.get("address", "").strip()
    client.notes = request.form.get("notes", "").strip()
    client.student_type = request.form.get("student_type", client.student_type)
    client.status = request.form.get("status", client.status)
    client.delivery_mode = request.form.get("delivery_mode", client.delivery_mode or "offline")
    client.monthly_fee = float(request.form.get("monthly_fee")) if request.form.get("monthly_fee") else None
    client.default_btw_rate = float(request.form.get("default_btw_rate", client.default_btw_rate or 21))
    db.session.commit()
    flash(f"Client '{client.name}' updated.", "success")
    return redirect(url_for("main.clients_page"))


@bp.route("/clients/delete/<int:client_id>", methods=["POST"])
def delete_client(client_id):
    client = Client.query.get_or_404(client_id)
    db.session.delete(client)
    db.session.commit()
    flash(f"Client deleted.", "success")
    return redirect(url_for("main.clients_page"))


@bp.route("/clients/import", methods=["POST"])
def import_clients():
    if "file" not in request.files:
        flash("No file selected.", "danger")
        return redirect(url_for("main.clients_page"))

    file = request.files["file"]
    if file.filename == "":
        flash("No file selected.", "danger")
        return redirect(url_for("main.clients_page"))

    if not file.filename.lower().endswith((".xlsx", ".xls")):
        flash("Please upload an Excel file (.xlsx or .xls).", "danger")
        return redirect(url_for("main.clients_page"))

    filename = secure_filename(file.filename)
    filepath = os.path.join(current_app.config["UPLOAD_FOLDER"], filename)
    file.save(filepath)

    try:
        count = import_clients_from_excel(filepath)
        flash(f"Successfully imported {count} clients from Excel.", "success")
    except Exception as e:
        flash(f"Import error: {str(e)}", "danger")
    finally:
        if os.path.exists(filepath):
            os.remove(filepath)

    return redirect(url_for("main.clients_page"))


# --------------- Income CRUD ---------------

@bp.route("/income/add", methods=["POST"])
def add_income():
    amount = float(request.form.get("amount", 0))
    btw_rate = float(request.form.get("btw_rate", 21))
    btw_amount = round(amount * btw_rate / 100, 2)
    total = round(amount + btw_amount, 2)

    income = Income(
        client_id=request.form.get("client_id") or None,
        date=datetime.strptime(request.form["date"], "%Y-%m-%d").date()
        if request.form.get("date")
        else date.today(),
        description=request.form.get("description", "").strip(),
        amount=amount,
        btw_rate=btw_rate,
        btw_amount=btw_amount,
        total=total,
        category=request.form.get("category", "Services").strip(),
        invoice_number=request.form.get("invoice_number", "").strip(),
    )
    db.session.add(income)
    db.session.commit()
    flash("Income entry added.", "success")
    return redirect(url_for("main.income_page"))


@bp.route("/income/edit/<int:income_id>", methods=["POST"])
def edit_income(income_id):
    income = Income.query.get_or_404(income_id)
    amount = float(request.form.get("amount", income.amount))
    btw_rate = float(request.form.get("btw_rate", income.btw_rate))
    btw_amount = round(amount * btw_rate / 100, 2)

    income.client_id = request.form.get("client_id") or None
    income.date = (
        datetime.strptime(request.form["date"], "%Y-%m-%d").date()
        if request.form.get("date")
        else income.date
    )
    income.description = request.form.get("description", income.description).strip()
    income.amount = amount
    income.btw_rate = btw_rate
    income.btw_amount = btw_amount
    income.total = round(amount + btw_amount, 2)
    income.category = request.form.get("category", income.category).strip()
    income.invoice_number = request.form.get("invoice_number", "").strip()
    db.session.commit()
    flash("Income entry updated.", "success")
    return redirect(url_for("main.income_page"))


@bp.route("/income/delete/<int:income_id>", methods=["POST"])
def delete_income(income_id):
    income = Income.query.get_or_404(income_id)
    db.session.delete(income)
    db.session.commit()
    flash("Income entry deleted.", "success")
    return redirect(url_for("main.income_page"))


# --------------- Expense CRUD ---------------

@bp.route("/expenses/add", methods=["POST"])
def add_expense():
    amount = float(request.form.get("amount", 0))
    btw_rate = float(request.form.get("btw_rate", 21))
    btw_amount = round(amount * btw_rate / 100, 2)
    total = round(amount + btw_amount, 2)

    expense = Expense(
        date=datetime.strptime(request.form["date"], "%Y-%m-%d").date()
        if request.form.get("date")
        else date.today(),
        description=request.form.get("description", "").strip(),
        amount=amount,
        btw_rate=btw_rate,
        btw_amount=btw_amount,
        total=total,
        category=request.form.get("category", "General").strip(),
        receipt_ref=request.form.get("receipt_ref", "").strip(),
    )
    db.session.add(expense)
    db.session.commit()
    flash("Expense entry added.", "success")
    return redirect(url_for("main.expenses_page"))


@bp.route("/expenses/edit/<int:expense_id>", methods=["POST"])
def edit_expense(expense_id):
    expense = Expense.query.get_or_404(expense_id)
    amount = float(request.form.get("amount", expense.amount))
    btw_rate = float(request.form.get("btw_rate", expense.btw_rate))
    btw_amount = round(amount * btw_rate / 100, 2)

    expense.date = (
        datetime.strptime(request.form["date"], "%Y-%m-%d").date()
        if request.form.get("date")
        else expense.date
    )
    expense.description = request.form.get(
        "description", expense.description
    ).strip()
    expense.amount = amount
    expense.btw_rate = btw_rate
    expense.btw_amount = btw_amount
    expense.total = round(amount + btw_amount, 2)
    expense.category = request.form.get("category", expense.category).strip()
    expense.receipt_ref = request.form.get("receipt_ref", "").strip()
    db.session.commit()
    flash("Expense entry updated.", "success")
    return redirect(url_for("main.expenses_page"))


@bp.route("/expenses/delete/<int:expense_id>", methods=["POST"])
def delete_expense(expense_id):
    expense = Expense.query.get_or_404(expense_id)
    db.session.delete(expense)
    db.session.commit()
    flash("Expense entry deleted.", "success")
    return redirect(url_for("main.expenses_page"))


# --------------- Hours CRUD ---------------

@bp.route("/hours/add", methods=["POST"])
def add_hours():
    entry = HourEntry(
        client_id=request.form.get("client_id") or None,
        date=datetime.strptime(request.form["date"], "%Y-%m-%d").date()
        if request.form.get("date")
        else date.today(),
        hours=float(request.form.get("hours", 0)),
        rate=float(request.form.get("rate", 0)),
        description=request.form.get("description", "").strip(),
        invoiced=bool(request.form.get("invoiced")),
    )
    db.session.add(entry)
    db.session.commit()
    flash("Hours entry added.", "success")
    return redirect(url_for("main.hours_page"))


@bp.route("/hours/edit/<int:entry_id>", methods=["POST"])
def edit_hours(entry_id):
    entry = HourEntry.query.get_or_404(entry_id)
    entry.client_id = request.form.get("client_id") or None
    entry.date = (
        datetime.strptime(request.form["date"], "%Y-%m-%d").date()
        if request.form.get("date")
        else entry.date
    )
    entry.hours = float(request.form.get("hours", entry.hours))
    entry.rate = float(request.form.get("rate", entry.rate))
    entry.description = request.form.get("description", entry.description).strip()
    entry.invoiced = bool(request.form.get("invoiced"))
    db.session.commit()
    flash("Hours entry updated.", "success")
    return redirect(url_for("main.hours_page"))


@bp.route("/hours/delete/<int:entry_id>", methods=["POST"])
def delete_hours(entry_id):
    entry = HourEntry.query.get_or_404(entry_id)
    db.session.delete(entry)
    db.session.commit()
    flash("Hours entry deleted.", "success")
    return redirect(url_for("main.hours_page"))


# --------------- Monthly Billing ---------------

_MONTH_NAMES = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December",
}

_CATEGORY_MAP = {
    "adult": "Dance Classes - Adults",
    "child": "Dance Classes - Kids",
    "individual": "Dance Classes - Individual",
    "online": "Dance Classes - Online",
}


def _billing_months():
    """Returns (month, year, label) for current month + past 12 months."""
    today = date.today()
    result = []
    y, m = today.year, today.month
    for _ in range(13):
        result.append((m, y, f"{_MONTH_NAMES[m]} {y}"))
        m -= 1
        if m < 1:
            m = 12
            y -= 1
    return result


def _next_invoice_number(month, year):
    count = MonthlyBilling.query.filter_by(month=month, year=year).count()
    return f"NAT-{year}-{month:02d}-{count + 1:03d}"


@bp.route("/billing")
def billing_page():
    my = request.args.get("my", f"{date.today().month}-{date.today().year}")
    try:
        parts = my.split("-")
        month, year = int(parts[0]), int(parts[1])
    except (ValueError, IndexError):
        month, year = date.today().month, date.today().year

    billings = (
        MonthlyBilling.query.filter_by(year=year, month=month)
        .join(Client)
        .order_by(Client.name)
        .all()
    )

    active_count = Client.query.filter_by(status="active").count()
    total_expected = sum(b.amount for b in billings)
    total_paid = sum(b.amount for b in billings if b.status == "paid")
    paid_count = sum(1 for b in billings if b.status == "paid")

    return render_template(
        "billing.html",
        billings=billings,
        year=year,
        month=month,
        month_name=_MONTH_NAMES.get(month, ""),
        active_count=active_count,
        billed_count=len(billings),
        total_expected=total_expected,
        total_paid=total_paid,
        paid_count=paid_count,
        months=_billing_months(),
        current_my=my,
    )


@bp.route("/billing/generate", methods=["POST"])
def billing_generate():
    my = request.form.get("my", f"{date.today().month}-{date.today().year}")
    parts = my.split("-")
    month, year = int(parts[0]), int(parts[1])

    active_clients = Client.query.filter_by(status="active").order_by(Client.name).all()
    defaults = FeeDefaults.get()
    fee_map = {
        "adult":      (defaults.adult_fee,      defaults.adult_btw),
        "child":      (defaults.child_fee,      defaults.child_btw),
        "individual": (defaults.individual_fee, defaults.individual_btw),
        "online":     (defaults.online_fee,     defaults.online_btw),
    }
    created = 0
    skipped = 0
    for c in active_clients:
        existing = MonthlyBilling.query.filter_by(client_id=c.id, month=month, year=year).first()
        if existing:
            skipped += 1
            continue
        def_fee, def_btw = fee_map.get(c.student_type, (0.0, 21.0))
        billing = MonthlyBilling(
            client_id=c.id,
            month=month,
            year=year,
            invoice_number=_next_invoice_number(month, year),
            amount=c.monthly_fee if c.monthly_fee is not None else def_fee,
            btw_rate=def_btw if def_btw is not None else 21.0,
        )
        db.session.add(billing)
        db.session.flush()
        created += 1

    db.session.commit()
    msg = f"Generated {created} billing entries for {_MONTH_NAMES.get(month, '')} {year}."
    if skipped:
        msg += f" ({skipped} already existed)"
    flash(msg, "success")
    return redirect(url_for("main.billing_page", my=my))


@bp.route("/billing/<int:billing_id>/paid", methods=["POST"])
def billing_mark_paid(billing_id):
    b = MonthlyBilling.query.get_or_404(billing_id)
    if b.status == "paid":
        flash("Already marked as paid.", "info")
        return redirect(url_for("main.billing_page", my=f"{b.month}-{b.year}"))

    btw_amount = round(b.amount * b.btw_rate / 100, 2)
    total = round(b.amount + btw_amount, 2)
    category = _CATEGORY_MAP.get(b.client.student_type, "Dance Classes")

    income = Income(
        client_id=b.client_id,
        date=date.today(),
        description=f"Dance fee {_MONTH_NAMES.get(b.month, '')} {b.year}",
        amount=b.amount,
        btw_rate=b.btw_rate,
        btw_amount=btw_amount,
        total=total,
        category=category,
        invoice_number=b.invoice_number,
    )
    db.session.add(income)
    db.session.flush()

    b.status = "paid"
    b.paid_date = date.today()
    b.income_id = income.id
    db.session.commit()
    flash(f"Payment recorded for {b.client.name}.", "success")
    return redirect(url_for("main.billing_page", my=f"{b.month}-{b.year}"))


@bp.route("/billing/<int:billing_id>/unpaid", methods=["POST"])
def billing_mark_unpaid(billing_id):
    b = MonthlyBilling.query.get_or_404(billing_id)
    if b.income_id:
        inc = Income.query.get(b.income_id)
        if inc:
            db.session.delete(inc)
    b.status = "unpaid"
    b.paid_date = None
    b.income_id = None
    db.session.commit()
    flash(f"Marked as unpaid for {b.client.name}.", "success")
    return redirect(url_for("main.billing_page", my=f"{b.month}-{b.year}"))


@bp.route("/billing/<int:billing_id>/edit", methods=["POST"])
def billing_edit(billing_id):
    b = MonthlyBilling.query.get_or_404(billing_id)
    b.amount = float(request.form.get("amount", b.amount))
    b.btw_rate = float(request.form.get("btw_rate", b.btw_rate))
    b.notes = request.form.get("notes", b.notes or "").strip()
    db.session.commit()
    flash(f"Updated billing for {b.client.name}.", "success")
    return redirect(url_for("main.billing_page", my=f"{b.month}-{b.year}"))


@bp.route("/billing/<int:billing_id>/delete", methods=["POST"])
def billing_delete(billing_id):
    b = MonthlyBilling.query.get_or_404(billing_id)
    my = f"{b.month}-{b.year}"
    if b.income_id:
        inc = Income.query.get(b.income_id)
        if inc:
            db.session.delete(inc)
    db.session.delete(b)
    db.session.commit()
    flash("Billing entry removed.", "success")
    return redirect(url_for("main.billing_page", my=my))


@bp.route("/billing/bulk-delete", methods=["POST"])
def billing_bulk_delete():
    ids = request.form.getlist("ids")
    my = request.form.get("my", f"{date.today().month}-{date.today().year}")
    if not ids:
        flash("No entries selected.", "warning")
        return redirect(url_for("main.billing_page", my=my))
    deleted = 0
    for bid in ids:
        b = MonthlyBilling.query.get(int(bid))
        if not b:
            continue
        if b.income_id:
            inc = Income.query.get(b.income_id)
            if inc:
                db.session.delete(inc)
        db.session.delete(b)
        deleted += 1
    db.session.commit()
    flash(f"Deleted {deleted} billing entr{'y' if deleted == 1 else 'ies'}.", "success")
    return redirect(url_for("main.billing_page", my=my))


@bp.route("/billing/<int:billing_id>/invoice")
def billing_invoice(billing_id):
    b = MonthlyBilling.query.get_or_404(billing_id)
    if b.status != "paid":
        flash("Invoice can only be downloaded for paid entries.", "warning")
        return redirect(url_for("main.billing_page", my=f"{b.month}-{b.year}"))
    from invoice_pdf import generate_invoice_pdf
    pdf_buf = generate_invoice_pdf(b)
    filename = f"Factuur_{b.invoice_number}_{b.client.name.replace(' ', '_')}.pdf"
    return send_file(pdf_buf, mimetype="application/pdf",
                     as_attachment=True, download_name=filename)


# --------------- Fee Settings ---------------

@bp.route("/settings", methods=["GET", "POST"])
def settings_page():
    defaults = FeeDefaults.get()
    if request.method == "POST":
        defaults.adult_fee      = float(request.form.get("adult_fee",      defaults.adult_fee))
        defaults.adult_btw      = float(request.form.get("adult_btw",      defaults.adult_btw))
        defaults.child_fee      = float(request.form.get("child_fee",      defaults.child_fee))
        defaults.child_btw      = float(request.form.get("child_btw",      defaults.child_btw))
        defaults.individual_fee = float(request.form.get("individual_fee", defaults.individual_fee))
        defaults.individual_btw = float(request.form.get("individual_btw", defaults.individual_btw))
        defaults.online_fee     = float(request.form.get("online_fee",     defaults.online_fee))
        defaults.online_btw     = float(request.form.get("online_btw",     defaults.online_btw))
        db.session.commit()
        flash("Default fees saved.", "success")
        return redirect(url_for("main.settings_page"))
    return render_template("settings.html", d=defaults)


# --------------- Year-End Report ---------------

def generate_year_report(year):
    incomes = Income.query.filter(db.extract("year", Income.date) == year).all()
    expenses = Expense.query.filter(db.extract("year", Expense.date) == year).all()
    hours = HourEntry.query.filter(db.extract("year", HourEntry.date) == year).all()

    monthly = {}
    for m in range(1, 13):
        m_incomes = [i for i in incomes if i.date.month == m]
        m_expenses = [e for e in expenses if e.date.month == m]
        m_hours = [h for h in hours if h.date.month == m]
        monthly[m] = {
            "income": sum(i.total for i in m_incomes),
            "income_excl": sum(i.amount for i in m_incomes),
            "expenses": sum(e.total for e in m_expenses),
            "expenses_excl": sum(e.amount for e in m_expenses),
            "btw_income": sum(i.btw_amount for i in m_incomes),
            "btw_expenses": sum(e.btw_amount for e in m_expenses),
            "hours": sum(h.hours for h in m_hours),
            "hour_earnings": sum(h.hours * h.rate for h in m_hours),
        }

    # Income by category
    income_categories = {}
    for i in incomes:
        cat = i.category or "Uncategorized"
        income_categories[cat] = income_categories.get(cat, 0) + i.total

    # Expense by category
    expense_categories = {}
    for e in expenses:
        cat = e.category or "Uncategorized"
        expense_categories[cat] = expense_categories.get(cat, 0) + e.total

    # Client breakdown
    client_income = {}
    for i in incomes:
        name = i.client.name if i.client else "No Client"
        client_income[name] = client_income.get(name, 0) + i.total

    total_income = sum(i.total for i in incomes)
    total_income_excl = sum(i.amount for i in incomes)
    total_expenses = sum(e.total for e in expenses)
    total_expenses_excl = sum(e.amount for e in expenses)
    total_btw_income = sum(i.btw_amount for i in incomes)
    total_btw_expenses = sum(e.btw_amount for e in expenses)
    total_hours = sum(h.hours for h in hours)

    return {
        "year": year,
        "monthly": monthly,
        "income_categories": income_categories,
        "expense_categories": expense_categories,
        "client_income": client_income,
        "total_income": total_income,
        "total_income_excl": total_income_excl,
        "total_expenses": total_expenses,
        "total_expenses_excl": total_expenses_excl,
        "total_btw_income": total_btw_income,
        "total_btw_expenses": total_btw_expenses,
        "btw_to_pay": total_btw_income - total_btw_expenses,
        "profit": total_income - total_expenses,
        "profit_excl": total_income_excl - total_expenses_excl,
        "total_hours": total_hours,
    }
