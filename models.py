from datetime import datetime, date
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class Client(db.Model):
    __tablename__ = "clients"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    company_name = db.Column(db.String(200))
    kvk_number = db.Column(db.String(20))
    btw_number = db.Column(db.String(20))
    email = db.Column(db.String(200))
    phone = db.Column(db.String(30))
    address = db.Column(db.Text)
    notes = db.Column(db.Text)
    # adult | child | individual | online
    student_type = db.Column(db.String(20), default="adult")
    # active | discontinued | pending
    status = db.Column(db.String(20), default="active")
    # offline | online
    delivery_mode = db.Column(db.String(10), default="offline")
    monthly_fee = db.Column(db.Float, nullable=True)
    default_btw_rate = db.Column(db.Float, default=21.0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    incomes = db.relationship("Income", backref="client", lazy=True)
    hours = db.relationship("HourEntry", backref="client", lazy=True)
    billings = db.relationship("MonthlyBilling", backref="client", lazy=True)

    def to_dict(self):
        return {
            "id": self.id,
            "name": self.name,
            "company_name": self.company_name or "",
            "kvk_number": self.kvk_number or "",
            "btw_number": self.btw_number or "",
            "email": self.email or "",
            "phone": self.phone or "",
            "address": self.address or "",
            "notes": self.notes or "",
            "student_type": self.student_type or "adult",
            "status": self.status or "active",
            "delivery_mode": self.delivery_mode or "offline",
            "monthly_fee": self.monthly_fee,
            "default_btw_rate": self.default_btw_rate or 21.0,
            "created_at": self.created_at.isoformat() if self.created_at else "",
        }


class Income(db.Model):
    __tablename__ = "incomes"

    id = db.Column(db.Integer, primary_key=True)
    client_id = db.Column(db.Integer, db.ForeignKey("clients.id"), nullable=True)
    date = db.Column(db.Date, nullable=False, default=date.today)
    description = db.Column(db.String(500), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    btw_rate = db.Column(db.Float, default=21.0)
    btw_amount = db.Column(db.Float, default=0.0)
    total = db.Column(db.Float, nullable=False)
    category = db.Column(db.String(100), default="Services")
    invoice_number = db.Column(db.String(50))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            "id": self.id,
            "client_id": self.client_id,
            "client_name": self.client.name if self.client else "",
            "date": self.date.isoformat() if self.date else "",
            "description": self.description,
            "amount": self.amount,
            "btw_rate": self.btw_rate,
            "btw_amount": self.btw_amount,
            "total": self.total,
            "category": self.category or "",
            "invoice_number": self.invoice_number or "",
        }


class Expense(db.Model):
    __tablename__ = "expenses"

    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.Date, nullable=False, default=date.today)
    description = db.Column(db.String(500), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    btw_rate = db.Column(db.Float, default=21.0)
    btw_amount = db.Column(db.Float, default=0.0)
    total = db.Column(db.Float, nullable=False)
    category = db.Column(db.String(100), default="General")
    receipt_ref = db.Column(db.String(100))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            "id": self.id,
            "date": self.date.isoformat() if self.date else "",
            "description": self.description,
            "amount": self.amount,
            "btw_rate": self.btw_rate,
            "btw_amount": self.btw_amount,
            "total": self.total,
            "category": self.category or "",
            "receipt_ref": self.receipt_ref or "",
        }


class HourEntry(db.Model):
    __tablename__ = "hours"

    id = db.Column(db.Integer, primary_key=True)
    client_id = db.Column(db.Integer, db.ForeignKey("clients.id"), nullable=True)
    date = db.Column(db.Date, nullable=False, default=date.today)
    hours = db.Column(db.Float, nullable=False)
    rate = db.Column(db.Float, nullable=False, default=0.0)
    description = db.Column(db.String(500))
    invoiced = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            "id": self.id,
            "client_id": self.client_id,
            "client_name": self.client.name if self.client else "",
            "date": self.date.isoformat() if self.date else "",
            "hours": self.hours,
            "rate": self.rate,
            "earnings": round(self.hours * self.rate, 2),
            "description": self.description or "",
            "invoiced": self.invoiced,
        }


class MonthlyBilling(db.Model):
    __tablename__ = "monthly_billing"

    id = db.Column(db.Integer, primary_key=True)
    client_id = db.Column(db.Integer, db.ForeignKey("clients.id"), nullable=False)
    month = db.Column(db.Integer, nullable=False)
    year = db.Column(db.Integer, nullable=False)
    invoice_number = db.Column(db.String(50), unique=True, nullable=False)
    amount = db.Column(db.Float, nullable=False, default=0.0)
    btw_rate = db.Column(db.Float, default=21.0)
    status = db.Column(db.String(20), default="unpaid")  # unpaid | paid
    paid_date = db.Column(db.Date, nullable=True)
    income_id = db.Column(db.Integer, nullable=True)  # ref to incomes.id
    notes = db.Column(db.String(200))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class MonthlyHoursLog(db.Model):
    """One record per generated month, storing how many classes were taught."""
    __tablename__ = "monthly_hours_log"

    id = db.Column(db.Integer, primary_key=True)
    month = db.Column(db.Integer, nullable=False)
    year = db.Column(db.Integer, nullable=False)
    group_classes = db.Column(db.Integer, default=0)
    individual_classes = db.Column(db.Integer, default=0)
    other_hours = db.Column(db.Float, default=0.0)
    total_hours = db.Column(db.Float, default=0.0)
    hour_entry_id = db.Column(db.Integer, nullable=True)  # ref to hours.id


class ClassPlanner(db.Model):
    """Planned classes — one row per date per planned event."""
    __tablename__ = "class_planner"

    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.Date, nullable=False)
    # venue options: Gymzaal hof van Monaco | Gymzaal 't Vijfspan | Online | Home
    venue = db.Column(db.String(200), nullable=False, default="Home")
    description = db.Column(db.Text)
    # group | individual | other
    class_type = db.Column(db.String(20), default="group")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class FeeDefaults(db.Model):
    """Single-row table holding default fee + BTW per student type."""
    __tablename__ = "fee_defaults"

    id = db.Column(db.Integer, primary_key=True)
    adult_fee = db.Column(db.Float, default=60.0)
    adult_btw = db.Column(db.Float, default=21.0)
    child_fee = db.Column(db.Float, default=45.0)
    child_btw = db.Column(db.Float, default=21.0)
    individual_fee = db.Column(db.Float, default=100.0)
    individual_btw = db.Column(db.Float, default=21.0)
    online_fee = db.Column(db.Float, default=40.0)
    online_btw = db.Column(db.Float, default=21.0)

    # Venue rent and travel expense
    monaco_rent = db.Column(db.Float, default=23.72)     # Gymzaal hof van Monaco rental cost per session
    vijfspan_rent = db.Column(db.Float, default=19.25)   # Gymzaal 't Vijfspan rental cost per session
    veenendaal_rent = db.Column(db.Float, default=0.0)   # Veenendaal Yoga rental cost per session
    monaco_distance = db.Column(db.Float, default=41.8)  # Distance to Gymzaal hof van Monaco (km)
    vijfspan_distance = db.Column(db.Float, default=40.5) # Distance to Gymzaal 't Vijfspan (km)
    veenendaal_distance = db.Column(db.Float, default=3.0) # Distance to Veenendaal Yoga (km)

    # Monthly expense rates
    internet_cost = db.Column(db.Float, default=15.0)     # Monthly internet cost
    website_hosting_cost = db.Column(db.Float, default=15.0) # Monthly website hosting cost

    @classmethod
    def get(cls):
        row = cls.query.first()
        if not row:
            row = cls()
            db.session.add(row)
            db.session.commit()
        return row
