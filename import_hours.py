"""
Populate hours tracking for Natyanjani based on class schedule.

Rules:
- Each class session = 3 hours of effort (prep + teaching)
- Intensive group: 8 classes/month (2 per week)
- Basic group: 4 classes/month (1 per week)
- Private: 4 sessions/month per active student
- Choreography: 8 hours/week (32 hours/month)
- Admin work: 8 hours/week (32 hours/month)

Hours are logged monthly per activity. Rates are set to 0 since
income is tracked separately — these hours count toward the
1,225-hour urencriterium for zelfstandigenaftrek.
"""
import os
import sys
from datetime import date
from calendar import monthrange

sys.path.insert(0, os.path.dirname(__file__))

from app import create_app
from models import db, Income, HourEntry

HOURS_PER_CLASS = 3
INTENSIVE_CLASSES_PER_MONTH = 8
BASIC_CLASSES_PER_MONTH = 4
PRIVATE_SESSIONS_PER_MONTH = 4
CHOREOGRAPHY_HOURS_PER_WEEK = 8
ADMIN_HOURS_PER_WEEK = 8

MONTHS = range(1, 10)  # Jan to Sep 2025
YEAR = 2025


def weeks_in_month(year, month):
    """Approximate weeks in a month (round to nearest 0.5)."""
    days = monthrange(year, month)[1]
    return round(days / 7 * 2) / 2  # e.g. 4.0 or 4.5


def main():
    app = create_app()
    with app.app_context():
        # Clear existing hours to avoid duplicates
        HourEntry.query.delete()
        db.session.flush()

        total_hours = 0
        entries_added = 0

        for month in MONTHS:
            first_of_month = date(YEAR, month, 1)
            weeks = weeks_in_month(YEAR, month)

            # Check which class types were active this month
            intensive_active = Income.query.filter(
                db.extract("year", Income.date) == YEAR,
                db.extract("month", Income.date) == month,
                Income.category == "Dance Classes - Intensive",
            ).count()

            basic_active = Income.query.filter(
                db.extract("year", Income.date) == YEAR,
                db.extract("month", Income.date) == month,
                Income.category == "Dance Classes - Basic",
            ).count()

            private_students = Income.query.filter(
                db.extract("year", Income.date) == YEAR,
                db.extract("month", Income.date) == month,
                Income.category == "Dance Classes - Private",
            ).count()

            month_hours = 0

            # 1. Intensive group classes: 8 classes × 3 hours
            if intensive_active > 0:
                h = INTENSIVE_CLASSES_PER_MONTH * HOURS_PER_CLASS
                entry = HourEntry(
                    date=first_of_month,
                    hours=h,
                    rate=0,
                    description=f"Intensive group classes ({INTENSIVE_CLASSES_PER_MONTH} classes × {HOURS_PER_CLASS}h, {intensive_active} students)",
                    invoiced=True,
                )
                db.session.add(entry)
                month_hours += h
                entries_added += 1

            # 2. Basic group classes: 4 classes × 3 hours
            if basic_active > 0:
                h = BASIC_CLASSES_PER_MONTH * HOURS_PER_CLASS
                entry = HourEntry(
                    date=first_of_month,
                    hours=h,
                    rate=0,
                    description=f"Basic group classes ({BASIC_CLASSES_PER_MONTH} classes × {HOURS_PER_CLASS}h, {basic_active} students)",
                    invoiced=True,
                )
                db.session.add(entry)
                month_hours += h
                entries_added += 1

            # 3. Private sessions: 4 sessions × 3 hours per student
            if private_students > 0:
                h = private_students * PRIVATE_SESSIONS_PER_MONTH * HOURS_PER_CLASS
                entry = HourEntry(
                    date=first_of_month,
                    hours=h,
                    rate=0,
                    description=f"Private classes ({private_students} students × {PRIVATE_SESSIONS_PER_MONTH} sessions × {HOURS_PER_CLASS}h)",
                    invoiced=True,
                )
                db.session.add(entry)
                month_hours += h
                entries_added += 1

            # 4. Choreography: 8 hours/week
            h = CHOREOGRAPHY_HOURS_PER_WEEK * weeks
            entry = HourEntry(
                date=first_of_month,
                hours=h,
                rate=0,
                description=f"Choreography ({CHOREOGRAPHY_HOURS_PER_WEEK}h/week × {weeks} weeks)",
                invoiced=False,
            )
            db.session.add(entry)
            month_hours += h
            entries_added += 1

            # 5. Admin work: 8 hours/week
            h = ADMIN_HOURS_PER_WEEK * weeks
            entry = HourEntry(
                date=first_of_month,
                hours=h,
                rate=0,
                description=f"Admin work ({ADMIN_HOURS_PER_WEEK}h/week × {weeks} weeks)",
                invoiced=False,
            )
            db.session.add(entry)
            month_hours += h
            entries_added += 1

            total_hours += month_hours
            print(f"  {first_of_month.strftime('%B %Y'):>18s}: {month_hours:6.1f} hours")

        db.session.commit()

        print(f"\n{'='*50}")
        print(f"HOURS IMPORT COMPLETE")
        print(f"{'='*50}")
        print(f"Entries added:       {entries_added}")
        print(f"Total hours (9 mo):  {total_hours:.1f}")
        print(f"Monthly average:     {total_hours / 9:.1f}")
        print(f"Projected annual:    {total_hours / 9 * 12:.0f}")
        print(f"{'='*50}")
        if total_hours / 9 * 12 >= 1225:
            print(f">> On track for urencriterium (1,225h)!")
        else:
            print(f">> Need {1225 - total_hours / 9 * 12:.0f} more hours/year for urencriterium.")


if __name__ == "__main__":
    main()
