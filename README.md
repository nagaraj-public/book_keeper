# BookKeeper NL

A simple bookkeeping application for entrepreneurs in the Netherlands.

## Prerequisites — What You Need to Install

Before running this app, make sure the following tools are installed on your PC:

### 1. Python 3.10 or higher
Download from https://www.python.org/downloads/  
During installation, check **"Add Python to PATH"**.

Verify after install:
```powershell
python --version
```

### 2. pip (comes with Python)
Verify:
```powershell
pip --version
```

### 3. Git (to clone/manage this repo)
Download from https://git-scm.com/downloads

Verify:
```powershell
git --version
```

> **No other tools needed.** Flask, SQLAlchemy, and all other dependencies are installed automatically via `pip install -r requirements.txt`.

---

## Features

- **Client Management** — Add, edit, delete clients. Import from Excel (.xlsx)
- **Income Tracking** — Record income with BTW (VAT) auto-calculation (21%, 9%, 0%)
- **Expense Tracking** — Categorize expenses with BTW support
- **Hours Tracking** — Log hours per client with hourly rates, track invoiced status
- **Dashboard** — Monthly overview of income, expenses, profit, and hours
- **Year-End Report** — Full annual summary with BTW totals, category breakdowns, and Dutch tax filing tips (zelfstandigenaftrek, MKB-winstvrijstelling)

## Quick Start

```bash
# Create virtual environment
python -m venv venv

# Activate (Windows)
.\venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the app
python app.py
```

Open http://localhost:5000 in your browser.

## Access from Your Phone (same WiFi)

The server is configured to accept connections from any device on your local network.

1. Find your PC's WiFi IP address:
   ```powershell
   ipconfig
   ```
   Look for the **IPv4 Address** under your WiFi adapter (e.g. `192.168.1.223`).

2. Open your phone's browser and navigate to:
   ```
   http://192.168.1.223:5000
   ```

3. (Optional) Allow the port through Windows Firewall if the phone can't connect:
   ```powershell
   netsh advfirewall firewall add rule name="BookKeeper Flask" dir=in action=allow protocol=TCP localport=5000
   ```

4. On Android (Chrome) or iOS (Safari) you can install it as a PWA — tap the browser menu → **Add to Home Screen**.

## Excel Import

The client import supports columns named in English or Dutch:
- Name / Naam
- Company / Bedrijfsnaam
- KVK / KVK Nummer
- BTW / BTW Nummer
- Email / E-mail
- Phone / Telefoon
- Address / Adres
- Notes / Opmerkingen

## Project Structure

```
book_keeper/
├── app.py              # Flask application entry point
├── config.py           # Configuration
├── models.py           # Database models (Client, Income, Expense, HourEntry)
├── routes.py           # All routes and report logic
├── excel_import.py     # Excel client import
├── requirements.txt    # Python dependencies
├── bookkeeper.db       # SQLite database (auto-created)
├── templates/
│   ├── base.html       # Layout with sidebar navigation
│   ├── dashboard.html  # Monthly dashboard
│   ├── clients.html    # Client management
│   ├── income.html     # Income entries
│   ├── expenses.html   # Expense entries
│   ├── hours.html      # Hours tracking
│   └── report.html     # Year-end tax report
└── uploads/            # Temp folder for Excel imports
```
