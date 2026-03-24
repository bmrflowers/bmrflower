# 🌸 Flower Shop Billing Software

A mobile-first Flask web application for managing billing in a flower shop, powered by **Supabase** as the backend database.

---

## Features

- **Login** — Session-based single admin login
- **Dashboard** — Quick-access cards with today's stats
- **Create Bill** — Select flowers, enter qty/rate, auto GST calculation
- **Invoice** — Printable invoice with PDF download (html2pdf)
- **Billing History** — Paginated history with card layout on mobile
- **Daily Sales Report** — Filter by date, summary stats
- **Export to Excel** — `.xlsx` export for history and daily reports
- **Mobile-First UI** — Bootstrap-like responsive light green theme

---

## Project Structure

```
BMR/
├── app.py                  # Main Flask application
├── requirements.txt
├── .env.example            # Copy to .env and fill credentials
├── static/
│   ├── css/
│   │   └── style.css       # Full custom stylesheet
│   └── js/
│       └── billing.js      # Dynamic billing calculations
└── templates/
    ├── login.html
    ├── dashboard.html
    ├── create_bill.html
    ├── invoice.html
    ├── history.html
    └── report.html
```

---

## Supabase Setup

### 1. Create a Supabase project
Go to [https://supabase.com](https://supabase.com), create a free project.

### 2. Create the `bills` table

Run the following SQL in your Supabase **SQL Editor**:

```sql
CREATE TABLE bills (
  id           BIGSERIAL PRIMARY KEY,
  customer_name TEXT NOT NULL,
  phone        TEXT,
  items        JSONB NOT NULL DEFAULT '[]',
  subtotal     NUMERIC(10,2) NOT NULL DEFAULT 0,
  gst          NUMERIC(5,2)  NOT NULL DEFAULT 5,
  gst_amount   NUMERIC(10,2) NOT NULL DEFAULT 0,
  total        NUMERIC(10,2) NOT NULL DEFAULT 0,
  bill_date    DATE,
  created_at   TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- Enable Row Level Security (RLS) — use your server key to bypass from backend
ALTER TABLE bills ENABLE ROW LEVEL SECURITY;

-- Allow all operations from service role (used by backend)
CREATE POLICY "Allow all for service role"
  ON bills
  USING (true)
  WITH CHECK (true);
```

> **Tip:** Copy your **Project URL** and **anon/public key** from  
> *Project Settings → API → Project API Keys*.

---

## Local Setup

### Prerequisites
- Python 3.10+ installed
- pip

### Steps

```bash
# 1. Clone / navigate to the project folder
cd "c:\Users\ELCOT\Documents\BMR"

# 2. Create virtual environment
python -m venv venv

# 3. Activate virtual environment (Windows)
venv\Scripts\activate

# 4. Install dependencies
pip install -r requirements.txt

# 5. Configure environment
copy .env.example .env
# Open .env and fill in your SUPABASE_URL, SUPABASE_KEY, etc.

# 6. Run the app
python app.py
```

Open your browser at **http://localhost:5000**

---

## Default Login

| Field    | Value       |
|----------|-------------|
| Username | `admin`     |
| Password | `flower@123` |

Change these in `.env` before going live.

---

## Environment Variables

| Variable         | Description                        |
|------------------|------------------------------------|
| `SUPABASE_URL`   | Your Supabase project URL          |
| `SUPABASE_KEY`   | Supabase anon/service key          |
| `SECRET_KEY`     | Flask session secret key           |
| `ADMIN_USERNAME` | Login username (default: admin)    |
| `ADMIN_PASSWORD` | Login password (default: flower@123) |
| `SHOP_NAME`      | Shop name shown on invoices        |
| `SHOP_ADDRESS`   | Shop address                       |
| `SHOP_PHONE`     | Phone number                       |
| `SHOP_EMAIL`     | Email                              |
| `SHOP_GST`       | GST registration number            |

---

## Security Notes

- Change `SECRET_KEY`, `ADMIN_PASSWORD` in `.env` before deploying.
- Never commit `.env` to version control — it's already in `.gitignore`.
- The Supabase **service role key** (if used) bypasses RLS — keep it server-side only.

---

## Tech Stack

| Layer      | Technology             |
|------------|------------------------|
| Backend    | Python / Flask         |
| Database   | Supabase (PostgreSQL)  |
| Frontend   | Custom CSS + Vanilla JS |
| PDF Export | html2pdf.js (CDN)      |
| Excel      | openpyxl               |
