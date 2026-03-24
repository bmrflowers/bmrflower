"""
Flower Shop Billing Software
Main Flask Application
"""

import os
import json
import smtplib
import threading
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, date
from functools import wraps

from dotenv import load_dotenv
load_dotenv()  # Load .env file if present

from flask import (
    Flask, render_template, request, redirect,
    url_for, session, flash, jsonify, send_file
)
from supabase import create_client, Client
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import io

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'flowershop-secret-2024-change-in-prod')

# Jinja2 custom filter: enumerate for templates
app.jinja_env.filters['enumerate'] = enumerate

# ── Supabase configuration ──────────────────────────────────────────────────
SUPABASE_URL = os.environ.get('SUPABASE_URL', '')
SUPABASE_KEY = os.environ.get('SUPABASE_KEY', '')

# Lazy singleton — created on first use so the app starts even with missing keys
_supabase_client: Client = None

def get_supabase() -> Client:
    global _supabase_client
    if _supabase_client is None:
        url = os.environ.get('SUPABASE_URL', SUPABASE_URL)
        key = os.environ.get('SUPABASE_KEY', SUPABASE_KEY)
        if not url or not key:
            raise RuntimeError(
                'SUPABASE_URL and SUPABASE_KEY must be set in .env before using the database.'
            )
        _supabase_client = create_client(url, key)
    return _supabase_client

# ── Admin credentials (change before production) ────────────────────────────
ADMIN_USERNAME = os.environ.get('ADMIN_USERNAME', 'admin')
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'flower@123')

# Shop info shown on invoices
SHOP_NAME = os.environ.get('SHOP_NAME', 'BMR FLOWERS')
SHOP_ADDRESS = os.environ.get('SHOP_ADDRESS', '12, Rose Garden Street, Chennai - 600001')
SHOP_PHONE = os.environ.get('SHOP_PHONE', '+91 98765 43210')
SHOP_EMAIL = os.environ.get('SHOP_EMAIL', 'info@bmrflowers.com')
SHOP_GST = os.environ.get('SHOP_GST', 'GST: 33ABCDE1234F1Z5')

# ── SMTP configuration ─────────────────────────────────────────────────────
SMTP_EMAIL       = os.environ.get('SMTP_EMAIL', '')
SMTP_APP_PASSWORD = os.environ.get('SMTP_APP_PASSWORD', '')


def _build_invoice_html(bill_id, customer_name, phone, email, bill_date,
                        items, subtotal, gst_percent, gst_amount, grand_total, shop):
    """Return a self-contained HTML email that looks like the printed invoice."""
    rows = ''
    for idx, item in enumerate(items, 1):
        rows += f"""
        <tr>
          <td style="padding:8px 10px;border-bottom:1px solid #e8f5e9;text-align:center">{idx}</td>
          <td style="padding:8px 10px;border-bottom:1px solid #e8f5e9">{item['name']}</td>
          <td style="padding:8px 10px;border-bottom:1px solid #e8f5e9;text-align:center">{item['qty']}</td>
          <td style="padding:8px 10px;border-bottom:1px solid #e8f5e9;text-align:right">&#8377;{item['rate']:.2f}</td>
          <td style="padding:8px 10px;border-bottom:1px solid #e8f5e9;text-align:right">&#8377;{item['total']:.2f}</td>
        </tr>"""

    phone_row = f'<p style="margin:2px 0;font-size:13px;color:#555">&#128241; {phone}</p>' if phone else ''
    email_row = f'<p style="margin:2px 0;font-size:13px;color:#555">&#9993; {email}</p>' if email else ''

    return f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8" /></head>
<body style="margin:0;padding:0;background:#f0f4f0;font-family:Arial,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f0f4f0;padding:30px 0">
  <tr><td align="center">
  <table width="580" cellpadding="0" cellspacing="0"
         style="background:#fff;border-radius:12px;overflow:hidden;
                box-shadow:0 4px 20px rgba(0,0,0,.10)">

    <!-- Header -->
    <tr><td style="background:#2e7d32;padding:24px 30px;text-align:center">
      <div style="font-size:26px;font-weight:800;color:#fff;letter-spacing:1px">
        &#127800; {shop['name']}
      </div>
      <div style="font-size:12px;color:#c8e6c9;margin-top:4px">{shop['address']}</div>
      <div style="font-size:12px;color:#c8e6c9;margin-top:2px">
        &#128222; {shop['phone']} &nbsp;|&nbsp; &#9993; {shop['email']}
      </div>
      <div style="font-size:11px;color:#a5d6a7;margin-top:2px">{shop['gst']}</div>
    </td></tr>

    <!-- TAX INVOICE label -->
    <tr><td style="padding:16px 30px 0;text-align:center">
      <div style="font-size:14px;font-weight:700;color:#2e7d32;
                  letter-spacing:2px;text-transform:uppercase;border-bottom:2px solid #c8e6c9;
                  padding-bottom:12px">Tax Invoice</div>
    </td></tr>

    <!-- Bill To / Invoice No -->
    <tr><td style="padding:16px 30px">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td style="vertical-align:top">
            <div style="font-size:10px;text-transform:uppercase;color:#888;letter-spacing:.08em;margin-bottom:4px">Bill To</div>
            <p style="margin:0;font-size:15px;font-weight:700;color:#222">{customer_name}</p>
            {phone_row}
            {email_row}
          </td>
          <td style="vertical-align:top;text-align:right">
            <div style="font-size:10px;text-transform:uppercase;color:#888;letter-spacing:.08em">Invoice No.</div>
            <div style="font-size:16px;font-weight:800;color:#222">#{bill_id:04d}</div>
            <div style="font-size:10px;text-transform:uppercase;color:#888;margin-top:6px">Date</div>
            <div style="font-size:13px;color:#444">{bill_date}</div>
          </td>
        </tr>
      </table>
    </td></tr>

    <!-- Items table -->
    <tr><td style="padding:0 30px">
      <table width="100%" cellpadding="0" cellspacing="0"
             style="border-radius:8px;overflow:hidden;border:1px solid #c8e6c9">
        <thead>
          <tr style="background:#2e7d32">
            <th style="padding:10px;color:#fff;font-size:12px;text-align:center">#</th>
            <th style="padding:10px;color:#fff;font-size:12px;text-align:left">Flower</th>
            <th style="padding:10px;color:#fff;font-size:12px;text-align:center">Qty</th>
            <th style="padding:10px;color:#fff;font-size:12px;text-align:right">Rate (&#8377;)</th>
            <th style="padding:10px;color:#fff;font-size:12px;text-align:right">Amount (&#8377;)</th>
          </tr>
        </thead>
        <tbody>{rows}</tbody>
      </table>
    </td></tr>

    <!-- Totals -->
    <tr><td style="padding:16px 30px">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td colspan="2"><table width="100%" cellpadding="0" cellspacing="0"
                style="max-width:260px;margin-left:auto">
            <tr>
              <td style="padding:5px 0;font-size:13px;color:#555">Subtotal</td>
              <td style="padding:5px 0;font-size:13px;color:#555;text-align:right">&#8377;{subtotal:.2f}</td>
            </tr>
            <tr>
              <td style="padding:5px 0;font-size:13px;color:#555">GST ({gst_percent}%)</td>
              <td style="padding:5px 0;font-size:13px;color:#555;text-align:right">&#8377;{gst_amount:.2f}</td>
            </tr>
            <tr style="border-top:2px solid #c8e6c9">
              <td style="padding:10px 0 5px;font-size:15px;font-weight:800;color:#2e7d32">Grand Total</td>
              <td style="padding:10px 0 5px;font-size:15px;font-weight:800;color:#2e7d32;text-align:right">&#8377;{grand_total:.2f}</td>
            </tr>
          </table></td>
        </tr>
      </table>
    </td></tr>

    <!-- Footer -->
    <tr><td style="background:#f1f8e9;padding:16px 30px;text-align:center;
                   font-size:13px;color:#558b2f;border-top:2px solid #c8e6c9">
      &#127800; Thank you for shopping with us! &nbsp;|&nbsp; Visit again!
    </td></tr>

  </table>
  </td></tr>
</table>
</body></html>"""


def send_invoice_email(to_addr, bill_id, customer_name, phone, bill_date,
                       items, subtotal, gst_percent, gst_amount, grand_total, shop):
    """Send the invoice as an HTML email. Called in a daemon thread."""
    if not SMTP_EMAIL or not SMTP_APP_PASSWORD or not to_addr:
        return
    try:
        html_body = _build_invoice_html(
            bill_id, customer_name, phone, to_addr, bill_date,
            items, subtotal, gst_percent, gst_amount, grand_total, shop
        )
        msg = MIMEMultipart('alternative')
        msg['Subject'] = f'Your Invoice #{bill_id:04d} from {shop["name"]}'
        msg['From']    = f'{shop["name"]} <{SMTP_EMAIL}>'
        msg['To']      = to_addr
        msg.attach(MIMEText(html_body, 'html', 'utf-8'))

        with smtplib.SMTP('smtp.gmail.com', 587, timeout=10) as server:
            server.ehlo()
            server.starttls()
            server.login(SMTP_EMAIL, SMTP_APP_PASSWORD)
            server.sendmail(SMTP_EMAIL, [to_addr], msg.as_string())
    except Exception as exc:
        # Log but never crash the main request
        print(f'[EMAIL] Failed to send invoice to {to_addr}: {exc}')


# ── Default flower catalogue (fallback if DB/flowers table unavailable) ──────
DEFAULT_FLOWERS = [
    {'name': 'Rose',     'emoji': '🌹'},
    {'name': 'Lily',     'emoji': '🌷'},
    {'name': 'Jasmine',  'emoji': '🤍'},
    {'name': 'Tulip',    'emoji': '🌸'},
    {'name': 'Lotus',    'emoji': '🪷'},
    {'name': 'Marigold', 'emoji': '🟡'},
    {'name': 'Orchid',   'emoji': '💜'},
]


def get_flowers():
    """Return active flower catalogue from Supabase, falling back to DEFAULT_FLOWERS."""
    try:
        res = (
            get_supabase().table('flowers')
            .select('id, name, emoji')
            .eq('is_active', True)
            .order('name')
            .execute()
        )
        return res.data if res.data else DEFAULT_FLOWERS
    except Exception:
        return DEFAULT_FLOWERS



def get_admin_user(username: str):
    """Fetch an admin user row by username. Returns dict or None."""
    try:
        res = get_supabase().table('admin_users').select('*').eq('username', username).execute()
        if res.data:
            return res.data[0]
        return None
    except Exception:
        return None


def create_admin_user(username: str, password: str) -> bool:
    """Create an admin user with a hashed password. Returns True on success."""
    try:
        get_supabase().table('admin_users').insert({
            'username': username,
            'password': password,
        }).execute()
        return True
    except Exception:
        return False


# ── Auth helper ─────────────────────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'logged_in' not in session:
            flash('Please log in to access this page.', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


def shop_info():
    return {
        'name': SHOP_NAME,
        'address': SHOP_ADDRESS,
        'phone': SHOP_PHONE,
        'email': SHOP_EMAIL,
        'gst': SHOP_GST,
    }


# ── Authentication routes ────────────────────────────────────────────────────
@app.route('/', methods=['GET', 'POST'])
def login():
    if session.get('logged_in'):
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()

        user = get_admin_user(username)

        # If user exists in DB, verify plaintext password
        if user:
            plain_pw = user.get('password')
            if plain_pw and password == plain_pw:
                session['logged_in'] = True
                session['username'] = username
                flash('Welcome back! 🌸', 'success')
                return redirect(url_for('dashboard'))
            else:
                flash('Invalid username or password.', 'danger')
        else:
            # Fallback: allow env-admin to login and create DB entry
            if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
                try:
                    create_admin_user(username, password)
                except Exception:
                    pass
                session['logged_in'] = True
                session['username'] = username
                flash('Welcome back! 🌸', 'success')
                return redirect(url_for('dashboard'))
            flash('Invalid username or password.', 'danger')

    return render_template('login.html', shop=shop_info())


@app.route('/logout')
def logout():
    session.clear()
    flash('You have been logged out.', 'info')
    return redirect(url_for('login'))


# ── Dashboard ────────────────────────────────────────────────────────────────
@app.route('/dashboard')
@login_required
def dashboard():
    today = date.today()
    today_str = today.isoformat()

    # Start of current week (Monday)
    week_start = (today - __import__('datetime').timedelta(days=today.weekday())).isoformat()
    # Start of current month
    month_start = today.replace(day=1).isoformat()

    def _fetch(start, end):
        try:
            res = (
                get_supabase().table('bills')
                .select('id, total, gst_amount')
                .gte('created_at', start + 'T00:00:00')
                .lte('created_at', end + 'T23:59:59')
                .execute()
            )
            bills = res.data
            return len(bills), round(sum(float(b.get('total', 0)) for b in bills), 2)
        except Exception:
            return 0, 0.0

    today_count, today_total = _fetch(today_str, today_str)
    week_count, week_total   = _fetch(week_start, today_str)
    month_count, month_total = _fetch(month_start, today_str)

    return render_template(
        'dashboard.html',
        shop=shop_info(),
        today_count=today_count,
        today_total=today_total,
        week_count=week_count,
        week_total=week_total,
        month_count=month_count,
        month_total=month_total,
    )


# ── Create Bill ──────────────────────────────────────────────────────────────
@app.route('/create-bill', methods=['GET', 'POST'])
@login_required
def create_bill():
    if request.method == 'POST':
        try:
            customer_name = request.form.get('customer_name', '').strip()
            phone         = request.form.get('phone', '').strip()
            email         = request.form.get('email', '').strip()
            bill_date     = request.form.get('bill_date', date.today().isoformat())
            gst_percent   = float(request.form.get('gst_percent', 5))

            # Collect item data sent from the form
            items = []
            for flower in get_flowers():
                key = flower['name'].lower()
                qty_val = request.form.get(f'qty_{key}', '').strip()
                rate_val = request.form.get(f'rate_{key}', '').strip()
                if qty_val and rate_val:
                    qty = float(qty_val)
                    rate = float(rate_val)
                    if qty > 0 and rate >= 0:
                        items.append({
                            'name': flower['name'],
                            'qty': qty,
                            'rate': rate,
                            'total': round(qty * rate, 2),
                        })

            if not customer_name:
                return jsonify({'success': False, 'error': 'Customer name is required.'}), 400

            if not items:
                return jsonify({'success': False, 'error': 'Select at least one flower item.'}), 400

            subtotal = round(sum(i['total'] for i in items), 2)
            gst_amount = round(subtotal * gst_percent / 100, 2)
            grand_total = round(subtotal + gst_amount, 2)

            # Save to Supabase
            result = get_supabase().table('bills').insert({
                'customer_name': customer_name,
                'phone': phone,
                'email': email,
                'items': items,
                'subtotal': subtotal,
                'gst': gst_percent,
                'gst_amount': gst_amount,
                'total': grand_total,
                'bill_date': bill_date,
                'created_at': datetime.utcnow().isoformat(),
            }).execute()

            bill_id = result.data[0]['id']

            # Send invoice email in background (non-blocking)
            if email:
                threading.Thread(
                    target=send_invoice_email,
                    args=(email, bill_id, customer_name, phone, bill_date,
                          items, subtotal, gst_percent, gst_amount, grand_total,
                          shop_info()),
                    daemon=True
                ).start()

            return jsonify({'success': True, 'bill_id': bill_id})

        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 500

    return render_template('create_bill.html', flowers=get_flowers(), shop=shop_info(), today=date.today().isoformat())


# ── Invoice view ─────────────────────────────────────────────────────────────
@app.route('/invoice/<int:bill_id>')
@login_required
def invoice(bill_id):
    try:
        result = get_supabase().table('bills').select('*').eq('id', bill_id).execute()
        if not result.data:
            flash('Bill not found.', 'danger')
            return redirect(url_for('history'))
        bill = result.data[0]
        # Ensure items is a list
        items = bill.get('items')
        if isinstance(items, str):
            try:
                items = json.loads(items)
            except Exception:
                items = []
        # ensure it's a list
        if items is None:
            items = []
        bill['items'] = items
        return render_template('invoice.html', bill=bill, items=items, shop=shop_info())
    except Exception as e:
        flash(f'Error loading invoice: {e}', 'danger')
        return redirect(url_for('history'))


# ── Billing History ──────────────────────────────────────────────────────────
@app.route('/history')
@login_required
def history():
    page = int(request.args.get('page', 1))
    per_page = 20
    offset = (page - 1) * per_page

    try:
        result = (
            get_supabase().table('bills')
            .select('id, customer_name, phone, email, total, created_at, bill_date')
            .order('created_at', desc=True)
            .range(offset, offset + per_page - 1)
            .execute()
        )
        bills = result.data
        has_next = len(bills) == per_page
    except Exception as e:
        flash(f'Error loading history: {e}', 'danger')
        bills = []
        has_next = False

    return render_template('history.html', bills=bills, page=page, has_next=has_next, shop=shop_info())


# ── Daily Sales Report ───────────────────────────────────────────────────────
@app.route('/report', methods=['GET'])
@login_required
def report():
    report_date = request.args.get('date', date.today().isoformat())

    try:
        result = (
            get_supabase().table('bills')
            .select('*')
            .gte('created_at', report_date + 'T00:00:00')
            .lte('created_at', report_date + 'T23:59:59')
            .order('created_at', desc=True)
            .execute()
        )
        bills = result.data
        total_sales = round(sum(float(b.get('total', 0)) for b in bills), 2)
        total_gst = round(sum(float(b.get('gst_amount', 0)) for b in bills), 2)
        total_subtotal = round(sum(float(b.get('subtotal', 0)) for b in bills), 2)
    except Exception as e:
        flash(f'Error loading report: {e}', 'danger')
        bills = []
        total_sales = total_gst = total_subtotal = 0.0

    return render_template(
        'report.html',
        bills=bills,
        report_date=report_date,
        total_sales=total_sales,
        total_gst=total_gst,
        total_subtotal=total_subtotal,
        shop=shop_info(),
    )


# ── Export to Excel ───────────────────────────────────────────────────────────
def _build_excel(bills: list, sheet_title: str) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title

    # Header style
    header_fill = PatternFill(start_color='4CAF50', end_color='4CAF50', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF')
    center = Alignment(horizontal='center')

    headers = ['Bill ID', 'Customer Name', 'Phone', 'Date', 'Subtotal (₹)', 'GST %', 'GST Amount (₹)', 'Total (₹)']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center

    for row_idx, b in enumerate(bills, 2):
        created = b.get('bill_date') or b.get('created_at', '')[:10]
        ws.cell(row=row_idx, column=1, value=b.get('id'))
        ws.cell(row=row_idx, column=2, value=b.get('customer_name', ''))
        ws.cell(row=row_idx, column=3, value=b.get('phone', ''))
        ws.cell(row=row_idx, column=4, value=created)
        ws.cell(row=row_idx, column=5, value=float(b.get('subtotal', 0)))
        ws.cell(row=row_idx, column=6, value=float(b.get('gst', 0)))
        ws.cell(row=row_idx, column=7, value=float(b.get('gst_amount', 0)))
        ws.cell(row=row_idx, column=8, value=float(b.get('total', 0)))

    # Auto column width
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = max(max_len + 4, 14)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@app.route('/export/history')
@login_required
def export_history():
    try:
        result = get_supabase().table('bills').select('*').order('created_at', desc=True).execute()
        buf = _build_excel(result.data, 'Billing History')
        filename = f'billing_history_{date.today().isoformat()}.xlsx'
        return send_file(buf, as_attachment=True, download_name=filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        flash(f'Export failed: {e}', 'danger')
        return redirect(url_for('history'))


@app.route('/export/report')
@login_required
def export_report():
    report_date = request.args.get('date', date.today().isoformat())
    try:
        result = (
            get_supabase().table('bills')
            .select('*')
            .gte('created_at', report_date + 'T00:00:00')
            .lte('created_at', report_date + 'T23:59:59')
            .execute()
        )
        buf = _build_excel(result.data, 'Daily Report')
        filename = f'daily_report_{report_date}.xlsx'
        return send_file(buf, as_attachment=True, download_name=filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        flash(f'Export failed: {e}', 'danger')
        return redirect(url_for('report'))


# ── Delete bill (optional utility) ───────────────────────────────────────────
@app.route('/bill/delete/<int:bill_id>', methods=['POST'])
@login_required
def delete_bill(bill_id):
    try:
        get_supabase().table('bills').delete().eq('id', bill_id).execute()
        flash('Bill deleted successfully.', 'success')
    except Exception as e:
        flash(f'Delete failed: {e}', 'danger')
    return redirect(url_for('history'))


# ── Manage Flowers ───────────────────────────────────────────────────────────
@app.route('/manage-flowers', methods=['GET', 'POST'])
@login_required
def manage_flowers():
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add':
            name = request.form.get('name', '').strip().title()
            emoji = request.form.get('emoji', '🌺').strip() or '🌺'
            if not name:
                flash('Flower name is required.', 'danger')
            else:
                try:
                    get_supabase().table('flowers').insert({
                        'name': name,
                        'emoji': emoji,
                        'is_active': True,
                    }).execute()
                    flash(f'{emoji} {name} added to the catalogue!', 'success')
                except Exception as e:
                    flash(f'Error adding flower: {e}', 'danger')
        elif action == 'delete':
            flower_id = request.form.get('flower_id', '').strip()
            if flower_id and flower_id.isdigit():
                try:
                    get_supabase().table('flowers').delete().eq('id', int(flower_id)).execute()
                    flash('Flower removed from catalogue.', 'success')
                except Exception as e:
                    flash(f'Error removing flower: {e}', 'danger')
        return redirect(url_for('manage_flowers'))

    db_ok = True
    try:
        res = (
            get_supabase().table('flowers')
            .select('*')
            .order('name')
            .execute()
        )
        flowers = res.data
    except Exception as e:
        flowers = DEFAULT_FLOWERS
        db_ok = False
        flash(f'DB unavailable — showing defaults. Run the setup SQL first. ({e})', 'warning')
    return render_template('manage_flowers.html', flowers=flowers, shop=shop_info(), db_ok=db_ok)


# ── Settings ─────────────────────────────────────────────────────────────────
@app.route('/settings', methods=['GET', 'POST'])
@login_required
def settings():
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'change_password':
            current_pw = request.form.get('current_password', '').strip()
            new_pw     = request.form.get('new_password', '').strip()
            confirm_pw = request.form.get('confirm_password', '').strip()

            username = session.get('username', ADMIN_USERNAME)
            user = get_admin_user(username)

            # Verify current password using plaintext column
            verified = False
            if user:
                plain_pw = user.get('password')
                if plain_pw and current_pw == plain_pw:
                    verified = True
            else:
                if username == ADMIN_USERNAME and current_pw == ADMIN_PASSWORD:
                    verified = True

            if not verified:
                flash('Current password is incorrect.', 'danger')
            elif len(new_pw) < 6:
                flash('New password must be at least 6 characters.', 'danger')
            elif new_pw != confirm_pw:
                flash('New passwords do not match.', 'danger')
            else:
                try:
                    if user:
                        get_supabase().table('admin_users').update({
                            'password': new_pw
                        }).eq('id', user.get('id')).execute()
                    else:
                        create_admin_user(username, new_pw)
                    flash('Password updated successfully! 🔐', 'success')
                except Exception as e:
                    flash(f'Error updating password: {e}', 'danger')
        return redirect(url_for('settings'))
    return render_template('settings.html', shop=shop_info())


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
