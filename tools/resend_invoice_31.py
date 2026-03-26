#!/usr/bin/env python3
"""Resend invoice 31 using stored email in the `bills` table.
Run from the project root where the venv is available.
"""
import os
import sys
from dotenv import load_dotenv
load_dotenv()

try:
    # Import app helpers
    from app import get_supabase, send_invoice_email
except Exception as e:
    print('Failed importing app:', e)
    sys.exit(2)

BILL_ID = 31

try:
    sb = get_supabase()
    res = sb.table('bills').select('id,email').eq('id', BILL_ID).execute()
    if not res.data:
        print(f'Bill {BILL_ID} not found in database.')
        sys.exit(1)
    to_addr = res.data[0].get('email')
    if not to_addr:
        print(f'No email stored for bill {BILL_ID}.')
        sys.exit(1)
    print(f'Resending invoice {BILL_ID} to {to_addr}...')
    send_invoice_email(to_addr, BILL_ID)
    print('Done (send attempted). Check outgoing_invoices/raw for saved .eml and server logs for status.')
except Exception as exc:
    print('Error during resend:', exc)
    sys.exit(3)
