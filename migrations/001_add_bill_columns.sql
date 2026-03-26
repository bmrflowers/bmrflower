-- Migration: add optional billing columns to `bills` table
-- Run this in Supabase SQL editor or via psql connected to your database.

-- Check existing columns
SELECT column_name, data_type
FROM information_schema.columns
WHERE table_name = 'bills'
ORDER BY ordinal_position;

-- Add columns if they do not already exist
ALTER TABLE bills
  ADD COLUMN IF NOT EXISTS luggage_qty INTEGER DEFAULT 0,
  ADD COLUMN IF NOT EXISTS luggage_rate NUMERIC(12,2) DEFAULT 0,
  ADD COLUMN IF NOT EXISTS luggage_total NUMERIC(12,2) DEFAULT 0,
  ADD COLUMN IF NOT EXISTS old_balance NUMERIC(12,2) DEFAULT 0,
  ADD COLUMN IF NOT EXISTS cash_paid NUMERIC(12,2) DEFAULT 0,
  ADD COLUMN IF NOT EXISTS amount_due NUMERIC(12,2) DEFAULT 0;

-- Optional: add index on bill_date for faster range queries
CREATE INDEX IF NOT EXISTS idx_bills_bill_date ON bills (bill_date);

-- Optional: vacuum analyze (if running psql with proper permissions)
-- VACUUM ANALYZE bills;
