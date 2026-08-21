-- ==============================================================================
-- Schema for Bank Recon Cloud Sync (Supabase / PostgreSQL)
-- Run this in Supabase SQL Editor (1-click setup / migration)
-- ==============================================================================

-- ── 1. Bank Discrepancies Table ──────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS bank_discrepancies (
    id BIGSERIAL PRIMARY KEY,
    company_key VARCHAR(50) NOT NULL DEFAULT 'eyerizz',
    company TEXT DEFAULT 'Eyerizz Eyewear',
    recon_date DATE NOT NULL,
    bank_name TEXT NOT NULL,
    journal TEXT NOT NULL,
    transaction_date DATE NOT NULL,
    bank_number TEXT,
    filename TEXT,
    amount NUMERIC(18, 2) NOT NULL DEFAULT 0.0,
    discrepancy_type VARCHAR(50) DEFAULT 'bank_only',
    odoo_number VARCHAR(100),
    odoo_reference VARCHAR(100),
    is_reconciled VARCHAR(20) DEFAULT 'Yes',
    recon_hash VARCHAR(64) UNIQUE,
    status VARCHAR(50) DEFAULT 'Pending',
    action_type VARCHAR(100),
    resolved_by VARCHAR(100),
    sales_notes TEXT,
    uploaded_by VARCHAR(100),
    device_id VARCHAR(100),
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_discrepancies_company ON bank_discrepancies(company_key);
CREATE INDEX IF NOT EXISTS idx_discrepancies_status ON bank_discrepancies(status);
CREATE INDEX IF NOT EXISTS idx_discrepancies_recon_date ON bank_discrepancies(recon_date);
CREATE INDEX IF NOT EXISTS idx_discrepancies_bank ON bank_discrepancies(bank_name);
CREATE INDEX IF NOT EXISTS idx_discrepancies_txn_date ON bank_discrepancies(transaction_date);
CREATE INDEX IF NOT EXISTS idx_discrepancies_hash ON bank_discrepancies(recon_hash);

ALTER TABLE bank_discrepancies ENABLE ROW LEVEL SECURITY;
DO $$
BEGIN
    IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename = 'bank_discrepancies' AND policyname = 'authenticated_read_discrepancies') THEN
        CREATE POLICY "authenticated_read_discrepancies" ON bank_discrepancies FOR SELECT TO authenticated USING (true);
    END IF;
    IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename = 'bank_discrepancies' AND policyname = 'deny_anon_discrepancies') THEN
        CREATE POLICY "deny_anon_discrepancies" ON bank_discrepancies FOR ALL TO anon USING (false) WITH CHECK (false);
    END IF;
END $$;

-- ── 2. Bank Merchant Transactions Table (EDC Reports) ─────────────────────────
CREATE TABLE IF NOT EXISTS bank_merchant_transactions (
    id BIGSERIAL PRIMARY KEY,
    company_key VARCHAR(50) NOT NULL DEFAULT 'eyerizz',
    company TEXT DEFAULT 'Eyerizz Eyewear',
    bank_name TEXT NOT NULL,
    store TEXT,
    transaction_date DATE NOT NULL,
    trace_number TEXT,
    card_type VARCHAR(50),
    gross_amount NUMERIC(18, 2) NOT NULL DEFAULT 0.0,
    net_amount NUMERIC(18, 2) DEFAULT 0.0,
    fee_amount NUMERIC(18, 2) DEFAULT 0.0,
    filename TEXT,
    recon_hash VARCHAR(64) UNIQUE NOT NULL,
    uploaded_by VARCHAR(100),
    device_id VARCHAR(100),
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_merchant_company_date ON bank_merchant_transactions(company_key, transaction_date);
CREATE INDEX IF NOT EXISTS idx_merchant_bank ON bank_merchant_transactions(bank_name);
CREATE INDEX IF NOT EXISTS idx_merchant_hash ON bank_merchant_transactions(recon_hash);

ALTER TABLE bank_merchant_transactions ENABLE ROW LEVEL SECURITY;
DO $$
BEGIN
    IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename = 'bank_merchant_transactions' AND policyname = 'deny_anon_merchant') THEN
        CREATE POLICY "deny_anon_merchant" ON bank_merchant_transactions FOR ALL TO anon USING (false) WITH CHECK (false);
    END IF;
    IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename = 'bank_merchant_transactions' AND policyname = 'deny_authenticated_merchant') THEN
        CREATE POLICY "deny_authenticated_merchant" ON bank_merchant_transactions FOR ALL TO authenticated USING (false) WITH CHECK (false);
    END IF;
END $$;

-- ── 3. Bank Mutation Transactions Table (Bank Statements) ──────────────────────
CREATE TABLE IF NOT EXISTS bank_mutation_transactions (
    id BIGSERIAL PRIMARY KEY,
    company_key VARCHAR(50) NOT NULL DEFAULT 'eyerizz',
    company TEXT DEFAULT 'Eyerizz Eyewear',
    bank_name TEXT NOT NULL,
    account_number TEXT,
    transaction_date DATE NOT NULL,
    description TEXT,
    amount NUMERIC(18, 2) NOT NULL DEFAULT 0.0,
    mutation_type VARCHAR(10) DEFAULT 'CR',
    balance NUMERIC(18, 2) DEFAULT 0.0,
    filename TEXT,
    recon_hash VARCHAR(64) UNIQUE NOT NULL,
    uploaded_by VARCHAR(100),
    device_id VARCHAR(100),
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_mutation_company_date ON bank_mutation_transactions(company_key, transaction_date);
CREATE INDEX IF NOT EXISTS idx_mutation_bank ON bank_mutation_transactions(bank_name);
CREATE INDEX IF NOT EXISTS idx_mutation_hash ON bank_mutation_transactions(recon_hash);

ALTER TABLE bank_mutation_transactions ENABLE ROW LEVEL SECURITY;
DO $$
BEGIN
    IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename = 'bank_mutation_transactions' AND policyname = 'deny_anon_mutation') THEN
        CREATE POLICY "deny_anon_mutation" ON bank_mutation_transactions FOR ALL TO anon USING (false) WITH CHECK (false);
    END IF;
    IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename = 'bank_mutation_transactions' AND policyname = 'deny_authenticated_mutation') THEN
        CREATE POLICY "deny_authenticated_mutation" ON bank_mutation_transactions FOR ALL TO authenticated USING (false) WITH CHECK (false);
    END IF;
END $$;
