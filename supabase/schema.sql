-- ═══════════════════════════════════════════════════════════════════
-- Hectar Control Tower — Supabase Schema
-- ═══════════════════════════════════════════════════════════════════
-- Run this in Supabase SQL Editor (Dashboard → SQL Editor → New Query)
-- This creates all tables mapped to your Excel workbook structure.
-- ═══════════════════════════════════════════════════════════════════

-- ── 1. Main Trades Table ─────────────────────────────────────────
-- Covers: Main Sheet, Sesame, Almonds, Soybeans, RCN
-- All commodity sheets share the same 85-column structure.
-- source_sheet identifies which sheet each row came from.

CREATE TABLE IF NOT EXISTS trades (
  id            BIGSERIAL PRIMARY KEY,
  source_sheet  TEXT NOT NULL DEFAULT 'Main Sheet',
  excel_row     INTEGER,

  -- ── Identity ──
  product                       TEXT,
  position                      TEXT,
  month                         TEXT,
  commodity_code                TEXT,
  origin                        TEXT,
  variety                       TEXT,
  packer                        TEXT,

  -- ── Specs ──
  yield                         TEXT,
  specs                         TEXT,
  price_per_lbs                 NUMERIC,

  -- ── Logistics ──
  port_of_loading               TEXT,
  port_of_discharge             TEXT,
  no_of_containers              NUMERIC,
  quantity_mt                    NUMERIC,

  -- ── Pricing ──
  purchase_price_per_mt         NUMERIC,
  sales_price_per_mt            NUMERIC,
  purchase_value                NUMERIC,
  sales_value                   NUMERIC,
  gross_margin                  NUMERIC,

  -- ── Expenses ──
  clearance_charges             NUMERIC,
  brokerage                     NUMERIC,
  warehouse_loss                NUMERIC,
  claims_paid                   NUMERIC,
  claims_received               NUMERIC,
  interest_loss                 NUMERIC,
  total_expenses                NUMERIC,

  -- ── P&L ──
  net_profit                    NUMERIC,
  profit_pct                    NUMERIC,

  -- ── Shipping ──
  bl_number                     TEXT,
  remarks                       TEXT,
  bl_date                       TEXT,
  etd                           TEXT,
  eta                           TEXT,
  transit_days                  NUMERIC,

  -- ── Seller Side ──
  payment_terms                 TEXT,
  payment_by                    TEXT,
  seller                        TEXT,
  contract_reference_number     TEXT,
  broker                        TEXT,
  shipper_shipment_period       TEXT,

  -- ── Buyer Side ──
  buyer                         TEXT,
  sales_contract_reference_number TEXT,
  buyer_broker                  TEXT,
  buyer_payment_term            TEXT,
  buyer_shipment_period         TEXT,

  -- ── Outward Payments ──
  advance_paid                  NUMERIC,
  advance_paid_on               TEXT,
  final_payment_paid            NUMERIC,
  final_payment_amount_paid_on  TEXT,
  total_outwards                NUMERIC,
  outward_remaining             NUMERIC,
  outward_adjustment            NUMERIC,
  outward                       TEXT,

  -- ── Inward Payments ──
  advance_from_buyer            NUMERIC,
  advance_received_on           TEXT,
  second_payment_from_buyer     NUMERIC,
  payment_received_on           TEXT,
  third_payment_from_buyer      NUMERIC,
  payment_received_on_2         TEXT,
  total_inwards                 NUMERIC,
  inward_remaining              NUMERIC,
  inward_adjustment             NUMERIC,
  inward                        TEXT,

  -- ── Working Capital ──
  working_capital_days          NUMERIC,

  -- ── Supplier Invoices (1st) ──
  supplier_1_invoice_number     TEXT,
  supplier_1_date               TEXT,
  supplier_1_sales_price        NUMERIC,
  supplier_1_invoice_value      NUMERIC,

  -- ── Supplier Invoices (2nd) ──
  supplier_2_invoice_number     TEXT,
  supplier_2_date               TEXT,
  supplier_2_sales_price        NUMERIC,
  supplier_2_invoice_value      NUMERIC,

  -- ── Supplier Invoices (3rd / Hectar Global) ──
  supplier_3_invoice_number     TEXT,
  supplier_3_date               TEXT,
  supplier_3_sales_price        NUMERIC,
  supplier_3_invoice_value      NUMERIC,

  -- ── Final Invoice ──
  final_invoice_no              TEXT,
  invoice_date                  TEXT,

  -- ── Quality ──
  buy_outturn_nut_count         TEXT,
  sell_outturn_nut_count        TEXT,
  outturn_nut_count_per_rbs     TEXT,
  recutting_outturn_nut_count   TEXT,

  -- ── Customs ──
  boe_date                      TEXT,
  exchange_rate                 NUMERIC,
  trade_no                      TEXT,

  -- ── Metadata ──
  created_at    TIMESTAMPTZ DEFAULT NOW(),
  updated_at    TIMESTAMPTZ DEFAULT NOW()
);

-- Indexes for common queries
CREATE INDEX IF NOT EXISTS idx_trades_source_sheet ON trades(source_sheet);
CREATE INDEX IF NOT EXISTS idx_trades_position ON trades(position);
CREATE INDEX IF NOT EXISTS idx_trades_product ON trades(product);
CREATE INDEX IF NOT EXISTS idx_trades_origin ON trades(origin);
CREATE INDEX IF NOT EXISTS idx_trades_month ON trades(month);
CREATE INDEX IF NOT EXISTS idx_trades_packer ON trades(packer);
CREATE INDEX IF NOT EXISTS idx_trades_trade_no ON trades(trade_no);

-- ── 2. NV Trades (Nhava Sheva port view) ────────────────────────

CREATE TABLE IF NOT EXISTS nv_trades (
  id              BIGSERIAL PRIMARY KEY,
  excel_row       INTEGER,
  position        TEXT,
  port_of_discharge TEXT,
  no_of_containers NUMERIC,
  quantity_mt     NUMERIC,
  purchase_price_per_mt NUMERIC,
  sales_price_per_mt NUMERIC,
  bl_number       TEXT,
  etd             TEXT,
  eta             TEXT,
  seller          TEXT,
  shipper_shipment_period TEXT,
  buyer           TEXT,
  buyer_shipment_period TEXT,
  remarks         TEXT,
  created_at      TIMESTAMPTZ DEFAULT NOW(),
  updated_at      TIMESTAMPTZ DEFAULT NOW()
);

-- ── 3. TUT Trades (Tuticorin port view) ─────────────────────────

CREATE TABLE IF NOT EXISTS tut_trades (
  id              BIGSERIAL PRIMARY KEY,
  excel_row       INTEGER,
  position        TEXT,
  no_of_containers NUMERIC,
  purchase_price_per_mt NUMERIC,
  sales_price_per_mt NUMERIC,
  bl_number       TEXT,
  etd             TEXT,
  eta             TEXT,
  seller          TEXT,
  shipper_shipment_period TEXT,
  buyer           TEXT,
  buyer_shipment_period TEXT,
  remarks         TEXT,
  created_at      TIMESTAMPTZ DEFAULT NOW(),
  updated_at      TIMESTAMPTZ DEFAULT NOW()
);

-- ── 4. Currency Rates ───────────────────────────────────────────

CREATE TABLE IF NOT EXISTS currency_rates (
  id          BIGSERIAL PRIMARY KEY,
  excel_row   INTEGER,
  rate_date   TEXT,
  usd_to_inr  NUMERIC,
  usd_to_tzs  NUMERIC,
  usd_to_xaf  NUMERIC,
  usd_to_ngn  NUMERIC,
  created_at  TIMESTAMPTZ DEFAULT NOW()
);

-- ── 5. Financials (pivot summary) ───────────────────────────────

CREATE TABLE IF NOT EXISTS financials (
  id                    BIGSERIAL PRIMARY KEY,
  excel_row             INTEGER,
  row_label             TEXT,
  sum_no_of_containers  NUMERIC,
  sum_quantity_mt       NUMERIC,
  sum_purchase_value    NUMERIC,
  sum_sales_value       NUMERIC,
  sum_gross_margin      NUMERIC,
  sum_net_profit        NUMERIC,
  created_at            TIMESTAMPTZ DEFAULT NOW()
);

-- ── 6. Mark-to-Market ───────────────────────────────────────────

CREATE TABLE IF NOT EXISTS mtm (
  id                BIGSERIAL PRIMARY KEY,
  excel_row         INTEGER,
  commodity         TEXT,
  position          TEXT,
  quantity          NUMERIC,
  no_of_containers  NUMERIC,
  destination       TEXT,
  average_price     NUMERIC,
  marked_at         NUMERIC,
  mark_to_market    NUMERIC,
  created_at        TIMESTAMPTZ DEFAULT NOW()
);

-- ── 7. Sync Log ─────────────────────────────────────────────────

CREATE TABLE IF NOT EXISTS sync_log (
  id              BIGSERIAL PRIMARY KEY,
  synced_at       TIMESTAMPTZ DEFAULT NOW(),
  status          TEXT NOT NULL DEFAULT 'success',
  sheets_synced   TEXT[],
  total_rows      INTEGER,
  duration_ms     INTEGER,
  error_message   TEXT
);

-- ── 8. Enable Row Level Security (public read, server write) ────

ALTER TABLE trades ENABLE ROW LEVEL SECURITY;
ALTER TABLE nv_trades ENABLE ROW LEVEL SECURITY;
ALTER TABLE tut_trades ENABLE ROW LEVEL SECURITY;
ALTER TABLE currency_rates ENABLE ROW LEVEL SECURITY;
ALTER TABLE financials ENABLE ROW LEVEL SECURITY;
ALTER TABLE mtm ENABLE ROW LEVEL SECURITY;
ALTER TABLE sync_log ENABLE ROW LEVEL SECURITY;

-- Allow public read access (anyone can view the dashboard)
CREATE POLICY "Public read trades" ON trades FOR SELECT USING (true);
CREATE POLICY "Public read nv_trades" ON nv_trades FOR SELECT USING (true);
CREATE POLICY "Public read tut_trades" ON tut_trades FOR SELECT USING (true);
CREATE POLICY "Public read currency_rates" ON currency_rates FOR SELECT USING (true);
CREATE POLICY "Public read financials" ON financials FOR SELECT USING (true);
CREATE POLICY "Public read mtm" ON mtm FOR SELECT USING (true);
CREATE POLICY "Public read sync_log" ON sync_log FOR SELECT USING (true);

-- Service role (used by sync endpoint) can do everything
-- No additional policy needed — service role bypasses RLS
