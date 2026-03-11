"use client";

import { useState, useEffect, useCallback, useMemo } from "react";

// ═══════════════════════════════════════════════════════════════════════════════
// TYPES
// ═══════════════════════════════════════════════════════════════════════════════

interface Trade {
  id: number;
  source_sheet: string;
  product: string | null;
  position: string | null;
  month: string | null;
  commodity_code: string | null;
  origin: string | null;
  variety: string | null;
  packer: string | null;
  quantity_mt: number | null;
  no_of_containers: number | null;
  purchase_price_per_mt: number | null;
  sales_price_per_mt: number | null;
  purchase_value: number | null;
  sales_value: number | null;
  gross_margin: number | null;
  net_profit: number | null;
  profit_pct: number | null;
  total_expenses: number | null;
  bl_number: string | null;
  bl_date: string | null;
  etd: string | null;
  eta: string | null;
  seller: string | null;
  buyer: string | null;
  port_of_loading: string | null;
  port_of_discharge: string | null;
  trade_no: string | null;
  advance_paid: number | null;
  advance_from_buyer: number | null;
  total_outwards: number | null;
  total_inwards: number | null;
  outward_remaining: number | null;
  inward_remaining: number | null;
  working_capital_days: number | null;
  clearance_charges: number | null;
  brokerage: number | null;
  warehouse_loss: number | null;
  claims_paid: number | null;
  claims_received: number | null;
  interest_loss: number | null;
  payment_terms: string | null;
  buyer_payment_term: string | null;
  transit_days: number | null;
  remarks: string | null;
  advance_paid_on: string | null;
  final_payment_paid: number | null;
  final_payment_amount_paid_on: string | null;
  advance_received_on: string | null;
  second_payment_from_buyer: number | null;
  payment_received_on: string | null;
  third_payment_from_buyer: number | null;
  contract_reference_number: string | null;
  sales_contract_reference_number: string | null;
  broker: string | null;
  buyer_broker: string | null;
  [key: string]: unknown;
}

interface SyncStatus {
  synced_at: string;
  status: string;
  total_rows: number;
  duration_ms: number;
  sheets_synced: string[];
}

type DealStage = "need_buyer" | "collecting_payment" | "waiting_to_ship" | "at_sea" | "done";
type Page = "today" | "all_trades" | "open_positions" | "payments_out" | "payments_in" | "pnl";

// ═══════════════════════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════════════════════

function n(v: unknown): number { return Number(v) || 0; }
function has(v: unknown): boolean { return v !== null && v !== undefined && String(v).trim() !== "" && String(v).trim() !== "—"; }

function fmtK(num: number | null | undefined): string {
  if (num === null || num === undefined || isNaN(num)) return "—";
  if (Math.abs(num) >= 1_000_000) return "$" + (num / 1_000_000).toFixed(2) + "M";
  if (Math.abs(num) >= 1_000) return "$" + (num / 1_000).toFixed(0) + "K";
  return "$" + num.toFixed(0);
}

function fmtFull(num: number | null | undefined): string {
  if (num === null || num === undefined || isNaN(num)) return "—";
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 }).format(num);
}

function fmt(num: number | null | undefined): string {
  if (num === null || num === undefined || isNaN(num)) return "—";
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(num);
}

function fmtDate(d: string | null | undefined): string {
  if (!d) return "—";
  try {
    const date = new Date(d);
    if (isNaN(date.getTime())) return d;
    return date.toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "2-digit" });
  } catch { return d; }
}

function timeAgo(dateStr: string): string {
  const diff = Date.now() - new Date(dateStr).getTime();
  const mins = Math.floor(diff / 60000);
  if (mins < 1) return "just now";
  if (mins < 60) return `${mins}m ago`;
  const hrs = Math.floor(mins / 60);
  if (hrs < 24) return `${hrs}h ago`;
  return `${Math.floor(hrs / 24)}d ago`;
}

function sumBy(trades: Trade[], field: keyof Trade): number {
  return trades.reduce((s, t) => s + n(t[field]), 0);
}

// ═══════════════════════════════════════════════════════════════════════════════
// DEAL LIFECYCLE CLASSIFICATION
// ═══════════════════════════════════════════════════════════════════════════════
//
// A commodity trade moves through these stages:
//   1. Need a Buyer — purchase contract exists but no sales contract yet
//   2. Waiting to Ship — both sides contracted, awaiting shipment (no B/L)
//   3. At Sea — B/L issued, cargo in transit
//   4. Collecting Payment — goods arrived/delivered, buyer owes money
//   5. Done — fully settled (position = "Inactive")
//
// Classification logic:
//   "Inactive" position → Done (all payments settled)
//   "Sold" position → Collecting Payment (at port, financial close)
//   No buyer → Need a Buyer (open exposure)
//   Has buyer + no B/L → Waiting to Ship
//   Has B/L → At Sea (in transit)

function classifyDeal(t: Trade): DealStage {
  const pos = (t.position || "").toLowerCase().trim();
  if (pos === "inactive") return "done";
  if (pos === "sold") return "collecting_payment";
  if (!has(t.buyer)) return "need_buyer";
  if (!has(t.bl_number)) return "waiting_to_ship";
  return "at_sea";
}

const STAGE_CONFIG: Record<DealStage, { label: string; color: string; dot: string; desc: string }> = {
  need_buyer:         { label: "Need a Buyer",       color: "text-violet-700", dot: "bg-violet-500", desc: "Open exposure" },
  collecting_payment: { label: "Collecting Payment",  color: "text-amber-700",  dot: "bg-amber-500",  desc: "Outstanding receivables" },
  waiting_to_ship:    { label: "Waiting to Ship",     color: "text-emerald-700", dot: "bg-emerald-500", desc: "B/L not yet issued" },
  at_sea:             { label: "At Sea",              color: "text-sky-700",    dot: "bg-sky-500",    desc: "" },
  done:               { label: "Done",                color: "text-gray-500",   dot: "bg-gray-400",   desc: "Fully settled" },
};

// ═══════════════════════════════════════════════════════════════════════════════
// PAYMENT TERM PARSING
// ═══════════════════════════════════════════════════════════════════════════════

interface PaymentScheduleItem {
  label: string;
  amount: number;
  dueDate: Date | null;
  dueLabel: string;
}

function parseDate(d: string | null | undefined): Date | null {
  if (!d) return null;
  const date = new Date(d);
  return isNaN(date.getTime()) ? null : date;
}

function addDays(d: Date, days: number): Date {
  const r = new Date(d);
  r.setDate(r.getDate() + days);
  return r;
}

function monthLabel(d: Date): string {
  return d.toLocaleDateString("en-US", { month: "short", year: "numeric" }).toUpperCase();
}

/**
 * Parse seller payment terms into approximate due dates.
 * Terms from "Payment Terms" column.
 */
function parseSellerPaymentSchedule(t: Trade): PaymentScheduleItem[] {
  const term = (t.payment_terms || "").trim();
  const purchaseVal = n(t.purchase_value);
  if (!term || purchaseVal === 0) return [];

  const etd = parseDate(t.etd);
  const eta = parseDate(t.eta);
  const blDate = parseDate(t.bl_date);
  const termLower = term.toLowerCase();

  // Split terms: "10+90", "5%+95%", "10%+90%", "15%+85%"
  const splitMatch = term.match(/(\d+)\s*[%]?\s*\+\s*(\d+)\s*[%]?/);
  if (splitMatch) {
    const pct1 = parseInt(splitMatch[1]) / 100;
    const pct2 = parseInt(splitMatch[2]) / 100;
    return [
      { label: `${splitMatch[1]}% advance`, amount: purchaseVal * pct1, dueDate: etd, dueLabel: etd ? `On ETD (${fmtDate(t.etd)})` : "On ETD" },
      { label: `${splitMatch[2]}% balance`, amount: purchaseVal * pct2, dueDate: eta, dueLabel: eta ? `On arrival (${fmtDate(t.eta)})` : "On arrival" },
    ];
  }

  // "100% CAD" / "CAD" / "100% on Loading" / "Pre Shipment"
  if (termLower.includes("cad") || termLower.includes("loading") || termLower.includes("pre shipment") || termLower.includes("pre-shipment")) {
    return [{ label: "100% CAD", amount: purchaseVal, dueDate: etd, dueLabel: etd ? `On ETD (${fmtDate(t.etd)})` : "On/before ETD" }];
  }

  // "100% TT N days Before vessel arrival"
  const ttBeforeMatch = termLower.match(/(\d+)\s*days?\s*before/);
  if (ttBeforeMatch && eta) {
    const daysBefore = parseInt(ttBeforeMatch[1]);
    const due = addDays(eta, -daysBefore);
    return [{ label: `TT ${daysBefore}d before arrival`, amount: purchaseVal, dueDate: due, dueLabel: `${fmtDate(due.toISOString())}` }];
  }

  // "On BL Copies" / "100% TT upon Copy BL"
  if (termLower.includes("bl cop") || termLower.includes("copy bl") || termLower.includes("upon bl")) {
    const due = blDate ? addDays(blDate, 2) : null;
    return [{ label: "On B/L copies", amount: purchaseVal, dueDate: due, dueLabel: due ? `BL+2d (${fmtDate(due.toISOString())})` : "On B/L receipt" }];
  }

  // "DP at Sight" / "through Bank"
  if (termLower.includes("dp at sight") || termLower.includes("through bank")) {
    return [{ label: "DP at sight", amount: purchaseVal, dueDate: eta, dueLabel: eta ? `On arrival (${fmtDate(t.eta)})` : "On arrival" }];
  }

  // Default: assume due on ETD
  return [{ label: term, amount: purchaseVal, dueDate: etd || eta, dueLabel: etd ? fmtDate(t.etd) : (eta ? fmtDate(t.eta) : "TBD") }];
}

/**
 * Parse buyer payment terms into approximate expected receipt dates.
 * Terms from "Buyer Payment Term" column.
 */
function parseBuyerPaymentSchedule(t: Trade): PaymentScheduleItem[] {
  const term = (t.buyer_payment_term || "").trim();
  const salesVal = n(t.sales_value);
  if (!term || salesVal === 0) return [];

  const eta = parseDate(t.eta);
  const blDate = parseDate(t.bl_date);
  const termLower = term.toLowerCase();

  // Split terms: "HSS 15%+85%", "10% ADVANCE AND 90% CAD"
  const splitMatch = term.match(/(\d+)\s*[%]?\s*(?:\+|and)\s*(\d+)\s*[%]?/i);
  if (splitMatch) {
    const pct1 = parseInt(splitMatch[1]) / 100;
    const pct2 = parseInt(splitMatch[2]) / 100;
    const advDate = parseDate(t.advance_received_on);
    return [
      { label: `${splitMatch[1]}% advance`, amount: salesVal * pct1, dueDate: advDate, dueLabel: advDate ? fmtDate(t.advance_received_on) : "On contract" },
      { label: `${splitMatch[2]}% balance`, amount: salesVal * pct2, dueDate: eta, dueLabel: eta ? `On arrival (${fmtDate(t.eta)})` : "On arrival" },
    ];
  }

  // "On Delivery" / "Payment after Delivery" / "After Delivery"
  if (termLower.includes("delivery") || termLower.includes("after delivery")) {
    const due = eta ? addDays(eta, 1) : null;
    return [{ label: "On delivery", amount: salesVal, dueDate: due, dueLabel: due ? fmtDate(due.toISOString()) : "On delivery" }];
  }

  // "N days after delivery"
  const daysAfterMatch = termLower.match(/(\d+)\s*days?\s*after/);
  if (daysAfterMatch && eta) {
    const days = parseInt(daysAfterMatch[1]);
    const due = addDays(eta, days);
    return [{ label: `${days}d after delivery`, amount: salesVal, dueDate: due, dueLabel: fmtDate(due.toISOString()) }];
  }

  // "100% CAD on Vessel Arrival" / "CAD"
  if (termLower.includes("cad") || termLower.includes("vessel arrival")) {
    return [{ label: "CAD on arrival", amount: salesVal, dueDate: eta, dueLabel: eta ? fmtDate(t.eta) : "On arrival" }];
  }

  // "DP at Sight"
  if (termLower.includes("dp at sight")) {
    return [{ label: "DP at sight", amount: salesVal, dueDate: eta, dueLabel: eta ? fmtDate(t.eta) : "On arrival" }];
  }

  // "100% on copy BL" / "TT Payment" / "TT upon BL"
  if (termLower.includes("copy bl") || termLower.includes("upon bl") || termLower.includes("tt payment")) {
    const due = blDate ? addDays(blDate, 4) : null;
    return [{ label: "On B/L copies", amount: salesVal, dueDate: due, dueLabel: due ? fmtDate(due.toISOString()) : "On B/L" }];
  }

  // "Through Bank" / "DA" / "DA 45 DAYS"
  const daMatch = termLower.match(/da\s*(\d+)/);
  if (daMatch && eta) {
    const days = parseInt(daMatch[1]);
    const due = addDays(eta, days);
    return [{ label: `DA ${days} days`, amount: salesVal, dueDate: due, dueLabel: fmtDate(due.toISOString()) }];
  }
  if (termLower.includes("through bank") || termLower === "da") {
    const due = eta ? addDays(eta, 5) : null;
    return [{ label: "Through bank", amount: salesVal, dueDate: due, dueLabel: due ? fmtDate(due.toISOString()) : "On arrival" }];
  }

  // Default
  return [{ label: term, amount: salesVal, dueDate: eta, dueLabel: eta ? fmtDate(t.eta) : "TBD" }];
}

// ═══════════════════════════════════════════════════════════════════════════════
// DEAL DETAIL MODAL
// ═══════════════════════════════════════════════════════════════════════════════

function DealModal({ trade: t, onClose }: { trade: Trade; onClose: () => void }) {
  const stage = classifyDeal(t);
  const cfg = STAGE_CONFIG[stage];
  const grossMargin = n(t.gross_margin);
  const totalExp = n(t.total_expenses);
  const netProfit = n(t.net_profit);
  const purchaseVal = n(t.purchase_value);
  const salesVal = n(t.sales_value);
  const marginPct = salesVal > 0 ? ((netProfit / salesVal) * 100).toFixed(2) : "—";

  const expenses = [
    { label: "Clearance Charges", val: n(t.clearance_charges) },
    { label: "Brokerage", val: n(t.brokerage) },
    { label: "Warehouse Loss", val: n(t.warehouse_loss) },
    { label: "Claims Paid", val: n(t.claims_paid) },
    { label: "Claims Received", val: -n(t.claims_received) },
    { label: "Interest Loss", val: n(t.interest_loss) },
  ].filter(e => e.val !== 0);

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center bg-black/30 backdrop-blur-sm" onClick={onClose}>
      <div className="bg-white rounded-t-2xl sm:rounded-2xl max-w-lg w-full mx-0 sm:mx-4 max-h-[90vh] overflow-y-auto shadow-2xl" onClick={e => e.stopPropagation()}>
        {/* Header */}
        <div className="sticky top-0 bg-white border-b border-gray-100 px-5 py-4 flex items-start justify-between rounded-t-2xl">
          <div>
            <div className="flex items-center gap-2 mb-0.5">
              <h2 className="text-base font-bold text-gray-900">{t.product || "Unknown"}</h2>
              <span className={`inline-flex items-center gap-1.5 px-2 py-0.5 rounded-full text-[11px] font-medium bg-gray-100 ${cfg.color}`}>
                <span className={`w-1.5 h-1.5 rounded-full ${cfg.dot}`} />
                {cfg.label}
              </span>
            </div>
            <div className="flex flex-wrap items-center gap-x-3 gap-y-0.5 text-xs text-gray-400">
              {t.trade_no && <span>{t.commodity_code || t.source_sheet}-{t.trade_no}</span>}
              {t.origin && <span>{t.origin}</span>}
              {t.variety && <span>{t.variety}</span>}
            </div>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 text-xl p-1">&times;</button>
        </div>

        <div className="px-5 py-4 space-y-5">
          {/* Counterparties */}
          <div className="grid grid-cols-2 gap-4">
            <div>
              <div className="text-[10px] text-gray-400 uppercase tracking-wider mb-1">Seller</div>
              <div className="text-sm font-medium text-gray-900">{t.seller || "—"}</div>
              {t.payment_terms && <div className="text-[11px] text-gray-400 mt-0.5">{t.payment_terms}</div>}
            </div>
            <div>
              <div className="text-[10px] text-gray-400 uppercase tracking-wider mb-1">Buyer</div>
              <div className="text-sm font-medium text-gray-900">{t.buyer || <span className="text-amber-600 italic">Not yet assigned</span>}</div>
              {t.buyer_payment_term && <div className="text-[11px] text-gray-400 mt-0.5">{t.buyer_payment_term}</div>}
            </div>
          </div>

          {/* Trade summary */}
          <div className="grid grid-cols-4 gap-2">
            {[
              { label: "Qty (MT)", value: fmt(t.quantity_mt) },
              { label: "Containers", value: fmt(t.no_of_containers) },
              { label: "Buy $/MT", value: fmtFull(t.purchase_price_per_mt) },
              { label: "Sell $/MT", value: fmtFull(t.sales_price_per_mt) },
            ].map(({ label, value }) => (
              <div key={label} className="bg-gray-50 rounded-lg px-3 py-2">
                <div className="text-[10px] text-gray-400">{label}</div>
                <div className="text-sm font-semibold text-gray-900">{value}</div>
              </div>
            ))}
          </div>

          {/* Margin Waterfall */}
          <div>
            <h3 className="text-xs font-semibold text-gray-700 mb-2">Margin Breakdown</h3>
            <div className="bg-gray-50 rounded-lg p-4 space-y-1.5 text-xs">
              <div className="flex justify-between"><span className="text-gray-500">Purchase Value</span><span className="font-mono text-gray-700">{fmtFull(purchaseVal)}</span></div>
              <div className="flex justify-between"><span className="text-gray-500">Sales Value</span><span className="font-mono text-gray-700">{fmtFull(salesVal)}</span></div>
              <div className="border-t border-gray-200 my-1" />
              <div className="flex justify-between font-medium"><span className="text-gray-700">Gross Margin</span><span className={`font-mono ${grossMargin >= 0 ? "text-emerald-700" : "text-red-600"}`}>{fmtFull(grossMargin)}</span></div>
              {expenses.length > 0 && expenses.map(({ label, val }) => (
                <div key={label} className="flex justify-between pl-3"><span className="text-gray-400">{label}</span><span className="font-mono text-red-500">{fmtFull(val)}</span></div>
              ))}
              {totalExp !== 0 && <div className="flex justify-between pl-3"><span className="text-gray-500">Total Expenses</span><span className="font-mono text-red-600">{fmtFull(totalExp)}</span></div>}
              <div className="border-t border-gray-200 my-1" />
              <div className="flex justify-between text-sm font-bold"><span className="text-gray-900">Net Profit</span><span className={`font-mono ${netProfit >= 0 ? "text-emerald-700" : "text-red-600"}`}>{fmtFull(netProfit)}</span></div>
              <div className="flex justify-between"><span className="text-gray-400">Net Margin</span><span className={`font-mono ${netProfit >= 0 ? "text-emerald-700" : "text-red-600"}`}>{marginPct}%</span></div>
            </div>
          </div>

          {/* Cash Flow */}
          {(n(t.total_outwards) > 0 || n(t.total_inwards) > 0) && (
            <div>
              <h3 className="text-xs font-semibold text-gray-700 mb-2">Cash Flow</h3>
              <div className="bg-gray-50 rounded-lg p-4 space-y-1.5 text-xs">
                {n(t.advance_paid) > 0 && <div className="flex justify-between"><span className="text-gray-500">Advance to Seller {t.advance_paid_on && <span className="text-gray-300">({fmtDate(t.advance_paid_on)})</span>}</span><span className="font-mono text-red-500">{fmtFull(t.advance_paid)}</span></div>}
                {n(t.final_payment_paid) > 0 && <div className="flex justify-between"><span className="text-gray-500">Final to Seller {t.final_payment_amount_paid_on && <span className="text-gray-300">({fmtDate(t.final_payment_amount_paid_on)})</span>}</span><span className="font-mono text-red-500">{fmtFull(t.final_payment_paid)}</span></div>}
                <div className="flex justify-between font-medium"><span className="text-gray-600">Total Paid Out</span><span className="font-mono text-red-600">{fmtFull(t.total_outwards)}</span></div>
                <div className="border-t border-gray-200 my-1" />
                {n(t.advance_from_buyer) > 0 && <div className="flex justify-between"><span className="text-gray-500">Advance from Buyer {t.advance_received_on && <span className="text-gray-300">({fmtDate(t.advance_received_on)})</span>}</span><span className="font-mono text-emerald-600">{fmtFull(t.advance_from_buyer)}</span></div>}
                {n(t.second_payment_from_buyer) > 0 && <div className="flex justify-between"><span className="text-gray-500">2nd from Buyer {t.payment_received_on && <span className="text-gray-300">({fmtDate(t.payment_received_on)})</span>}</span><span className="font-mono text-emerald-600">{fmtFull(t.second_payment_from_buyer)}</span></div>}
                {n(t.third_payment_from_buyer) > 0 && <div className="flex justify-between"><span className="text-gray-500">3rd from Buyer</span><span className="font-mono text-emerald-600">{fmtFull(t.third_payment_from_buyer)}</span></div>}
                <div className="flex justify-between font-medium"><span className="text-gray-600">Total Received</span><span className="font-mono text-emerald-700">{fmtFull(t.total_inwards)}</span></div>
                <div className="border-t border-gray-200 my-1" />
                {n(t.outward_remaining) > 0 && <div className="flex justify-between"><span className="text-gray-500">Owed to Seller</span><span className="font-mono text-amber-600 font-medium">{fmtFull(t.outward_remaining)}</span></div>}
                {n(t.inward_remaining) > 0 && <div className="flex justify-between"><span className="text-gray-500">Owed by Buyer</span><span className="font-mono text-amber-600 font-medium">{fmtFull(t.inward_remaining)}</span></div>}
                {n(t.working_capital_days) > 0 && <div className="flex justify-between"><span className="text-gray-500">Working Capital Days</span><span className="font-mono text-gray-700">{fmt(t.working_capital_days)}d</span></div>}
              </div>
            </div>
          )}

          {/* Logistics */}
          {(has(t.bl_number) || has(t.etd) || has(t.eta)) && (
            <div>
              <h3 className="text-xs font-semibold text-gray-700 mb-2">Logistics</h3>
              <div className="bg-gray-50 rounded-lg p-4 grid grid-cols-2 gap-2 text-xs">
                {has(t.port_of_loading) && <div><span className="text-gray-400">Load: </span><span className="text-gray-700">{t.port_of_loading}</span></div>}
                {has(t.port_of_discharge) && <div><span className="text-gray-400">Discharge: </span><span className="text-gray-700">{t.port_of_discharge}</span></div>}
                {has(t.bl_number) && <div><span className="text-gray-400">B/L: </span><span className="text-gray-700 font-mono">{t.bl_number}</span></div>}
                {has(t.bl_date) && <div><span className="text-gray-400">B/L Date: </span><span className="text-gray-700">{fmtDate(t.bl_date)}</span></div>}
                {has(t.etd) && <div><span className="text-gray-400">ETD: </span><span className="text-gray-700">{fmtDate(t.etd)}</span></div>}
                {has(t.eta) && <div><span className="text-gray-400">ETA: </span><span className="text-gray-700">{fmtDate(t.eta)}</span></div>}
                {has(t.transit_days) && <div><span className="text-gray-400">Transit: </span><span className="text-gray-700">{t.transit_days}d</span></div>}
              </div>
            </div>
          )}

          {/* Contract refs */}
          {(has(t.contract_reference_number) || has(t.sales_contract_reference_number)) && (
            <div className="text-xs text-gray-400 space-y-0.5">
              {has(t.contract_reference_number) && <div>Purchase Contract: {t.contract_reference_number}</div>}
              {has(t.sales_contract_reference_number) && <div>Sales Contract: {t.sales_contract_reference_number}</div>}
            </div>
          )}

          {t.remarks && <div className="text-xs text-gray-500 bg-gray-50 rounded-lg p-3 italic">{t.remarks}</div>}
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// DEAL ROW (mobile-friendly card for deal lists)
// ═══════════════════════════════════════════════════════════════════════════════

function DealRow({ trade: t, onClick, showAmount }: { trade: Trade; onClick: () => void; showAmount?: "net_profit" | "purchase_value" | "inward_remaining" | "outward_remaining" }) {
  const amt = showAmount ? n(t[showAmount]) : null;
  return (
    <button onClick={onClick} className="w-full text-left flex items-center gap-3 px-4 py-3 hover:bg-gray-50 border-b border-gray-100 transition-colors group">
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2 mb-0.5">
          <span className="text-sm font-semibold text-gray-900">{t.product || "—"}</span>
          {t.trade_no && <span className="text-[11px] text-gray-400 font-mono">{t.commodity_code || t.source_sheet}-{t.trade_no}</span>}
        </div>
        <div className="text-xs text-gray-500 truncate">
          {t.seller || "—"} → {t.buyer || <span className="text-amber-600">No buyer</span>}
        </div>
        <div className="flex flex-wrap gap-x-3 gap-y-0 text-[11px] text-gray-400 mt-0.5">
          {has(t.quantity_mt) && <span>{fmt(t.quantity_mt)} MT</span>}
          {has(t.etd) && <span>ETD {fmtDate(t.etd)}</span>}
          {has(t.eta) && <span>ETA {fmtDate(t.eta)}</span>}
          {has(t.origin) && <span>{t.origin}</span>}
        </div>
      </div>
      {amt !== null && amt !== 0 && (
        <div className="text-right shrink-0">
          <div className={`text-sm font-bold font-mono ${showAmount === "net_profit" ? (amt >= 0 ? "text-emerald-700" : "text-red-600") : "text-gray-900"}`}>{fmtK(amt)}</div>
          {has(t.eta) && showAmount !== "net_profit" && <div className="text-[10px] text-gray-400">{fmtDate(t.eta)}</div>}
        </div>
      )}
      <svg className="w-4 h-4 text-gray-300 group-hover:text-gray-500 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" /></svg>
    </button>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// PAGES
// ═══════════════════════════════════════════════════════════════════════════════

// ─── Today's Work ─────────────────────────────────────────────────

function TodaysWorkPage({ trades, stageGroups, onSelect }: {
  trades: Trade[];
  stageGroups: Record<DealStage, Trade[]>;
  onSelect: (t: Trade) => void;
}) {
  const [activeStage, setActiveStage] = useState<DealStage | null>(null);
  const needAttention = trades.filter(t => classifyDeal(t) !== "done");

  // Collecting payment trades needing follow-up
  const collectingPayment = stageGroups.collecting_payment.filter(t => n(t.inward_remaining) > 0);
  const outstandingAmount = sumBy(collectingPayment, "inward_remaining");

  // Find ETA range for at_sea
  const atSeaETAs = stageGroups.at_sea
    .map(t => t.eta).filter(Boolean)
    .map(d => new Date(d!))
    .filter(d => !isNaN(d.getTime()))
    .sort((a, b) => a.getTime() - b.getTime());
  const etaRange = atSeaETAs.length > 0
    ? `ETA ${atSeaETAs[0].toLocaleDateString("en-US", { month: "short" })} – ${atSeaETAs[atSeaETAs.length - 1].toLocaleDateString("en-US", { month: "short" })}`
    : "";

  const stageDesc: Record<DealStage, string> = {
    need_buyer: `${fmtK(sumBy(stageGroups.need_buyer, "purchase_value"))} exposure`,
    collecting_payment: outstandingAmount > 0 ? `${fmtK(outstandingAmount)} outstanding` : "All collected",
    waiting_to_ship: "B/L not yet issued",
    at_sea: etaRange,
    done: "Fully settled",
  };

  const stages: DealStage[] = ["need_buyer", "collecting_payment", "waiting_to_ship", "at_sea", "done"];

  const visibleTrades = activeStage ? stageGroups[activeStage] : needAttention;
  const showAmount = activeStage === "collecting_payment" ? "inward_remaining" as const
    : activeStage === "need_buyer" ? "purchase_value" as const
    : "net_profit" as const;

  return (
    <div>
      <div className="px-4 sm:px-6 pt-5 pb-3">
        <h1 className="text-xl font-bold text-gray-900">Today&apos;s Work</h1>
        <p className="text-sm text-gray-400 mt-0.5">{needAttention.length} trades need attention · updated {new Date().toLocaleTimeString("en-US", { hour: "2-digit", minute: "2-digit" })}</p>
      </div>

      {/* Stage summary cards */}
      <div className="px-4 sm:px-6 pb-4 overflow-x-auto">
        <div className="flex gap-3 min-w-max">
          {stages.map(stage => {
            const count = stageGroups[stage].length;
            const cfg = STAGE_CONFIG[stage];
            const isActive = activeStage === stage;
            return (
              <button key={stage} onClick={() => setActiveStage(isActive ? null : stage)}
                className={`flex-shrink-0 rounded-xl px-5 py-4 text-left transition-all border ${isActive ? "border-gray-300 bg-white shadow-sm" : "border-gray-100 bg-gray-50 hover:bg-white hover:border-gray-200"}`}
                style={{ minWidth: 150 }}>
                <div className={`flex items-center gap-1.5 mb-1`}>
                  <span className={`w-2 h-2 rounded-full ${cfg.dot}`} />
                  <span className={`text-[11px] font-medium ${cfg.color}`}>{cfg.label}</span>
                </div>
                <div className="text-2xl font-bold text-gray-900">{count}</div>
                <div className="text-[11px] text-gray-400 mt-0.5">{stageDesc[stage]}</div>
              </button>
            );
          })}
        </div>
      </div>

      {/* Actionable trades */}
      {!activeStage && collectingPayment.length > 0 && (
        <div className="border-t border-gray-100">
          <div className="px-4 sm:px-6 py-3 flex items-center justify-between">
            <div className="flex items-center gap-2">
              <span className="w-1 h-4 bg-amber-500 rounded-full" />
              <span className="text-sm font-semibold text-gray-700">Collecting Payment</span>
              <span className="text-xs text-gray-400">{collectingPayment.length} trades</span>
            </div>
            <span className="text-sm font-bold text-gray-900">{fmtK(outstandingAmount)}</span>
          </div>
          {collectingPayment.map((t, i) => (
            <DealRow key={t.id || i} trade={t} onClick={() => onSelect(t)} showAmount="inward_remaining" />
          ))}
        </div>
      )}

      {/* Active stage filter list */}
      {activeStage && (
        <div className="border-t border-gray-100">
          <div className="px-4 sm:px-6 py-3 flex items-center gap-2">
            <span className={`w-1 h-4 rounded-full ${STAGE_CONFIG[activeStage].dot}`} />
            <span className="text-sm font-semibold text-gray-700">{STAGE_CONFIG[activeStage].label}</span>
            <span className="text-xs text-gray-400">{visibleTrades.length} trades</span>
          </div>
          {visibleTrades.map((t, i) => (
            <DealRow key={t.id || i} trade={t} onClick={() => onSelect(t)} showAmount={showAmount} />
          ))}
        </div>
      )}

      {/* If no active stage and no collecting payment, show all needing attention */}
      {!activeStage && collectingPayment.length === 0 && needAttention.length > 0 && (
        <div className="border-t border-gray-100">
          {needAttention.slice(0, 20).map((t, i) => (
            <DealRow key={t.id || i} trade={t} onClick={() => onSelect(t)} showAmount="purchase_value" />
          ))}
        </div>
      )}
    </div>
  );
}

// ─── All Trades / Open Positions ──────────────────────────────────

function TradesListPage({ trades, title, subtitle, onSelect }: {
  trades: Trade[];
  title: string;
  subtitle: string;
  onSelect: (t: Trade) => void;
}) {
  const [search, setSearch] = useState("");
  const [sortBy, setSortBy] = useState<"default" | "profit" | "value">("default");

  const filtered = useMemo(() => {
    let list = [...trades];
    if (search) {
      const q = search.toLowerCase();
      list = list.filter(t =>
        (t.product || "").toLowerCase().includes(q) ||
        (t.seller || "").toLowerCase().includes(q) ||
        (t.buyer || "").toLowerCase().includes(q) ||
        (t.origin || "").toLowerCase().includes(q) ||
        (t.trade_no || "").toLowerCase().includes(q)
      );
    }
    if (sortBy === "profit") list.sort((a, b) => n(b.net_profit) - n(a.net_profit));
    if (sortBy === "value") list.sort((a, b) => n(b.purchase_value) - n(a.purchase_value));
    return list;
  }, [trades, search, sortBy]);

  return (
    <div>
      <div className="px-4 sm:px-6 pt-5 pb-3">
        <h1 className="text-xl font-bold text-gray-900">{title}</h1>
        <p className="text-sm text-gray-400 mt-0.5">{subtitle}</p>
      </div>
      <div className="px-4 sm:px-6 pb-3 flex gap-2">
        <input type="text" placeholder="Search commodity, seller, buyer..." value={search} onChange={e => setSearch(e.target.value)}
          className="flex-1 px-3 py-2 text-sm border border-gray-200 rounded-lg bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent" />
        <select value={sortBy} onChange={e => setSortBy(e.target.value as typeof sortBy)}
          className="px-3 py-2 text-sm border border-gray-200 rounded-lg bg-white text-gray-600">
          <option value="default">Default</option>
          <option value="profit">By Profit</option>
          <option value="value">By Value</option>
        </select>
      </div>
      <div className="border-t border-gray-100">
        {filtered.map((t, i) => (
          <DealRow key={t.id || i} trade={t} onClick={() => onSelect(t)} showAmount="net_profit" />
        ))}
        {filtered.length === 0 && <div className="px-6 py-8 text-center text-sm text-gray-400">No trades found</div>}
      </div>
    </div>
  );
}

// ─── Payments Out ─────────────────────────────────────────────────

function PaymentsOutPage({ trades, onSelect }: { trades: Trade[]; onSelect: (t: Trade) => void }) {
  const active = trades.filter(t => classifyDeal(t) !== "done");

  // Build payment schedule by supplier by month
  const scheduleBySupplier = useMemo(() => {
    const map = new Map<string, { supplier: string; months: Map<string, { amount: number; terms: string; trades: Trade[] }> }>();

    for (const t of active) {
      const supplier = t.seller || "Unknown";
      const items = parseSellerPaymentSchedule(t);
      if (!map.has(supplier)) map.set(supplier, { supplier, months: new Map() });
      const entry = map.get(supplier)!;

      for (const item of items) {
        const month = item.dueDate ? monthLabel(item.dueDate) : "TBD";
        if (!entry.months.has(month)) entry.months.set(month, { amount: 0, terms: "", trades: [] });
        const m = entry.months.get(month)!;
        m.amount += item.amount;
        m.terms = m.terms || item.label;
        m.trades.push(t);
      }
    }

    return [...map.values()].sort((a, b) => {
      const totalA = [...a.months.values()].reduce((s, m) => s + m.amount, 0);
      const totalB = [...b.months.values()].reduce((s, m) => s + m.amount, 0);
      return totalB - totalA;
    });
  }, [active]);

  // Get all months
  const allMonths = useMemo(() => {
    const months = new Set<string>();
    for (const s of scheduleBySupplier) {
      for (const m of s.months.keys()) months.add(m);
    }
    return [...months].sort((a, b) => {
      const da = new Date(a);
      const db = new Date(b);
      if (isNaN(da.getTime()) || isNaN(db.getTime())) return a.localeCompare(b);
      return da.getTime() - db.getTime();
    });
  }, [scheduleBySupplier]);

  // Monthly totals
  const monthTotals = allMonths.map(m => {
    let total = 0;
    for (const s of scheduleBySupplier) {
      total += s.months.get(m)?.amount || 0;
    }
    return total;
  });

  const grandTotal = monthTotals.reduce((s, v) => s + v, 0);

  // Top-level summary: This month, next month, rest
  const now = new Date();
  const thisMonth = monthLabel(now);
  const nextMonthDate = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const nextMonth = monthLabel(nextMonthDate);

  const thisMonthTotal = monthTotals[allMonths.indexOf(thisMonth)] || 0;
  const nextMonthTotal = monthTotals[allMonths.indexOf(nextMonth)] || 0;
  const restTotal = grandTotal - thisMonthTotal - nextMonthTotal;

  return (
    <div>
      <div className="px-4 sm:px-6 pt-5 pb-3">
        <h1 className="text-xl font-bold text-gray-900">Payments Out</h1>
        <p className="text-sm text-gray-400 mt-0.5">Supplier payment schedule · Approximate dates from payment terms</p>
      </div>

      {/* Summary cards */}
      <div className="px-4 sm:px-6 pb-4 grid grid-cols-3 gap-3">
        <div className="bg-gray-50 border border-gray-100 rounded-xl px-4 py-3">
          <div className="text-[10px] text-gray-400 uppercase tracking-wider">Due by end of {thisMonth}</div>
          <div className="text-xl font-bold text-gray-900 mt-1">{fmtK(thisMonthTotal)}</div>
          <div className="text-[11px] text-gray-400">Pre-shipment & B/L-triggered</div>
        </div>
        <div className="bg-gray-50 border border-gray-100 rounded-xl px-4 py-3">
          <div className="text-[10px] text-gray-400 uppercase tracking-wider">Due in {nextMonth}</div>
          <div className="text-xl font-bold text-gray-900 mt-1">{fmtK(nextMonthTotal)}</div>
          <div className="text-[11px] text-gray-400">On-arrival & document payments</div>
        </div>
        <div className="bg-gray-50 border border-gray-100 rounded-xl px-4 py-3">
          <div className="text-[10px] text-gray-400 uppercase tracking-wider">Later</div>
          <div className="text-xl font-bold text-gray-900 mt-1">{fmtK(restTotal)}</div>
          <div className="text-[11px] text-gray-400">Future shipment obligations</div>
        </div>
      </div>

      {/* Schedule table */}
      <div className="px-4 sm:px-6 pb-4">
        <h2 className="text-sm font-semibold text-gray-700 mb-2">Supplier Payment Schedule <span className="font-normal text-gray-400">By month · Approximate dates from payment terms</span></h2>
        <div className="border border-gray-200 rounded-xl overflow-x-auto">
          <table className="w-full text-xs">
            <thead>
              <tr className="bg-gray-50 border-b border-gray-200">
                <th className="px-4 py-2.5 text-left font-semibold text-gray-600">Supplier</th>
                {allMonths.map(m => <th key={m} className="px-4 py-2.5 text-right font-semibold text-gray-600 whitespace-nowrap">{m}</th>)}
                <th className="px-4 py-2.5 text-right font-semibold text-gray-600">Total</th>
              </tr>
            </thead>
            <tbody>
              {scheduleBySupplier.map((s, i) => {
                const total = [...s.months.values()].reduce((sum, m) => sum + m.amount, 0);
                return (
                  <tr key={i} className="border-b border-gray-100 hover:bg-gray-50">
                    <td className="px-4 py-2.5 font-medium text-gray-900">{s.supplier}</td>
                    {allMonths.map(m => {
                      const cell = s.months.get(m);
                      return (
                        <td key={m} className="px-4 py-2.5 text-right">
                          {cell ? (
                            <div>
                              <div className="font-bold text-amber-700">{fmtK(cell.amount)}</div>
                              <div className="text-[10px] text-gray-400">{cell.terms}</div>
                            </div>
                          ) : null}
                        </td>
                      );
                    })}
                    <td className="px-4 py-2.5 text-right font-bold text-gray-900">{fmtK(total)}</td>
                  </tr>
                );
              })}
            </tbody>
            <tfoot>
              <tr className="bg-gray-900 text-white">
                <td className="px-4 py-2.5 font-semibold">Total</td>
                {monthTotals.map((t, i) => <td key={i} className="px-4 py-2.5 text-right font-bold">{fmtK(t)}</td>)}
                <td className="px-4 py-2.5 text-right font-bold">{fmtK(grandTotal)}</td>
              </tr>
            </tfoot>
          </table>
        </div>
        <p className="text-[11px] text-gray-400 mt-2">
          Payment dates are approximated from payment term logic (CAD = on ETD, TT-15d = ETA minus 15 days, DP = on arrival, split terms prorated). Confirm exact due dates with the ops team.
        </p>
      </div>
    </div>
  );
}

// ─── Payments In ──────────────────────────────────────────────────

function PaymentsInPage({ trades, onSelect }: { trades: Trade[]; onSelect: (t: Trade) => void }) {
  const active = trades.filter(t => classifyDeal(t) !== "done");

  const scheduleByBuyer = useMemo(() => {
    const map = new Map<string, { buyer: string; months: Map<string, { amount: number; terms: string; trades: Trade[] }> }>();

    for (const t of active) {
      if (!has(t.buyer)) continue;
      const buyer = t.buyer!;
      const items = parseBuyerPaymentSchedule(t);
      if (!map.has(buyer)) map.set(buyer, { buyer, months: new Map() });
      const entry = map.get(buyer)!;

      for (const item of items) {
        const month = item.dueDate ? monthLabel(item.dueDate) : "TBD";
        if (!entry.months.has(month)) entry.months.set(month, { amount: 0, terms: "", trades: [] });
        const m = entry.months.get(month)!;
        m.amount += item.amount;
        m.terms = m.terms || item.label;
        m.trades.push(t);
      }
    }

    return [...map.values()].sort((a, b) => {
      const totalA = [...a.months.values()].reduce((s, m) => s + m.amount, 0);
      const totalB = [...b.months.values()].reduce((s, m) => s + m.amount, 0);
      return totalB - totalA;
    });
  }, [active]);

  const allMonths = useMemo(() => {
    const months = new Set<string>();
    for (const s of scheduleByBuyer) {
      for (const m of s.months.keys()) months.add(m);
    }
    return [...months].sort((a, b) => {
      const da = new Date(a);
      const db = new Date(b);
      if (isNaN(da.getTime()) || isNaN(db.getTime())) return a.localeCompare(b);
      return da.getTime() - db.getTime();
    });
  }, [scheduleByBuyer]);

  const monthTotals = allMonths.map(m => {
    let total = 0;
    for (const s of scheduleByBuyer) {
      total += s.months.get(m)?.amount || 0;
    }
    return total;
  });

  const grandTotal = monthTotals.reduce((s, v) => s + v, 0);

  return (
    <div>
      <div className="px-4 sm:px-6 pt-5 pb-3">
        <h1 className="text-xl font-bold text-gray-900">Payments In</h1>
        <p className="text-sm text-gray-400 mt-0.5">Expected buyer receipts · Approximate dates from payment terms</p>
      </div>

      <div className="px-4 sm:px-6 pb-4">
        <div className="border border-gray-200 rounded-xl overflow-x-auto">
          <table className="w-full text-xs">
            <thead>
              <tr className="bg-gray-50 border-b border-gray-200">
                <th className="px-4 py-2.5 text-left font-semibold text-gray-600">Buyer</th>
                {allMonths.map(m => <th key={m} className="px-4 py-2.5 text-right font-semibold text-gray-600 whitespace-nowrap">{m}</th>)}
                <th className="px-4 py-2.5 text-right font-semibold text-gray-600">Total</th>
              </tr>
            </thead>
            <tbody>
              {scheduleByBuyer.map((s, i) => {
                const total = [...s.months.values()].reduce((sum, m) => sum + m.amount, 0);
                return (
                  <tr key={i} className="border-b border-gray-100 hover:bg-gray-50">
                    <td className="px-4 py-2.5 font-medium text-gray-900">{s.buyer}</td>
                    {allMonths.map(m => {
                      const cell = s.months.get(m);
                      return (
                        <td key={m} className="px-4 py-2.5 text-right">
                          {cell ? (
                            <div>
                              <div className="font-bold text-emerald-700">{fmtK(cell.amount)}</div>
                              <div className="text-[10px] text-gray-400">{cell.terms}</div>
                            </div>
                          ) : null}
                        </td>
                      );
                    })}
                    <td className="px-4 py-2.5 text-right font-bold text-gray-900">{fmtK(total)}</td>
                  </tr>
                );
              })}
            </tbody>
            <tfoot>
              <tr className="bg-gray-900 text-white">
                <td className="px-4 py-2.5 font-semibold">Total</td>
                {monthTotals.map((t, i) => <td key={i} className="px-4 py-2.5 text-right font-bold">{fmtK(t)}</td>)}
                <td className="px-4 py-2.5 text-right font-bold">{fmtK(grandTotal)}</td>
              </tr>
            </tfoot>
          </table>
        </div>
        <p className="text-[11px] text-gray-400 mt-2">
          Receipt dates are approximated from buyer payment terms (On Delivery = ETA+1d, CAD = ETA, DP = ETA, TT on BL = BL+3-5d). Confirm with operations.
        </p>
      </div>
    </div>
  );
}

// ─── P&L Page ─────────────────────────────────────────────────────

function PnLPage({ trades }: { trades: Trade[] }) {
  const settled = trades.filter(t => classifyDeal(t) === "done");

  const totalSales = sumBy(settled, "sales_value");
  const totalPurchase = sumBy(settled, "purchase_value");
  const totalGrossMargin = sumBy(settled, "gross_margin");
  const totalNetProfit = sumBy(settled, "net_profit");
  const netMarginPct = totalSales > 0 ? (totalNetProfit / totalSales * 100).toFixed(2) : "0";

  // By product
  const byProduct = useMemo(() => {
    const map = new Map<string, { product: string; trades: number; containers: number; mt: number; purchase: number; sales: number; grossMargin: number; netProfit: number }>();
    for (const t of settled) {
      const p = t.product || "Other";
      if (!map.has(p)) map.set(p, { product: p, trades: 0, containers: 0, mt: 0, purchase: 0, sales: 0, grossMargin: 0, netProfit: 0 });
      const e = map.get(p)!;
      e.trades++;
      e.containers += n(t.no_of_containers);
      e.mt += n(t.quantity_mt);
      e.purchase += n(t.purchase_value);
      e.sales += n(t.sales_value);
      e.grossMargin += n(t.gross_margin);
      e.netProfit += n(t.net_profit);
    }
    return [...map.values()].sort((a, b) => b.sales - a.sales);
  }, [settled]);

  // Expense breakdown
  const clearanceTotal = sumBy(settled, "clearance_charges");
  const brokerageTotal = sumBy(settled, "brokerage");
  const interestTotal = sumBy(settled, "interest_loss") + sumBy(settled, "warehouse_loss") + sumBy(settled, "claims_paid") - sumBy(settled, "claims_received");
  const maxExpense = Math.max(clearanceTotal, brokerageTotal, Math.abs(interestTotal), 1);

  // Date range
  const months = settled.map(t => t.month).filter(Boolean);
  const dateRange = months.length > 0 ? `${months[0]} – ${months[months.length - 1]}` : "All time";

  return (
    <div>
      <div className="px-4 sm:px-6 pt-5 pb-3">
        <h1 className="text-xl font-bold text-gray-900">P&L</h1>
        <p className="text-sm text-gray-400 mt-0.5">{settled.length} settled trades · {dateRange}</p>
      </div>

      {/* Summary cards */}
      <div className="px-4 sm:px-6 pb-4 grid grid-cols-2 sm:grid-cols-4 gap-3">
        <div className="bg-gray-50 border border-gray-100 rounded-xl px-4 py-3">
          <div className="text-[10px] text-gray-400 uppercase tracking-wider">Sales Value</div>
          <div className="text-xl font-bold text-gray-900 mt-1">{fmtK(totalSales)}</div>
          <div className="text-[11px] text-gray-400">{settled.length} settled trades</div>
        </div>
        <div className="bg-gray-50 border border-gray-100 rounded-xl px-4 py-3">
          <div className="text-[10px] text-gray-400 uppercase tracking-wider">Purchase Value</div>
          <div className="text-xl font-bold text-gray-900 mt-1">{fmtK(totalPurchase)}</div>
          <div className="text-[11px] text-gray-400">Cost of goods sold</div>
        </div>
        <div className="bg-gray-50 border border-gray-100 rounded-xl px-4 py-3">
          <div className="text-[10px] text-gray-400 uppercase tracking-wider">Net Profit</div>
          <div className={`text-xl font-bold mt-1 ${totalNetProfit >= 0 ? "text-emerald-700" : "text-red-600"}`}>{fmtFull(totalNetProfit)}</div>
          <div className="text-[11px] text-gray-400">After all expenses</div>
        </div>
        <div className="bg-gray-50 border border-gray-100 rounded-xl px-4 py-3">
          <div className="text-[10px] text-gray-400 uppercase tracking-wider">Net Margin</div>
          <div className={`text-xl font-bold mt-1 ${totalNetProfit >= 0 ? "text-emerald-700" : "text-red-600"}`}>{netMarginPct}%</div>
          <div className="text-[11px] text-gray-400">Range {byProduct.length > 0 ? `${Math.min(...byProduct.map(p => p.sales > 0 ? p.netProfit / p.sales * 100 : 0)).toFixed(1)}–${Math.max(...byProduct.map(p => p.sales > 0 ? p.netProfit / p.sales * 100 : 0)).toFixed(1)}%` : "—"} by product</div>
        </div>
      </div>

      {/* By product table */}
      <div className="px-4 sm:px-6 pb-4">
        <h2 className="text-sm font-semibold text-gray-700 mb-2">By Product <span className="font-normal text-gray-400">{settled.length} settled trades · {dateRange}</span></h2>
        <div className="border border-gray-200 rounded-xl overflow-x-auto">
          <table className="w-full text-xs">
            <thead>
              <tr className="bg-gray-50 border-b border-gray-200">
                <th className="px-4 py-2.5 text-left font-semibold text-gray-600">Product</th>
                <th className="px-4 py-2.5 text-right font-semibold text-gray-600">Trades</th>
                <th className="px-4 py-2.5 text-right font-semibold text-gray-600">Ctns</th>
                <th className="px-4 py-2.5 text-right font-semibold text-gray-600">MT</th>
                <th className="px-4 py-2.5 text-right font-semibold text-gray-600">Purchase ($)</th>
                <th className="px-4 py-2.5 text-right font-semibold text-gray-600">Sales ($)</th>
                <th className="px-4 py-2.5 text-right font-semibold text-gray-600">Gross Margin</th>
                <th className="px-4 py-2.5 text-right font-semibold text-gray-600">Net Profit</th>
                <th className="px-4 py-2.5 text-right font-semibold text-gray-600">Margin %</th>
              </tr>
            </thead>
            <tbody>
              {byProduct.map((p, i) => (
                <tr key={i} className="border-b border-gray-100 hover:bg-gray-50">
                  <td className="px-4 py-2.5 font-semibold text-gray-900">{p.product}</td>
                  <td className="px-4 py-2.5 text-right text-gray-600">{p.trades}</td>
                  <td className="px-4 py-2.5 text-right text-gray-600">{fmt(p.containers)}</td>
                  <td className="px-4 py-2.5 text-right text-gray-600">{fmt(p.mt)}</td>
                  <td className="px-4 py-2.5 text-right text-gray-600">{fmtK(p.purchase)}</td>
                  <td className="px-4 py-2.5 text-right text-gray-600">{fmtK(p.sales)}</td>
                  <td className="px-4 py-2.5 text-right text-gray-600">{fmtK(p.grossMargin)}</td>
                  <td className={`px-4 py-2.5 text-right font-bold ${p.netProfit >= 0 ? "text-emerald-700" : "text-red-600"}`}>{fmtK(p.netProfit)}</td>
                  <td className="px-4 py-2.5 text-right text-gray-600">{p.sales > 0 ? (p.netProfit / p.sales * 100).toFixed(2) + "%" : "—"}</td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr className="bg-gray-900 text-white">
                <td className="px-4 py-2.5 font-semibold">Total</td>
                <td className="px-4 py-2.5 text-right font-bold">{settled.length}</td>
                <td className="px-4 py-2.5 text-right font-bold">{fmt(byProduct.reduce((s, p) => s + p.containers, 0))}</td>
                <td className="px-4 py-2.5 text-right font-bold">{fmt(byProduct.reduce((s, p) => s + p.mt, 0))}</td>
                <td className="px-4 py-2.5 text-right font-bold">{fmtK(totalPurchase)}</td>
                <td className="px-4 py-2.5 text-right font-bold">{fmtK(totalSales)}</td>
                <td className="px-4 py-2.5 text-right font-bold">{fmtK(totalGrossMargin)}</td>
                <td className="px-4 py-2.5 text-right font-bold">{fmtK(totalNetProfit)}</td>
                <td className="px-4 py-2.5 text-right font-bold">{netMarginPct}%</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>

      {/* Expense Breakdown */}
      <div className="px-4 sm:px-6 pb-6">
        <h2 className="text-sm font-semibold text-gray-700 mb-2">Expense Breakdown <span className="font-normal text-gray-400">What reduces gross margin to net</span></h2>
        <div className="space-y-3">
          {[
            { label: "Clearance Charges", val: clearanceTotal },
            { label: "Brokerage", val: brokerageTotal },
            { label: "Interest & Other", val: interestTotal },
          ].filter(e => e.val !== 0).map(({ label, val }) => (
            <div key={label} className="flex items-center gap-3">
              <span className="w-36 text-xs text-gray-600 shrink-0">{label}</span>
              <div className="flex-1 h-6 bg-gray-100 rounded overflow-hidden">
                <div className="h-full bg-teal-600 rounded" style={{ width: `${Math.max(4, Math.abs(val) / maxExpense * 100)}%` }} />
              </div>
              <span className="w-16 text-xs font-bold text-gray-900 text-right shrink-0">{fmtK(Math.abs(val))}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN COMPONENT
// ═══════════════════════════════════════════════════════════════════════════════

const NAV_ITEMS: { page: Page; label: string; section: string; icon: string }[] = [
  { page: "today", label: "Today's Work", section: "OVERVIEW", icon: "📋" },
  { page: "all_trades", label: "All Trades", section: "SHIPMENTS", icon: "📦" },
  { page: "open_positions", label: "Open Positions", section: "SHIPMENTS", icon: "🔓" },
  { page: "payments_out", label: "Payments Out", section: "PAYMENTS", icon: "💸" },
  { page: "payments_in", label: "Payments In", section: "PAYMENTS", icon: "💰" },
  { page: "pnl", label: "P&L", section: "ANALYSIS", icon: "📊" },
];

export default function ControlTower() {
  const [trades, setTrades] = useState<Trade[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [syncStatus, setSyncStatus] = useState<SyncStatus | null>(null);
  const [syncing, setSyncing] = useState(false);
  const [selectedDeal, setSelectedDeal] = useState<Trade | null>(null);
  const [activePage, setActivePage] = useState<Page>("today");
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);

  // ─── Fetch ──────────────────────────────────────────────────────

  const fetchTrades = useCallback(async () => {
    try {
      const res = await fetch("/api/data?table=trades");
      const result = await res.json();
      if (result.success) { setTrades(result.data); setError(null); }
      else setError(result.error);
    } catch { setError("Failed to load data"); }
    finally { setLoading(false); }
  }, []);

  const fetchSyncStatus = useCallback(async () => {
    try {
      const res = await fetch("/api/data?table=sync");
      const result = await res.json();
      if (result.success && result.lastSync) setSyncStatus(result.lastSync);
    } catch { /* ignore */ }
  }, []);

  const triggerSync = async () => {
    setSyncing(true); setError(null);
    try {
      const res = await fetch("/api/sync", { method: "POST" });
      const result = await res.json();
      if (result.success) { await fetchTrades(); await fetchSyncStatus(); }
      else setError(result.error);
    } catch { setError("Sync failed — check console"); }
    finally { setSyncing(false); }
  };

  useEffect(() => { fetchTrades(); fetchSyncStatus(); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (loading || syncing) return;
    const stale = trades.length === 0 || !syncStatus || (Date.now() - new Date(syncStatus.synced_at).getTime() > 5 * 60 * 1000);
    if (stale) triggerSync();
  }, [loading, trades.length, syncStatus]); // eslint-disable-line react-hooks/exhaustive-deps

  // ─── Stage Groups ──────────────────────────────────────────────

  const stageGroups = useMemo(() => {
    const groups: Record<DealStage, Trade[]> = {
      need_buyer: [], collecting_payment: [], waiting_to_ship: [], at_sea: [], done: [],
    };
    for (const t of trades) {
      groups[classifyDeal(t)].push(t);
    }
    return groups;
  }, [trades]);

  // Portfolio metrics
  const totalPortfolio = sumBy(trades, "purchase_value");
  const totalNetProfit = sumBy(trades.filter(t => classifyDeal(t) === "done"), "net_profit");
  const openExposure = sumBy(trades.filter(t => classifyDeal(t) === "need_buyer"), "purchase_value");

  const openPositions = trades.filter(t => classifyDeal(t) !== "done");
  const needAttentionCount = openPositions.length;

  function handlePageChange(page: Page) {
    setActivePage(page);
    setMobileMenuOpen(false);
  }

  // ─── Render ─────────────────────────────────────────────────────

  return (
    <div className="min-h-screen bg-white text-gray-900 flex flex-col sm:flex-row">
      {/* Deal Modal */}
      {selectedDeal && <DealModal trade={selectedDeal} onClose={() => setSelectedDeal(null)} />}

      {/* Mobile Header */}
      <header className="sm:hidden sticky top-0 z-20 bg-white border-b border-gray-100 px-4 py-3 flex items-center justify-between">
        <button onClick={() => setMobileMenuOpen(!mobileMenuOpen)} className="text-gray-500 p-1">
          <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 6h16M4 12h16M4 18h16" /></svg>
        </button>
        <div className="text-center">
          <h1 className="text-sm font-bold">Hectar</h1>
        </div>
        <button onClick={triggerSync} disabled={syncing} className="text-xs text-gray-400 hover:text-gray-600">
          {syncing ? "..." : "↻ Refresh"}
        </button>
      </header>

      {/* Mobile menu overlay */}
      {mobileMenuOpen && (
        <div className="sm:hidden fixed inset-0 z-30 bg-black/20" onClick={() => setMobileMenuOpen(false)}>
          <nav className="bg-white w-64 h-full shadow-xl p-4" onClick={e => e.stopPropagation()}>
            <div className="mb-6">
              <h1 className="text-lg font-bold text-gray-900">Hectar</h1>
              <p className="text-[11px] text-gray-400 uppercase tracking-wider">Operations</p>
            </div>
            {NAV_ITEMS.map((item, i, arr) => {
              const showSection = i === 0 || arr[i - 1].section !== item.section;
              return (
                <div key={item.page}>
                  {showSection && <div className="text-[10px] text-gray-400 uppercase tracking-wider mt-4 mb-1 px-2">{item.section}</div>}
                  <button onClick={() => handlePageChange(item.page)}
                    className={`w-full text-left px-3 py-2 rounded-lg text-sm flex items-center justify-between ${activePage === item.page ? "bg-blue-50 text-blue-700 font-medium" : "text-gray-600 hover:bg-gray-50"}`}>
                    <span>{item.label}</span>
                    {item.page === "today" && <span className="text-xs text-gray-400">{needAttentionCount}</span>}
                    {item.page === "all_trades" && <span className="text-xs text-gray-400">{trades.length}</span>}
                    {item.page === "open_positions" && <span className="text-xs text-gray-400">{openPositions.length}</span>}
                  </button>
                </div>
              );
            })}
          </nav>
        </div>
      )}

      {/* Desktop Sidebar */}
      <aside className="hidden sm:flex sm:flex-col sm:w-52 lg:w-56 border-r border-gray-100 bg-gray-50/50 min-h-screen sticky top-0 shrink-0">
        <div className="px-5 pt-5 pb-4">
          <h1 className="text-lg font-bold text-gray-900">Hectar</h1>
          <p className="text-[10px] text-gray-400 uppercase tracking-wider">Operations</p>
        </div>

        <nav className="flex-1 px-3 space-y-0.5">
          {NAV_ITEMS.map((item, i, arr) => {
            const showSection = i === 0 || arr[i - 1].section !== item.section;
            return (
              <div key={item.page}>
                {showSection && <div className="text-[10px] text-gray-400 uppercase tracking-wider mt-4 mb-1 px-2">{item.section}</div>}
                <button onClick={() => setActivePage(item.page)}
                  className={`w-full text-left px-3 py-2 rounded-lg text-sm flex items-center justify-between transition-colors ${activePage === item.page ? "bg-white text-blue-700 font-medium shadow-sm border border-gray-200" : "text-gray-600 hover:bg-white hover:text-gray-900"}`}>
                  <span>{item.label}</span>
                  {item.page === "today" && <span className={`text-xs ${activePage === item.page ? "text-blue-500" : "text-gray-400"}`}>{needAttentionCount}</span>}
                  {item.page === "all_trades" && <span className="text-xs text-gray-400">{trades.length}</span>}
                  {item.page === "open_positions" && <span className="text-xs text-gray-400">{openPositions.length}</span>}
                </button>
              </div>
            );
          })}
        </nav>

        {/* Sidebar footer stats */}
        <div className="px-5 py-4 border-t border-gray-100 space-y-1.5 text-xs">
          <div className="flex justify-between"><span className="text-gray-400">Portfolio</span><span className="font-semibold text-gray-900">{fmtK(totalPortfolio)}</span></div>
          <div className="flex justify-between"><span className="text-gray-400">Net Profit (YTD)</span><span className={`font-semibold ${totalNetProfit >= 0 ? "text-emerald-700" : "text-red-600"}`}>{fmtK(totalNetProfit)}</span></div>
          <div className="flex justify-between"><span className="text-gray-400">Open Exposure</span><span className="font-semibold text-gray-900">{fmtK(openExposure)}</span></div>
          <div className="text-[10px] text-gray-300 mt-2">{new Date().toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric" })}</div>
        </div>
      </aside>

      {/* Main content */}
      <main className="flex-1 min-w-0">
        {/* Desktop top bar */}
        <div className="hidden sm:flex items-center justify-end px-6 py-3 border-b border-gray-100">
          {syncStatus && <span className="text-[11px] text-gray-400 mr-3">updated {timeAgo(syncStatus.synced_at)}</span>}
          <button onClick={triggerSync} disabled={syncing}
            className="px-3 py-1.5 text-xs border border-gray-200 rounded-lg text-gray-600 hover:bg-gray-50 transition-colors disabled:opacity-50">
            {syncing ? "Syncing..." : "↻ Refresh"}
          </button>
        </div>

        {/* Error */}
        {error && (
          <div className="mx-4 sm:mx-6 mt-4 p-3 bg-red-50 border border-red-200 rounded-lg text-xs text-red-700">{error}</div>
        )}

        {/* Loading */}
        {loading && (
          <div className="px-6 py-16 text-center">
            <div className="text-sm text-gray-400">Loading trades...</div>
          </div>
        )}

        {/* Empty state */}
        {!loading && trades.length === 0 && !error && (
          <div className="px-6 py-16 text-center space-y-4">
            <h2 className="text-lg font-semibold text-gray-900">Ready to sync</h2>
            <p className="text-sm text-gray-400">Pull the latest data from your Excel spreadsheet.</p>
            <button onClick={triggerSync} disabled={syncing}
              className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-gray-300 text-white text-sm font-medium rounded-lg transition-colors">
              {syncing ? "Syncing..." : "Run First Sync"}
            </button>
          </div>
        )}

        {/* Pages */}
        {trades.length > 0 && (
          <>
            {activePage === "today" && <TodaysWorkPage trades={trades} stageGroups={stageGroups} onSelect={setSelectedDeal} />}
            {activePage === "all_trades" && <TradesListPage trades={trades} title="All Trades" subtitle={`${trades.length} trades across all sheets`} onSelect={setSelectedDeal} />}
            {activePage === "open_positions" && <TradesListPage trades={openPositions} title="Open Positions" subtitle={`${openPositions.length} non-settled trades`} onSelect={setSelectedDeal} />}
            {activePage === "payments_out" && <PaymentsOutPage trades={trades} onSelect={setSelectedDeal} />}
            {activePage === "payments_in" && <PaymentsInPage trades={trades} onSelect={setSelectedDeal} />}
            {activePage === "pnl" && <PnLPage trades={trades} />}
          </>
        )}
      </main>
    </div>
  );
}
