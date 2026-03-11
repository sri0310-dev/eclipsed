"use client";

import { useState, useEffect, useCallback, useMemo } from "react";

// ─── Types ────────────────────────────────────────────────────────────────────

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
  [key: string]: unknown;
}

interface SyncStatus {
  synced_at: string;
  status: string;
  total_rows: number;
  duration_ms: number;
  sheets_synced: string[];
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function fmtCurrency(n: number | null | undefined): string {
  if (n === null || n === undefined || isNaN(n)) return "—";
  if (Math.abs(n) >= 1_000_000) return "$" + (n / 1_000_000).toFixed(1) + "M";
  if (Math.abs(n) >= 1_000) return "$" + (n / 1_000).toFixed(0) + "K";
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 }).format(n);
}

function fmtFull(n: number | null | undefined): string {
  if (n === null || n === undefined || isNaN(n)) return "—";
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 }).format(n);
}

function fmt(n: number | null | undefined): string {
  if (n === null || n === undefined || isNaN(n)) return "—";
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(n);
}

function pctStr(part: number, total: number): string {
  return total > 0 ? ((part / total) * 100).toFixed(1) + "%" : "—";
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

function n(v: unknown): number {
  return Number(v) || 0;
}

/**
 * Position semantics (user-defined):
 * - "Inactive" = fully completed, all transactions settled (gray in Excel)
 * - "Sold" = at port, financial close in progress (blue in Excel) — ONLY "Sold"
 * - "Long" = open position (same as other open positions)
 * - Everything else = open / in transit / at origin / contracts
 */
function positionStage(pos: string): "completed" | "at_port" | "open" {
  const p = (pos || "").toLowerCase().trim();
  if (p === "inactive") return "completed";
  if (p === "sold") return "at_port";
  return "open"; // "long", "open", blank, everything else = open/active
}

function stageBadge(pos: string) {
  const stage = positionStage(pos);
  if (stage === "completed") return { label: "Completed", cls: "bg-zinc-700/50 text-zinc-400" };
  if (stage === "at_port") return { label: "At Port", cls: "bg-blue-900/40 text-blue-300" };
  return { label: pos || "Open", cls: "bg-amber-900/30 text-amber-300" };
}

function sumBy(trades: Trade[], field: keyof Trade): number {
  return trades.reduce((s, t) => s + n(t[field]), 0);
}

function groupSum(trades: Trade[], groupField: keyof Trade, sumField: keyof Trade): [string, number][] {
  const map = new Map<string, number>();
  for (const t of trades) {
    const key = String(t[groupField] ?? "").trim();
    if (!key || key === "—") continue;
    map.set(key, (map.get(key) || 0) + n(t[sumField]));
  }
  return [...map.entries()].sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]));
}

// ─── Bar component ────────────────────────────────────────────────────────────

function HBar({ label, value, max, color, prefix = "$" }: { label: string; value: number; max: number; color: string; prefix?: string }) {
  const w = max > 0 ? Math.max(4, (Math.abs(value) / Math.abs(max)) * 100) : 0;
  return (
    <div className="flex items-center gap-2 text-xs group">
      <span className="w-28 truncate text-zinc-400 text-right shrink-0">{label}</span>
      <div className="flex-1 h-5 bg-zinc-800/50 rounded overflow-hidden">
        <div className={`h-full ${color} rounded`} style={{ width: `${w}%` }} />
      </div>
      <span className="w-20 text-zinc-400 text-right font-mono shrink-0 text-[11px]">
        {prefix === "$" ? fmtCurrency(value) : fmt(value)}
      </span>
    </div>
  );
}

// ─── Deal Drilldown Modal ─────────────────────────────────────────────────────

function DealModal({ trade, onClose }: { trade: Trade; onClose: () => void }) {
  const badge = stageBadge(trade.position || "");
  const grossMargin = n(trade.gross_margin);
  const totalExp = n(trade.total_expenses);
  const netProfit = n(trade.net_profit);
  const purchaseVal = n(trade.purchase_value);
  const salesVal = n(trade.sales_value);
  const marginPct = salesVal > 0 ? ((netProfit / salesVal) * 100).toFixed(2) : "—";

  const expenses = [
    { label: "Clearance Charges", val: n(trade.clearance_charges) },
    { label: "Brokerage", val: n(trade.brokerage) },
    { label: "Warehouse Loss", val: n(trade.warehouse_loss) },
    { label: "Claims Paid", val: n(trade.claims_paid) },
    { label: "Claims Received", val: n(trade.claims_received) },
    { label: "Interest Loss", val: n(trade.interest_loss) },
  ].filter(e => e.val !== 0);

  const cashFlowItems = [
    { label: "Advance Paid to Seller", val: n(trade.advance_paid), color: "text-red-400", date: trade.advance_paid_on },
    { label: "Final Payment to Seller", val: n(trade.final_payment_paid), color: "text-red-400", date: trade.final_payment_amount_paid_on },
    { label: "Total Paid (Outwards)", val: n(trade.total_outwards), color: "text-red-400", date: null },
    { label: "Advance from Buyer", val: n(trade.advance_from_buyer), color: "text-emerald-400", date: trade.advance_received_on },
    { label: "2nd Payment from Buyer", val: n(trade.second_payment_from_buyer), color: "text-emerald-400", date: trade.payment_received_on },
    { label: "3rd Payment from Buyer", val: n(trade.third_payment_from_buyer), color: "text-emerald-400", date: null },
    { label: "Total Received (Inwards)", val: n(trade.total_inwards), color: "text-emerald-400", date: null },
  ].filter(e => e.val !== 0);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-sm" onClick={onClose}>
      <div className="bg-zinc-900 border border-zinc-700 rounded-xl max-w-2xl w-full mx-4 max-h-[85vh] overflow-y-auto shadow-2xl" onClick={e => e.stopPropagation()}>
        {/* Header */}
        <div className="sticky top-0 bg-zinc-900 border-b border-zinc-800 px-6 py-4 flex items-start justify-between">
          <div>
            <div className="flex items-center gap-3 mb-1">
              <h2 className="text-lg font-bold text-zinc-100">{trade.product || "Unknown"}</h2>
              <span className={`px-2 py-0.5 rounded text-[11px] font-medium ${badge.cls}`}>{badge.label}</span>
            </div>
            <div className="flex items-center gap-4 text-xs text-zinc-500">
              {trade.trade_no && <span>Trade #{trade.trade_no}</span>}
              {trade.origin && <span>{trade.origin}</span>}
              {trade.variety && <span>{trade.variety}</span>}
              {trade.source_sheet && <span>{trade.source_sheet}</span>}
            </div>
          </div>
          <button onClick={onClose} className="text-zinc-500 hover:text-zinc-300 text-xl leading-none p-1">&times;</button>
        </div>

        <div className="px-6 py-4 space-y-5">
          {/* Counterparties */}
          <div className="grid grid-cols-2 gap-4">
            <div>
              <div className="text-[10px] text-zinc-500 uppercase tracking-wider mb-1">Seller (We Buy From)</div>
              <div className="text-sm text-zinc-200">{trade.seller || "—"}</div>
              {trade.payment_terms && <div className="text-[11px] text-zinc-500 mt-0.5">Terms: {trade.payment_terms}</div>}
            </div>
            <div>
              <div className="text-[10px] text-zinc-500 uppercase tracking-wider mb-1">Buyer (We Sell To)</div>
              <div className="text-sm text-zinc-200">{trade.buyer || "—"}</div>
              {trade.buyer_payment_term && <div className="text-[11px] text-zinc-500 mt-0.5">Terms: {trade.buyer_payment_term}</div>}
            </div>
          </div>

          {/* Trade Economics */}
          <div>
            <div className="text-[10px] text-zinc-500 uppercase tracking-wider mb-2">Trade Economics</div>
            <div className="grid grid-cols-4 gap-3">
              <div className="bg-zinc-800/50 rounded-lg px-3 py-2">
                <div className="text-xs text-zinc-500">Qty (MT)</div>
                <div className="text-sm font-bold text-zinc-200">{fmt(trade.quantity_mt)}</div>
              </div>
              <div className="bg-zinc-800/50 rounded-lg px-3 py-2">
                <div className="text-xs text-zinc-500">Containers</div>
                <div className="text-sm font-bold text-zinc-200">{fmt(trade.no_of_containers)}</div>
              </div>
              <div className="bg-zinc-800/50 rounded-lg px-3 py-2">
                <div className="text-xs text-zinc-500">Buy $/MT</div>
                <div className="text-sm font-bold text-zinc-200">{fmtFull(trade.purchase_price_per_mt)}</div>
              </div>
              <div className="bg-zinc-800/50 rounded-lg px-3 py-2">
                <div className="text-xs text-zinc-500">Sell $/MT</div>
                <div className="text-sm font-bold text-zinc-200">{fmtFull(trade.sales_price_per_mt)}</div>
              </div>
            </div>
          </div>

          {/* Margin Waterfall */}
          <div>
            <div className="text-[10px] text-zinc-500 uppercase tracking-wider mb-2">Margin Breakdown</div>
            <div className="bg-zinc-800/30 rounded-lg p-4 space-y-2">
              <div className="flex justify-between text-xs">
                <span className="text-zinc-400">Purchase Value</span>
                <span className="text-red-400 font-mono">{fmtFull(purchaseVal)}</span>
              </div>
              <div className="flex justify-between text-xs">
                <span className="text-zinc-400">Sales Value</span>
                <span className="text-emerald-400 font-mono">{fmtFull(salesVal)}</span>
              </div>
              <div className="border-t border-zinc-700 my-1" />
              <div className="flex justify-between text-xs font-medium">
                <span className="text-zinc-300">Gross Margin</span>
                <span className={`font-mono ${grossMargin >= 0 ? "text-emerald-400" : "text-red-400"}`}>{fmtFull(grossMargin)}</span>
              </div>
              {expenses.length > 0 && (
                <>
                  <div className="border-t border-zinc-700/50 my-1" />
                  {expenses.map(({ label, val }) => (
                    <div key={label} className="flex justify-between text-xs">
                      <span className="text-zinc-500 pl-2">- {label}</span>
                      <span className="text-red-400/70 font-mono">{fmtFull(Math.abs(val))}</span>
                    </div>
                  ))}
                  <div className="flex justify-between text-xs">
                    <span className="text-zinc-400 pl-2">Total Expenses</span>
                    <span className="text-red-400 font-mono">{fmtFull(totalExp)}</span>
                  </div>
                </>
              )}
              <div className="border-t border-zinc-700 my-1" />
              <div className="flex justify-between text-sm font-bold">
                <span className="text-zinc-200">Net Profit</span>
                <span className={`font-mono ${netProfit >= 0 ? "text-emerald-400" : "text-red-400"}`}>{fmtFull(netProfit)}</span>
              </div>
              <div className="flex justify-between text-xs">
                <span className="text-zinc-500">Net Margin %</span>
                <span className={`font-mono ${netProfit >= 0 ? "text-emerald-400" : "text-red-400"}`}>{marginPct}%</span>
              </div>
            </div>
          </div>

          {/* Cash Flow Timeline */}
          {cashFlowItems.length > 0 && (
            <div>
              <div className="text-[10px] text-zinc-500 uppercase tracking-wider mb-2">Cash Flow</div>
              <div className="bg-zinc-800/30 rounded-lg p-4 space-y-2">
                {cashFlowItems.map(({ label, val, color, date }) => (
                  <div key={label} className="flex items-center justify-between text-xs">
                    <div className="flex items-center gap-2">
                      <span className="text-zinc-400">{label}</span>
                      {date && <span className="text-[10px] text-zinc-600">{date}</span>}
                    </div>
                    <span className={`font-mono ${color}`}>{fmtFull(val)}</span>
                  </div>
                ))}
                <div className="border-t border-zinc-700 my-1" />
                <div className="flex justify-between text-xs">
                  <span className="text-zinc-400">Still Owed to Seller</span>
                  <span className="text-amber-400 font-mono">{fmtFull(n(trade.outward_remaining))}</span>
                </div>
                <div className="flex justify-between text-xs">
                  <span className="text-zinc-400">Still Owed by Buyer</span>
                  <span className="text-amber-400 font-mono">{fmtFull(n(trade.inward_remaining))}</span>
                </div>
                {n(trade.working_capital_days) > 0 && (
                  <div className="flex justify-between text-xs">
                    <span className="text-zinc-400">Working Capital Days</span>
                    <span className="text-zinc-300 font-mono">{fmt(trade.working_capital_days)}d</span>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* Logistics */}
          {(trade.bl_number || trade.etd || trade.eta) && (
            <div>
              <div className="text-[10px] text-zinc-500 uppercase tracking-wider mb-2">Logistics</div>
              <div className="grid grid-cols-2 gap-3 text-xs">
                {trade.port_of_loading && <div><span className="text-zinc-500">Port of Loading: </span><span className="text-zinc-300">{trade.port_of_loading}</span></div>}
                {trade.port_of_discharge && <div><span className="text-zinc-500">Port of Discharge: </span><span className="text-zinc-300">{trade.port_of_discharge}</span></div>}
                {trade.bl_number && <div><span className="text-zinc-500">B/L Number: </span><span className="text-zinc-300 font-mono">{trade.bl_number}</span></div>}
                {trade.etd && <div><span className="text-zinc-500">ETD: </span><span className="text-zinc-300">{trade.etd}</span></div>}
                {trade.eta && <div><span className="text-zinc-500">ETA: </span><span className="text-zinc-300">{trade.eta}</span></div>}
                {trade.transit_days && <div><span className="text-zinc-500">Transit Days: </span><span className="text-zinc-300">{trade.transit_days}d</span></div>}
              </div>
            </div>
          )}

          {/* Remarks */}
          {trade.remarks && (
            <div>
              <div className="text-[10px] text-zinc-500 uppercase tracking-wider mb-1">Remarks</div>
              <p className="text-xs text-zinc-400 bg-zinc-800/30 rounded-lg p-3">{trade.remarks}</p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── Clickable Deal Card ──────────────────────────────────────────────────────

function DealCard({ trade, onClick }: { trade: Trade; onClick: () => void }) {
  const badge = stageBadge(trade.position || "");
  const netProfit = n(trade.net_profit);
  const outRem = n(trade.outward_remaining);
  const inRem = n(trade.inward_remaining);
  const hasOpenItems = outRem > 0 || inRem > 0;

  return (
    <button onClick={onClick} className="w-full text-left bg-zinc-800/30 hover:bg-zinc-800/60 border border-zinc-800 hover:border-zinc-700 rounded-lg px-4 py-3 transition-all group">
      <div className="flex items-start justify-between mb-1.5">
        <div className="flex items-center gap-2">
          <span className="text-sm font-medium text-zinc-200 group-hover:text-white">{trade.product || "—"}</span>
          <span className={`px-1.5 py-0.5 rounded text-[10px] ${badge.cls}`}>{badge.label}</span>
        </div>
        <span className={`text-sm font-bold font-mono ${netProfit >= 0 ? "text-emerald-400" : "text-red-400"}`}>
          {fmtCurrency(netProfit)}
        </span>
      </div>
      <div className="flex items-center justify-between text-[11px]">
        <span className="text-zinc-500">
          {trade.origin || "—"} · {trade.buyer || "No buyer"} · {fmt(trade.quantity_mt)} MT
        </span>
        {hasOpenItems && (
          <span className="text-amber-400/70">
            {outRem > 0 && `Pay: ${fmtCurrency(outRem)}`}
            {outRem > 0 && inRem > 0 && " · "}
            {inRem > 0 && `Recv: ${fmtCurrency(inRem)}`}
          </span>
        )}
      </div>
    </button>
  );
}

// ─── Section header ──────────────────────────────────────────────────────────

function Section({ title, subtitle, children }: { title: string; subtitle?: string; children: React.ReactNode }) {
  return (
    <div>
      <div className="pt-2 pb-2">
        <h2 className="text-sm font-semibold text-zinc-200">{title}</h2>
        {subtitle && <p className="text-[11px] text-zinc-500">{subtitle}</p>}
      </div>
      {children}
    </div>
  );
}

// ─── Metric Card ──────────────────────────────────────────────────────────────

function Metric({ label, value, color = "text-zinc-200", sub }: { label: string; value: string; color?: string; sub?: string }) {
  return (
    <div className="bg-zinc-900/60 border border-zinc-800 rounded-lg px-4 py-3">
      <div className={`text-lg font-bold ${color}`}>{value}</div>
      <div className="text-[11px] text-zinc-500">{label}</div>
      {sub && <div className="text-[10px] text-zinc-600 mt-0.5">{sub}</div>}
    </div>
  );
}

// ─── Main Component ───────────────────────────────────────────────────────────

export default function ControlTower() {
  const [trades, setTrades] = useState<Trade[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [syncStatus, setSyncStatus] = useState<SyncStatus | null>(null);
  const [syncing, setSyncing] = useState(false);
  const [selectedDeal, setSelectedDeal] = useState<Trade | null>(null);
  const [expandedSection, setExpandedSection] = useState<string | null>(null);

  const toggleSection = (id: string) => setExpandedSection(prev => prev === id ? null : id);

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
  useEffect(() => {
    if (trades.length === 0) return;
    const id = setInterval(() => { fetchTrades(); fetchSyncStatus(); }, 60_000);
    return () => clearInterval(id);
  }, [trades.length, fetchTrades, fetchSyncStatus]);

  // ─── Analytics ──────────────────────────────────────────────────

  const a = useMemo(() => {
    if (trades.length === 0) return null;

    // Stage classification
    const completed = trades.filter(t => positionStage(t.position ?? "") === "completed");
    const atPort = trades.filter(t => positionStage(t.position ?? "") === "at_port");
    const open = trades.filter(t => positionStage(t.position ?? "") === "open");
    const active = [...open, ...atPort]; // all non-completed

    // ── Company Health KPIs ──
    const totalPurchase = sumBy(trades, "purchase_value");
    const totalSales = sumBy(trades, "sales_value");
    const totalGrossMargin = sumBy(trades, "gross_margin");
    const totalNetProfit = sumBy(trades, "net_profit");
    const totalExpenses = sumBy(trades, "total_expenses");
    const totalOutwards = sumBy(trades, "total_outwards");
    const totalInwards = sumBy(trades, "total_inwards");
    const outwardRemaining = sumBy(active, "outward_remaining");
    const inwardRemaining = sumBy(active, "inward_remaining");
    const netCash = totalInwards - totalOutwards;

    // Completed deals margin analysis
    const completedGrossMargin = sumBy(completed, "gross_margin");
    const completedNetProfit = sumBy(completed, "net_profit");
    const completedSales = sumBy(completed, "sales_value");
    const completedMarginPct = completedSales > 0 ? (completedNetProfit / completedSales) * 100 : 0;

    // Active deals value at risk
    const activePurchaseValue = sumBy(active, "purchase_value");
    const activeSalesValue = sumBy(active, "sales_value");
    const activeExpectedProfit = sumBy(active, "net_profit");

    // Working capital
    const wcDays = active.map(t => n(t.working_capital_days)).filter(d => d > 0);
    const avgWcDays = wcDays.length > 0 ? wcDays.reduce((a, b) => a + b, 0) / wcDays.length : 0;

    // ── Risk Signals ──
    // Deals with negative profit
    const losingDeals = active.filter(t => n(t.net_profit) < 0);
    const totalLoss = losingDeals.reduce((s, t) => s + n(t.net_profit), 0);

    // Overdue receivables (at_port deals with open receivables — money should be in already)
    const overdueReceivables = atPort.filter(t => n(t.inward_remaining) > 0);
    const overdueReceivableAmount = sumBy(overdueReceivables, "inward_remaining");

    // Large open payables
    const openPayables = active.filter(t => n(t.outward_remaining) > 0).sort((a, b) => n(b.outward_remaining) - n(a.outward_remaining));
    const openReceivables = active.filter(t => n(t.inward_remaining) > 0).sort((a, b) => n(b.inward_remaining) - n(a.inward_remaining));

    // Concentration by value
    const buyerExposure = groupSum(active, "buyer", "sales_value");
    const sellerExposure = groupSum(active, "seller", "purchase_value");
    const commodityExposure = groupSum(active, "product", "purchase_value");
    const originExposure = groupSum(active, "origin", "purchase_value");

    const topBuyer = buyerExposure[0];
    const topSeller = sellerExposure[0];
    const topBuyerPct = topBuyer && activeSalesValue > 0 ? (topBuyer[1] / activeSalesValue) * 100 : 0;
    const topSellerPct = topSeller && activePurchaseValue > 0 ? (topSeller[1] / activePurchaseValue) * 100 : 0;

    // ── P&L by Commodity ──
    const profitByCommodity = groupSum(trades, "product", "net_profit");

    // ── P&L Heatmap ──
    const products = [...new Set(trades.map(t => t.product).filter(Boolean))] as string[];
    const origins = [...new Set(trades.map(t => t.origin).filter(Boolean))] as string[];
    const topProducts = profitByCommodity.slice(0, 6).map(([n]) => n);
    const topOrigins = groupSum(trades, "origin", "purchase_value").slice(0, 6).map(([n]) => n);
    const heatmap = topProducts.map(p =>
      topOrigins.map(o => {
        const matching = trades.filter(t => t.product === p && t.origin === o);
        return { count: matching.length, profit: matching.reduce((s, t) => s + n(t.net_profit), 0) };
      })
    );

    // ── Expense breakdown (all trades) ──
    const expenseBreakdown = [
      { label: "Clearance Charges", val: sumBy(trades, "clearance_charges") },
      { label: "Brokerage", val: sumBy(trades, "brokerage") },
      { label: "Warehouse Loss", val: sumBy(trades, "warehouse_loss") },
      { label: "Claims Paid", val: sumBy(trades, "claims_paid") },
      { label: "Claims Received", val: sumBy(trades, "claims_received") },
      { label: "Interest Loss", val: sumBy(trades, "interest_loss") },
    ].filter(e => e.val !== 0);

    // Working capital distribution
    const wcBuckets = [
      { label: "0–30d", count: wcDays.filter(d => d <= 30).length },
      { label: "31–60d", count: wcDays.filter(d => d > 30 && d <= 60).length },
      { label: "61–90d", count: wcDays.filter(d => d > 60 && d <= 90).length },
      { label: "90d+", count: wcDays.filter(d => d > 90).length },
    ].filter(b => b.count > 0);

    return {
      completed, atPort, open, active, trades,
      totalPurchase, totalSales, totalGrossMargin, totalNetProfit, totalExpenses,
      totalOutwards, totalInwards, outwardRemaining, inwardRemaining, netCash,
      completedGrossMargin, completedNetProfit, completedSales, completedMarginPct,
      activePurchaseValue, activeSalesValue, activeExpectedProfit,
      avgWcDays, losingDeals, totalLoss,
      overdueReceivables, overdueReceivableAmount,
      openPayables, openReceivables,
      buyerExposure, sellerExposure, commodityExposure, originExposure,
      topBuyer, topSeller, topBuyerPct, topSellerPct,
      profitByCommodity, heatmap, topProducts, topOrigins,
      expenseBreakdown, wcBuckets,
      products, origins,
    };
  }, [trades]);

  // ═══════════════════════════════════════════════════════════════
  // RENDER
  // ═══════════════════════════════════════════════════════════════

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100">
      {/* Deal Drilldown Modal */}
      {selectedDeal && <DealModal trade={selectedDeal} onClose={() => setSelectedDeal(null)} />}

      {/* Header */}
      <header className="border-b border-zinc-800 bg-zinc-900/50 backdrop-blur-sm sticky top-0 z-10">
        <div className="max-w-[1400px] mx-auto px-6 py-3 flex items-center justify-between">
          <div>
            <h1 className="text-lg font-bold tracking-tight">Hectar Control Tower</h1>
            <p className="text-[11px] text-zinc-500">Commodity Trading & Risk Management</p>
          </div>
          <div className="flex items-center gap-4">
            {syncStatus && (
              <span className="text-[11px] text-zinc-600">
                Synced {timeAgo(syncStatus.synced_at)} · {syncStatus.total_rows} rows
              </span>
            )}
            <button onClick={triggerSync} disabled={syncing}
              className="px-3 py-1 text-xs bg-zinc-800 hover:bg-zinc-700 text-zinc-300 rounded border border-zinc-700 transition-colors disabled:opacity-50">
              {syncing ? "Syncing..." : "Refresh"}
            </button>
            {trades.length > 0 && (
              <div className="flex items-center gap-1.5">
                <div className="h-2 w-2 rounded-full bg-emerald-500 animate-pulse" />
                <span className="text-[11px] text-zinc-400">Live</span>
              </div>
            )}
          </div>
        </div>
      </header>

      <main className="max-w-[1400px] mx-auto px-6 py-5 space-y-5">
        {/* Loading */}
        {loading && (
          <div className="space-y-4">
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              {[1, 2, 3, 4].map(i => (
                <div key={i} className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4 animate-pulse">
                  <div className="h-7 bg-zinc-800 rounded w-16 mb-1.5" />
                  <div className="h-3 bg-zinc-800/50 rounded w-20" />
                </div>
              ))}
            </div>
            <div className="text-center text-sm text-zinc-500 py-2">Loading...</div>
          </div>
        )}

        {/* Empty */}
        {!loading && trades.length === 0 && !error && (
          <div className="max-w-md mx-auto py-16 text-center space-y-5">
            <h2 className="text-lg font-semibold">Ready to sync</h2>
            <p className="text-sm text-zinc-400">Pull the latest data from your Excel spreadsheet.</p>
            <button onClick={triggerSync} disabled={syncing}
              className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-zinc-700 text-white text-sm font-medium rounded-lg transition-colors">
              {syncing ? "Syncing..." : "Run First Sync"}
            </button>
          </div>
        )}

        {/* Error */}
        {error && (
          <div className="p-3 bg-red-900/20 border border-red-800 rounded-lg text-xs text-red-400">{error}</div>
        )}

        {/* ═══ INSIGHTS DASHBOARD ═══ */}
        {a && (
          <>
            {/* ═══ 1. PORTFOLIO HEALTH AT A GLANCE ═══ */}
            <div className="grid grid-cols-3 gap-3">
              <div className="rounded-xl border border-amber-800/30 bg-gradient-to-br from-amber-950/20 to-zinc-900/60 px-5 py-4">
                <div className="flex items-center gap-2 mb-2">
                  <div className="h-2.5 w-2.5 rounded-full bg-amber-500" />
                  <span className="text-xs text-amber-300/70">Active Deals</span>
                </div>
                <div className="text-3xl font-bold text-amber-400">{a.open.length}</div>
                <div className="text-[11px] text-zinc-500 mt-1">
                  {fmtCurrency(a.activePurchaseValue)} deployed · {fmtCurrency(a.activeExpectedProfit)} expected P&L
                </div>
              </div>
              <div className="rounded-xl border border-blue-800/30 bg-gradient-to-br from-blue-950/20 to-zinc-900/60 px-5 py-4">
                <div className="flex items-center gap-2 mb-2">
                  <div className="h-2.5 w-2.5 rounded-full bg-blue-500" />
                  <span className="text-xs text-blue-300/70">At Port (Sold)</span>
                </div>
                <div className="text-3xl font-bold text-blue-400">{a.atPort.length}</div>
                <div className="text-[11px] text-zinc-500 mt-1">
                  Financial close in progress · {a.overdueReceivables.length > 0 && <span className="text-amber-400">{a.overdueReceivables.length} with open receivables</span>}
                  {a.overdueReceivables.length === 0 && "All collections on track"}
                </div>
              </div>
              <div className="rounded-xl border border-zinc-800 bg-gradient-to-br from-zinc-900/60 to-zinc-900/30 px-5 py-4">
                <div className="flex items-center gap-2 mb-2">
                  <div className="h-2.5 w-2.5 rounded-full bg-zinc-600" />
                  <span className="text-xs text-zinc-500">Completed</span>
                </div>
                <div className="text-3xl font-bold text-zinc-500">{a.completed.length}</div>
                <div className="text-[11px] text-zinc-600 mt-1">
                  Realized margin: {a.completedMarginPct.toFixed(1)}% · {fmtCurrency(a.completedNetProfit)} net
                </div>
              </div>
            </div>

            {/* ═══ 2. FINANCIAL HEALTH KPIs ═══ */}
            <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-3">
              <Metric label="Total Revenue" value={fmtCurrency(a.totalSales)} color="text-emerald-400" />
              <Metric label="Net Profit" value={fmtCurrency(a.totalNetProfit)} color={a.totalNetProfit >= 0 ? "text-emerald-400" : "text-red-400"} sub={`${pctStr(a.totalNetProfit, a.totalSales)} margin`} />
              <Metric label="Net Cash Position" value={fmtCurrency(a.netCash)} color={a.netCash >= 0 ? "text-emerald-400" : "text-red-400"} />
              <Metric label="Open Payables" value={fmtCurrency(a.outwardRemaining)} color="text-red-400" sub={`${a.openPayables.length} deals`} />
              <Metric label="Open Receivables" value={fmtCurrency(a.inwardRemaining)} color="text-amber-400" sub={`${a.openReceivables.length} deals`} />
              <Metric label="Avg WC Days" value={a.avgWcDays > 0 ? `${Math.round(a.avgWcDays)}d` : "—"} color="text-zinc-300" />
            </div>

            {/* ═══ 3. RISK ALERTS ═══ */}
            {(a.losingDeals.length > 0 || a.overdueReceivables.length > 0 || a.topBuyerPct > 30 || a.topSellerPct > 30) && (
              <div className="rounded-xl border border-red-900/30 bg-red-950/10 p-4">
                <h2 className="text-xs font-semibold text-red-400/80 mb-3">Risk Alerts</h2>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                  {a.losingDeals.length > 0 && (
                    <div className="flex items-start gap-3 text-xs">
                      <span className="text-red-500 mt-0.5">●</span>
                      <div>
                        <span className="text-zinc-300 font-medium">{a.losingDeals.length} deals with negative P&L</span>
                        <span className="text-red-400/70 ml-2">({fmtCurrency(a.totalLoss)} total loss)</span>
                      </div>
                    </div>
                  )}
                  {a.overdueReceivables.length > 0 && (
                    <div className="flex items-start gap-3 text-xs">
                      <span className="text-amber-500 mt-0.5">●</span>
                      <div>
                        <span className="text-zinc-300 font-medium">{a.overdueReceivables.length} at-port deals with uncollected payments</span>
                        <span className="text-amber-400/70 ml-2">({fmtCurrency(a.overdueReceivableAmount)})</span>
                      </div>
                    </div>
                  )}
                  {a.topBuyerPct > 30 && a.topBuyer && (
                    <div className="flex items-start gap-3 text-xs">
                      <span className="text-orange-500 mt-0.5">●</span>
                      <div>
                        <span className="text-zinc-300 font-medium">Buyer concentration: {a.topBuyer[0]}</span>
                        <span className="text-orange-400/70 ml-2">({a.topBuyerPct.toFixed(0)}% of active sales)</span>
                      </div>
                    </div>
                  )}
                  {a.topSellerPct > 30 && a.topSeller && (
                    <div className="flex items-start gap-3 text-xs">
                      <span className="text-orange-500 mt-0.5">●</span>
                      <div>
                        <span className="text-zinc-300 font-medium">Seller concentration: {a.topSeller[0]}</span>
                        <span className="text-orange-400/70 ml-2">({a.topSellerPct.toFixed(0)}% of active purchases)</span>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* ═══ 4. DEALS AT PORT — REQUIRING ATTENTION ═══ */}
            {a.atPort.length > 0 && (
              <Section title="At Port — Financial Close in Progress" subtitle={`${a.atPort.length} deals with shipment at port. Click any deal for full breakdown.`}>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                  {a.atPort.map((t, i) => (
                    <DealCard key={t.id || i} trade={t} onClick={() => setSelectedDeal(t)} />
                  ))}
                </div>
              </Section>
            )}

            {/* ═══ 5. OPEN POSITIONS — MONITORING REQUIRED ═══ */}
            <Section title="Active Positions" subtitle={`${a.open.length} open deals. Click any deal for cash flow & margin breakdown.`}>
              <div className="space-y-1">
                {/* Summary strip */}
                <div className="flex flex-wrap gap-2 mb-2">
                  {a.open.filter(t => !t.buyer || !t.seller).length > 0 && (
                    <span className="px-2 py-1 rounded bg-amber-900/20 border border-amber-800/30 text-[11px] text-amber-400">
                      {a.open.filter(t => !t.buyer || !t.seller).length} without counterparty
                    </span>
                  )}
                  {a.open.filter(t => n(t.net_profit) < 0).length > 0 && (
                    <span className="px-2 py-1 rounded bg-red-900/20 border border-red-800/30 text-[11px] text-red-400">
                      {a.open.filter(t => n(t.net_profit) < 0).length} with negative P&L
                    </span>
                  )}
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                  {a.open.sort((x, y) => Math.abs(n(y.purchase_value)) - Math.abs(n(x.purchase_value))).map((t, i) => (
                    <DealCard key={t.id || i} trade={t} onClick={() => setSelectedDeal(t)} />
                  ))}
                </div>
              </div>
            </Section>

            {/* ═══ 6. P&L PERFORMANCE ═══ */}
            <Section title="P&L by Commodity" subtitle="Net profit contribution across all trades">
              <div className="rounded-xl border border-zinc-800 bg-zinc-900/30 p-4">
                <div className="space-y-1.5">
                  {a.profitByCommodity.slice(0, 10).map(([name, profit]) => {
                    const maxVal = Math.max(...a.profitByCommodity.map(([, v]) => Math.abs(v)), 1);
                    const w = Math.max(4, (Math.abs(profit) / maxVal) * 100);
                    return (
                      <div key={name} className="flex items-center gap-2 text-xs">
                        <span className="w-28 truncate text-zinc-300 text-right shrink-0">{name}</span>
                        <div className="flex-1 h-5 bg-zinc-800/50 rounded overflow-hidden">
                          <div className={`h-full rounded ${profit >= 0 ? "bg-emerald-600" : "bg-red-600"}`} style={{ width: `${w}%` }} />
                        </div>
                        <span className={`w-20 text-right font-mono shrink-0 text-[11px] ${profit >= 0 ? "text-emerald-400" : "text-red-400"}`}>
                          {fmtCurrency(profit)}
                        </span>
                      </div>
                    );
                  })}
                </div>
              </div>
            </Section>

            {/* ═══ 7. P&L HEATMAP ═══ */}
            {a.topProducts.length > 0 && a.topOrigins.length > 0 && (
              <Section title="P&L Heatmap — Commodity x Origin" subtitle="Net profit by trade corridor. Green = profitable, Red = losing">
                <div className="rounded-xl border border-zinc-800 bg-zinc-900/30 p-4 overflow-x-auto">
                  <table className="text-xs border-collapse">
                    <thead>
                      <tr>
                        <th className="px-3 py-2 text-left text-zinc-500 font-medium" />
                        {a.topOrigins.map(o => <th key={o} className="px-3 py-2 text-zinc-400 font-medium whitespace-nowrap">{o}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {a.topProducts.map((p, pi) => (
                        <tr key={p}>
                          <td className="px-3 py-2 text-zinc-300 font-medium whitespace-nowrap">{p}</td>
                          {a.heatmap[pi].map((cell, oi) => {
                            if (cell.count === 0) return <td key={oi} className="px-3 py-2 text-center text-zinc-700">·</td>;
                            const pos = cell.profit >= 0;
                            const maxP = Math.max(...a.heatmap.flat().map(c => Math.abs(c.profit)), 1);
                            const intensity = Math.min(Math.abs(cell.profit) / maxP, 1);
                            return (
                              <td key={oi} className="px-3 py-2 text-center" style={{
                                backgroundColor: pos
                                  ? `rgba(16, 185, 129, ${0.08 + intensity * 0.35})`
                                  : `rgba(239, 68, 68, ${0.08 + intensity * 0.35})`,
                              }}>
                                <div className={`font-mono font-medium ${pos ? "text-emerald-300" : "text-red-300"}`}>{fmtCurrency(cell.profit)}</div>
                                <div className="text-[10px] text-zinc-500">{cell.count} trades</div>
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </Section>
            )}

            {/* ═══ 8. CASH FLOW & EXPENSES ═══ */}
            <Section title="Cash Flow & Expenses" subtitle="Company-wide financial position">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="rounded-xl border border-zinc-800 bg-zinc-900/30 p-4">
                  <h3 className="text-xs font-medium text-zinc-400 mb-3">Cash Flow</h3>
                  <div className="space-y-2.5">
                    <div className="flex justify-between text-xs">
                      <span className="text-zinc-400">Total Paid to Sellers</span>
                      <span className="text-red-400 font-mono">{fmtFull(a.totalOutwards)}</span>
                    </div>
                    <div className="flex justify-between text-xs">
                      <span className="text-zinc-400">Total Received from Buyers</span>
                      <span className="text-emerald-400 font-mono">{fmtFull(a.totalInwards)}</span>
                    </div>
                    <div className="border-t border-zinc-800 my-1" />
                    <div className="flex justify-between text-xs font-medium">
                      <span className="text-zinc-300">Net Cash Position</span>
                      <span className={`font-mono ${a.netCash >= 0 ? "text-emerald-400" : "text-red-400"}`}>{fmtFull(a.netCash)}</span>
                    </div>
                    <div className="border-t border-zinc-800 my-1" />
                    <div className="flex justify-between text-xs">
                      <span className="text-zinc-400">Still Owed to Sellers</span>
                      <span className="text-amber-400 font-mono">{fmtFull(a.outwardRemaining)}</span>
                    </div>
                    <div className="flex justify-between text-xs">
                      <span className="text-zinc-400">Still Owed by Buyers</span>
                      <span className="text-amber-400 font-mono">{fmtFull(a.inwardRemaining)}</span>
                    </div>
                  </div>
                </div>
                <div className="rounded-xl border border-zinc-800 bg-zinc-900/30 p-4">
                  <h3 className="text-xs font-medium text-zinc-400 mb-3">Expense Breakdown</h3>
                  <div className="space-y-2.5">
                    {a.expenseBreakdown.map(({ label, val }) => (
                      <div key={label} className="flex justify-between text-xs">
                        <span className="text-zinc-400">{label}</span>
                        <span className={`font-mono ${val > 0 ? "text-red-400" : "text-emerald-400"}`}>{fmtFull(Math.abs(val))}</span>
                      </div>
                    ))}
                    <div className="border-t border-zinc-800 my-1" />
                    <div className="flex justify-between text-xs font-medium">
                      <span className="text-zinc-300">Total Expenses</span>
                      <span className="text-red-400 font-mono">{fmtFull(a.totalExpenses)}</span>
                    </div>
                  </div>
                </div>
              </div>
            </Section>

            {/* ═══ 9. EXPOSURE ANALYSIS ═══ */}
            <Section title="Exposure Analysis" subtitle="Dollar-weighted concentration across counterparties, commodities, origins">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="rounded-xl border border-zinc-800 bg-zinc-900/30 p-4">
                  <h3 className="text-xs font-medium text-zinc-400 mb-3">Buyer Exposure (Sales Value)</h3>
                  <div className="space-y-1.5">
                    {a.buyerExposure.slice(0, 8).map(([name, val]) => (
                      <HBar key={name} label={name} value={val} max={a.buyerExposure[0]?.[1] || 1} color="bg-cyan-600" />
                    ))}
                  </div>
                </div>
                <div className="rounded-xl border border-zinc-800 bg-zinc-900/30 p-4">
                  <h3 className="text-xs font-medium text-zinc-400 mb-3">Seller Exposure (Purchase Value)</h3>
                  <div className="space-y-1.5">
                    {a.sellerExposure.slice(0, 8).map(([name, val]) => (
                      <HBar key={name} label={name} value={val} max={a.sellerExposure[0]?.[1] || 1} color="bg-purple-600" />
                    ))}
                  </div>
                </div>
                <div className="rounded-xl border border-zinc-800 bg-zinc-900/30 p-4">
                  <h3 className="text-xs font-medium text-zinc-400 mb-3">Commodity Exposure</h3>
                  <div className="space-y-1.5">
                    {a.commodityExposure.slice(0, 8).map(([name, val]) => (
                      <HBar key={name} label={name} value={val} max={a.commodityExposure[0]?.[1] || 1} color="bg-blue-600" />
                    ))}
                  </div>
                </div>
                <div className="rounded-xl border border-zinc-800 bg-zinc-900/30 p-4">
                  <h3 className="text-xs font-medium text-zinc-400 mb-3">Origin Exposure</h3>
                  <div className="space-y-1.5">
                    {a.originExposure.slice(0, 8).map(([name, val]) => (
                      <HBar key={name} label={name} value={val} max={a.originExposure[0]?.[1] || 1} color="bg-amber-600" />
                    ))}
                  </div>
                </div>
              </div>
            </Section>

            {/* ═══ 10. WORKING CAPITAL ═══ */}
            {a.wcBuckets.length > 0 && (
              <Section title="Working Capital Distribution" subtitle={`Average: ${Math.round(a.avgWcDays)} days across active deals`}>
                <div className="rounded-xl border border-zinc-800 bg-zinc-900/30 p-4">
                  <div className="space-y-1.5">
                    {a.wcBuckets.map(({ label, count }) => (
                      <HBar key={label} label={label} value={count} max={Math.max(...a.wcBuckets.map(b => b.count), 1)} color="bg-orange-600" prefix="" />
                    ))}
                  </div>
                </div>
              </Section>
            )}

            {/* ═══ 11. COMPLETED DEALS PERFORMANCE ═══ */}
            {a.completed.length > 0 && (
              <Section title="Completed Deals — Historical Performance" subtitle={`${a.completed.length} settled trades. Click any for full margin breakdown.`}>
                <div className="grid grid-cols-4 gap-3 mb-3">
                  <Metric label="Realized Revenue" value={fmtCurrency(a.completedSales)} color="text-emerald-400" />
                  <Metric label="Gross Margin" value={fmtCurrency(a.completedGrossMargin)} color={a.completedGrossMargin >= 0 ? "text-emerald-400" : "text-red-400"} />
                  <Metric label="Net Profit" value={fmtCurrency(a.completedNetProfit)} color={a.completedNetProfit >= 0 ? "text-emerald-400" : "text-red-400"} />
                  <Metric label="Net Margin %" value={a.completedMarginPct.toFixed(1) + "%"} color={a.completedMarginPct >= 0 ? "text-emerald-400" : "text-red-400"} />
                </div>
                <button onClick={() => toggleSection("completed")}
                  className="text-xs text-zinc-500 hover:text-zinc-300 transition-colors mb-2">
                  {expandedSection === "completed" ? "Hide deals ▾" : `Show ${a.completed.length} completed deals ▸`}
                </button>
                {expandedSection === "completed" && (
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                    {a.completed.sort((x, y) => Math.abs(n(y.net_profit)) - Math.abs(n(x.net_profit))).map((t, i) => (
                      <DealCard key={t.id || i} trade={t} onClick={() => setSelectedDeal(t)} />
                    ))}
                  </div>
                )}
              </Section>
            )}
          </>
        )}

        <div className="text-center text-[11px] text-zinc-700 pt-2 pb-4">
          Hectar CTRM · Supabase + OneDrive · Auto-sync every 5 min
        </div>
      </main>
    </div>
  );
}
