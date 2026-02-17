import ExcelJS from "exceljs";

// ─── Column Mappings (Excel header → DB column) ──────────────────

export const TRADES_COLUMN_MAP: Record<string, string> = {
  "Product": "product",
  "Position": "position",
  "Month": "month",
  "Commodity Code": "commodity_code",
  "Origin": "origin",
  "Variety": "variety",
  "Packer": "packer",
  "Yield": "yield",
  "Specs": "specs",
  "Price ($/Lbs)": "price_per_lbs",
  "Port of Loading": "port_of_loading",
  "Port of Discharge": "port_of_discharge",
  "No of Containers": "no_of_containers",
  "Quantity (MT)": "quantity_mt",
  "Purchase Price/MT": "purchase_price_per_mt",
  "Sales Price/MT": "sales_price_per_mt",
  "Purchase Value": "purchase_value",
  "Sales Value": "sales_value",
  "Gross Margin": "gross_margin",
  "Clearnce Charges": "clearance_charges",
  "Clearance Charges": "clearance_charges",
  "Brokerage ": "brokerage",
  "Brokerage": "brokerage",
  "Warehouse Loss": "warehouse_loss",
  "Claims Paid": "claims_paid",
  "Claims Received": "claims_received",
  "Interest Loss": "interest_loss",
  "Total Expenses": "total_expenses",
  "Net Profit": "net_profit",
  "Profit %": "profit_pct",
  "BL Number": "bl_number",
  "Remarks": "remarks",
  "BL Date": "bl_date",
  "ETD": "etd",
  "ETA": "eta",
  "Transit Days": "transit_days",
  "Payment Terms": "payment_terms",
  "Payment by": "payment_by",
  "Seller": "seller",
  "Contract Reference Number": "contract_reference_number",
  "Broker": "broker",
  "Shipper Shipment Period": "shipper_shipment_period",
  "Buyer": "buyer",
  "Sales Contract Reference Number": "sales_contract_reference_number",
  "Broker.1": "buyer_broker",
  "Buyer Payment Term": "buyer_payment_term",
  "Buyer Shipment Period": "buyer_shipment_period",
  "Advance Paid": "advance_paid",
  "Advance Paid on": "advance_paid_on",
  "Final Payment paid": "final_payment_paid",
  "Final Payment Amount Paid on": "final_payment_amount_paid_on",
  "Total outwards": "total_outwards",
  "Remaining": "outward_remaining",
  "Adjustment": "outward_adjustment",
  "Outward": "outward",
  "Advance from\nBuyer": "advance_from_buyer",
  "Advance Received on": "advance_received_on",
  "2nd Payment from Buyer": "second_payment_from_buyer",
  "Payment Received on": "payment_received_on",
  "3rd Payment from Buyer": "third_payment_from_buyer",
  "Payment Received on.1": "payment_received_on_2",
  "Total Inwards": "total_inwards",
  "Remaining.1": "inward_remaining",
  "Adjustment.1": "inward_adjustment",
  "Inward": "inward",
  "Working Capital Days": "working_capital_days",
  "1st Supplier Invoice Number": "supplier_1_invoice_number",
  "Date": "supplier_1_date",
  "1st Supplier Sales Price": "supplier_1_sales_price",
  "Invoice Value": "supplier_1_invoice_value",
  "2nd Supplier Centum Invoice Number": "supplier_2_invoice_number",
  "Date.1": "supplier_2_date",
  "2nd Supplier Sales Price": "supplier_2_sales_price",
  "Invoice Value.1": "supplier_2_invoice_value",
  "3rd Supplier Hectar Global Invoice Number": "supplier_3_invoice_number",
  "Date.2": "supplier_3_date",
  "3rd Supplier Sales Price": "supplier_3_sales_price",
  "Invoice Value.2": "supplier_3_invoice_value",
  "Final Invoice No": "final_invoice_no",
  "Invoice Date": "invoice_date",
  "Buy\nOutturn/Nut Count": "buy_outturn_nut_count",
  "Sell\nOutturn/ Nut count": "sell_outturn_nut_count",
  "Outturn/ Nut Count as per the RBS": "outturn_nut_count_per_rbs",
  "Recutting Outturn/Nut Count": "recutting_outturn_nut_count",
  "BOE Date": "boe_date",
  "Exchange Rate": "exchange_rate",
  "Trade No": "trade_no",
};

const NV_COLUMN_MAP: Record<string, string> = {
  "Position": "position",
  "Port of Discharge": "port_of_discharge",
  "No of Containers": "no_of_containers",
  "Quantity (MT)": "quantity_mt",
  "Purchase Price/MT": "purchase_price_per_mt",
  "Sales Price/MT": "sales_price_per_mt",
  "BL Number": "bl_number",
  "ETD": "etd",
  "ETA": "eta",
  "Seller": "seller",
  "Shipper Shipment Period": "shipper_shipment_period",
  "Buyer": "buyer",
  "Buyer Shipment Period": "buyer_shipment_period",
  "Remarks": "remarks",
};

const TUT_COLUMN_MAP: Record<string, string> = {
  "Position": "position",
  "No of Containers": "no_of_containers",
  "Purchase Price/MT": "purchase_price_per_mt",
  "Sales Price/MT": "sales_price_per_mt",
  "BL Number": "bl_number",
  "ETD": "etd",
  "ETA": "eta",
  "Seller": "seller",
  "Shipper Shipment Period": "shipper_shipment_period",
  "Buyer": "buyer",
  "Buyer Shipment Period": "buyer_shipment_period",
  "Remarks": "remarks",
};

const CURRENCY_COLUMN_MAP: Record<string, string> = {
  "Date": "rate_date",
  "USD to INR": "usd_to_inr",
  "USD to TZS": "usd_to_tzs",
  "USD to XAF": "usd_to_xaf",
  "USD to NGN": "usd_to_ngn",
};

// Sheets that use the main trades structure
const COMMODITY_SHEETS = ["Main Sheet", "Sesame ", "Sesame", "Almonds", "Soybeans", "RCN"];

// ─── Cell Value Extraction ───────────────────────────────────────

function cellValue(cell: ExcelJS.Cell): string | number | boolean | null {
  const val = cell.value;
  if (val === null || val === undefined) return null;

  if (typeof val === "object" && "result" in val) {
    return (val as { result: unknown }).result as string | number | boolean | null;
  }
  if (typeof val === "object" && "richText" in val) {
    return (val as { richText: { text: string }[] }).richText.map((t) => t.text).join("");
  }
  if (val instanceof Date) {
    return val.toISOString().split("T")[0]; // YYYY-MM-DD
  }
  if (typeof val === "object" && "hyperlink" in val) {
    return (val as { text?: string; hyperlink: string }).text || (val as { hyperlink: string }).hyperlink;
  }
  return val as string | number | boolean;
}

function cleanForDb(val: string | number | boolean | null): string | number | null {
  if (val === null || val === undefined) return null;
  if (typeof val === "boolean") return val ? 1 : 0;
  if (typeof val === "number") return isNaN(val) ? null : val;
  const s = String(val).trim();
  if (!s || s === "-" || s === "NaN" || s === "nan" || s === "NaT" || s === "undefined") return null;
  return s;
}

// ─── OneDrive Download (Badger Token + Redeem) ──────────────────
//
// Microsoft changed their sharing infrastructure in 2024-2025.
// The old base64 shares API (api.onedrive.com/v1.0/shares/u!...)
// no longer works for the new URL format (/x/c/{cid}/{resid}).
//
// The working approach for new-format personal OneDrive links:
// 1. Follow the 1drv.ms redirect to get the resolved URL with ?redeem= param
// 2. Get a "Badger" token from api-badgerp.svc.ms (no OAuth needed)
// 3. Use the Badger token to query the OneDrive API for @content.downloadUrl
// 4. Download the file from that URL

const BADGER_TOKEN_URL = "https://api-badgerp.svc.ms/v1.0/token";
const BADGER_APP_UUID = "5cbed6ac-a083-4e14-b191-b4ba07653de2";
const BADGER_APP_ID = "1141147648";
const PERSONAL_API = "https://my.microsoftpersonalcontent.com/_api/v2.0/shares";
const BROWSER_UA =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36";

// Simple in-memory Badger token cache (valid ~1 week)
let cachedBadgerToken: { token: string; expiresAt: number } | null = null;

async function getBadgerToken(): Promise<string> {
  if (cachedBadgerToken && Date.now() < cachedBadgerToken.expiresAt) {
    return cachedBadgerToken.token;
  }

  console.log("[download] Requesting new Badger token...");
  const res = await fetch(BADGER_TOKEN_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      AppId: BADGER_APP_ID,
      "User-Agent": BROWSER_UA,
    },
    body: JSON.stringify({ appId: BADGER_APP_UUID }),
  });

  if (!res.ok) {
    throw new Error(`Badger token request failed: HTTP ${res.status}`);
  }

  const data = await res.json();
  // Cache for 6 days (token is valid for ~1 week)
  cachedBadgerToken = {
    token: data.token,
    expiresAt: Date.now() + 6 * 24 * 60 * 60 * 1000,
  };
  return data.token;
}

/**
 * Follow the 1drv.ms redirect chain to get the resolved URL
 * which contains the ?redeem= parameter needed for the API.
 *
 * Tries two approaches:
 * 1. Follow all redirects and check final URL for ?redeem= param
 * 2. If no redeem in URL, check intermediate redirect (manual mode)
 */
async function resolveShareLink(shareUrl: string): Promise<string> {
  // Approach 1: Follow all redirects, check final URL
  const res = await fetch(shareUrl, {
    redirect: "follow",
    headers: { "User-Agent": BROWSER_UA },
  });
  const finalUrl = res.url;

  // Check if the redeem param is in the final URL
  if (finalUrl.includes("redeem=")) {
    return finalUrl;
  }

  // Approach 2: Check intermediate redirects (manual mode)
  // Sometimes the redeem param appears in the first redirect
  const manualRes = await fetch(shareUrl, {
    redirect: "manual",
    headers: { "User-Agent": BROWSER_UA },
  });
  const location = manualRes.headers.get("location");
  if (location && location.includes("redeem=")) {
    return location;
  }

  // If the HTML body contains the redeem value, try to extract it
  const html = await res.text().catch(() => "");
  const redeemMatch = html.match(/[?&]redeem=([^&"'\s]+)/);
  if (redeemMatch) {
    // Reconstruct the URL with the redeem parameter
    const urlObj = new URL(finalUrl);
    urlObj.searchParams.set("redeem", decodeURIComponent(redeemMatch[1]));
    return urlObj.toString();
  }

  // Return whatever we have — the caller will handle the missing redeem
  console.log("[download] No redeem found. Final URL:", finalUrl.substring(0, 200));
  console.log("[download] Redirect location:", location?.substring(0, 200) || "none");
  return finalUrl;
}

/**
 * Download a file using the Badger Token + Redeem method.
 * Works for new OneDrive personal links (/x/c/{cid}/{resid}).
 */
async function downloadViaBadger(shareUrl: string): Promise<ArrayBuffer> {
  // Step 1: Resolve the sharing URL to get the redeem parameter
  console.log("[download] Resolving share link...");
  const resolvedUrl = await resolveShareLink(shareUrl);
  console.log("[download] Resolved to:", resolvedUrl.substring(0, 120) + "...");

  const url = new URL(resolvedUrl);
  const redeem = url.searchParams.get("redeem");

  if (!redeem) {
    throw new Error(
      `No 'redeem' parameter found in resolved URL. ` +
        `This may be an old-format link. Resolved: ${resolvedUrl.substring(0, 200)}`
    );
  }

  // Step 2: Get Badger token (cached)
  const token = await getBadgerToken();

  // Step 3: Query the OneDrive API for file metadata
  // The redeem value is already URL-safe base64 — use it directly
  const apiUrl = `${PERSONAL_API}/u!${redeem}/driveitem`;
  console.log("[download] Querying OneDrive API for download URL...");

  const metaRes = await fetch(apiUrl, {
    headers: {
      "User-Agent": BROWSER_UA,
      Authorization: `Badger ${token}`,
      Prefer: "autoredeem",
    },
  });

  if (!metaRes.ok) {
    const body = await metaRes.text().catch(() => "");
    throw new Error(
      `OneDrive API returned HTTP ${metaRes.status}: ${body.substring(0, 500)}`
    );
  }

  const metadata = await metaRes.json();
  const downloadUrl = metadata["@content.downloadUrl"];

  if (!downloadUrl) {
    throw new Error(
      `No @content.downloadUrl in API response. Keys: ${Object.keys(metadata).join(", ")}`
    );
  }

  console.log("[download] Got download URL, fetching file...");

  // Step 4: Download the actual file
  const fileRes = await fetch(downloadUrl, {
    headers: { "User-Agent": BROWSER_UA },
  });

  if (!fileRes.ok) {
    throw new Error(`File download failed: HTTP ${fileRes.status}`);
  }

  return fileRes.arrayBuffer();
}

/**
 * Fallback: try the old authkey-based approach for legacy links.
 */
async function downloadViaAuthKey(shareUrl: string): Promise<ArrayBuffer> {
  console.log("[download] Trying legacy authkey method...");
  const resolvedUrl = await resolveShareLink(shareUrl);
  const url = new URL(resolvedUrl);

  const resid = url.searchParams.get("resid") || url.searchParams.get("id");
  const authkey = url.searchParams.get("authkey");
  const cid = url.searchParams.get("cid") || resid?.split("!")[0];

  if (!resid || !authkey || !cid) {
    throw new Error(
      `Could not extract resid/authkey/cid from resolved URL. ` +
        `Params: ${url.search}`
    );
  }

  const apiUrl = `https://api.onedrive.com/v1.0/drives/${cid}/items/${resid}?authkey=${authkey}`;
  const metaRes = await fetch(apiUrl, {
    headers: { "User-Agent": BROWSER_UA },
  });

  if (!metaRes.ok) {
    throw new Error(`Legacy API returned HTTP ${metaRes.status}`);
  }

  const metadata = await metaRes.json();
  const downloadUrl = metadata["@content.downloadUrl"];

  if (!downloadUrl) {
    throw new Error("No @content.downloadUrl in legacy API response");
  }

  const fileRes = await fetch(downloadUrl, {
    headers: { "User-Agent": BROWSER_UA },
  });

  if (!fileRes.ok) {
    throw new Error(`Legacy file download failed: HTTP ${fileRes.status}`);
  }

  return fileRes.arrayBuffer();
}

export interface ParsedWorkbook {
  trades: Record<string, unknown>[];
  nvTrades: Record<string, unknown>[];
  tutTrades: Record<string, unknown>[];
  currencyRates: Record<string, unknown>[];
  financials: Record<string, unknown>[];
  mtm: Record<string, unknown>[];
  sheetsFound: string[];
}

/**
 * Download an Excel file from a OneDrive sharing URL and parse all worksheets
 * into structured records ready for Supabase insertion.
 *
 * Tries multiple strategies:
 * 1. Badger Token + Redeem (works for new /c/ format links)
 * 2. Legacy authkey method (works for old /s! format links)
 */
export async function downloadAndParse(shareUrl: string): Promise<ParsedWorkbook> {
  const errors: string[] = [];

  // Strategy 1: Badger Token + Redeem (primary — for new-format links)
  try {
    console.log("[sync] Strategy 1: Badger Token + Redeem");
    const arrayBuffer = await downloadViaBadger(shareUrl);
    console.log(`[sync] Downloaded ${arrayBuffer.byteLength} bytes via Badger method`);
    return parseWorkbook(arrayBuffer);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error("[sync] Strategy 1 failed:", msg);
    errors.push(`Badger Token: ${msg}`);
  }

  // Strategy 2: Legacy authkey method (for old-format links)
  try {
    console.log("[sync] Strategy 2: Legacy authkey method");
    const arrayBuffer = await downloadViaAuthKey(shareUrl);
    console.log(`[sync] Downloaded ${arrayBuffer.byteLength} bytes via legacy method`);
    return parseWorkbook(arrayBuffer);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error("[sync] Strategy 2 failed:", msg);
    errors.push(`Legacy authkey: ${msg}`);
  }

  throw new Error(
    `Failed to download Excel file. All strategies failed:\n` +
      errors.map((e, i) => `  ${i + 1}. ${e}`).join("\n") +
      `\n\nTroubleshooting:\n` +
      `  - Verify the sharing URL opens in an incognito browser window\n` +
      `  - Ensure sharing is "Anyone with the link" (not "People in your org")\n` +
      `  - Try generating a new sharing link`
  );
}

async function parseWorkbook(arrayBuffer: ArrayBuffer): Promise<ParsedWorkbook> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);

  const sheetsFound = workbook.worksheets.map((ws) => ws.name);

  const result: ParsedWorkbook = {
    trades: [],
    nvTrades: [],
    tutTrades: [],
    currencyRates: [],
    financials: [],
    mtm: [],
    sheetsFound,
  };

  // ── Parse commodity/trade sheets ──
  for (const sheetName of COMMODITY_SHEETS) {
    const ws = workbook.getWorksheet(sheetName);
    if (!ws) continue;
    const records = parseSheet(ws, TRADES_COLUMN_MAP);
    for (const rec of records) {
      rec.source_sheet = sheetName.trim();
    }
    result.trades.push(...records);
  }

  // ── Parse NV sheet ──
  const nvWs = workbook.getWorksheet("NV");
  if (nvWs) {
    result.nvTrades = parseSheet(nvWs, NV_COLUMN_MAP);
  }

  // ── Parse TUT sheet ──
  const tutWs = workbook.getWorksheet("TUT");
  if (tutWs) {
    result.tutTrades = parseSheet(tutWs, TUT_COLUMN_MAP);
  }

  // ── Parse Currency sheet ──
  const currWs = workbook.getWorksheet("Currency");
  if (currWs) {
    result.currencyRates = parseSheet(currWs, CURRENCY_COLUMN_MAP);
  }

  // ── Parse Financials (pivot format) ──
  const finWs = workbook.getWorksheet("Financials");
  if (finWs) {
    result.financials = parseFinancials(finWs);
  }

  // ── Parse MTM ──
  const mtmWs = workbook.getWorksheet("MTM");
  if (mtmWs) {
    result.mtm = parseMtm(mtmWs);
  }

  return result;
}

// ─── Generic Sheet Parser ────────────────────────────────────────

function parseSheet(
  ws: ExcelJS.Worksheet,
  columnMap: Record<string, string>
): Record<string, unknown>[] {
  const records: Record<string, unknown>[] = [];

  // Read header row
  const headerRow = ws.getRow(1);
  const headers: string[] = [];
  const colCount = ws.columnCount;

  // Track duplicate header names (Date, Date.1, Date.2, etc.)
  const headerCounts: Record<string, number> = {};

  for (let c = 1; c <= colCount; c++) {
    let h = String(cellValue(headerRow.getCell(c)) ?? "").trim();
    if (!h) {
      headers.push("");
      continue;
    }

    // Handle duplicates by appending .N suffix
    if (headerCounts[h] !== undefined) {
      headerCounts[h]++;
      h = `${h}.${headerCounts[h]}`;
    } else {
      headerCounts[h] = 0;
    }
    headers.push(h);
  }

  // Read data rows
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const record: Record<string, unknown> = { excel_row: r };
    let hasData = false;

    for (let c = 1; c <= colCount; c++) {
      const header = headers[c - 1];
      if (!header) continue;

      const dbCol = columnMap[header];
      if (!dbCol) continue;

      const val = cleanForDb(cellValue(row.getCell(c)));
      if (val !== null) hasData = true;
      record[dbCol] = val;
    }

    if (hasData) records.push(record);
  }

  return records;
}

// ─── Financials Parser (pivot table format) ──────────────────────

function parseFinancials(ws: ExcelJS.Worksheet): Record<string, unknown>[] {
  const records: Record<string, unknown>[] = [];
  let headerFound = false;

  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const firstCell = String(cellValue(row.getCell(1)) ?? "").trim();

    if (firstCell === "Row Labels") {
      headerFound = true;
      continue;
    }

    if (headerFound && firstCell) {
      records.push({
        excel_row: r,
        row_label: cleanForDb(cellValue(row.getCell(1))),
        sum_no_of_containers: cleanForDb(cellValue(row.getCell(2))),
        sum_quantity_mt: cleanForDb(cellValue(row.getCell(3))),
        sum_purchase_value: cleanForDb(cellValue(row.getCell(4))),
        sum_sales_value: cleanForDb(cellValue(row.getCell(5))),
        sum_gross_margin: cleanForDb(cellValue(row.getCell(6))),
        sum_net_profit: cleanForDb(cellValue(row.getCell(7))),
      });
    }
  }

  return records;
}

// ─── MTM Parser ──────────────────────────────────────────────────

function parseMtm(ws: ExcelJS.Worksheet): Record<string, unknown>[] {
  const records: Record<string, unknown>[] = [];
  let headerFound = false;

  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);

    // Check all cells in row for "Commodity" header
    for (let c = 1; c <= ws.columnCount; c++) {
      if (String(cellValue(row.getCell(c)) ?? "").trim() === "Commodity") {
        headerFound = true;
        break;
      }
    }

    if (!headerFound) continue;
    if (String(cellValue(row.getCell(1)) ?? "").trim() === "Commodity") continue;

    const commodity = cleanForDb(cellValue(row.getCell(2)));
    if (!commodity) continue;

    records.push({
      excel_row: r,
      commodity,
      position: cleanForDb(cellValue(row.getCell(3))),
      quantity: cleanForDb(cellValue(row.getCell(4))),
      no_of_containers: cleanForDb(cellValue(row.getCell(5))),
      destination: cleanForDb(cellValue(row.getCell(6))),
      average_price: cleanForDb(cellValue(row.getCell(7))),
      marked_at: cleanForDb(cellValue(row.getCell(8))),
      mark_to_market: cleanForDb(cellValue(row.getCell(9))),
    });
  }

  return records;
}
