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

// ─── Download & Parse ────────────────────────────────────────────

function toDownloadUrl(shareUrl: string): string {
  const url = new URL(shareUrl);
  if (url.hostname === "1drv.ms") {
    url.searchParams.set("download", "1");
    return url.toString();
  }
  if (url.hostname.includes("onedrive.live.com")) {
    url.pathname = url.pathname
      .replace(/\/edit\.aspx/i, "/download.aspx")
      .replace(/\/embed/i, "/download");
    url.searchParams.set("download", "1");
    return url.toString();
  }
  if (url.hostname.includes("sharepoint.com")) {
    url.searchParams.set("download", "1");
    return url.toString();
  }
  url.searchParams.set("download", "1");
  return url.toString();
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
 */
export async function downloadAndParse(shareUrl: string): Promise<ParsedWorkbook> {
  const downloadUrl = toDownloadUrl(shareUrl);

  const response = await fetch(downloadUrl, {
    redirect: "follow",
    headers: { "User-Agent": "HectarControlTower/2.0" },
  });

  if (!response.ok) {
    throw new Error(
      `Failed to download Excel (HTTP ${response.status}). ` +
      `Ensure sharing is set to "Anyone with the link".`
    );
  }

  const arrayBuffer = await response.arrayBuffer();
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
