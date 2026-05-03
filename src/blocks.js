/*
 * Thoughtrons Excel add-in — block definitions.
 *
 * Each block is a function that uses the Excel API to write a brand-styled
 * range starting at the active cell. Includes header styling, RAG conditional
 * formatting, dropdowns where appropriate, formulas where useful.
 *
 * Adding a new block: append to BLOCKS, then add to LIBRARY in taskpane.html.
 *
 * Brand tokens — keep in sync with the .xlsx templates and PowerPoint add-in.
 */

window.ThoughtronsBrand = {
  ORANGE: "#ED8B1F",
  TEXT:   "#111111",
  TEXT2:  "#555555",
  TEXT3:  "#888888",
  RULE:   "#DDDDDD",
  HEADER_FILL: "#1F1F1F",
  OK_FILL:    "#DCF5DC",  OK_FONT:    "#2D7A2D",
  WARN_FILL:  "#FFF1CC",  WARN_FONT:  "#9B7A00",
  ALERT_FILL: "#FFD6D6",  ALERT_FONT: "#A12626",
};
const C = window.ThoughtronsBrand;

// ----- Helpers -----
function styleHeaderRow(range) {
  range.format.fill.color = C.HEADER_FILL;
  range.format.font.color = "#FFFFFF";
  range.format.font.name = "Consolas";
  range.format.font.size = 10;
  range.format.font.bold = true;
  range.format.borders.getItem("EdgeBottom").style = "Continuous";
  range.format.borders.getItem("EdgeTop").style = "Continuous";
  range.format.horizontalAlignment = "Left";
  range.format.verticalAlignment = "Center";
  range.format.rowHeight = 22;
}

function styleBodyRange(range) {
  range.format.font.name = "Calibri";
  range.format.font.size = 11;
  range.format.font.color = C.TEXT;
  range.format.verticalAlignment = "Top";
  range.format.borders.getItem("EdgeBottom").style = "Continuous";
  range.format.borders.getItem("EdgeBottom").color = C.RULE;
  range.format.wrapText = true;
}

function styleEyebrow(cell) {
  cell.format.font.name = "Consolas";
  cell.format.font.size = 9;
  cell.format.font.bold = true;
  cell.format.font.color = C.ORANGE;
}

function styleTitle(cell) {
  cell.format.font.name = "Calibri";
  cell.format.font.size = 22;
  cell.format.font.bold = true;
  cell.format.font.color = C.TEXT;
  cell.format.rowHeight = 32;
}

function styleStatNumber(cell) {
  cell.format.font.name = "Calibri";
  cell.format.font.size = 28;
  cell.format.font.bold = true;
  cell.format.font.color = C.ORANGE;
  cell.format.rowHeight = 36;
}

function getActiveCellAddress(ctx, ws) {
  // Resolve the active cell on the active sheet.
  const sel = ctx.workbook.getSelectedRange();
  ctx.load(sel, "address");
  return sel;
}

function ofs(addr, dr, dc) {
  // address like "Sheet1!B5" → returns just the row/col modified.
  // Strip sheet prefix if present
  const a = addr.split("!").pop().replace(/\$/g, "");
  const m = a.match(/([A-Z]+)(\d+)/);
  const colLetters = m[1], row = parseInt(m[2], 10);
  // col letters → number
  let col = 0;
  for (const ch of colLetters) col = col * 26 + (ch.charCodeAt(0) - 64);
  const newCol = col + dc;
  const newRow = row + dr;
  // number → letters
  let s = "", n = newCol;
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s + newRow;
}

// Apply RAG text conditional formatting on a range
function applyRagFormat(ws, rangeStr, choices) {
  // choices: array of { text, fill, font }
  for (const ch of choices) {
    const cf = ws.getRange(rangeStr).conditionalFormats.add("ContainsText");
    cf.textComparison.format.fill.color = ch.fill;
    cf.textComparison.format.font.color = ch.font;
    cf.textComparison.format.font.bold = true;
    cf.textComparison.rule = { operator: "Contains", text: ch.text };
  }
}

// Add a dropdown via data validation
function addDropdown(ws, rangeStr, choices) {
  const range = ws.getRange(rangeStr);
  range.dataValidation.rule = {
    list: { inCellDropDown: true, source: choices.join(",") }
  };
}

// ============================================================
// BLOCKS
// ============================================================
window.ThoughtronsBlocks = {

  // ---------- Headers / chrome ----------
  brand_header: async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    const sel = getActiveCellAddress(ctx, ws);
    await ctx.sync();
    const start = sel.address;
    const a = ofs(start, 0, 0);
    const b = ofs(start, 1, 0);
    const c = ofs(start, 2, 0);
    ws.getRange(a).values = [["// THOUGHTRONS // SHEET TITLE"]];
    styleEyebrow(ws.getRange(a));
    ws.getRange(b).values = [["Sheet title goes here."]];
    styleTitle(ws.getRange(b));
    ws.getRange(c).values = [["Subtitle / one-line description."]];
    ws.getRange(c).format.font.name = "Calibri";
    ws.getRange(c).format.font.size = 11;
    ws.getRange(c).format.font.italic = true;
    ws.getRange(c).format.font.color = C.TEXT2;
    return ctx.sync();
  },

  // ---------- KPI strip (4 KPIs) ----------
  kpi_strip: async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    const sel = getActiveCellAddress(ctx, ws);
    await ctx.sync();
    const start = sel.address.split("!").pop().replace(/\$/g, "");
    // Labels in row 1, values in row 2
    const labelRange = ws.getRange(start + ":" + ofs(start, 0, 3));
    labelRange.values = [["KPI ONE", "KPI TWO", "KPI THREE", "KPI FOUR"]];
    labelRange.format.font.name = "Consolas";
    labelRange.format.font.size = 9;
    labelRange.format.font.color = C.TEXT3;
    const valueRange = ws.getRange(ofs(start, 1, 0) + ":" + ofs(start, 1, 3));
    valueRange.values = [["00", "00", "00%", "₹00"]];
    valueRange.format.font.name = "Calibri";
    valueRange.format.font.size = 28;
    valueRange.format.font.bold = true;
    valueRange.format.font.color = C.ORANGE;
    valueRange.format.rowHeight = 36;
    return ctx.sync();
  },

  // ---------- Tables ----------
  rag_dashboard: async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    const sel = getActiveCellAddress(ctx, ws);
    await ctx.sync();
    const start = sel.address.split("!").pop().replace(/\$/g, "");
    const headers = ["ID", "Programme", "Owner", "Start", "Due", "% complete", "RAG", "Notes"];
    const sample  = [
      ["P-01", "Sample programme", "Owner", "2026-01-01", "2026-09-30", 0.45, "GREEN", "On track."],
      ["P-02", "Sample programme", "Owner", "2026-02-01", "2026-12-15", 0.30, "AMBER", "Watching closely."],
      ["P-03", "Sample programme", "Owner", "2025-11-01", "2026-06-30", 0.10, "RED",   "Off track; intervention needed."],
    ];
    const headerR = ws.getRange(start + ":" + ofs(start, 0, headers.length - 1));
    headerR.values = [headers];
    styleHeaderRow(headerR);
    const bodyR = ws.getRange(ofs(start, 1, 0) + ":" + ofs(start, sample.length, headers.length - 1));
    bodyR.values = sample;
    styleBodyRange(bodyR);
    bodyR.getColumn(5).numberFormat = [["0%"]];
    // Dropdown on RAG column
    const ragCol = ofs(start, 1, 6) + ":" + ofs(start, 30, 6);
    addDropdown(ws, ragCol, ["GREEN", "AMBER", "RED"]);
    return ctx.sync();
  },

  risk_block: async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    const sel = getActiveCellAddress(ctx, ws);
    await ctx.sync();
    const start = sel.address.split("!").pop().replace(/\$/g, "");
    const headers = ["ID", "Risk", "Category", "Owner", "Likelihood (1-5)", "Impact (1-5)", "Score", "Mitigation", "Status"];
    const sample = [
      ["R-01", "Single-supplier dependence", "Supply chain", "Owner", 4, 5, null, "Qualify second supplier.", "OPEN"],
      ["R-02", "Test schedule slip",         "Schedule",     "Owner", 3, 4, null, "Book ranges 2 quarters ahead.", "OPEN"],
    ];
    const hR = ws.getRange(start + ":" + ofs(start, 0, headers.length - 1));
    hR.values = [headers];
    styleHeaderRow(hR);
    const bR = ws.getRange(ofs(start, 1, 0) + ":" + ofs(start, sample.length, headers.length - 1));
    bR.values = sample;
    styleBodyRange(bR);
    // Score formula = Likelihood × Impact in column G
    for (let i = 1; i <= sample.length; i++) {
      const sc = ofs(start, i, 6);
      const li = ofs(start, i, 4);
      const im = ofs(start, i, 5);
      ws.getRange(sc).formulas = [[`=${li}*${im}`]];
    }
    addDropdown(ws, ofs(start,1,8) + ":" + ofs(start,30,8), ["OPEN", "MITIGATING", "ACCEPTED", "CLOSED"]);
    return ctx.sync();
  },

  action_items_block: async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    const sel = getActiveCellAddress(ctx, ws);
    await ctx.sync();
    const start = sel.address.split("!").pop().replace(/\$/g, "");
    const headers = ["#", "Action", "Owner", "Due", "Status"];
    const sample = [
      ["A1", "Action — verb-first, specific.", "Name", "DD MMM", "OPEN"],
      ["A2", "Action.", "Name", "DD MMM", "OPEN"],
      ["A3", "Action.", "Name", "DD MMM", "OPEN"],
    ];
    const hR = ws.getRange(start + ":" + ofs(start, 0, headers.length - 1));
    hR.values = [headers];
    styleHeaderRow(hR);
    const bR = ws.getRange(ofs(start, 1, 0) + ":" + ofs(start, sample.length, headers.length - 1));
    bR.values = sample;
    styleBodyRange(bR);
    addDropdown(ws, ofs(start,1,4) + ":" + ofs(start,30,4), ["OPEN", "IN PROG", "DONE", "BLOCKED"]);
    return ctx.sync();
  },

  raci_matrix: async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    const sel = getActiveCellAddress(ctx, ws);
    await ctx.sync();
    const start = sel.address.split("!").pop().replace(/\$/g, "");
    const headers = ["Activity", "Person 1", "Person 2", "Person 3", "Person 4", "Person 5"];
    const sample  = [
      ["Activity 1", "R", "A", "C", "I", ""],
      ["Activity 2", "A", "R", "C", "C", "I"],
      ["Activity 3", "C", "C", "R", "A", ""],
      ["Activity 4", "I", "C", "A", "R", "C"],
    ];
    const hR = ws.getRange(start + ":" + ofs(start, 0, headers.length - 1));
    hR.values = [headers];
    styleHeaderRow(hR);
    const bR = ws.getRange(ofs(start, 1, 0) + ":" + ofs(start, sample.length, headers.length - 1));
    bR.values = sample;
    styleBodyRange(bR);
    // Centre-align the R/A/C/I letters
    const rangeR = ws.getRange(ofs(start, 1, 1) + ":" + ofs(start, sample.length, headers.length - 1));
    rangeR.format.horizontalAlignment = "Center";
    rangeR.format.font.bold = true;
    return ctx.sync();
  },

  pipeline_block: async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    const sel = getActiveCellAddress(ctx, ws);
    await ctx.sync();
    const start = sel.address.split("!").pop().replace(/\$/g, "");
    const headers = ["ID", "Opportunity", "Client", "Owner", "Value (₹)", "Stage", "Probability", "Weighted (₹)", "Next action"];
    const sample = [
      ["O-01", "Opportunity name", "Client", "Owner", 2500000, "Qualified",   0.30, null, "Send proposal"],
      ["O-02", "Opportunity name", "Client", "Owner", 1800000, "Proposal",    0.50, null, "Reviewing T&Cs"],
      ["O-03", "Opportunity name", "Client", "Owner", 1200000, "Negotiation", 0.70, null, "Final round"],
    ];
    const hR = ws.getRange(start + ":" + ofs(start, 0, headers.length - 1));
    hR.values = [headers];
    styleHeaderRow(hR);
    const bR = ws.getRange(ofs(start, 1, 0) + ":" + ofs(start, sample.length, headers.length - 1));
    bR.values = sample;
    styleBodyRange(bR);
    bR.getColumn(4).numberFormat = [['"₹"#,##0']];
    bR.getColumn(6).numberFormat = [["0%"]];
    bR.getColumn(7).numberFormat = [['"₹"#,##0']];
    // Weighted = Value × Probability
    for (let i = 1; i <= sample.length; i++) {
      const w  = ofs(start, i, 7);
      const v  = ofs(start, i, 4);
      const pr = ofs(start, i, 6);
      ws.getRange(w).formulas = [[`=${v}*${pr}`]];
    }
    addDropdown(ws, ofs(start,1,5) + ":" + ofs(start,30,5),
      ["Lead", "Qualified", "Proposal", "Negotiation", "Closed-Won", "Closed-Lost"]);
    return ctx.sync();
  },

  // ---------- Classification banner row at top of sheet ----------
  classify_unclassified: (ctx) => bannerTop(ctx, "UNCLASSIFIED",  C.OK_FILL,    C.OK_FONT),
  classify_internal:     (ctx) => bannerTop(ctx, "INTERNAL",      "#FFE8C2",    "#9B5A00"),
  classify_restricted:   (ctx) => bannerTop(ctx, "RESTRICTED",    C.ALERT_FILL, C.ALERT_FONT),
  classify_confidential: (ctx) => bannerTop(ctx, "CONFIDENTIAL",  C.ALERT_FILL, C.ALERT_FONT),
};

async function bannerTop(ctx, label, fill, fontColor) {
  const ws = ctx.workbook.worksheets.getActiveWorksheet();
  // Insert a row at the top of the sheet
  ws.getRange("1:1").insert("Down");
  const banner = ws.getRange("A1:Z1");
  banner.merge(false);
  banner.values = [[label]];
  banner.format.font.name = "Consolas";
  banner.format.font.size = 11;
  banner.format.font.bold = true;
  banner.format.font.color = fontColor;
  banner.format.fill.color = fill;
  banner.format.horizontalAlignment = "Center";
  banner.format.verticalAlignment = "Center";
  banner.format.rowHeight = 22;
  return ctx.sync();
}
