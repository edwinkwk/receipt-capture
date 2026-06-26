/**
 * ═══════════════════════════════════════════════════════════════════════
 *  Shopping Receipt Tracker — Google Apps Script Webhook
 * ═══════════════════════════════════════════════════════════════════════
 *
 *  SETUP:
 *  1. Open Google Sheets → Extensions → Apps Script
 *  2. Paste this code into Code.gs
 *  3. Deploy → New deployment → Web app
 *       Execute as: Me
 *       Who has access: Anyone  (or "Anyone with Google account" for tighter auth)
 *  4. Copy the deployment URL — that is your webhook endpoint
 *
 *  AUTHENTICATION (optional but recommended):
 *  Set a secret token in Script Properties:
 *      Project Settings → Script Properties → Add: WEBHOOK_SECRET = <your-token>
 *      Then include header  X-Webhook-Secret: <your-token>  on every request.
 *
 * ═══════════════════════════════════════════════════════════════════════
 *  SUPPORTED ACTIONS  (POST body must be JSON)
 * ═══════════════════════════════════════════════════════════════════════
 *
 *  1. findItem        — search for an item across all sheets
 *  2. updateItem       — update an item's description (and optionally other fields)
 *  3. addReceiptSheet  — add a new receipt sheet with line items
 *  4. getOverview      — return the Overview sheet as JSON
 *  5. getSummary       — return the Category Summary sheet as JSON
 *  6. getSheet         — return any named sheet as JSON
 *
 */

// ── Constants ───────────────────────────────────────────────────────────
const SPREADSHEET_ID  = "1Md91UVc4aaH7to6iMkQx8TaeAfryeA1dwEVXx6LXOsE";
const OVERVIEW_SHEET  = "Overview";
const SUMMARY_SHEET   = "Category Summary";

const COLORS = { dark: "1B3A4B", accent: "4A90D9", light: "E8F0FE" };

const NUM_COLS   = 9;
const COL_DATE   = 1;
const COL_NUM    = 2;
const COL_DESC   = 3;
const COL_CAT    = 4;
const COL_QTY    = 5;
const COL_UNIT   = 6;
const COL_AMOUNT = 7;
const COL_PROMO  = 8;
const COL_NOTES  = 9;

// ── GET handler ─────────────────────────────────────────────────────────
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || "";

  if (action === "getConfig") {
    const key = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY") || "";
    return ContentService.createTextOutput(JSON.stringify({ k: key }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "getResult") {
    const id = e.parameter.id || "";
    const cached = CacheService.getScriptCache().get("extract_" + id);
    if (!cached) {
      return ContentService.createTextOutput(JSON.stringify({ status: "pending" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(cached)
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true, message: "Receipt Tracker webhook is running." }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST handler ────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Optional auth check
    const secret = PropertiesService.getScriptProperties().getProperty("WEBHOOK_SECRET");
    if (secret) {
      const provided = (e.parameter && e.parameter.secret) || body.secret || "";
      if (provided !== secret) {
        return ContentService.createTextOutput(JSON.stringify({ ok: false, error: "Unauthorized" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    let result;
    switch (body.action) {
      case "findItem":
        result = findItem(ss, body);
        break;
      case "updateItem":
        result = updateItem(ss, body);
        break;
      case "addReceiptSheet":
        result = addReceiptSheet(ss, body);
        break;
      case "getOverview":
        result = getSheetData(ss, OVERVIEW_SHEET);
        break;
      case "getSummary":
        result = getSheetData(ss, SUMMARY_SHEET);
        break;
      case "getSheet":
        result = getSheetData(ss, body.sheet || "");
        break;
      default:
        result = { ok: false, error: "Unknown action: " + body.action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ═══════════════════════════════════════════════════════════════════════
//  ACTION: addReceiptSheet
//  Creates or appends to a store-based worksheet, updates Overview
//  and Category Summary.
// ═══════════════════════════════════════════════════════════════════════
function addReceiptSheet(ss, body) {
  const items     = body.items     || [];
  const title     = body.title     || body.sheet;
  const subtitle  = body.subtitle  || "";
  const footer    = body.footer    || "";
  const total     = body.total     || 0;
  const storeName = (body.sheet || "").trim();
  const receiptDate = (body.overviewEntry && body.overviewEntry.date) || "";
  if (!storeName) return { ok: false, error: "sheet (store name) is required" };

  // ── Find existing sheet by partial match ─────────────────────────────
  const reserved = [OVERVIEW_SHEET, SUMMARY_SHEET];
  let bestMatch = null;
  let bestLen = 0;
  ss.getSheets().forEach(function(s) {
    var name = s.getName();
    if (reserved.indexOf(name) > -1) return;
    var sLow = name.toLowerCase();
    var storeLow = storeName.toLowerCase();
    if (storeLow.startsWith(sLow) || sLow.startsWith(storeLow)) {
      var matchLen = Math.min(sLow.length, storeLow.length);
      if (matchLen > bestLen) { bestLen = matchLen; bestMatch = s; }
    }
  });

  var ws;
  var isAppend = false;

  if (bestMatch) {
    ws = bestMatch;
    isAppend = true;
  } else {
    // Insert new sheet after Category Summary
    var summarySheet = ss.getSheetByName(SUMMARY_SHEET);
    var insertIdx = summarySheet
      ? ss.getSheets().map(function(s) { return s.getName(); }).indexOf(SUMMARY_SHEET) + 1
      : ss.getNumSheets();
    ws = ss.insertSheet(storeName, insertIdx);
  }

  var HEADERS = ["Date", "#", "Item Description", "Category", "Qty", "Unit Price ($)", "Amount ($)", "Promo?", "Notes"];

  if (isAppend) {
    // ── Append receipt block to existing sheet ─────────────────────────
    var lastRow = ws.getLastRow();
    var startRow = lastRow + 2; // blank separator row

    // Receipt header bar
    ws.setRowHeight(startRow, 28);
    ws.getRange(startRow, 1, 1, NUM_COLS).merge()
      .setValue(subtitle || title)
      .setFontFamily("Arial").setFontSize(11).setFontWeight("bold").setFontColor("#FFFFFF")
      .setBackground("#" + COLORS.dark).setHorizontalAlignment("center").setVerticalAlignment("middle");

    // Column headers
    var hdrRow = startRow + 1;
    ws.setRowHeight(hdrRow, 28);
    HEADERS.forEach(function(h, i) {
      ws.getRange(hdrRow, i + 1).setValue(h)
        .setFontFamily("Arial").setFontSize(10).setFontWeight("bold").setFontColor("#FFFFFF")
        .setBackground("#" + COLORS.accent).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
    });

    // Item rows
    items.forEach(function(item, idx) {
      var row = hdrRow + 1 + idx;
      var bg = idx % 2 === 1 ? "#F5F5F5" : "#FFFFFF";
      writeItemRow(ws, row, idx + 1, item, bg, receiptDate);
    });

    // Total row
    var totalRow = hdrRow + 1 + items.length;
    writeTotalRow(ws, totalRow, total);

    // Footer
    if (footer) writeFooterRow(ws, totalRow + 1, footer);

  } else {
    // ── Create new sheet ───────────────────────────────────────────────
    // Title row
    ws.setRowHeight(1, 36);
    ws.getRange(1, 1, 1, NUM_COLS).merge().setValue(title)
      .setFontFamily("Arial").setFontSize(14).setFontWeight("bold").setFontColor("#FFFFFF")
      .setBackground("#" + COLORS.dark).setHorizontalAlignment("center").setVerticalAlignment("middle");

    // Subtitle row
    ws.setRowHeight(2, 28);
    ws.getRange(2, 1, 1, NUM_COLS).merge().setValue(subtitle)
      .setFontFamily("Arial").setFontSize(9).setFontColor("#555555")
      .setBackground("#EEEEEE").setHorizontalAlignment("center").setVerticalAlignment("middle");

    // Header row
    ws.setRowHeight(3, 36);
    HEADERS.forEach(function(h, i) {
      ws.getRange(3, i + 1).setValue(h)
        .setFontFamily("Arial").setFontSize(11).setFontWeight("bold").setFontColor("#FFFFFF")
        .setBackground("#" + COLORS.dark).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
    });

    // Item rows
    items.forEach(function(item, idx) {
      var row = idx + 4;
      var bg = idx % 2 === 1 ? "#F5F5F5" : "#FFFFFF";
      writeItemRow(ws, row, idx + 1, item, bg, receiptDate);
    });

    // Total row
    var totalRow = items.length + 4;
    writeTotalRow(ws, totalRow, total);

    // Footer
    if (footer) writeFooterRow(ws, totalRow + 1, footer);

    // Column widths
    ws.setColumnWidth(COL_DATE, 90);
    ws.setColumnWidth(COL_NUM, 40);
    ws.setColumnWidth(COL_DESC, 280);
    ws.setColumnWidth(COL_CAT, 120);
    ws.setColumnWidth(COL_QTY, 50);
    ws.setColumnWidth(COL_UNIT, 100);
    ws.setColumnWidth(COL_AMOUNT, 100);
    ws.setColumnWidth(COL_PROMO, 65);
    ws.setColumnWidth(COL_NOTES, 180);
  }

  // ── Append to Overview ────────────────────────────────────────────────
  var overviewUpdated = false;
  var ov = body.overviewEntry;
  if (ov) {
    var ovSheet = ss.getSheetByName(OVERVIEW_SHEET);
    if (ovSheet) {
      var lastOvRow = ovSheet.getLastRow();
      ovSheet.insertRowBefore(lastOvRow);
      var insertRow = lastOvRow;
      var rowData = [ov.date, ov.store, ov.location, ov.items, ov.total, ov.currency, ov.payment];
      var bg = insertRow % 2 === 0 ? "#F5F5F5" : "#FFFFFF";
      rowData.forEach(function(val, i) {
        var cell = ovSheet.getRange(insertRow, i + 1);
        cell.setValue(val).setFontFamily("Arial").setFontSize(10).setFontColor("#000000").setBackground(bg)
          .setHorizontalAlignment(i === 0 || i === 3 || i === 4 || i === 5 ? "center" : "left")
          .setVerticalAlignment("middle");
        if (i === 4) cell.setNumberFormat("$#,##0.00");
      });
      overviewUpdated = true;
    }
  }

  // ── Update Category Summary ───────────────────────────────────────────
  var summaryUpdated = false;
  var sumSheet = ss.getSheetByName(SUMMARY_SHEET);
  if (sumSheet && items.length > 0) {
    var sumData = sumSheet.getDataRange().getValues();
    var headerRow = -1;
    var grandRow = -1;
    for (var r = 0; r < sumData.length; r++) {
      if (String(sumData[r][0]).toLowerCase() === "category") headerRow = r;
      if (String(sumData[r][0]).toUpperCase() === "GRAND TOTAL") grandRow = r;
    }
    if (headerRow > -1 && grandRow > -1) {
      // Build category totals from items
      var catTotals = {};
      items.forEach(function(it) {
        var cat = (it.category || "Uncategorized").trim();
        var amt = (parseFloat(it.qty) || 1) * (parseFloat(it.unitPrice) || 0);
        catTotals[cat] = (catTotals[cat] || 0) + amt;
      });

      // Update existing or insert new categories
      var insertedRows = 0;
      Object.keys(catTotals).forEach(function(cat) {
        var found = false;
        for (var r = headerRow + 1; r < grandRow + insertedRows; r++) {
          var existingCat = String(sumSheet.getRange(r + 1, 1).getValue()).trim();
          if (existingCat === cat) {
            var oldVal = parseFloat(sumSheet.getRange(r + 1, 2).getValue()) || 0;
            sumSheet.getRange(r + 1, 2).setValue(oldVal + catTotals[cat]).setNumberFormat("$#,##0.00");
            found = true;
            break;
          }
        }
        if (!found) {
          var insertAt = grandRow + insertedRows + 1;
          sumSheet.insertRowBefore(insertAt);
          sumSheet.getRange(insertAt, 1).setValue(cat).setFontFamily("Arial").setFontSize(10);
          sumSheet.getRange(insertAt, 2).setValue(catTotals[cat]).setFontFamily("Arial").setFontSize(10).setNumberFormat("$#,##0.00");
          insertedRows++;
        }
      });

      // Recalculate grand total and percentages
      var freshData = sumSheet.getDataRange().getValues();
      var newHeaderRow = -1, newGrandRow = -1;
      for (var r = 0; r < freshData.length; r++) {
        if (String(freshData[r][0]).toLowerCase() === "category") newHeaderRow = r;
        if (String(freshData[r][0]).toUpperCase() === "GRAND TOTAL") newGrandRow = r;
      }
      if (newHeaderRow > -1 && newGrandRow > -1) {
        var grandTotal = 0;
        for (var r = newHeaderRow + 1; r < newGrandRow; r++) {
          grandTotal += (parseFloat(freshData[r][1]) || 0);
        }
        sumSheet.getRange(newGrandRow + 1, 2).setValue(grandTotal).setNumberFormat("$#,##0.00");
        sumSheet.getRange(newHeaderRow + 1, 3).setValue("% of $" + grandTotal.toFixed(2));
        for (var r = newHeaderRow + 1; r < newGrandRow; r++) {
          var catTotal = parseFloat(freshData[r][1]) || 0;
          sumSheet.getRange(r + 1, 3).setValue(grandTotal > 0 ? catTotal / grandTotal : 0)
            .setNumberFormat("0.0%");
        }
        sumSheet.getRange(newGrandRow + 1, 3).setValue(1).setNumberFormat("0.0%");
      }
      summaryUpdated = true;
    }
  }

  return {
    ok: true,
    sheet: ws.getName(),
    appended: isAppend,
    itemsAdded: items.length,
    total: total,
    overviewUpdated: overviewUpdated,
    summaryUpdated: summaryUpdated
  };
}

// ── Shared formatting helpers ───────────────────────────────────────────
function writeItemRow(ws, row, num, item, bg, date) {
  ws.getRange(row, COL_DATE).setValue(date || "").setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontColor("#000000").setHorizontalAlignment("center");
  ws.getRange(row, COL_NUM).setValue(num).setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontColor("#000000").setHorizontalAlignment("center");
  ws.getRange(row, COL_DESC).setValue(item.description || "").setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontColor("#000000");
  ws.getRange(row, COL_CAT).setValue(item.category || "").setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontColor("#000000");
  ws.getRange(row, COL_QTY).setValue(item.qty || 1).setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontColor("#000000").setHorizontalAlignment("center");
  ws.getRange(row, COL_UNIT).setValue(item.unitPrice || 0).setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontColor("#000000").setNumberFormat("$#,##0.00").setHorizontalAlignment("right");
  ws.getRange(row, COL_AMOUNT).setFormula("=E" + row + "*F" + row).setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontWeight("bold").setFontColor("#000000").setNumberFormat("$#,##0.00").setHorizontalAlignment("right");
  ws.getRange(row, COL_PROMO).setValue(item.promo ? "Yes" : "").setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontColor("#000000").setHorizontalAlignment("center");
  ws.getRange(row, COL_NOTES).setValue(item.notes || "").setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontColor("#000000");
}

function writeTotalRow(ws, row, total) {
  ws.setRowHeight(row, 28);
  ws.getRange(row, 1, 1, 4).merge();
  ws.getRange(row, COL_QTY).setValue("TOTAL").setFontFamily("Arial").setFontSize(11).setFontWeight("bold").setHorizontalAlignment("right");
  ws.getRange(row, COL_UNIT).setValue(total).setFontFamily("Arial").setFontSize(11).setFontWeight("bold").setNumberFormat("$#,##0.00").setHorizontalAlignment("right");
}

function writeFooterRow(ws, row, footer) {
  ws.getRange(row, 1, 1, NUM_COLS).merge().setValue(footer)
    .setFontFamily("Arial").setFontSize(9).setFontColor("#888888").setHorizontalAlignment("left");
}

// ═══════════════════════════════════════════════════════════════════════
//  MIGRATION: consolidateSheets
//  Run once from the Apps Script editor (Run → migrateSheets) to merge
//  old per-receipt tabs into consolidated per-store tabs with Date column.
//  Deletes old tabs after migration. Safe to run multiple times.
// ═══════════════════════════════════════════════════════════════════════
function migrateSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var reserved = [OVERVIEW_SHEET, SUMMARY_SHEET];
  var allSheets = ss.getSheets();

  // ── Group sheets by store base name ──────────────────────────────────
  // Old sheet names are like "Kmart 04-06-2026", "Kmart Cockburn 07-06-2026"
  // Extract the store part by removing the trailing date pattern (DD-MM-YYYY)
  var groups = {}; // { baseStore: [{ sheet, storeFull, date, subtitle }] }

  allSheets.forEach(function(s) {
    var name = s.getName();
    if (reserved.indexOf(name) > -1) return;
    if (name.startsWith("_")) return; // skip test sheets

    // Try to extract date from sheet name (DD-MM-YYYY at end)
    var dateMatch = name.match(/\s(\d{2}-\d{2}-\d{4})$/);
    if (!dateMatch) return; // not an old-format sheet, skip

    var dateStr = dateMatch[1]; // "04-06-2026"
    var storePart = name.replace(/\s\d{2}-\d{2}-\d{4}$/, "").trim();

    // Get base store name: find the shortest prefix that groups related stores
    // e.g. "Kmart", "Kmart Cockburn", "Kmart Booragoon" → base is "Kmart"
    // "FairPrice City Square Mall" → base is "FairPrice"
    // For single-word stores like "Woolworths" → base is "Woolworths"
    var baseStore = storePart;

    // Check if this store's first word matches an existing group
    var firstWord = storePart.split(/\s+/)[0].toLowerCase();
    var matchedBase = null;
    Object.keys(groups).forEach(function(existing) {
      var existFirst = existing.split(/\s+/)[0].toLowerCase();
      if (firstWord === existFirst) matchedBase = existing;
    });

    if (matchedBase) {
      // Use the shorter name as the base
      if (storePart.length < matchedBase.length) {
        groups[storePart] = groups[matchedBase];
        delete groups[matchedBase];
        baseStore = storePart;
      } else {
        baseStore = matchedBase;
      }
    }

    if (!groups[baseStore]) groups[baseStore] = [];

    // Read subtitle from row 2 for receipt details
    var subtitle = "";
    try { subtitle = s.getRange(2, 1).getValue() || ""; } catch(e) {}

    // Convert date DD-MM-YYYY to DD/MM/YYYY for display
    var dateParts = dateStr.split("-");
    var dateFmt = dateParts[0] + "/" + dateParts[1] + "/" + dateParts[2];

    groups[baseStore].push({
      sheet: s,
      name: name,
      storeFull: storePart,
      date: dateFmt,
      subtitle: subtitle
    });
  });

  // ── Sort each group by date ──────────────────────────────────────────
  Object.keys(groups).forEach(function(base) {
    groups[base].sort(function(a, b) {
      var dA = a.date.split("/"); // DD/MM/YYYY
      var dB = b.date.split("/");
      var tA = new Date(dA[2], dA[1] - 1, dA[0]).getTime();
      var tB = new Date(dB[2], dB[1] - 1, dB[0]).getTime();
      return tA - tB;
    });
  });

  // ── Create consolidated sheets ───────────────────────────────────────
  var HEADERS = ["Date", "#", "Item Description", "Category", "Qty", "Unit Price ($)", "Amount ($)", "Promo?", "Notes"];
  var summarySheet = ss.getSheetByName(SUMMARY_SHEET);
  var insertIdx = summarySheet
    ? ss.getSheets().map(function(s) { return s.getName(); }).indexOf(SUMMARY_SHEET) + 1
    : ss.getNumSheets();

  var migrated = 0;

  Object.keys(groups).forEach(function(baseStore) {
    var entries = groups[baseStore];
    if (entries.length === 0) return;

    // Create new consolidated sheet
    var newSheet = ss.insertSheet(baseStore + " (new)", insertIdx);
    insertIdx++;

    // Title row
    newSheet.setRowHeight(1, 36);
    newSheet.getRange(1, 1, 1, NUM_COLS).merge()
      .setValue(baseStore + " — Purchase Tracker")
      .setFontFamily("Arial").setFontSize(14).setFontWeight("bold").setFontColor("#FFFFFF")
      .setBackground("#" + COLORS.dark).setHorizontalAlignment("center").setVerticalAlignment("middle");

    // Subtitle
    newSheet.setRowHeight(2, 28);
    newSheet.getRange(2, 1, 1, NUM_COLS).merge()
      .setValue("Consolidated receipts for " + baseStore)
      .setFontFamily("Arial").setFontSize(9).setFontColor("#555555")
      .setBackground("#EEEEEE").setHorizontalAlignment("center").setVerticalAlignment("middle");

    // Header row
    newSheet.setRowHeight(3, 36);
    HEADERS.forEach(function(h, i) {
      newSheet.getRange(3, i + 1).setValue(h)
        .setFontFamily("Arial").setFontSize(11).setFontWeight("bold").setFontColor("#FFFFFF")
        .setBackground("#" + COLORS.dark).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
    });

    var currentRow = 4;

    // Process each old receipt sheet
    entries.forEach(function(entry, entryIdx) {
      var oldSheet = entry.sheet;
      var data = oldSheet.getDataRange().getValues();

      // Find old header row (row with "#" in column A or B)
      var oldHeaderIdx = -1;
      for (var r = 0; r < data.length; r++) {
        if (String(data[r][0]).trim() === "#" || String(data[r][1]).trim() === "#") {
          oldHeaderIdx = r;
          break;
        }
      }
      if (oldHeaderIdx === -1) return;

      // Determine if old format has Date column (9 cols) or not (8 cols)
      var oldHasDate = String(data[oldHeaderIdx][0]).toLowerCase().trim() === "date";
      var colOffset = oldHasDate ? 0 : -1; // if no date col, old # is col 0 (A)

      // Add separator + receipt header bar between receipts (not before first)
      if (entryIdx > 0) {
        currentRow++; // blank separator row
        newSheet.setRowHeight(currentRow, 28);
        newSheet.getRange(currentRow, 1, 1, NUM_COLS).merge()
          .setValue(entry.subtitle || (entry.storeFull + " | " + entry.date))
          .setFontFamily("Arial").setFontSize(11).setFontWeight("bold").setFontColor("#FFFFFF")
          .setBackground("#" + COLORS.dark).setHorizontalAlignment("center").setVerticalAlignment("middle");
        currentRow++;

        // Column headers for appended blocks
        newSheet.setRowHeight(currentRow, 28);
        HEADERS.forEach(function(h, i) {
          newSheet.getRange(currentRow, i + 1).setValue(h)
            .setFontFamily("Arial").setFontSize(10).setFontWeight("bold").setFontColor("#FFFFFF")
            .setBackground("#" + COLORS.accent).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
        });
        currentRow++;
      }

      // Copy item rows
      var itemNum = 1;
      for (var r = oldHeaderIdx + 1; r < data.length; r++) {
        var row = data[r];
        // Stop at total/footer rows
        var firstCell = String(row[0] || "").trim().toUpperCase();
        var secondCell = String(row[1] || "").trim().toUpperCase();
        if (firstCell === "TOTAL" || firstCell === "TOTAL PAID" || firstCell === "" && secondCell === "") {
          // Check if this is a total row
          var hasTotalKeyword = false;
          for (var c = 0; c < row.length; c++) {
            if (String(row[c] || "").toUpperCase().indexOf("TOTAL") > -1) { hasTotalKeyword = true; break; }
          }
          if (hasTotalKeyword) break;
          if (firstCell === "" && secondCell === "") break;
        }

        // Read values based on old column layout
        var desc, cat, qty, unitPrice, amount, promo, notes;
        if (oldHasDate) {
          // Old 9-col: Date, #, Desc, Cat, Qty, Unit, Amount, Promo, Notes
          desc = row[2]; cat = row[3]; qty = row[4]; unitPrice = row[5]; amount = row[6]; promo = row[7]; notes = row[8];
        } else {
          // Old 8-col: #, Desc, Cat, Qty, Unit, Amount, Promo, Notes
          desc = row[1]; cat = row[2]; qty = row[3]; unitPrice = row[4]; amount = row[5]; promo = row[6]; notes = row[7];
        }

        // Skip non-item rows (header text, empty rows)
        if (!desc || String(desc).trim() === "" || String(desc).trim() === "Item Description") continue;

        var bg = (itemNum - 1) % 2 === 1 ? "#F5F5F5" : "#FFFFFF";
        newSheet.getRange(currentRow, COL_DATE).setValue(entry.date).setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontColor("#000000").setHorizontalAlignment("center");
        newSheet.getRange(currentRow, COL_NUM).setValue(itemNum).setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontColor("#000000").setHorizontalAlignment("center");
        newSheet.getRange(currentRow, COL_DESC).setValue(desc || "").setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontColor("#000000");
        newSheet.getRange(currentRow, COL_CAT).setValue(cat || "").setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontColor("#000000");
        newSheet.getRange(currentRow, COL_QTY).setValue(qty || 1).setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontColor("#000000").setHorizontalAlignment("center");
        newSheet.getRange(currentRow, COL_UNIT).setValue(unitPrice || 0).setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontColor("#000000").setNumberFormat("$#,##0.00").setHorizontalAlignment("right");
        newSheet.getRange(currentRow, COL_AMOUNT).setFormula("=E" + currentRow + "*F" + currentRow).setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontWeight("bold").setFontColor("#000000").setNumberFormat("$#,##0.00").setHorizontalAlignment("right");
        newSheet.getRange(currentRow, COL_PROMO).setValue(String(promo || "").trim()).setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontColor("#000000").setHorizontalAlignment("center");
        newSheet.getRange(currentRow, COL_NOTES).setValue(notes || "").setBackground(bg).setFontFamily("Arial").setFontSize(10).setFontColor("#000000");

        currentRow++;
        itemNum++;
      }

      // Find and write total for this receipt
      var receiptTotal = 0;
      for (var r = oldHeaderIdx + 1; r < data.length; r++) {
        for (var c = 0; c < data[r].length; c++) {
          if (String(data[r][c] || "").toUpperCase().indexOf("TOTAL") > -1) {
            // Look for the numeric value in adjacent cells
            for (var cc = c + 1; cc < data[r].length; cc++) {
              var val = parseFloat(data[r][cc]);
              if (!isNaN(val) && val > 0) { receiptTotal = val; break; }
            }
            if (receiptTotal > 0) break;
          }
        }
        if (receiptTotal > 0) break;
      }
      writeTotalRow(newSheet, currentRow, receiptTotal);
      currentRow++;

      // Find and write footer
      for (var r = data.length - 1; r > oldHeaderIdx; r--) {
        var footerText = String(data[r][0] || "").trim();
        if (footerText.startsWith("*") || footerText.startsWith("Payment")) {
          writeFooterRow(newSheet, currentRow, footerText);
          currentRow++;
          break;
        }
      }

      migrated++;
    });

    // Column widths
    newSheet.setColumnWidth(COL_DATE, 90);
    newSheet.setColumnWidth(COL_NUM, 40);
    newSheet.setColumnWidth(COL_DESC, 280);
    newSheet.setColumnWidth(COL_CAT, 120);
    newSheet.setColumnWidth(COL_QTY, 50);
    newSheet.setColumnWidth(COL_UNIT, 100);
    newSheet.setColumnWidth(COL_AMOUNT, 100);
    newSheet.setColumnWidth(COL_PROMO, 65);
    newSheet.setColumnWidth(COL_NOTES, 180);

    // Delete old sheets
    entries.forEach(function(entry) {
      ss.deleteSheet(entry.sheet);
    });

    // Rename new sheet (remove " (new)" suffix)
    newSheet.setName(baseStore);
  });

  // Delete test sheets
  ss.getSheets().forEach(function(s) {
    if (s.getName().startsWith("_test")) {
      ss.deleteSheet(s);
    }
  });

  Logger.log("Migration complete. Consolidated " + migrated + " receipts into " + Object.keys(groups).length + " store sheets.");
}

// ═══════════════════════════════════════════════════════════════════════
//  ACTION: findItem — search across all receipt sheets
// ═══════════════════════════════════════════════════════════════════════
function findItem(ss, body) {
  var query = (body.query || "").toLowerCase().trim();
  if (!query) return { ok: false, error: "query is required" };

  var results = [];
  var reserved = [OVERVIEW_SHEET, SUMMARY_SHEET];
  ss.getSheets().forEach(function(sheet) {
    if (reserved.indexOf(sheet.getName()) > -1) return;
    var data = sheet.getDataRange().getValues();
    data.forEach(function(row, rowIdx) {
      var desc = String(row[COL_DESC - 1] || "").toLowerCase();
      if (desc.indexOf(query) > -1) {
        results.push({
          sheet: sheet.getName(),
          row: rowIdx + 1,
          date: row[COL_DATE - 1],
          description: row[COL_DESC - 1],
          category: row[COL_CAT - 1],
          qty: row[COL_QTY - 1],
          unitPrice: row[COL_UNIT - 1],
          amount: row[COL_AMOUNT - 1],
          promo: row[COL_PROMO - 1],
          notes: row[COL_NOTES - 1]
        });
      }
    });
  });

  return { ok: true, query: query, matches: results.length, results: results };
}

// ═══════════════════════════════════════════════════════════════════════
//  ACTION: updateItem — update an item's fields
// ═══════════════════════════════════════════════════════════════════════
function updateItem(ss, body) {
  var query = (body.query || "").toLowerCase().trim();
  if (!query) return { ok: false, error: "query is required" };

  var updates = body.updates || {};
  var reserved = [OVERVIEW_SHEET, SUMMARY_SHEET];
  var updated = 0;

  ss.getSheets().forEach(function(sheet) {
    if (reserved.indexOf(sheet.getName()) > -1) return;
    var data = sheet.getDataRange().getValues();
    data.forEach(function(row, rowIdx) {
      var desc = String(row[COL_DESC - 1] || "").toLowerCase();
      if (desc.indexOf(query) > -1) {
        if (updates.description !== undefined) sheet.getRange(rowIdx + 1, COL_DESC).setValue(updates.description);
        if (updates.category    !== undefined) sheet.getRange(rowIdx + 1, COL_CAT).setValue(updates.category);
        if (updates.qty         !== undefined) sheet.getRange(rowIdx + 1, COL_QTY).setValue(updates.qty);
        if (updates.unitPrice   !== undefined) sheet.getRange(rowIdx + 1, COL_UNIT).setValue(updates.unitPrice);
        if (updates.notes       !== undefined) sheet.getRange(rowIdx + 1, COL_NOTES).setValue(updates.notes);
        if (updates.promo       !== undefined) sheet.getRange(rowIdx + 1, COL_PROMO).setValue(updates.promo ? "Yes" : "");
        updated++;
      }
    });
  });

  return { ok: true, query: query, updated: updated };
}

// ═══════════════════════════════════════════════════════════════════════
//  ACTION: getSheetData — return any sheet as JSON
// ═══════════════════════════════════════════════════════════════════════
function getSheetData(ss, sheetName) {
  if (!sheetName) return { ok: false, error: "sheet name is required" };
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { ok: false, error: "Sheet not found: " + sheetName };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, sheet: sheetName, rows: [] };

  // Find the header row (first row with content in column A)
  var headerIdx = 0;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] !== "" && data[i][0] !== null) { headerIdx = i; break; }
  }

  var headers = data[headerIdx].map(function(h) {
    return String(h).toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_|_$/g, "");
  });

  var rows = [];
  for (var r = headerIdx + 1; r < data.length; r++) {
    var obj = {};
    var hasContent = false;
    headers.forEach(function(key, i) {
      if (key) {
        obj[key] = data[r][i];
        if (data[r][i] !== "" && data[r][i] !== null) hasContent = true;
      }
    });
    if (hasContent) rows.push(obj);
  }

  return { ok: true, sheet: sheetName, rows: rows };
}
