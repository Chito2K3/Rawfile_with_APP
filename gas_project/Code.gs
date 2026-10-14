/**
 * ============================================================================
 * PHARMACY STORAGE INVENTORY & EXPIRY MANAGEMENT SYSTEM
 * Google Apps Script Backend (Code.gs)
 * Amang Rodriguez Memorial Medical Center (ARMMC)
 * ============================================================================
 */

// Target or active sheet configuration
const CONFIG = {
  MAIN_SHEET_NAME: null, // null will auto-detect active inventory sheet or look for standard headers
  MASTER_DCI_SHEET_NAME: 'Sheet1', // Master catalog tab with Col A: DCI Code, Col B: Generic Description
  LOG_SHEET_NAME: 'STOCK_OUT_LOG',
  COLUMN_KEYS: {
    CODE: 'ITEM CODE',
    DESC: 'ITEM DESCRIPTION',
    BRAND: 'BRAND NAME',
    QTY: 'QTY STORAGE',
    EXP: 'EXP DATE',
    PRICE: 'ACQUISITION PRICE',
    DATE_OUT: 'DATE OUT'
  }
};

/**
 * Automatically creates custom menu on spreadsheet open
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏥 Pharmacy Expiry Monitoring')
    .addItem('🚀 Open Inventory Manager (Sidebar)', 'showSidebar')
    .addItem('🖥️ Open Inventory Manager (Full Dialog)', 'showDialog')
    .addSeparator()
    .addItem('➖ Quick Stock-Out / Dispense', 'showStockOutDialog')
    .addItem('➕ Add Medicine or New Batch', 'showAddBatchDialog')
    .addItem('🔄 Restock 0-Qty Items', 'showRestockDialog')
    .addSeparator()
    .addItem('📋 View Stock-Out Audit Log', 'openLogSheet')
    .addItem('🛠️ Ensure / Upgrade Sheet Columns (Option A)', 'migrateSheetToOptionA')
    .addToUi();
}

/**
 * Web App entrypoint (for standalone web app deployment)
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Pharmacy Expiry Monitoring - ARMMC')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Shows the application in the Google Sheets Sidebar
 */
function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Pharmacy Expiry Monitoring - ARMMC');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows the application in a large Modal Dialog
 */
function showDialog() {
  const html = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setWidth(1200)
    .setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, '🏥 Pharmacy Expiry Monitoring - Amang Rodriguez Memorial Medical Center');
}

function showStockOutDialog() {
  showDialog();
}
function showAddBatchDialog() {
  showDialog();
}
function showRestockDialog() {
  showDialog();
}

/**
 * Helper to include HTML sub-files (CSS, JS)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Navigates or focuses on the STOCK_OUT_LOG sheet
 */
function openLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = initLogSheet(ss);
  }
  ss.setActiveSheet(logSheet);
}

/**
 * Resolves the primary Inventory sheet and locates its headers
 * Option A Standard: ITEM CODE | ITEM DESCRIPTION | BRAND NAME | QTY STORAGE | EXP DATE | ACQUISITION PRICE
 */
function getInventorySheetInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = null;

  if (CONFIG.MAIN_SHEET_NAME) {
    sheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  }
  
  if (!sheet) {
    // Check if 'STORAGE' inventory sheet exists
    sheet = ss.getSheetByName('STORAGE') || ss.getSheetByName('Storage');
  }

  if (!sheet) {
    // Try to find the inventory sheet that has ITEM CODE and QTY STORAGE, avoiding the master DCI catalog tab
    const sheets = ss.getSheets();
    for (let s of sheets) {
      const name = s.getName();
      if (name === CONFIG.LOG_SHEET_NAME) continue;
      if (CONFIG.MASTER_DCI_SHEET_NAME && name === CONFIG.MASTER_DCI_SHEET_NAME) continue;
      
      const data = s.getDataRange().getValues();
      if (data.length > 0) {
        const headerRow = data[0].map(c => String(c).trim().toUpperCase());
        if (headerRow.some(h => h.includes('ITEM CODE') || h.includes('ITEM_CODE'))) {
          sheet = s;
          break;
        }
      }
    }
  }

  if (!sheet) {
    // Fallback: check any sheet with ITEM CODE header
    const sheets = ss.getSheets();
    for (let s of sheets) {
      if (s.getName() === CONFIG.LOG_SHEET_NAME) continue;
      const data = s.getDataRange().getValues();
      if (data.length > 0) {
        const headerRow = data[0].map(c => String(c).trim().toUpperCase());
        if (headerRow.some(h => h.includes('ITEM CODE') || h.includes('ITEM_CODE'))) {
          sheet = s;
          break;
        }
      }
    }
  }

  if (!sheet) {
    sheet = ss.getSheets()[0]; // Fallback to first sheet
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  if (values.length === 0) {
    throw new Error("Inventory sheet is empty.");
  }

  // Find column indices
  const headerRow = values[0];
  let colCode = -1, colDesc = -1, colBrand = -1, colQty = -1, colExp = -1, colPrice = -1, colDateOut = -1;

  for (let j = 0; j < headerRow.length; j++) {
    const h = String(headerRow[j]).trim().toUpperCase();
    if (colCode === -1 && (h.includes('CODE') || h.includes('ITEM CODE'))) colCode = j;
    if (colDesc === -1 && (h.includes('DESC') || h.includes('DESCRIPTION') || h.includes('GENERIC'))) colDesc = j;
    if (colBrand === -1 && (h.includes('BRAND') || h.includes('TRADE'))) colBrand = j;
    if (colQty === -1 && (h.includes('QTY') || h.includes('STORAGE') || h.includes('QUANTITY'))) colQty = j;
    if (colExp === -1 && (h.includes('EXP') || h.includes('DATE') || h.includes('EXPIRY')) && !h.includes('OUT')) colExp = j;
    if (colPrice === -1 && (h.includes('PRICE') || h.includes('COST') || h.includes('ACQUISITION'))) colPrice = j;
    if (colDateOut === -1 && (h.includes('DATE OUT') || h.includes('DATE_OUT') || h.includes('OUT DATE') || h.includes('DISPENSED DATE'))) colDateOut = j;
  }

  // Fallbacks if columns are not explicitly labeled
  if (colCode === -1) colCode = 0;
  if (colDesc === -1) colDesc = 1;

  // If Option A layout is present or Brand was found:
  if (colBrand !== -1) {
    if (colQty === -1) colQty = 3;
    if (colExp === -1) colExp = 4;
    if (colPrice === -1) colPrice = 5;
    if (colDateOut === -1 && headerRow.length >= 7) colDateOut = 6;
  } else {
    // Legacy 4-column fallback: Code(0), Desc(1), Qty(2), Exp(3)
    if (colQty === -1) colQty = 2;
    if (colExp === -1) colExp = 3;
  }

  return {
    sheet: sheet,
    sheetName: sheet.getName(),
    values: values,
    colCode: colCode,
    colDesc: colDesc,
    colBrand: colBrand,
    colQty: colQty,
    colExp: colExp,
    colPrice: colPrice,
    colDateOut: colDateOut
  };
}

/**
 * Migration Helper: Upgrades the current inventory sheet to Option A layout:
 * Column A: ITEM CODE
 * Column B: ITEM DESCRIPTION
 * Column C: BRAND NAME
 * Column D: QTY STORAGE
 * Column E: EXP DATE
 * Column F: ACQUISITION PRICE (Repository only)
 * Column G: DATE OUT (Last Dispensed Date)
 */
function migrateSheetToOptionA() {
  const ui = SpreadsheetApp.getUi();
  try {
    const info = getInventorySheetInfo();
    const sheet = info.sheet;

    // Ensure sheet has at least 7 columns
    while (sheet.getMaxColumns() < 7) {
      sheet.insertColumnAfter(sheet.getMaxColumns());
    }

    const checkCols = Math.max(sheet.getLastColumn(), 12);
    const headerRow = sheet.getRange(1, 1, 1, checkCols).getValues()[0];
    const upperHeaders = headerRow.map(h => String(h).trim().toUpperCase());

    // 1. Ensure BRAND NAME is at Column C (index 3)
    const brandColIdx = upperHeaders.findIndex(h => h.includes('BRAND'));
    if (brandColIdx === -1) {
      // Insert BRAND NAME after Column B (Description)
      sheet.insertColumnAfter(2);
      sheet.getRange(1, 3).setValue('BRAND NAME').setFontWeight('bold');
    } else if (brandColIdx !== 2) {
      // Ensure Column C has BRAND NAME
      sheet.getRange(1, 3).setValue('BRAND NAME').setFontWeight('bold');
    }

    // 2. Clear any stray ACQUISITION PRICE headers that may have landed in Column H, I, J
    for (let c = 8; c <= Math.max(sheet.getLastColumn(), 12); c++) {
      const val = String(sheet.getRange(1, c).getValue()).trim().toUpperCase();
      if (val.includes('PRICE') || val.includes('ACQUISITION')) {
        sheet.getRange(1, c).clearContent();
      }
    }

    // 3. Explicitly place ACQUISITION PRICE in Column F (Column 6)
    sheet.getRange(1, 6).setValue('ACQUISITION PRICE').setFontWeight('bold');

    // 4. Format Column F as Currency (₱#,##0.00)
    const maxRows = sheet.getMaxRows();
    if (maxRows > 1) {
      sheet.getRange(2, 6, maxRows - 1, 1).setNumberFormat('"₱"#,##0.00');
    }

    // 5. Explicitly place DATE OUT in Column G (Column 7)
    sheet.getRange(1, 7).setValue('DATE OUT').setFontWeight('bold').setHorizontalAlignment('center');
    if (maxRows > 1) {
      sheet.getRange(2, 7, maxRows - 1, 1).setNumberFormat('@').setHorizontalAlignment('center');
    }

    ui.alert(
      'Option A Setup Complete',
      'The sheet columns have been confirmed:\n\n' +
      '• Column A: ITEM CODE\n' +
      '• Column B: ITEM DESCRIPTION\n' +
      '• Column C: BRAND NAME\n' +
      '• Column D: QTY STORAGE\n' +
      '• Column E: EXP DATE\n' +
      '• Column F: ACQUISITION PRICE (Repository only)\n' +
      '• Column G: DATE OUT (Last Dispensed Date)\n\n' +
      'Brand Name, Acquisition Price, and Date Out are now active!',
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('Error', 'Migration failed: ' + err.message, ui.ButtonSet.OK);
  }
}

/**
 * Fetches all inventory records with parsed quantities, row numbers, brand name, price, and expiry status
 */
function getInventoryData() {
  try {
    const info = getInventorySheetInfo();
    const rows = [];
    const ssTz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

    for (let i = 1; i < info.values.length; i++) {
      const row = info.values[i];
      const code = String(row[info.colCode] || '').trim();
      if (!code) continue; // Skip empty rows

      const desc = String(row[info.colDesc] || '').trim();
      const brand = info.colBrand !== -1 ? String(row[info.colBrand] || '').trim() : '';
      const rawQty = row[info.colQty];
      const rawExpVal = row[info.colExp];
      const rawPrice = info.colPrice !== -1 ? row[info.colPrice] : 0;

      // Clean quantity
      let cleanQty = 0;
      if (typeof rawQty === 'number') {
        cleanQty = rawQty;
      } else {
        const parsed = String(rawQty).replace(/[^0-9.-]/g, '');
        cleanQty = parsed ? parseFloat(parsed) : 0;
      }

      // Clean acquisition price
      let cleanPrice = 0;
      if (typeof rawPrice === 'number') {
        cleanPrice = rawPrice;
      } else {
        const parsedPrice = String(rawPrice).replace(/[^0-9.-]/g, '');
        cleanPrice = parsedPrice ? parseFloat(parsedPrice) : 0;
      }

      // Parse expiration robustly (handles Date objects, timestamps, strings like 'Feb-27' or '2027-02-01')
      let expYear = null;
      let expLabel = 'NO EXPIRY';

      if (rawExpVal instanceof Date) {
        expLabel = Utilities.formatDate(rawExpVal, ssTz, 'MMM-yy');
        expYear = parseInt(Utilities.formatDate(rawExpVal, ssTz, 'yyyy'), 10);
      } else if (rawExpVal) {
        const strExp = String(rawExpVal).trim();
        // Check for format MMM-yy (e.g. Feb-27, May-30)
        const matchShort = strExp.match(/^([A-Za-z]{3})-(\d{2})$/);
        if (matchShort) {
          expYear = 2000 + parseInt(matchShort[2], 10);
          const m = matchShort[1].charAt(0).toUpperCase() + matchShort[1].slice(1).toLowerCase();
          expLabel = `${m}-${matchShort[2]}`;
        } else {
          const parsedD = new Date(strExp);
          if (!isNaN(parsedD.getTime()) && strExp.length > 8) {
            expLabel = Utilities.formatDate(parsedD, ssTz, 'MMM-yy');
            expYear = parseInt(Utilities.formatDate(parsedD, ssTz, 'yyyy'), 10);
          } else {
            const matchYr = strExp.match(/20(\d{2})/);
            if (matchYr) {
              expYear = 2000 + parseInt(matchYr[1], 10);
            }
            expLabel = strExp || 'NO EXPIRY';
          }
        }
      }

      let expStatus = 'SAFE'; // CRITICAL_2026, NEAR_2027, SAFE, EXPIRED, NO_EXPIRY
      if (!expYear) {
        expStatus = 'NO_EXPIRY';
      } else if (expYear < 2026) {
        expStatus = 'EXPIRED';
      } else if (expYear === 2026) {
        expStatus = 'CRITICAL_2026';
      } else if (expYear === 2027) {
        expStatus = 'NEAR_2027';
      } else {
        expStatus = 'SAFE';
      }

      const isDonation = code.startsWith('DMDON') || code.startsWith('DMCSP');

      // Format Date Out cleanly (YYYY-MM-DD)
      let dateOutStr = '';
      if (info.colDateOut !== -1 && row[info.colDateOut]) {
        const rawDateOut = row[info.colDateOut];
        if (rawDateOut instanceof Date) {
          dateOutStr = Utilities.formatDate(rawDateOut, ssTz, 'yyyy-MM-dd');
        } else {
          dateOutStr = String(rawDateOut).trim();
        }
      }

      rows.push({
        rowIndex: i + 1, // 1-based row index in Google Sheet
        itemCode: code,
        itemDesc: desc,
        brandName: brand,
        qty: cleanQty,
        expDate: expLabel,
        expYear: expYear,
        expStatus: expStatus,
        acquisitionPrice: cleanPrice,
        dateOut: dateOutStr,
        isDonation: isDonation,
        isZeroStock: cleanQty <= 0
      });
    }

    const masterCatalogRes = getMasterDciCatalog();

    return {
      success: true,
      sheetName: info.sheetName,
      totalCount: rows.length,
      hasBrandColumn: info.colBrand !== -1,
      hasPriceColumn: info.colPrice !== -1,
      hasDateOutColumn: info.colDateOut !== -1,
      items: rows,
      catalog: masterCatalogRes.catalog || []
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Fetches the master list of DCI Codes and Generic Descriptions from 'Sheet1'
 * Column A: DCI Code (Item Code)
 * Column B: Generic Description & Strength
 */
function getMasterDciCatalog() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.MASTER_DCI_SHEET_NAME || 'Sheet1');
    if (!sheet) {
      const sheets = ss.getSheets();
      for (let s of sheets) {
        const sName = s.getName().trim().toLowerCase();
        if (sName === 'sheet1' || sName === 'sheet 1' || sName.includes('catalog') || sName.includes('dci')) {
          sheet = s;
          break;
        }
      }
    }

    if (!sheet) {
      return { success: false, catalog: [], message: "Master DCI tab ('Sheet1') not found." };
    }

    const data = sheet.getDataRange().getValues();
    if (!data || data.length === 0) {
      return { success: true, catalog: [] };
    }

    const catalog = [];
    const seenCodes = new Set();

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rawCode = String(row[0] || '').trim();
      const rawDesc = String(row[1] || '').trim();

      if (!rawCode) continue;

      // Skip header row if it contains header text like "CODE" or "DESCRIPTION"
      if (i === 0 && (rawCode.toUpperCase().includes('CODE') || rawDesc.toUpperCase().includes('DESC'))) {
        continue;
      }

      const upperCode = rawCode.toUpperCase();
      if (!seenCodes.has(upperCode)) {
        seenCodes.add(upperCode);
        catalog.push({
          code: rawCode,
          desc: rawDesc
        });
      }
    }

    // Sort alphabetically by Generic Description
    catalog.sort((a, b) => a.desc.localeCompare(b.desc));

    return {
      success: true,
      catalog: catalog,
      totalCount: catalog.length
    };
  } catch (err) {
    return { success: false, catalog: [], error: err.message };
  }
}

/**
 * Smart Add Medicine or New Batch
 * If itemCode exists, finds the last occurrence and inserts a row directly below it!
 * If itemCode does not exist, appends to the end of the inventory list.
 * Supports Brand Name and Acquisition Price per batch.
 */
function addNewMedicineOrBatch(payload) {
  try {
    const { itemCode, itemDesc, brandName, qty, expDate, acquisitionPrice } = payload;
    if (!itemCode || !itemDesc) {
      throw new Error("Item Code (DCI) and Generic Description are required.");
    }
    if (!brandName || String(brandName).trim() === '') {
      throw new Error("Brand Name is required for every batch delivery.");
    }

    const cleanCode = String(itemCode).trim();
    const cleanDesc = String(itemDesc).trim();
    const cleanBrand = String(brandName).trim();
    const cleanQty = parseFloat(String(qty).replace(/[^0-9.-]/g, '')) || 0;
    const cleanExp = String(expDate || '').trim();
    const cleanPrice = parseFloat(String(acquisitionPrice || '0').replace(/[^0-9.-]/g, '')) || 0;

    const info = getInventorySheetInfo();
    const sheet = info.sheet;
    const values = info.values;

    // Search for existing item occurrences
    let lastMatchingRowIndex = -1;
    let existingCount = 0;

    for (let i = 1; i < values.length; i++) {
      const code = String(values[i][info.colCode] || '').trim().toUpperCase();
      if (code === cleanCode.toUpperCase()) {
        lastMatchingRowIndex = i + 1; // 1-based index in sheet
        existingCount++;
      }
    }

    let insertedRowIndex = -1;
    let mode = '';

    if (lastMatchingRowIndex !== -1) {
      // Existing item: Insert directly below the last existing batch row
      sheet.insertRowAfter(lastMatchingRowIndex);
      insertedRowIndex = lastMatchingRowIndex + 1;
      mode = `Added as new batch #${existingCount + 1} below row ${lastMatchingRowIndex}`;
    } else {
      // New item: Append to the end of sheet
      insertedRowIndex = sheet.getLastRow() + 1;
      mode = `Added as brand new medicine at row ${insertedRowIndex}`;
    }

    // Set values and explicit alignments for the new row
    const ssTz = sheet.getParent().getSpreadsheetTimeZone();
    const formattedExp = normalizeExpiryString(cleanExp, ssTz);

    // 1. Item Code (Left)
    sheet.getRange(insertedRowIndex, info.colCode + 1)
      .setValue(cleanCode)
      .setHorizontalAlignment('left');

    // 2. Generic Name (Left)
    sheet.getRange(insertedRowIndex, info.colDesc + 1)
      .setValue(cleanDesc)
      .setHorizontalAlignment('left');

    // 3. Brand Name (Left)
    if (info.colBrand !== -1) {
      sheet.getRange(insertedRowIndex, info.colBrand + 1)
        .setValue(cleanBrand)
        .setHorizontalAlignment('left');
    }

    // 4. Storage Quantity (Right aligned, formatted as #,##0)
    sheet.getRange(insertedRowIndex, info.colQty + 1)
      .setValue(cleanQty)
      .setNumberFormat('#,##0')
      .setHorizontalAlignment('right');
    
    // 5. Expiry Date (Plain text '@' centered to avoid timezone date shifts)
    sheet.getRange(insertedRowIndex, info.colExp + 1)
      .setNumberFormat('@')
      .setValue(formattedExp)
      .setHorizontalAlignment('center');

    // 6. Acquisition Price (Right aligned, formatted as ₱#,##0.00)
    if (info.colPrice !== -1) {
      sheet.getRange(insertedRowIndex, info.colPrice + 1)
        .setValue(cleanPrice)
        .setNumberFormat('"₱"#,##0.00')
        .setHorizontalAlignment('right');
    }

    // 7. Date Out (Empty for newly added batch)
    if (info.colDateOut !== -1) {
      sheet.getRange(insertedRowIndex, info.colDateOut + 1)
        .setValue('')
        .setHorizontalAlignment('center');
    }

    return {
      success: true,
      message: `Successfully saved: [${cleanCode}] ${cleanDesc} (${cleanBrand})`,
      mode: mode,
      insertedRowIndex: insertedRowIndex
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Normalizes expiry strings (e.g. "may-30" -> "May-30") and handles Date objects using Spreadsheet timezone
 */
function normalizeExpiryString(inputVal, tz) {
  if (!inputVal) return 'NO EXPIRY';
  if (inputVal instanceof Date) {
    return Utilities.formatDate(inputVal, tz || Session.getScriptTimeZone(), 'MMM-yy');
  }
  const str = String(inputVal).trim();
  if (!str) return 'NO EXPIRY';

  // Format: MMM-yy (e.g. May-30, Feb-27, Dec-28)
  const match = str.match(/^([A-Za-z]{3})-(\d{2})$/);
  if (match) {
    const m = match[1].charAt(0).toUpperCase() + match[1].slice(1).toLowerCase();
    return `${m}-${match[2]}`;
  }

  // Format: MMM-yyyy (e.g. May-2030)
  const match4 = str.match(/^([A-Za-z]{3})-(\d{4})$/);
  if (match4) {
    const m = match4[1].charAt(0).toUpperCase() + match4[1].slice(1).toLowerCase();
    return `${m}-${match4[2].slice(-2)}`;
  }

  return str;
}

/**
 * Records Stock-Out (Deduction of Quantity from specific Brand & Expiry batch)
 * Reduces QTY STORAGE and logs transaction with financial details into STOCK_OUT_LOG sheet
 */
function recordStockOut(payload) {
  try {
    const { rowIndex, itemCode, qtyOut, dateOut, remarks, user } = payload;
    const deductQty = parseFloat(qtyOut);

    if (isNaN(deductQty) || deductQty <= 0) {
      throw new Error("Quantity out must be a positive number.");
    }

    const info = getInventorySheetInfo();
    const sheet = info.sheet;

    // Validate row matches item code
    const targetCode = String(sheet.getRange(rowIndex, info.colCode + 1).getValue()).trim();
    if (targetCode.toUpperCase() !== String(itemCode).trim().toUpperCase()) {
      throw new Error(`Row mismatch! Expected ${itemCode} but found ${targetCode} at row ${rowIndex}. Please refresh list.`);
    }

    const desc = String(sheet.getRange(rowIndex, info.colDesc + 1).getValue()).trim();
    const brandName = info.colBrand !== -1 ? String(sheet.getRange(rowIndex, info.colBrand + 1).getValue()).trim() : '';
    const rawExp = sheet.getRange(rowIndex, info.colExp + 1).getValue();
    const ssTz = sheet.getParent().getSpreadsheetTimeZone();
    const expDate = normalizeExpiryString(rawExp, ssTz);
    
    // Read batch unit price (repository only, not for computation)
    const rawPrice = info.colPrice !== -1 ? sheet.getRange(rowIndex, info.colPrice + 1).getValue() : 0;
    const price = typeof rawPrice === 'number' ? rawPrice : (parseFloat(String(rawPrice).replace(/[^0-9.-]/g, '')) || 0);

    const rawCurrentQty = sheet.getRange(rowIndex, info.colQty + 1).getValue();
    const currentQty = typeof rawCurrentQty === 'number' ? rawCurrentQty : (parseFloat(String(rawCurrentQty).replace(/[^0-9.-]/g, '')) || 0);

    if (deductQty > currentQty) {
      throw new Error(`Cannot deduct ${deductQty.toLocaleString()} units! Current storage only has ${currentQty.toLocaleString()} units.`);
    }

    const newQty = currentQty - deductQty;
    const cleanDateOut = String(dateOut || Utilities.formatDate(new Date(), ssTz, 'yyyy-MM-dd')).trim();

    // Update inventory sheet quantity
    sheet.getRange(rowIndex, info.colQty + 1)
      .setValue(newQty)
      .setNumberFormat('#,##0')
      .setHorizontalAlignment('right');

    // Update DATE OUT in inventory sheet
    if (info.colDateOut !== -1) {
      sheet.getRange(rowIndex, info.colDateOut + 1)
        .setNumberFormat('@')
        .setValue(cleanDateOut)
        .setHorizontalAlignment('center');
    }

    // Log the transaction
    logStockOutTransaction({
      timestamp: new Date(),
      dateOut: cleanDateOut,
      itemCode: targetCode,
      itemDesc: desc,
      brandName: brandName,
      expDate: expDate,
      unitPrice: price,
      qtyOut: deductQty,
      previousQty: currentQty,
      remainingBalance: newQty,
      remarks: remarks || 'Normal Dispense',
      user: user || Session.getActiveUser().getEmail() || 'User'
    });

    return {
      success: true,
      itemCode: targetCode,
      itemDesc: desc,
      brandName: brandName,
      expDate: expDate,
      unitPrice: price,
      qtyDeducted: deductQty,
      dateOut: cleanDateOut,
      previousQty: currentQty,
      newQty: newQty,
      message: `Deducted ${deductQty.toLocaleString()} units of ${desc} [${brandName || 'Batch'}]. Date Out: ${cleanDateOut}. Remaining: ${newQty.toLocaleString()}`
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Updates an existing row (useful for restocking 0-qty lines with new brand, expiry, price, and qty)
 */
function updateMedicineBatch(payload) {
  try {
    const { rowIndex, itemCode, itemDesc, brandName, newQty, newExpDate, acquisitionPrice } = payload;
    const info = getInventorySheetInfo();
    const sheet = info.sheet;

    const targetCode = String(sheet.getRange(rowIndex, info.colCode + 1).getValue()).trim();
    if (targetCode.toUpperCase() !== String(itemCode).trim().toUpperCase()) {
      throw new Error(`Row mismatch! Row ${rowIndex} does not match ${itemCode}.`);
    }

    const cleanQty = parseFloat(String(newQty).replace(/[^0-9.-]/g, '')) || 0;
    const cleanExp = String(newExpDate || '').trim();
    const cleanBrand = String(brandName || '').trim();
    const cleanPrice = parseFloat(String(acquisitionPrice || '0').replace(/[^0-9.-]/g, '')) || 0;

    const ssTz = sheet.getParent().getSpreadsheetTimeZone();
    const formattedExp = normalizeExpiryString(cleanExp, ssTz);

    if (itemDesc) {
      sheet.getRange(rowIndex, info.colDesc + 1)
        .setValue(String(itemDesc).trim())
        .setHorizontalAlignment('left');
    }

    if (info.colBrand !== -1 && cleanBrand) {
      sheet.getRange(rowIndex, info.colBrand + 1)
        .setValue(cleanBrand)
        .setHorizontalAlignment('left');
    }

    // Set Quantity (Right-aligned, number format #,##0)
    sheet.getRange(rowIndex, info.colQty + 1)
      .setValue(cleanQty)
      .setNumberFormat('#,##0')
      .setHorizontalAlignment('right');
    
    // Set Expiry Date (Plain text '@' centered)
    sheet.getRange(rowIndex, info.colExp + 1)
      .setNumberFormat('@')
      .setValue(formattedExp)
      .setHorizontalAlignment('center');

    // Set Acquisition Price (Right-aligned, currency format ₱#,##0.00)
    if (info.colPrice !== -1) {
      sheet.getRange(rowIndex, info.colPrice + 1)
        .setValue(cleanPrice)
        .setNumberFormat('"₱"#,##0.00')
        .setHorizontalAlignment('right');
    }

    return {
      success: true,
      message: `Updated [${itemCode}] at row ${rowIndex}: Brand: ${cleanBrand || 'N/A'}, Qty: ${cleanQty.toLocaleString()}, Expiry: ${formattedExp}, Price: ₱${cleanPrice.toFixed(2)}`
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Initializes and formats the STOCK_OUT_LOG sheet with Brand and Date Out columns (No computed total cost)
 */
function initLogSheet(ss) {
  const logSheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
  const headers = [
    'TIMESTAMP',
    'DATE OUT',
    'ITEM CODE',
    'GENERIC NAME',
    'BRAND NAME',
    'EXPIRY DATE',
    'UNIT PRICE (₱)',
    'QTY OUT',
    'PREVIOUS QTY',
    'REMAINING BALANCE',
    'REMARKS',
    'DISPENSED BY'
  ];
  
  logSheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1e293b')
    .setFontColor('#ffffff');

  logSheet.setFrozenRows(1);
  logSheet.setColumnWidth(1, 160); // Timestamp
  logSheet.setColumnWidth(2, 110); // Date Out
  logSheet.setColumnWidth(3, 130); // Code
  logSheet.setColumnWidth(4, 260); // Generic
  logSheet.setColumnWidth(5, 150); // Brand
  logSheet.setColumnWidth(6, 110); // Expiry
  logSheet.setColumnWidth(7, 120); // Unit Price (Repository)
  logSheet.setColumnWidth(8, 90);  // Qty Out
  logSheet.setColumnWidth(9, 100); // Prev
  logSheet.setColumnWidth(10, 120); // Rem
  logSheet.setColumnWidth(11, 180); // Remarks
  logSheet.setColumnWidth(12, 150); // User

  return logSheet;
}

/**
 * Appends transaction record to STOCK_OUT_LOG with Date Out and Brand Name
 */
function logStockOutTransaction(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = initLogSheet(ss);
    }

    const row = [
      Utilities.formatDate(data.timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      data.dateOut || Utilities.formatDate(data.timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      data.itemCode,
      data.itemDesc,
      data.brandName || '',
      data.expDate,
      data.unitPrice || 0,
      data.qtyOut,
      data.previousQty,
      data.remainingBalance,
      data.remarks,
      data.user
    ];

    logSheet.appendRow(row);
    const lastRow = logSheet.getLastRow();
    
    // Formatting:
    // Col 2: Date Out
    logSheet.getRange(lastRow, 2).setNumberFormat('@').setHorizontalAlignment('center');
    // Col 7: Unit Price
    logSheet.getRange(lastRow, 7).setNumberFormat('"₱"#,##0.00');
    // Col 8: Qty Out
    logSheet.getRange(lastRow, 8).setNumberFormat('#,##0');
    // Col 9-10: Prev & Remaining
    logSheet.getRange(lastRow, 9, 1, 2).setNumberFormat('#,##0');
  } catch (err) {
    Logger.log("Error logging stock-out: " + err.message);
  }
}

/**
 * Fetches recent audit logs from STOCK_OUT_LOG
 */
function getRecentLogs(limit) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
    if (!logSheet || logSheet.getLastRow() <= 1) {
      return { success: true, logs: [] };
    }

    const max = limit || 50;
    const values = logSheet.getDataRange().getValues();
    const logs = [];

    const headerRow = values[0].map(h => String(h).toUpperCase());
    const hasDateOut = headerRow.includes('DATE OUT') || headerRow.includes('DATE_OUT');

    // Read in reverse chronological order
    for (let i = values.length - 1; i >= 1 && logs.length < max; i--) {
      const r = values[i];
      if (hasDateOut) {
        logs.push({
          timestamp: r[0],
          dateOut: r[1] ? (r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[1])) : '',
          itemCode: r[2],
          itemDesc: r[3],
          brandName: r[4] || '',
          expDate: r[5],
          unitPrice: r[6] || 0,
          qtyOut: r[7],
          previousQty: r[8],
          remainingBalance: r[9],
          remarks: r[10],
          user: r[11]
        });
      } else {
        // Fallback for older log schema
        logs.push({
          timestamp: r[0],
          dateOut: '',
          itemCode: r[1],
          itemDesc: r[2],
          brandName: r[3] || '',
          expDate: r[4],
          unitPrice: r[5] || 0,
          qtyOut: r[6],
          previousQty: r[8] || r[7],
          remainingBalance: r[9] || r[8],
          remarks: r[10] || r[9],
          user: r[11] || r[10]
        });
      }
    }

    return { success: true, logs: logs };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
