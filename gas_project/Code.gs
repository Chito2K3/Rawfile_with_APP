/**
 * ============================================================================
 * PHARMACY STORAGE INVENTORY & EXPIRY MANAGEMENT SYSTEM
 * Google Apps Script Backend (Code.gs)
 * Amang Rodriguez Memorial Medical Center (ARMMC)
 * ============================================================================
 */

// Target or active sheet configuration
// Target or active sheet configuration
const CONFIG = {
  MAIN_SHEET_NAME: 'Stocks', // Primary inventory sheet name (supports 'Stocks', 'STORAGE', etc.)
  STORAGE_SHEET_ALIASES: ['Stocks', 'STOCKS', 'stocks', 'Stock', 'STORAGE', 'Storage', 'Inventory', 'INVENTORY'],
  MASTER_DCI_SHEET_NAME: 'Medicine', // Master catalog tab with Col A: DCI Code, Col B: Generic Description (supports 'Medicine', 'Sheet1', etc.)
  LOG_SHEET_NAME: 'STOCK_OUT_LOG',
  COLUMN_KEYS: {
    CODE: 'ITEM CODE',
    DESC: 'ITEM DESCRIPTION',
    BRAND: 'BRAND NAME',
    UOM: 'UOM',
    QTY: 'QTY STORAGE',
    EXP: 'EXP DATE',
    PRICE: 'ACQUISITION PRICE',
    DATE_OUT: 'DATE OUT'
  },
  STANDARD_UOM_LIST: [
    'Tablet',
    'Capsule',
    'Vial',
    'Ampule',
    'Bottle',
    'Tube',
    'Nebule',
    'Syringe',
    'Sachet'
  ]
};

/**
 * Automatically infers standard hospital UOM from medicine description
 */
function inferUomFromDescription(desc) {
  if (!desc) return 'Tablet';
  const u = String(desc).toUpperCase();

  if (u.includes('CAPSULE') || u.includes(' CAP ') || u.endsWith(' CAP')) return 'Capsule';
  if (u.includes('VIAL')) return 'Vial';
  if (u.includes('AMPULE') || u.includes('AMP ') || u.endsWith(' AMP')) return 'Ampule';
  if (u.includes('NEBULE') || u.includes('RESPIRATORY SOLUTION')) return 'Nebule';
  if (u.includes('SYRINGE') || u.includes('PRE-FILLED') || u.includes('PREFILLED')) return 'Syringe';
  if (u.includes('SACHET') || u.includes('POWDER FOR SUSPENSION')) return 'Sachet';
  if (u.includes('TUBE') || u.includes('OINTMENT') || u.includes('CREAM') || u.includes('GEL')) return 'Tube';
  if (u.includes('BOTTLE') || u.includes('SYRUP') || u.includes('SUSPENSION') || u.includes('ELIXIR') || u.includes('DROPS') || u.includes('OPHTHALMIC') || u.includes('SOLUTION') || u.includes('SHAMPOO')) return 'Bottle';
  if (u.includes('TABLET') || u.includes('TAB ') || u.endsWith(' TAB') || u.includes('CAPLET')) return 'Tablet';

  return 'Tablet';
}

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
    .addItem('📑 Register Item in Master Formulary', 'showAddCatalogDialog')
    .addItem('🔄 Restock 0-Qty Items', 'showRestockDialog')
    .addSeparator()
    .addItem('📋 View Stock-Out Audit Log', 'openLogSheet')
    .addItem('🛠️ Fix & Upgrade Stock-Out Log Sheet (12 Cols)', 'repairAndMigrateLogSheet')
    .addItem('🛠️ Ensure / Upgrade Sheet Columns (Option A + UOM)', 'migrateSheetToOptionA')
    .addItem('🧹 Clean Blank Rows & Consolidate Sheet', 'cleanAndConsolidateSheet')
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
function showAddCatalogDialog() {
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
 * Option A Standard with UOM:
 * ITEM CODE | ITEM DESCRIPTION | BRAND NAME | UOM | QTY STORAGE | EXP DATE | ACQUISITION PRICE | DATE OUT
 */
function getInventorySheetInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = null;

  // 1. Try CONFIG.MAIN_SHEET_NAME if defined
  if (CONFIG.MAIN_SHEET_NAME) {
    sheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  }
  
  // 2. Try known aliases (Stocks, STORAGE, Stock, Inventory, etc.)
  if (!sheet) {
    const aliases = CONFIG.STORAGE_SHEET_ALIASES || ['Stocks', 'STOCKS', 'stocks', 'Stock', 'STORAGE', 'Storage', 'Inventory', 'INVENTORY'];
    for (let alias of aliases) {
      sheet = ss.getSheetByName(alias);
      if (sheet) break;
    }
  }

  // 3. Fallback: Scan sheets, skipping logs, catalogs, warehouses, and backups
  if (!sheet) {
    const sheets = ss.getSheets();
    for (let s of sheets) {
      const name = s.getName();
      const lower = name.trim().toLowerCase();
      if (name === CONFIG.LOG_SHEET_NAME || lower.includes('log')) continue;
      if (CONFIG.MASTER_DCI_SHEET_NAME && name === CONFIG.MASTER_DCI_SHEET_NAME) continue;
      if (lower === 'medicine' || lower === 'medicines' || lower === 'sheet1' || lower === 'sheet 1' || lower.includes('catalog') || lower.includes('dci') || lower.includes('formulary')) continue;
      if (lower.includes('backup') || lower.includes('copy') || lower.includes('old') || lower.includes('archive') || lower.includes('warehouse') || lower.includes('temp')) continue;
      
      const data = s.getDataRange().getValues();
      if (data.length > 0) {
        const headerRow = data[0].map(c => String(c).trim().toUpperCase());
        if (headerRow.some(h => h.includes('ITEM CODE') || h.includes('ITEM_CODE') || h.includes('CODE'))) {
          sheet = s;
          break;
        }
      }
    }
  }

  if (!sheet) {
    sheet = ss.getSheets()[0];
  }

  const dataRange = sheet.getDataRange();
  let values = dataRange.getValues();
  
  // If sheet is completely blank or Row 1 is empty, auto-initialize standard headers
  const isSheetEmpty = values.length === 0 || (values.length === 1 && values[0].every(c => !String(c).trim()));
  if (isSheetEmpty) {
    const standardHeaders = [
      'ITEM CODE', 'ITEM DESCRIPTION', 'BRAND NAME', 'UOM',
      'QTY STORAGE', 'EXP DATE', 'ACQUISITION PRICE', 'DATE OUT', 'BATCH ID'
    ];
    sheet.getRange(1, 1, 1, 9).setValues([standardHeaders])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#f3f4f6');
    sheet.setFrozenRows(1);
    values = sheet.getDataRange().getValues();
  }

  // Find column indices
  let headerRow = values[0];
  const hasCodeHeader = headerRow.some(h => String(h).toUpperCase().includes('CODE'));
  if (!hasCodeHeader && headerRow.every(h => !String(h).trim())) {
    const standardHeaders = [
      'ITEM CODE', 'ITEM DESCRIPTION', 'BRAND NAME', 'UOM',
      'QTY STORAGE', 'EXP DATE', 'ACQUISITION PRICE', 'DATE OUT', 'BATCH ID'
    ];
    sheet.getRange(1, 1, 1, 9).setValues([standardHeaders])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#f3f4f6');
    sheet.setFrozenRows(1);
    values = sheet.getDataRange().getValues();
    headerRow = values[0];
  }
  let colCode = -1, colDesc = -1, colBrand = -1, colUom = -1, colQty = -1, colExp = -1, colPrice = -1, colDateOut = -1, colBatchId = -1;

  for (let j = 0; j < headerRow.length; j++) {
    const h = String(headerRow[j]).trim().toUpperCase();
    if (colCode === -1 && (h.includes('CODE') || h.includes('ITEM CODE'))) colCode = j;
    if (colDesc === -1 && (h.includes('DESC') || h.includes('DESCRIPTION') || h.includes('GENERIC'))) colDesc = j;
    if (colBrand === -1 && (h.includes('BRAND') || h.includes('TRADE'))) colBrand = j;
    if (colUom === -1 && (h === 'UOM' || h.includes('UNIT OF MEASURE') || h === 'UNIT' || h.includes('MEASURE'))) colUom = j;
    if (colQty === -1 && (h.includes('QTY') || h.includes('STORAGE') || h.includes('QUANTITY') || h.includes('STOCK') || h.includes('COUNT') || h.includes('ACTUAL'))) colQty = j;
    if (colExp === -1 && (h.includes('EXP') || h.includes('DATE') || h.includes('EXPIRY')) && !h.includes('OUT')) colExp = j;
    if (colPrice === -1 && (h.includes('PRICE') || h.includes('COST') || h.includes('ACQUISITION'))) colPrice = j;
    if (colDateOut === -1 && (h.includes('DATE OUT') || h.includes('DATE_OUT') || h.includes('OUT DATE') || h.includes('DISPENSED DATE'))) colDateOut = j;
    if (colBatchId === -1 && (h.includes('BATCH ID') || h.includes('BATCH_ID') || h === 'UID' || h.includes('UNIQUE ID'))) colBatchId = j;
  }

  // Fallbacks if columns are not explicitly labeled
  if (colCode === -1) colCode = 0;
  if (colDesc === -1) colDesc = 1;

  // If Option A layout with UOM is present:
  if (colBrand !== -1 && colUom !== -1) {
    if (colQty === -1) colQty = 4;
    if (colExp === -1) colExp = 5;
    if (colPrice === -1) colPrice = 6;
    if (colDateOut === -1 && headerRow.length >= 8) colDateOut = 7;
    if (colBatchId === -1 && headerRow.length >= 9) colBatchId = 8;
  } else if (colBrand !== -1) {
    // Option A without UOM fallback
    if (colQty === -1) colQty = 3;
    if (colExp === -1) colExp = 4;
    if (colPrice === -1) colPrice = 5;
    if (colDateOut === -1 && headerRow.length >= 7) colDateOut = 6;
    if (colBatchId === -1 && headerRow.length >= 8) colBatchId = 7;
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
    colUom: colUom,
    colQty: colQty,
    colExp: colExp,
    colPrice: colPrice,
    colDateOut: colDateOut,
    colBatchId: colBatchId
  };
}

/**
 * AppSheet-style Unique ID generator (UNIQUEID())
 * Produces an 8-character unique alphanumeric key (e.g., BID-7K9F2A41)
 */
function generateUniqueBatchId() {
  const chars = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ';
  let uid = 'BID-';
  for (let i = 0; i < 8; i++) {
    uid += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return uid;
}

/**
 * Migration Helper: Upgrades the current inventory sheet to Option A layout with UOM:
 * Column A: ITEM CODE
 * Column B: ITEM DESCRIPTION
 * Column C: BRAND NAME
 * Column D: UOM (Dropdown list: Tablet, Capsule, Vial, etc.)
 * Column E: QTY STORAGE
 * Column F: EXP DATE
 * Column G: ACQUISITION PRICE (Repository only)
 * Column H: DATE OUT (Last Dispensed Date)
 */
function migrateSheetToOptionA() {
  const ui = SpreadsheetApp.getUi();
  try {
    const info = getInventorySheetInfo();
    const sheet = info.sheet;

    // Ensure sheet has at least 8 columns
    while (sheet.getMaxColumns() < 8) {
      sheet.insertColumnAfter(sheet.getMaxColumns());
    }

    const checkCols = Math.max(sheet.getLastColumn(), 14);
    const headerRow = sheet.getRange(1, 1, 1, checkCols).getValues()[0];
    const upperHeaders = headerRow.map(h => String(h).trim().toUpperCase());

    // 1. Ensure BRAND NAME is at Column C (index 3)
    const brandColIdx = upperHeaders.findIndex(h => h.includes('BRAND'));
    if (brandColIdx === -1) {
      sheet.insertColumnAfter(2);
      sheet.getRange(1, 3).setValue('BRAND NAME').setFontWeight('bold');
    } else if (brandColIdx !== 2) {
      sheet.getRange(1, 3).setValue('BRAND NAME').setFontWeight('bold');
    }

    // Refresh headers after possible brand insertion
    const curHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 14)).getValues()[0];
    const curUpper = curHeaders.map(h => String(h).trim().toUpperCase());

    // 2. Ensure UOM is at Column D (index 4)
    const uomColIdx = curUpper.findIndex(h => h === 'UOM' || h.includes('UNIT OF MEASURE') || h === 'UNIT');
    if (uomColIdx === -1) {
      sheet.insertColumnAfter(3);
      sheet.getRange(1, 4).setValue('UOM').setFontWeight('bold').setHorizontalAlignment('center');
    } else if (uomColIdx !== 3) {
      sheet.getRange(1, 4).setValue('UOM').setFontWeight('bold').setHorizontalAlignment('center');
    }

    // Apply native Google Sheets Data Validation dropdown (Tablet, Capsule, Vial, etc.) to Column D
    const maxRows = sheet.getMaxRows();
    if (maxRows > 1) {
      const uomRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(CONFIG.STANDARD_UOM_LIST, true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, 4, maxRows - 1, 1).setDataValidation(uomRule).setHorizontalAlignment('center');

      // Pre-fill empty UOMs for existing rows by auto-detecting from Description
      const dataVals = sheet.getRange(2, 1, maxRows - 1, 4).getValues();
      for (let r = 0; r < dataVals.length; r++) {
        const itemCode = String(dataVals[r][0] || '').trim();
        if (!itemCode) continue;
        const currentUom = String(dataVals[r][3] || '').trim();
        if (!currentUom) {
          const desc = String(dataVals[r][1] || '').trim();
          const detected = inferUomFromDescription(desc);
          if (detected) {
            sheet.getRange(r + 2, 4).setValue(detected);
          }
        }
      }
    }

    // 3. Clear any stray ACQUISITION PRICE headers that may have landed in Column I, J, K
    for (let c = 9; c <= Math.max(sheet.getLastColumn(), 14); c++) {
      const val = String(sheet.getRange(1, c).getValue()).trim().toUpperCase();
      if (val.includes('PRICE') || val.includes('ACQUISITION')) {
        sheet.getRange(1, c).clearContent();
      }
    }

    // 4. Ensure Column E is QTY STORAGE
    sheet.getRange(1, 5).setValue('QTY STORAGE').setFontWeight('bold').setHorizontalAlignment('right');
    if (maxRows > 1) {
      sheet.getRange(2, 5, maxRows - 1, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
    }

    // 5. Ensure Column F is EXP DATE with mmm-yy date format and upgrade existing text dates to true Date objects
    sheet.getRange(1, 6).setValue('EXP DATE').setFontWeight('bold').setHorizontalAlignment('center');
    if (maxRows > 1) {
      sheet.getRange(2, 6, maxRows - 1, 1).setNumberFormat('mmm-yy').setHorizontalAlignment('center');
      const ssTz = sheet.getParent().getSpreadsheetTimeZone();
      const expRange = sheet.getRange(2, 6, maxRows - 1, 1);
      const expValues = expRange.getValues();
      for (let r = 0; r < expValues.length; r++) {
        const raw = expValues[r][0];
        if (raw && !(raw instanceof Date)) {
          const parsed = getExpiryDateObject(raw, ssTz);
          if (parsed) {
            sheet.getRange(r + 2, 6).setValue(parsed);
          }
        }
      }
    }

    // 6. Explicitly place ACQUISITION PRICE in Column G (Column 7)
    sheet.getRange(1, 7).setValue('ACQUISITION PRICE').setFontWeight('bold').setHorizontalAlignment('right');
    if (maxRows > 1) {
      sheet.getRange(2, 7, maxRows - 1, 1).setNumberFormat('"₱"#,##0.00');
    }

    // 7. Explicitly place DATE OUT in Column H (Column 8)
    sheet.getRange(1, 8).setValue('DATE OUT').setFontWeight('bold').setHorizontalAlignment('center');
    if (maxRows > 1) {
      sheet.getRange(2, 8, maxRows - 1, 1).setNumberFormat('@').setHorizontalAlignment('center');
    }

    ui.alert(
      'Option A Setup Complete',
      'The sheet columns have been confirmed:\n\n' +
      '• Column A: ITEM CODE\n' +
      '• Column B: ITEM DESCRIPTION\n' +
      '• Column C: BRAND NAME\n' +
      '• Column D: UOM (Dropdown list: Tablet, Capsule, Vial, etc.)\n' +
      '• Column E: QTY STORAGE\n' +
      '• Column F: EXP DATE\n' +
      '• Column G: ACQUISITION PRICE (Repository only)\n' +
      '• Column H: DATE OUT (Last Dispensed Date)\n\n' +
      'Unit of Measure (UOM) dropdown is now active in Google Sheets!',
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
      const uom = info.colUom !== -1 ? String(row[info.colUom] || '').trim() : '';
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

      let batchId = '';
      if (info.colBatchId !== -1 && row[info.colBatchId]) {
        batchId = String(row[info.colBatchId]).trim();
      }
      if (!batchId) {
        batchId = `BID-${code.replace(/[^A-Za-z0-9]/g, '').slice(0, 8)}-${(brand || 'GEN').replace(/[^A-Za-z0-9]/g, '').slice(0, 3)}-${i + 1}`;
      }

      rows.push({
        batchId: batchId,
        rowIndex: i + 1, // 1-based row index in Google Sheet
        itemCode: code,
        itemDesc: desc,
        brandName: brand,
        uom: uom || inferUomFromDescription(desc),
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
      hasUomColumn: info.colUom !== -1,
      hasPriceColumn: info.colPrice !== -1,
      hasDateOutColumn: info.colDateOut !== -1,
      items: rows,
      catalog: masterCatalogRes.catalog || [],
      uomList: CONFIG.STANDARD_UOM_LIST
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Fetches the master list of DCI Codes and Generic Descriptions from 'Medicine' (or 'Sheet1')
 * Column A: DCI Code (Item Code)
 * Column B: Generic Description & Strength
 */
function getMasterDciCatalog() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = null;

    if (CONFIG.MASTER_DCI_SHEET_NAME) {
      sheet = ss.getSheetByName(CONFIG.MASTER_DCI_SHEET_NAME);
    }
    if (!sheet) {
      sheet = ss.getSheetByName('Medicine') || ss.getSheetByName('medicine') || ss.getSheetByName('Sheet1') || ss.getSheetByName('Sheet 1');
    }
    if (!sheet) {
      const sheets = ss.getSheets();
      for (let s of sheets) {
        const sName = s.getName().trim().toLowerCase();
        if (sName === 'medicine' || sName === 'medicines' || sName === 'sheet1' || sName === 'sheet 1' || sName.includes('catalog') || sName.includes('dci') || sName.includes('formulary')) {
          sheet = s;
          break;
        }
      }
    }

    if (!sheet) {
      return { success: false, catalog: [], message: "Master DCI tab ('Medicine' or 'Sheet1') not found." };
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
 * Adds a new medicine (DCI Code and Generic Description) to the 'Medicine' master catalog tab.
 * Column A: DCI Code (Item Code)
 * Column B: Generic Description & Strength
 * Prevents duplicates by case-insensitive DCI Code check.
 */
function addMedicineToMasterCatalog(payload) {
  try {
    const { itemCode, itemDesc } = payload || {};
    if (!itemCode || !String(itemCode).trim()) {
      throw new Error("Item Code / DCI Code is required.");
    }
    if (!itemDesc || !String(itemDesc).trim()) {
      throw new Error("Generic Description & Strength is required.");
    }

    const cleanCode = String(itemCode).trim().toUpperCase();
    const cleanDesc = String(itemDesc).trim().toUpperCase();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = null;
    if (CONFIG.MASTER_DCI_SHEET_NAME) {
      sheet = ss.getSheetByName(CONFIG.MASTER_DCI_SHEET_NAME);
    }
    if (!sheet) {
      sheet = ss.getSheetByName('Medicine') || ss.getSheetByName('medicine') || ss.getSheetByName('Sheet1') || ss.getSheetByName('Sheet 1');
    }
    if (!sheet) {
      const sheets = ss.getSheets();
      for (let s of sheets) {
        const sName = s.getName().trim().toLowerCase();
        if (sName === 'medicine' || sName === 'medicines' || sName === 'sheet1' || sName === 'sheet 1' || sName.includes('catalog') || sName.includes('dci') || sName.includes('formulary')) {
          sheet = s;
          break;
        }
      }
    }

    if (!sheet) {
      // If no catalog sheet exists at all, create 'Medicine' tab
      sheet = ss.insertSheet('Medicine');
      sheet.getRange(1, 1, 1, 2).setValues([['DCI CODE', 'GENERIC DESCRIPTION & STRENGTH']])
        .setFontWeight('bold')
        .setBackground('#f3f4f6');
      sheet.setFrozenRows(1);
    }

    const data = sheet.getDataRange().getValues();
    let codeColIdx = 0;
    let descColIdx = 1;
    let startRow = 0;

    if (data.length > 0) {
      const h0 = String(data[0][0] || '').trim().toUpperCase();
      const h1 = String(data[0][1] || '').trim().toUpperCase();
      if (h0.includes('CODE') || h1.includes('DESC')) {
        startRow = 1;
      }

      // Check for duplicate code
      for (let i = startRow; i < data.length; i++) {
        const existingCode = String(data[i][codeColIdx] || '').trim().toUpperCase();
        if (existingCode === cleanCode) {
          const existingDesc = String(data[i][descColIdx] || '').trim();
          return {
            success: true,
            alreadyExisted: true,
            code: cleanCode,
            desc: existingDesc || cleanDesc,
            message: `[${cleanCode}] already exists in Master Catalog: ${existingDesc || cleanDesc}`
          };
        }
      }
    }

    // Append to the last row of the catalog sheet
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1).setValue(cleanCode).setHorizontalAlignment('left');
    sheet.getRange(nextRow, 2).setValue(cleanDesc).setHorizontalAlignment('left');

    return {
      success: true,
      alreadyExisted: false,
      code: cleanCode,
      desc: cleanDesc,
      message: `Successfully registered [${cleanCode}] ${cleanDesc} into Medicine Master Catalog!`
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Smart Add Medicine or New Batch
 * If itemCode exists, finds the last occurrence and inserts a row directly below it!
 * If itemCode does not exist, appends to the end of the inventory list.
 * Supports Brand Name and Acquisition Price per batch.
 */
function addNewMedicineOrBatch(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Concurrency guard for multi-mobile simultaneous encoders
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
    const cleanUom = String(payload.uom || inferUomFromDescription(cleanDesc)).trim();
    const cleanQty = parseFloat(String(qty).replace(/[^0-9.-]/g, '')) || 0;
    const cleanExp = String(expDate || '').trim();
    const cleanPrice = parseFloat(String(acquisitionPrice || '0').replace(/[^0-9.-]/g, '')) || 0;

    // Auto-enroll into Medicine Master Catalog if not already registered
    try {
      addMedicineToMasterCatalog({ itemCode: cleanCode, itemDesc: cleanDesc });
    } catch (catErr) {
      console.warn("Could not auto-enroll into master catalog: " + catErr.message);
    }

    const info = getInventorySheetInfo();
    const sheet = info.sheet;
    const values = info.values;

    // Search for existing item occurrences and their expiry scores for FEFO ordering
    const existingBatches = [];
    const ssTz = sheet.getParent().getSpreadsheetTimeZone();
    const formattedExp = normalizeExpiryString(cleanExp, ssTz);
    const newExpScore = getExpiryComparableScore(cleanExp, ssTz);

    for (let i = 1; i < values.length; i++) {
      const code = String(values[i][info.colCode] || '').trim().toUpperCase();
      if (code === cleanCode.toUpperCase()) {
        const rawExp = values[i][info.colExp];
        const score = getExpiryComparableScore(rawExp, ssTz);
        existingBatches.push({
          rowIndex: i + 1, // 1-based index in sheet
          expScore: score,
          brand: info.colBrand !== -1 ? String(values[i][info.colBrand] || '').trim() : '',
          uom: info.colUom !== -1 ? String(values[i][info.colUom] || '').trim() : '',
          expStr: normalizeExpiryString(rawExp, ssTz)
        });
      }
    }

    let insertedRowIndex = -1;
    let mode = '';

    if (existingBatches.length > 0) {
      // FEFO (First-Expired, First-Out): Find first existing batch that expires LATER than new incoming batch
      const laterBatch = existingBatches.find(b => b.expScore > newExpScore);

      if (laterBatch) {
        // Insert directly ABOVE the later expiring batch so nearest expiry appears first
        sheet.insertRowBefore(laterBatch.rowIndex);
        insertedRowIndex = laterBatch.rowIndex;
        mode = `FEFO Priority: Placed at Row ${insertedRowIndex} above ${laterBatch.brand || 'existing batch'} [${laterBatch.expStr}]`;
      } else {
        // If new batch expires later than or equal to all existing batches, insert below the last batch
        const lastBatch = existingBatches[existingBatches.length - 1];
        sheet.insertRowAfter(lastBatch.rowIndex);
        insertedRowIndex = lastBatch.rowIndex + 1;
        mode = `FEFO Placed: Inserted at Row ${insertedRowIndex} below existing batches`;
      }
    } else {
      // New item: Find first available empty data row (starts at Row 2, never skips down)
      insertedRowIndex = findNextAvailableRow(sheet, info);
      mode = `Added as brand new medicine at Row ${insertedRowIndex}`;
    }

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

    // 4. Unit of Measure (UOM) (Center)
    if (info.colUom !== -1) {
      const uomCell = sheet.getRange(insertedRowIndex, info.colUom + 1);
      uomCell.setValue(cleanUom).setHorizontalAlignment('center');
      const uomRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(CONFIG.STANDARD_UOM_LIST, true)
        .setAllowInvalid(true)
        .build();
      uomCell.setDataValidation(uomRule);
    }

    // 5. Storage Quantity (Right aligned, formatted as #,##0)
    sheet.getRange(insertedRowIndex, info.colQty + 1)
      .setValue(cleanQty)
      .setNumberFormat('#,##0')
      .setHorizontalAlignment('right');
    
    // 6. Expiry Date (Proper Date object with mmm-yy number format for true Sheet date representation)
    const expDateObj = getExpiryDateObject(cleanExp, ssTz);
    if (expDateObj) {
      sheet.getRange(insertedRowIndex, info.colExp + 1)
        .setValue(expDateObj)
        .setNumberFormat('mmm-yy')
        .setHorizontalAlignment('center');
    } else {
      sheet.getRange(insertedRowIndex, info.colExp + 1)
        .setNumberFormat('@')
        .setValue(formattedExp)
        .setHorizontalAlignment('center');
    }

    // 7. Acquisition Price (Right aligned, formatted as ₱#,##0.00)
    if (info.colPrice !== -1) {
      sheet.getRange(insertedRowIndex, info.colPrice + 1)
        .setValue(cleanPrice)
        .setNumberFormat('"₱"#,##0.00')
        .setHorizontalAlignment('right');
    }

    // 8. Date Out (Empty for newly added batch)
    if (info.colDateOut !== -1) {
      sheet.getRange(insertedRowIndex, info.colDateOut + 1)
        .setValue('')
        .setHorizontalAlignment('center');
    }

    // 9. Unique Batch Identifier (AppSheet UNIQUEID style)
    const newBatchId = generateUniqueBatchId();
    if (info.colBatchId !== -1) {
      sheet.getRange(insertedRowIndex, info.colBatchId + 1)
        .setValue(newBatchId)
        .setHorizontalAlignment('center');
    } else if (sheet.getMaxColumns() >= 9) {
      sheet.getRange(insertedRowIndex, 9)
        .setValue(newBatchId)
        .setHorizontalAlignment('center');
    }

    return {
      success: true,
      message: `Successfully saved: [${cleanCode}] ${cleanDesc} (${cleanBrand} - ${cleanUom})`,
      mode: mode,
      insertedRowIndex: insertedRowIndex,
      batchId: newBatchId
    };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Converts any expiry value (Date object, 'YYYY-MM', 'YYYY-MM-DD', 'MMM-yy') into a JavaScript Date object.
 * By pharmaceutical convention, expiration is anchored to the last day of the given month.
 */
function getExpiryDateObject(val, tz) {
  if (!val) return null;
  if (val instanceof Date) return val;
  const s = String(val).trim();
  if (!s || s.toUpperCase() === 'NO EXPIRY' || s.toUpperCase() === 'N/A') return null;

  const monthMap = {
    jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
    jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12
  };

  // 1. Check YYYY-MM (e.g. '2028-12')
  const mIsoMonth = s.match(/^(\d{4})-(\d{1,2})$/);
  if (mIsoMonth) {
    const yr = parseInt(mIsoMonth[1], 10);
    const mo = parseInt(mIsoMonth[2], 10);
    const lastDay = new Date(yr, mo, 0).getDate();
    return new Date(yr, mo - 1, lastDay, 12, 0, 0); // midday avoids timezone boundary drift
  }

  // 2. Check YYYY-MM-DD
  const mIsoDay = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (mIsoDay) {
    const yr = parseInt(mIsoDay[1], 10);
    const mo = parseInt(mIsoDay[2], 10);
    const day = parseInt(mIsoDay[3], 10);
    return new Date(yr, mo - 1, day, 12, 0, 0);
  }

  // 3. Check MMM-yy (e.g. 'Feb-28', 'Dec-28')
  const mShort = s.match(/^([A-Za-z]{3})-(\d{2})$/);
  if (mShort) {
    const mo = monthMap[mShort[1].toLowerCase()];
    if (mo) {
      const yr = 2000 + parseInt(mShort[2], 10);
      const lastDay = new Date(yr, mo, 0).getDate();
      return new Date(yr, mo - 1, lastDay, 12, 0, 0);
    }
  }

  // 4. Check MMM-yyyy (e.g. 'Feb-2028')
  const mLong = s.match(/^([A-Za-z]{3})-(\d{4})$/);
  if (mLong) {
    const mo = monthMap[mLong[1].toLowerCase()];
    if (mo) {
      const yr = parseInt(mLong[2], 10);
      const lastDay = new Date(yr, mo, 0).getDate();
      return new Date(yr, mo - 1, lastDay, 12, 0, 0);
    }
  }

  // 5. Fallback date parse
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return d;
  }

  return null;
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
 * Converts expiry values (Date object or strings like 'Feb-28', 'Jan-30', '2028-02-01')
 * into a comparable integer score YYYYMM (e.g. Feb-28 -> 202802, Jan-30 -> 203001)
 * for FEFO (First-Expired, First-Out) sorting.
 */
function getExpiryComparableScore(val, tz) {
  if (!val) return 999999;
  if (val instanceof Date) {
    const y = parseInt(Utilities.formatDate(val, tz || Session.getScriptTimeZone(), 'yyyy'), 10);
    const m = parseInt(Utilities.formatDate(val, tz || Session.getScriptTimeZone(), 'MM'), 10);
    return y * 100 + m;
  }
  const s = String(val).trim();
  if (!s || s.toUpperCase() === 'NO EXPIRY' || s.toUpperCase() === 'N/A') return 999999;

  const monthMap = {
    jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
    jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12
  };

  // Format MMM-yy (e.g. Feb-28, Jan-30)
  const mShort = s.match(/^([A-Za-z]{3})-(\d{2})$/);
  if (mShort) {
    const mo = monthMap[mShort[1].toLowerCase()] || 0;
    const yr = 2000 + parseInt(mShort[2], 10);
    return yr * 100 + mo;
  }

  // Format MMM-yyyy (e.g. Feb-2028)
  const mLong = s.match(/^([A-Za-z]{3})-(\d{4})$/);
  if (mLong) {
    const mo = monthMap[mLong[1].toLowerCase()] || 0;
    const yr = parseInt(mLong[2], 10);
    return yr * 100 + mo;
  }

  // Format yyyy-MM or yyyy-MM-dd
  const mIso = s.match(/^(\d{4})-(\d{1,2})/);
  if (mIso) {
    const yr = parseInt(mIso[1], 10);
    const mo = parseInt(mIso[2], 10);
    return yr * 100 + mo;
  }

  // Standard date parsing fallback
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return d.getFullYear() * 100 + (d.getMonth() + 1);
  }

  // Fallback: match 4-digit year
  const mYr = s.match(/20(\d{2})/);
  if (mYr) {
    return (2000 + parseInt(mYr[1], 10)) * 100 + 12;
  }

  return 999999;
}

/**
 * Records Stock-Out (Deduction of Quantity from specific Brand & Expiry batch)
 * Reduces QTY STORAGE and logs transaction with financial details into STOCK_OUT_LOG sheet
 */
function recordStockOut(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // 30-second atomic lock to prevent race conditions among mobile users
  try {
    const { batchId, rowIndex, itemCode, qtyOut, dateOut, remarks } = payload;
    const deductQty = parseFloat(qtyOut);

    if (isNaN(deductQty) || deductQty <= 0) {
      throw new Error("Quantity out must be a positive number.");
    }

    const info = getInventorySheetInfo();
    const sheet = info.sheet;
    const values = sheet.getDataRange().getValues();

    // AppSheet-style UID Resolution: Locate row by BATCH_ID first, preventing index shift collisions
    let targetRowIndex = -1;

    if (batchId && info.colBatchId !== -1) {
      for (let r = 1; r < values.length; r++) {
        if (String(values[r][info.colBatchId] || '').trim() === String(batchId).trim()) {
          targetRowIndex = r + 1;
          break;
        }
      }
    }

    // Fallback: Verify rowIndex against itemCode
    if (targetRowIndex === -1 && rowIndex && rowIndex <= values.length) {
      const codeAtRow = String(values[rowIndex - 1][info.colCode] || '').trim();
      if (codeAtRow.toUpperCase() === String(itemCode).trim().toUpperCase()) {
        targetRowIndex = rowIndex;
      }
    }

    // Fallback: Scan sheet for itemCode match
    if (targetRowIndex === -1) {
      for (let r = 1; r < values.length; r++) {
        const c = String(values[r][info.colCode] || '').trim().toUpperCase();
        if (c === String(itemCode).trim().toUpperCase()) {
          targetRowIndex = r + 1;
          break;
        }
      }
    }

    if (targetRowIndex === -1) {
      throw new Error(`Item [${itemCode}] could not be found. Another mobile user may have modified the row. Please refresh the inventory.`);
    }

    const targetRowData = values[targetRowIndex - 1];
    const targetCode = String(targetRowData[info.colCode] || '').trim();
    const desc = String(targetRowData[info.colDesc] || '').trim();
    const brandName = info.colBrand !== -1 ? String(targetRowData[info.colBrand] || '').trim() : '';
    const uom = info.colUom !== -1 ? String(targetRowData[info.colUom] || '').trim() : '';
    const rawExp = targetRowData[info.colExp];
    const ssTz = sheet.getParent().getSpreadsheetTimeZone();
    const expDate = normalizeExpiryString(rawExp, ssTz);
    
    // Read batch unit price
    const rawPrice = info.colPrice !== -1 ? targetRowData[info.colPrice] : 0;
    const price = typeof rawPrice === 'number' ? rawPrice : (parseFloat(String(rawPrice).replace(/[^0-9.-]/g, '')) || 0);

    const rawCurrentQty = targetRowData[info.colQty];
    const currentQty = typeof rawCurrentQty === 'number' ? rawCurrentQty : (parseFloat(String(rawCurrentQty).replace(/[^0-9.-]/g, '')) || 0);

    if (deductQty > currentQty) {
      throw new Error(`Cannot deduct ${deductQty.toLocaleString()} units! Current storage only has ${currentQty.toLocaleString()} units.`);
    }

    const newQty = currentQty - deductQty;
    const cleanDateOut = String(dateOut || Utilities.formatDate(new Date(), ssTz, 'yyyy-MM-dd')).trim();
    const resolvedUom = uom || inferUomFromDescription(desc);

    // Update inventory sheet quantity
    sheet.getRange(targetRowIndex, info.colQty + 1)
      .setValue(newQty)
      .setNumberFormat('#,##0')
      .setHorizontalAlignment('right');

    // Update DATE OUT in inventory sheet
    if (info.colDateOut !== -1) {
      sheet.getRange(targetRowIndex, info.colDateOut + 1)
        .setNumberFormat('@')
        .setValue(cleanDateOut)
        .setHorizontalAlignment('center');
    }

    // Log the transaction (Canonical 12-column audit log, NO DISPENSED BY)
    logStockOutTransaction({
      timestamp: new Date(),
      dateOut: cleanDateOut,
      itemCode: targetCode,
      itemDesc: desc,
      brandName: brandName,
      uom: resolvedUom,
      expDate: expDate,
      unitPrice: price,
      qtyOut: deductQty,
      previousQty: currentQty,
      remainingBalance: newQty,
      remarks: remarks || 'Normal Dispense'
    });

    return {
      success: true,
      itemCode: targetCode,
      itemDesc: desc,
      brandName: brandName,
      uom: resolvedUom,
      expDate: expDate,
      unitPrice: price,
      qtyDeducted: deductQty,
      dateOut: cleanDateOut,
      previousQty: currentQty,
      newQty: newQty,
      message: `Deducted ${deductQty.toLocaleString()} ${resolvedUom} of ${desc} [${brandName || 'Batch'}]. Date Out: ${cleanDateOut}. Balance: ${newQty.toLocaleString()}`
    };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Updates an existing row (useful for restocking 0-qty lines with new brand, uom, expiry, price, and qty)
 */
function updateMedicineBatch(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const { batchId, rowIndex, itemCode, itemDesc, brandName, uom, newQty, newExpDate, acquisitionPrice } = payload;
    const info = getInventorySheetInfo();
    const sheet = info.sheet;
    const values = sheet.getDataRange().getValues();

    let targetRowIndex = -1;
    if (batchId && info.colBatchId !== -1) {
      for (let r = 1; r < values.length; r++) {
        if (String(values[r][info.colBatchId] || '').trim() === String(batchId).trim()) {
          targetRowIndex = r + 1;
          break;
        }
      }
    }

    if (targetRowIndex === -1 && rowIndex && rowIndex <= values.length) {
      const codeAtRow = String(values[rowIndex - 1][info.colCode] || '').trim();
      if (codeAtRow.toUpperCase() === String(itemCode).trim().toUpperCase()) {
        targetRowIndex = rowIndex;
      }
    }

    if (targetRowIndex === -1) {
      throw new Error(`Row mismatch! Item [${itemCode}] could not be located. Please refresh.`);
    }

    const cleanQty = parseFloat(String(newQty).replace(/[^0-9.-]/g, '')) || 0;
    const cleanExp = String(newExpDate || '').trim();
    const cleanBrand = String(brandName || '').trim();
    const cleanPrice = parseFloat(String(acquisitionPrice || '0').replace(/[^0-9.-]/g, '')) || 0;

    const ssTz = sheet.getParent().getSpreadsheetTimeZone();
    const formattedExp = normalizeExpiryString(cleanExp, ssTz);

    if (itemDesc) {
      sheet.getRange(targetRowIndex, info.colDesc + 1)
        .setValue(String(itemDesc).trim())
        .setHorizontalAlignment('left');
    }

    if (info.colBrand !== -1 && cleanBrand) {
      sheet.getRange(targetRowIndex, info.colBrand + 1)
        .setValue(cleanBrand)
        .setHorizontalAlignment('left');
    }

    if (info.colUom !== -1 && (uom || itemDesc)) {
      const targetUom = String(uom || inferUomFromDescription(itemDesc)).trim();
      sheet.getRange(targetRowIndex, info.colUom + 1)
        .setValue(targetUom)
        .setHorizontalAlignment('center');
    }

    // Set Quantity
    sheet.getRange(targetRowIndex, info.colQty + 1)
      .setValue(cleanQty)
      .setNumberFormat('#,##0')
      .setHorizontalAlignment('right');
    
    // Set Expiration Date
    const expDateObj = getExpiryDateObject(cleanExp, ssTz);
    if (expDateObj) {
      sheet.getRange(targetRowIndex, info.colExp + 1)
        .setValue(expDateObj)
        .setNumberFormat('mmm-yy')
        .setHorizontalAlignment('center');
    } else {
      sheet.getRange(targetRowIndex, info.colExp + 1)
        .setNumberFormat('@')
        .setValue(formattedExp)
        .setHorizontalAlignment('center');
    }

    // Set Acquisition Price
    if (info.colPrice !== -1) {
      sheet.getRange(targetRowIndex, info.colPrice + 1)
        .setValue(cleanPrice)
        .setNumberFormat('"₱"#,##0.00')
        .setHorizontalAlignment('right');
    }

    return {
      success: true,
      message: `Updated [${itemCode}] at row ${targetRowIndex}: Brand: ${cleanBrand || 'N/A'}, Qty: ${cleanQty.toLocaleString()}, Expiry: ${formattedExp}, Price: ₱${cleanPrice.toFixed(2)}`
    };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Initializes and formats the STOCK_OUT_LOG sheet with Brand, UOM, and Date Out columns
 */
function initLogSheet(ss) {
  const logSheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
  const headers = [
    'TIMESTAMP',
    'DATE OUT',
    'ITEM CODE',
    'GENERIC NAME',
    'BRAND NAME',
    'UOM',
    'EXPIRY DATE',
    'UNIT PRICE (₱)',
    'QTY OUT',
    'PREVIOUS QTY',
    'REMAINING BALANCE',
    'REMARKS'
  ];
  
  logSheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1e293b')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  logSheet.setFrozenRows(1);
  logSheet.setColumnWidth(1, 160); // Timestamp
  logSheet.setColumnWidth(2, 110); // Date Out
  logSheet.setColumnWidth(3, 120); // Code
  logSheet.setColumnWidth(4, 260); // Generic
  logSheet.setColumnWidth(5, 130); // Brand
  logSheet.setColumnWidth(6, 85);  // UOM
  logSheet.setColumnWidth(7, 95);  // Expiry
  logSheet.setColumnWidth(8, 110); // Unit Price
  logSheet.setColumnWidth(9, 90);  // Qty Out
  logSheet.setColumnWidth(10, 100); // Prev Qty
  logSheet.setColumnWidth(11, 110); // Remaining Balance
  logSheet.setColumnWidth(12, 180); // Remarks

  return logSheet;
}

/**
 * Appends transaction record to STOCK_OUT_LOG with canonical 12 columns
 * Automatically runs self-healing check to repair headers if outdated
 */
/**
 * Robust Expiry Normalizer for Audit Logs:
 * 1. Cross-references active inventory sheet (Stocks) by itemCode and brandName
 * 2. Recovers original 'MMM-yy' (e.g. 'Jun-30', 'Dec-30') from Google Sheets date auto-coercion
 *    (where Sheets coerced 'Jun-30' into June 30 of the current year 2026)
 */
function cleanLogExpiryValue(rawVal, itemCode, brandName, inventoryInfo) {
  if (!rawVal) return 'NO EXPIRY';

  // 1. Cross-reference with active inventory sheet (Stocks) if available
  if (inventoryInfo && inventoryInfo.values && itemCode) {
    const cleanCode = String(itemCode).trim().toUpperCase();
    const cleanBrand = String(brandName || '').trim().toUpperCase();
    for (let r = 1; r < inventoryInfo.values.length; r++) {
      const row = inventoryInfo.values[r];
      const c = String(row[inventoryInfo.colCode] || '').trim().toUpperCase();
      if (c === cleanCode) {
        const b = inventoryInfo.colBrand !== -1 ? String(row[inventoryInfo.colBrand] || '').trim().toUpperCase() : '';
        if (!cleanBrand || b === cleanBrand || !b) {
          const invExp = row[inventoryInfo.colExp];
          if (invExp) {
            const ssTz = inventoryInfo.sheet.getParent().getSpreadsheetTimeZone();
            return normalizeExpiryString(invExp, ssTz);
          }
        }
      }
    }
  }

  // 2. If rawVal is a Date object (coerced by Google Sheets date parser)
  if (rawVal instanceof Date) {
    const yr = rawVal.getFullYear();
    const day = rawVal.getDate();
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const mon = months[rawVal.getMonth()];
    // If year was defaulted to 2026 and day is 26..31, the day was originally the 2-digit year (e.g. Jun-30 -> June 2030)
    if (yr === 2026 && day >= 26) {
      return `${mon}-${day}`;
    }
    return `${mon}-${String(yr).slice(-2)}`;
  }

  // 3. If rawVal is a string
  const s = String(rawVal).trim();
  if (!s || s.toUpperCase() === 'NO EXPIRY' || s.toUpperCase() === 'N/A') return 'NO EXPIRY';

  // Format MMM-yy (e.g. Jun-30, Feb-27)
  const mShort = s.match(/^([A-Za-z]{3})-(\d{2})$/);
  if (mShort) {
    return mShort[1].charAt(0).toUpperCase() + mShort[1].slice(1).toLowerCase() + '-' + mShort[2];
  }

  // Long Date string like 'Tue Jun 30 2026 15:00:00 GMT+0800 (Singapore Standard Time)'
  const mLongTs = s.match(/^[A-Za-z]{3}\s+([A-Za-z]{3})\s+(\d{1,2})\s+(\d{4})/);
  if (mLongTs) {
    const mon = mLongTs[1].charAt(0).toUpperCase() + mLongTs[1].slice(1).toLowerCase();
    const day = parseInt(mLongTs[2], 10);
    const yr = parseInt(mLongTs[3], 10);
    if (yr === 2026 && day >= 26) {
      return `${mon}-${day}`;
    }
    return `${mon}-${String(yr).slice(-2)}`;
  }

  // Fallback Date parser
  const d = new Date(s);
  if (!isNaN(d.getTime()) && s.length > 8) {
    const yr = d.getFullYear();
    const day = d.getDate();
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const mon = months[d.getMonth()];
    if (yr === 2026 && day >= 26) {
      return `${mon}-${day}`;
    }
    return `${mon}-${String(yr).slice(-2)}`;
  }

  return s;
}

/**
 * Appends transaction record to STOCK_OUT_LOG with canonical 12 columns
 * Automatically runs self-healing check to repair headers if outdated
 */
function logStockOutTransaction(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = initLogSheet(ss);
    } else {
      // Self-healing check: Verify Row 1 has the modern 12 headers
      const curHeaders = logSheet.getRange(1, 1, 1, Math.max(logSheet.getLastColumn(), 12)).getValues()[0];
      const hStr = curHeaders.map(h => String(h).toUpperCase()).join('|');
      if (!hStr.includes('DATE OUT') || !hStr.includes('UOM') || hStr.includes('DISPENSED BY') || curHeaders.length < 12) {
        repairAndMigrateLogSheet(false); // Repair silently without blocking popup
        logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
      }
    }

    const ssTz = ss.getSpreadsheetTimeZone();
    const cleanExpDate = cleanLogExpiryValue(data.expDate, data.itemCode, data.brandName, null);
    const nextRow = logSheet.getLastRow() + 1;

    const row = [
      Utilities.formatDate(data.timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      data.dateOut || Utilities.formatDate(data.timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      data.itemCode,
      data.itemDesc,
      data.brandName || '',
      data.uom || '',
      cleanExpDate,
      data.unitPrice || 0,
      data.qtyOut,
      data.previousQty,
      data.remainingBalance,
      data.remarks
    ];

    // Pre-format column 7 (Expiry Date) as Plain Text '@' to permanently lock clean string (e.g. 'Jun-30')
    logSheet.getRange(nextRow, 7).setNumberFormat('@');
    logSheet.getRange(nextRow, 1, 1, 12).setValues([row]);
    
    // Formatting:
    logSheet.getRange(nextRow, 1).setHorizontalAlignment('center');
    logSheet.getRange(nextRow, 2).setNumberFormat('@').setHorizontalAlignment('center');
    logSheet.getRange(nextRow, 3).setHorizontalAlignment('center');
    logSheet.getRange(nextRow, 4).setHorizontalAlignment('left');
    logSheet.getRange(nextRow, 5).setHorizontalAlignment('left');
    logSheet.getRange(nextRow, 6).setHorizontalAlignment('center');
    logSheet.getRange(nextRow, 7).setNumberFormat('@').setHorizontalAlignment('center');
    logSheet.getRange(nextRow, 8).setNumberFormat('"₱"#,##0.00').setHorizontalAlignment('right');
    logSheet.getRange(nextRow, 9, 1, 3).setNumberFormat('#,##0').setHorizontalAlignment('right');
    logSheet.getRange(nextRow, 12).setHorizontalAlignment('left');
  } catch (err) {
    Logger.log("Error logging stock-out: " + err.message);
  }
}

/**
 * Repairs, normalizes, and migrates STOCK_OUT_LOG to the canonical 12-column layout:
 * TIMESTAMP | DATE OUT | ITEM CODE | GENERIC NAME | BRAND NAME | UOM | EXPIRY DATE | UNIT PRICE (₱) | QTY OUT | PREVIOUS QTY | REMAINING BALANCE | REMARKS
 * Removes 'DISPENSED BY', eliminates total cost, and shifts all misplaced historical entries.
 */
function repairAndMigrateLogSheet(showAlert) {
  let ui = null;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = initLogSheet(ss);
      if (showAlert !== false && ui) {
        ui.alert('Stock-Out Log', 'STOCK_OUT_LOG tab was created with the canonical 12-column layout.', ui.ButtonSet.OK);
      }
      return { success: true, message: 'Created new log sheet.' };
    }

    const headers12 = [
      'TIMESTAMP',
      'DATE OUT',
      'ITEM CODE',
      'GENERIC NAME',
      'BRAND NAME',
      'UOM',
      'EXPIRY DATE',
      'UNIT PRICE (₱)',
      'QTY OUT',
      'PREVIOUS QTY',
      'REMAINING BALANCE',
      'REMARKS'
    ];

    const lastRow = logSheet.getLastRow();
    const lastCol = Math.max(logSheet.getLastColumn(), 15);
    const ssTz = ss.getSpreadsheetTimeZone();

    let inventoryInfo = null;
    try {
      inventoryInfo = getInventorySheetInfo();
    } catch (e) {
      Logger.log("Notice: inventory sheet could not be pre-loaded: " + e.message);
    }

    const normalizedRows = [];

    if (lastRow > 1) {
      const dataRange = logSheet.getRange(2, 1, lastRow - 1, lastCol);
      const rawValues = dataRange.getValues();

      for (let i = 0; i < rawValues.length; i++) {
        const r = rawValues[i];
        const tsRaw = r[0];
        if (!tsRaw && !r[1] && !r[2]) continue; // skip blank

        let timestamp = tsRaw instanceof Date ? Utilities.formatDate(tsRaw, ssTz, 'yyyy-MM-dd HH:mm:ss') : String(tsRaw || '').trim();
        let dateOut = '';
        let itemCode = '';
        let genericName = '';
        let brandName = '';
        let uom = '';
        let expDate = '';
        let unitPrice = 0;
        let qtyOut = 0;
        let prevQty = 0;
        let remBal = 0;
        let remarks = '';

        const col1Str = String(r[1] || '').trim();
        const col2Str = String(r[2] || '').trim();
        // Check if r[1] is an Item Code (Legacy 12-column row where r[1] is 'DMR...', r[2] is description)
        const isLegacyCol1Code = /^[A-Z]{2,}[0-9]+/.test(col1Str);

        if (isLegacyCol1Code) {
          // Format like Row 2 from screenshot:
          dateOut = timestamp.split(' ')[0] || Utilities.formatDate(new Date(), ssTz, 'yyyy-MM-dd');
          itemCode = col1Str;
          genericName = col2Str;
          brandName = String(r[3] || '').trim();
          expDate = cleanLogExpiryValue(r[4], itemCode, brandName, inventoryInfo);
          unitPrice = typeof r[5] === 'number' ? r[5] : (parseFloat(String(r[5] || '0').replace(/[^0-9.-]/g, '')) || 0);
          qtyOut = typeof r[6] === 'number' ? r[6] : (parseFloat(String(r[6] || '0').replace(/[^0-9.-]/g, '')) || 0);
          prevQty = typeof r[8] === 'number' ? r[8] : (parseFloat(String(r[8] || '0').replace(/[^0-9.-]/g, '')) || 0);
          remBal = typeof r[9] === 'number' ? r[9] : (parseFloat(String(r[9] || '0').replace(/[^0-9.-]/g, '')) || 0);
          remarks = String(r[10] || '').trim();
          uom = inferUomFromDescription(genericName);
        } else {
          // Modern row format (or Row 3 from screenshot):
          dateOut = col1Str instanceof Date ? Utilities.formatDate(col1Str, ssTz, 'yyyy-MM-dd') : col1Str;
          itemCode = col2Str;
          genericName = String(r[3] || '').trim();
          brandName = String(r[4] || '').trim();
          uom = String(r[5] || '').trim() || inferUomFromDescription(genericName);
          expDate = cleanLogExpiryValue(r[6], itemCode, brandName, inventoryInfo);
          unitPrice = typeof r[7] === 'number' ? r[7] : (parseFloat(String(r[7] || '0').replace(/[^0-9.-]/g, '')) || 0);
          qtyOut = typeof r[8] === 'number' ? r[8] : (parseFloat(String(r[8] || '0').replace(/[^0-9.-]/g, '')) || 0);
          prevQty = typeof r[9] === 'number' ? r[9] : (parseFloat(String(r[9] || '0').replace(/[^0-9.-]/g, '')) || 0);
          remBal = typeof r[10] === 'number' ? r[10] : (parseFloat(String(r[10] || '0').replace(/[^0-9.-]/g, '')) || 0);
          remarks = String(r[11] || '').trim();
        }

        normalizedRows.push([
          timestamp,
          dateOut,
          itemCode,
          genericName,
          brandName,
          uom,
          expDate,
          unitPrice,
          qtyOut,
          prevQty,
          remBal,
          remarks
        ]);
      }
    }

    // Clear entire sheet and rebuild cleanly
    logSheet.clear();
    while (logSheet.getMaxColumns() < 12) {
      logSheet.insertColumnAfter(logSheet.getMaxColumns());
    }

    // Set canonical headers
    logSheet.getRange(1, 1, 1, 12)
      .setValues([headers12])
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    logSheet.setFrozenRows(1);

    // Set Column Widths
    logSheet.setColumnWidth(1, 160); // Timestamp
    logSheet.setColumnWidth(2, 110); // Date Out
    logSheet.setColumnWidth(3, 120); // Item Code
    logSheet.setColumnWidth(4, 260); // Generic Name
    logSheet.setColumnWidth(5, 130); // Brand Name
    logSheet.setColumnWidth(6, 85);  // UOM
    logSheet.setColumnWidth(7, 95);  // Expiry Date
    logSheet.setColumnWidth(8, 110); // Unit Price
    logSheet.setColumnWidth(9, 90);  // Qty Out
    logSheet.setColumnWidth(10, 100); // Prev Qty
    logSheet.setColumnWidth(11, 110); // Remaining Balance
    logSheet.setColumnWidth(12, 180); // Remarks

    // Write back normalized data rows
    if (normalizedRows.length > 0) {
      // Pre-format column 7 (Expiry Date) as Plain Text '@' BEFORE writing values to prevent Google Sheets date coercion
      logSheet.getRange(2, 7, normalizedRows.length, 1).setNumberFormat('@');
      logSheet.getRange(2, 1, normalizedRows.length, 12).setValues(normalizedRows);

      // Formatting
      logSheet.getRange(2, 1, normalizedRows.length, 1).setHorizontalAlignment('center');
      logSheet.getRange(2, 2, normalizedRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center');
      logSheet.getRange(2, 3, normalizedRows.length, 1).setHorizontalAlignment('center');
      logSheet.getRange(2, 4, normalizedRows.length, 1).setHorizontalAlignment('left');
      logSheet.getRange(2, 5, normalizedRows.length, 1).setHorizontalAlignment('left');
      logSheet.getRange(2, 6, normalizedRows.length, 1).setHorizontalAlignment('center');
      logSheet.getRange(2, 7, normalizedRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center');
      logSheet.getRange(2, 8, normalizedRows.length, 1).setNumberFormat('"₱"#,##0.00').setHorizontalAlignment('right');
      logSheet.getRange(2, 9, normalizedRows.length, 3).setNumberFormat('#,##0').setHorizontalAlignment('right');
      logSheet.getRange(2, 12, normalizedRows.length, 1).setHorizontalAlignment('left');
    }

    // Trim any extra columns beyond Col 12 (delete legacy Column M, N, O etc.)
    const curMaxCols = logSheet.getMaxColumns();
    if (curMaxCols > 12) {
      try {
        logSheet.deleteColumns(13, curMaxCols - 12);
      } catch (colErr) {}
    }

    const msg = `Stock-Out Log Upgraded!\n\n` +
      `• Normalized: ${normalizedRows.length} audit records\n` +
      `• Canonical layout: 12 Columns (A to L)\n` +
      `• Removed 'DISPENSED BY' and legacy extra columns\n` +
      `• All dates, codes, and balances are now properly aligned.`;

    if (showAlert !== false && ui) {
      ui.alert('Stock-Out Log Upgrade', msg, ui.ButtonSet.OK);
    }

    return { success: true, message: msg, count: normalizedRows.length };
  } catch (err) {
    if (showAlert !== false && ui) {
      ui.alert('Error', 'Failed to migrate log sheet: ' + err.message, ui.ButtonSet.OK);
    }
    return { success: false, error: err.message };
  }
}

/**
 * Fetches recent audit logs from STOCK_OUT_LOG (Canonical 12-column layout)
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

    // Read in reverse chronological order
    for (let i = values.length - 1; i >= 1 && logs.length < max; i--) {
      const r = values[i];
      if (!r[0] && !r[1] && !r[2]) continue;

      logs.push({
        timestamp: r[0],
        dateOut: r[1] ? (r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[1])) : '',
        itemCode: r[2],
        itemDesc: r[3],
        brandName: r[4] || '',
        uom: r[5] || '',
        expDate: cleanLogExpiryValue(r[6], r[2], r[4], null),
        unitPrice: r[7] || 0,
        qtyOut: r[8],
        previousQty: r[9],
        remainingBalance: r[10],
        remarks: r[11] || ''
      });
    }

    return { success: true, logs: logs };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Finds the first available row for a new medicine:
 * 1. Checks if Row 2 is empty. If Row 2 is empty, returns 2.
 * 2. Checks subsequent rows for the first row with empty Item Code AND Description.
 * 3. If no empty rows exist within the data bounds, appends to the row after the last actual data row.
 */
function findNextAvailableRow(sheet, info) {
  const maxRows = sheet.getMaxRows();
  if (maxRows <= 1) {
    if (maxRows === 1) sheet.insertRowAfter(1);
    return 2;
  }

  const colCodeIdx = (info && info.colCode !== undefined && info.colCode !== -1) ? info.colCode + 1 : 1;
  const colDescIdx = (info && info.colDesc !== undefined && info.colDesc !== -1) ? info.colDesc + 1 : 2;

  const colCodeVals = sheet.getRange(1, colCodeIdx, maxRows, 1).getValues();
  const colDescVals = sheet.getRange(1, colDescIdx, maxRows, 1).getValues();

  // Find the first completely empty data row below header (Row 2 onwards)
  for (let r = 1; r < maxRows; r++) { // r=1 corresponds to Row 2
    const code = String(colCodeVals[r][0] || '').trim();
    const desc = String(colDescVals[r][0] || '').trim();
    if (code === '' && desc === '') {
      return r + 1; // 1-based index in sheet (e.g. Row 2)
    }
  }

  // If all rows down to maxRows are filled, insert a new row
  sheet.insertRowAfter(maxRows);
  return maxRows + 1;
}

/**
 * Finds the actual last row containing medicine data in the sheet.
 * Ignores empty rows with borders, data validation, or ghost formatting.
 * Returns 1 (header row) if the sheet has no medicine entries.
 */
function getLastActualDataRow(sheet, info) {
  const maxRows = sheet.getMaxRows();
  if (maxRows <= 1) return 1;

  const colCodeIdx = (info && info.colCode !== undefined && info.colCode !== -1) ? info.colCode + 1 : 1;
  const colDescIdx = (info && info.colDesc !== undefined && info.colDesc !== -1) ? info.colDesc + 1 : 2;

  const colCodeVals = sheet.getRange(1, colCodeIdx, maxRows, 1).getValues();
  const colDescVals = sheet.getRange(1, colDescIdx, maxRows, 1).getValues();

  for (let r = maxRows - 1; r >= 1; r--) {
    const code = String(colCodeVals[r][0] || '').trim();
    const desc = String(colDescVals[r][0] || '').trim();
    if (code !== '' || desc !== '') {
      return r + 1; // 1-based row index in sheet
    }
  }

  return 1; // Only header row exists
}

/**
 * Clean Blank Rows & Consolidate Sheet
 * Removes all empty ghost rows and shifts entries (e.g., stranded at Row 75)
 * up to Row 2 so the sheet is cleanly packed without any blank row gaps.
 */
function cleanAndConsolidateSheet() {
  let ui = null;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    // Running in standalone web app without spreadsheet UI
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const info = getInventorySheetInfo();
    const sheet = info.sheet;
    const maxRows = sheet.getMaxRows();
    if (maxRows <= 1) {
      if (ui) ui.alert('Consolidation', 'The inventory sheet is empty. Nothing to consolidate.', ui.ButtonSet.OK);
      return { success: true, message: 'Sheet is empty.' };
    }

    const lastCol = Math.max(sheet.getLastColumn(), 8);
    const range = sheet.getRange(2, 1, maxRows - 1, lastCol);
    const values = range.getValues();

    const cleanRows = [];
    let emptyCount = 0;

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const code = String(row[info.colCode] || '').trim();
      const desc = String(row[info.colDesc] || '').trim();

      if (code !== '' || desc !== '') {
        cleanRows.push(row);
      } else {
        emptyCount++;
      }
    }

    if (emptyCount === 0) {
      if (ui) ui.alert('Consolidation', 'The sheet is already clean! No empty rows were found.', ui.ButtonSet.OK);
      return { success: true, message: 'No blank rows found.', consolidatedCount: cleanRows.length };
    }

    // Clear existing data range below headers
    range.clearContent();
    range.clearDataValidations();

    // Write back non-empty rows starting at Row 2
    if (cleanRows.length > 0) {
      sheet.getRange(2, 1, cleanRows.length, lastCol).setValues(cleanRows);

      // Re-apply UOM dropdown validation to Column D
      if (info.colUom !== -1) {
        const uomRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(CONFIG.STANDARD_UOM_LIST, true)
          .setAllowInvalid(true)
          .build();
        sheet.getRange(2, info.colUom + 1, cleanRows.length, 1)
          .setDataValidation(uomRule)
          .setHorizontalAlignment('center');
      }

      // Re-apply Qty format
      if (info.colQty !== -1) {
        sheet.getRange(2, info.colQty + 1, cleanRows.length, 1)
          .setNumberFormat('#,##0')
          .setHorizontalAlignment('right');
      }

      // Re-apply Expiry format
      if (info.colExp !== -1) {
        sheet.getRange(2, info.colExp + 1, cleanRows.length, 1)
          .setNumberFormat('mmm-yy')
          .setHorizontalAlignment('center');
      }

      // Re-apply Price format
      if (info.colPrice !== -1) {
        sheet.getRange(2, info.colPrice + 1, cleanRows.length, 1)
          .setNumberFormat('"₱"#,##0.00')
          .setHorizontalAlignment('right');
      }

      // Re-apply Date Out format
      if (info.colDateOut !== -1) {
        sheet.getRange(2, info.colDateOut + 1, cleanRows.length, 1)
          .setNumberFormat('@')
          .setHorizontalAlignment('center');
      }
    }

    // Optionally trim excessive trailing blank rows to keep the sheet compact
    const targetRowCount = Math.max(cleanRows.length + 25, 50);
    const curMaxRows = sheet.getMaxRows();
    if (curMaxRows > targetRowCount + 15) {
      try {
        sheet.deleteRows(targetRowCount + 1, curMaxRows - targetRowCount);
      } catch (e) {}
    }

    const msg = 'Consolidation Complete!\n\n' +
      '• Removed: ' + emptyCount + ' blank/ghost rows\n' +
      '• Consolidated: ' + cleanRows.length + ' medicine records\n' +
      '• Inventory is now organized starting cleanly at Row 2.';

    if (ui) ui.alert('Clean & Consolidate', msg, ui.ButtonSet.OK);

    return {
      success: true,
      message: msg,
      emptyCount: emptyCount,
      activeCount: cleanRows.length
    };
  } catch (err) {
    if (ui) ui.alert('Error', 'Failed to consolidate sheet: ' + err.message, ui.ButtonSet.OK);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}
