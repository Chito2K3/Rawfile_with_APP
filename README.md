# 🏥 Amang Rodriguez Memorial Medical Center - Pharmacy Expiry Monitoring System

An automated Google Apps Script (GAS) application and data analysis suite for managing pharmacy storage inventory, FEFO (First-Expired, First-Out) expiration tracking, **Brand Name and Acquisition Price batch management**, stock-out dispensing, and automated batch grouping for **Amang Rodriguez Memorial Medical Center (ARMMC)**.

---

## 🌟 Key Features

1. **Brand Name, UOM & Acquisition Price Batch Tracking**:
   - Every batch under a DCI Code and Generic Name is individually distinguished by its **Brand Name**, **Unit of Measure (UOM)**, **Expiry Date**, and **Acquisition Price** (₱ / unit).
2. **Standard Hospital UOM Dropdowns & Auto-Detection**:
   - Native dropdown selection in both the Web App and Google Sheet with the 9 standard hospital units: `Tablet`, `Capsule`, `Vial`, `Ampule`, `Bottle`, `Tube`, `Nebule`, `Syringe`, and `Sachet`.
   - Smart keyword auto-detection automatically selects the matching UOM as you pick or type a medicine (e.g. `FLUCONAZOLE 200MG CAPSULE` -> `Capsule`, `ZOLEDRONIC ACID 4MG/5ML VIAL` -> `Vial`).
3. **Sheet1 Master Formulary Auto-Sync & Dual Comboboxes**:
   - Bi-directional auto-sync between **DCI Item Code** and **Generic Description** sourced directly from the **`Sheet1`** master formulary catalog tab.
   - Interactive searchable floating dropdowns with keyboard navigation, match highlighting, and 1-click field reset (`✕`).
4. **Smart FEFO Batch Grouping & Insertion (`sheet.insertRowBefore` / `insertRowAfter`)**:
   - When registering a new delivery for an existing medicine (e.g., `PNSS 1L` or `FLUCONAZOLE`), the system detects existing batches and inserts the new batch in **FEFO order (First-Expired, First-Out)** directly inside that medicine's group in Google Sheets!
   - Batches with earlier expiration dates (e.g., `Feb-28`) are automatically placed in the first row above later expiring batches (e.g., `Jan-30`).
   - Brand new medicines are automatically appended to the inventory master list.
5. **Stock-Out Dispensing & Real-Time Balance**:
   - 1-click modal to deduct dispensed quantities from the specific brand batch with stock overdraft prevention and UOM unit badges.
   - Automatically computes total dispensed monetary valuation (`Qty Out × Acquisition Price`).
   - Instantly updates `QTY STORAGE` and `DATE OUT` in Google Sheets.
6. **Automated Financial Audit Logging (`STOCK_OUT_LOG`)**:
   - Records every transaction with timestamps, date out, DCI code, generic name, brand name, UOM, batch expiry, unit acquisition price, quantity deducted, remaining balance, department/remarks, and user credentials.
7. **Zero-Stock Replenishment & In-Place Editing**:
   - Dedicated tab to review 0-quantity items, update the brand name, UOM, incoming expiry date, acquisition price, and replenish storage quantity.
8. **Intelligent Row Placement & One-Click Sheet Consolidation**:
   - Brand new medicines automatically append to the first available row starting at **Row 2** directly beneath the headers—never skipping down to row 75 or leaving ghost rows.
   - Built-in `🧹 Clean Blank Rows & Consolidate Sheet` utility removes all empty/ghost rows from cleared data and consolidates inventory contiguously.
9. **FEFO Expiry Risk Classification**:
   - 🔴 **2026 Critical / Immediate Risk** (FEFO priority dispensing)
   - 🟠 **2027 Near-Term Risk**
   - 🟢 **2028+ Safe Stock**
   - ⚠️ **Missing Expiry Audit Alerts**

---

## 📁 Repository Structure

```
├── gas_project/                      # Google Apps Script Source Files
│   ├── Code.gs                       # Backend controller, menu, Option A schema & sheet APIs
│   ├── Index.html                    # Main responsive Single-Page UI with Brand & Price fields
│   ├── Styles.html                   # Clinical CSS stylesheet, brand badges & cost callouts
│   ├── JavaScript.html               # Frontend logic, live search & financial calculation
│   └── README_SETUP.md               # Step-by-step GAS setup guide & Option A column layout
├── analyze.ps1                       # PowerShell Inventory Analysis Script
├── run_analysis.ps1                  # Comprehensive Expiry Profiling Report
├── sheet_1609045579.csv              # Raw inventory export snapshot
└── NEW RAW FILE__EXPIRY_2026.xlsx    # Raw warehouse inventory workbook
```

---

## 🚀 Setup & Deployment Guide

For full instructions on configuring the Google Sheet columns and deploying this to Google Sheets or as a standalone Web App, see [`gas_project/README_SETUP.md`](gas_project/README_SETUP.md).

---

## 👥 Institution

**Amang Rodriguez Memorial Medical Center (ARMMC)**  
*Pharmacy Storage & Inventory Management*
