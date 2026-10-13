# 🏥 Amang Rodriguez Memorial Medical Center - Pharmacy Expiry Monitoring System

An automated Google Apps Script (GAS) application and data analysis suite for managing pharmacy storage inventory, FEFO (First-Expired, First-Out) expiration tracking, **Brand Name and Acquisition Price batch management**, stock-out dispensing, and automated batch grouping for **Amang Rodriguez Memorial Medical Center (ARMMC)**.

---

## 🌟 Key Features

1. **Brand Name & Acquisition Price Batch Tracking**:
   - Every batch under a DCI Code and Generic Name is individually distinguished by its **Brand Name**, **Expiry Date**, and **Acquisition Price** (₱ / unit).
2. **Smart Batch Grouping & Insertion (`sheet.insertRowAfter`)**:
   - When registering a new delivery for an existing medicine (e.g., `DMR000451 FLUCONAZOLE`), the system detects existing batches and inserts the new row **directly below the latest batch** of that medicine in Google Sheets.
   - Brand new medicines are automatically appended to the inventory master list.
3. **Stock-Out Dispensing & Real-Time Balance**:
   - 1-click modal to deduct dispensed quantities from the specific brand batch with stock overdraft prevention.
   - Automatically computes total dispensed monetary valuation (`Qty Out × Acquisition Price`).
   - Instantly updates `QTY STORAGE` in Google Sheets.
4. **Automated Financial Audit Logging (`STOCK_OUT_LOG`)**:
   - Records every transaction with timestamps, DCI code, generic name, brand name, batch expiry, unit acquisition price, quantity deducted, total value dispensed, remaining balance, department/remarks, and user credentials.
5. **Zero-Stock Replenishment & In-Place Editing**:
   - Dedicated tab to review 0-quantity items, update the brand name, incoming expiry date, acquisition price, and replenish storage quantity.
6. **FEFO Expiry Risk Classification**:
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
