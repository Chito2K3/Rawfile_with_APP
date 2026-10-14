# 🏥 Pharmacy Expiry Monitoring (Amang Rodriguez Memorial Medical Center)

This custom Google Apps Script application connects directly to the ARMMC pharmacy inventory spreadsheet to monitor expiration dates (2026/2027/2028+), track batch deliveries with distinct **Brand Names** and **Acquisition Prices**, record stock-out dispensations with financial impact, restock 0-quantity items, and automatically group new batches below existing medicines.

---

## 📋 Recommended Sheet Column Layout (Option A)

For optimal tracking, your master inventory sheet is organized as follows:

| Column A | Column B | Column C | Column D | Column E | Column F | Column G |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **ITEM CODE** | **ITEM DESCRIPTION** | **BRAND NAME** | **QTY STORAGE** | **EXP DATE** | **ACQUISITION PRICE** | **DATE OUT** |
| `DMR000451` | `FLUCONAZOLE 200MG CAPSULE` | `Diflucan` | `200` | `Jun-28` | `₱125.00` | `2026-10-14` |
| `DMR000451` | `FLUCONAZOLE 200MG CAPSULE` | `Flucan` | `2,000` | `Apr-29` | `₱45.00` | `2026-10-10` |

> [!TIP]
> **1-Click Column Upgrade:**  
> If your existing sheet only has 4 columns (`ITEM CODE`, `ITEM DESCRIPTION`, `QTY STORAGE`, `EXP DATE`), simply click **`🏥 Pharmacy Expiry Monitoring` > `🛠️ Ensure / Upgrade Sheet Columns (Option A)`** in the Google Sheets menu. The system will automatically insert the `BRAND NAME`, `ACQUISITION PRICE`, and `DATE OUT` columns for you without overwriting any existing data!

---

## 🚀 How to Install in Google Sheets

### Step 1: Open Google Apps Script Editor
1. Open your Google Sheet: [Pharmacy Storage Spreadsheet](https://docs.google.com/spreadsheets/d/1sKJcZG-rnNN9_VknrbkxcuomegK1-CuY_iB81AQOFpc/edit?gid=1609045579#gid=1609045579).
2. In the top menu, click **Extensions** > **Apps Script**.

---

### Step 2: Create / Update the 4 Script Files
Inside the Apps Script editor, copy and paste the code from each of the files in the `gas_project/` directory:

#### 1. `Code.gs` (Script File)
* Replace the contents with the code inside [`gas_project/Code.gs`](./Code.gs).

#### 2. `Index.html` (HTML File)
* Click **➕ (Add a file)** > **HTML**, name it `Index`.
* Replace the contents with the code inside [`gas_project/Index.html`](./Index.html).

#### 3. `Styles.html` (HTML File)
* Click **➕ (Add a file)** > **HTML**, name it `Styles`.
* Replace the contents with the code inside [`gas_project/Styles.html`](./Styles.html).

#### 4. `JavaScript.html` (HTML File)
* Click **➕ (Add a file)** > **HTML**, name it `JavaScript`.
* Replace the contents with the code inside [`gas_project/JavaScript.html`](./JavaScript.html).

---

### Step 3: Save and Refresh
1. Click the **💾 Save project** icon at the top.
2. Go back to your Google Sheet tab and **refresh the browser page (F5)**.
3. You will now see the menu: **`🏥 Pharmacy Expiry Monitoring`**.
4. (Optional) Run **`🛠️ Ensure / Upgrade Sheet Columns (Option A)`** to verify your column headers.

---

## 🎯 How to Use the Features

### 1. ➕ Adding a Medicine or New Batch Delivery
* Go to the **Add Medicine / New Batch** tab.
* **Master Formulary Dropdowns (`Sheet1` Sync)**:
  * Both **Item Code (DCI)** and **Generic Description** are now **searchable dropdowns** dynamically populated from the **`Sheet1`** master catalog tab.
  * Selecting or typing a **DCI Code** (e.g. `DMR000002`) immediately auto-fills the matching **Generic Description**.
  * Selecting or typing a **Generic Description** immediately auto-fills the matching **DCI Code**.
  * Click the **✕** button beside either input to quickly clear and pick another medicine.
* **Smart Batch Grouping & Insertion**:
  * **If the medicine already exists in storage** (e.g. `DMR000451` Fluconazole): The system automatically detects previous batches and shows existing brands, quantities, and prices. The new batch is **automatically inserted directly below the last batch** in Google Sheets!
  * **If brand new**: Appends to the master list.
* Input the **Brand Name** (required), **Storage Quantity**, **Expiry Date**, and **Acquisition Price** (₱ / unit).
* Click **Save & Insert into Sheet**.

### 2. ➖ Quick Stock-Out / Dispensing by Brand & Batch
* In the **Inventory & Stock-Out** tab, search by Generic Name, Item Code, Brand Name (e.g. `Diflucan`), or Expiry.
* Click **➖ Out** on the target batch.
* The modal clearly displays the **Brand Name**, **Batch Expiry**, and **Current Storage Stock**.
* Enter how many units were dispensed.
* Select the **Date When Item Was Out** using the date picker (defaults to today's date).
* Click **Confirm Stock-Out**:
  * Updates `QTY STORAGE` in Google Sheets.
  * Records the selected date into the **`DATE OUT`** column (Column G) in Google Sheets.
  * Writes a complete audit record to **`STOCK_OUT_LOG`**.

### 3. 🔄 Restocking 0-Quantity Items
* Go to the **Restock 0-Qty Items** tab.
* Click **Restock & Update** on any out-of-stock medicine.
* Enter the incoming **Brand Name**, **New Expiry Date**, **Updated Price**, and **Quantity**.

### 4. 📋 Audit Log Sheet (`STOCK_OUT_LOG`)
* Automatically maintains an audit trail with:
  * `TIMESTAMP`
  * `DATE OUT`
  * `ITEM CODE`
  * `GENERIC NAME`
  * `BRAND NAME`
  * `EXPIRY DATE`
  * `UNIT PRICE (₱)` *(Repository reference)*
  * `QTY OUT`
  * `PREVIOUS QTY` & `REMAINING BALANCE`
  * `REMARKS / DEPARTMENT`
  * `DISPENSED BY`

---

## 🌐 Deploy / Update Standalone Web App
If you already deployed a Web App:
1. In Apps Script, click **Deploy** > **Manage deployments**.
2. Click the **✏️ Edit (pencil)** icon on your active deployment.
3. In the **Version** dropdown, select **New version**.
4. Click **Deploy**.
5. Refresh your Web App browser tab to see the updated Brand & Price fields!
