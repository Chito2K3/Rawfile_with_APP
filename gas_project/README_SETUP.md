# 🏥 Pharmacy Expiry Monitoring (Amang Rodriguez Memorial Medical Center)

This custom Google Apps Script application connects directly to the ARMMC pharmacy inventory spreadsheet to monitor expiration dates (2026/2027/2028+), track batch deliveries with distinct **Brand Names** and **Acquisition Prices**, record stock-out dispensations with financial impact, restock 0-quantity items, and automatically group new batches below existing medicines.

---

## 📋 Recommended Sheet Column Layout (Option A with UOM)

For optimal tracking, your master inventory sheet is organized with an 8-column schema:

| Column A | Column B | Column C | Column D | Column E | Column F | Column G | Column H |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **ITEM CODE** | **ITEM DESCRIPTION** | **BRAND NAME** | **UOM** | **QTY STORAGE** | **EXP DATE** | **ACQUISITION PRICE** | **DATE OUT** |
| `DMR000451` | `FLUCONAZOLE 200MG CAPSULE` | `Diflucan` | `Capsule` | `200` | `Jun-28` | `₱125.00` | `2026-10-14` |
| `DMR000451` | `FLUCONAZOLE 200MG CAPSULE` | `Flucan` | `Capsule` | `2,000` | `Apr-29` | `₱45.00` | `2026-10-10` |
| `DMR001125` | `ZOLEDRONIC ACID 4MG/5ML VIAL` | `Zometa` | `Vial` | `50` | `Oct-27` | `₱3,500.00` | `--` |

> [!TIP]
> **1-Click Column Upgrade & Dropdown Chips:**  
> Simply click **`🏥 Pharmacy Expiry Monitoring` > `🛠️ Ensure / Upgrade Sheet Columns (Option A)`** in the Google Sheets menu. The system will:
> 1. Automatically insert the `BRAND NAME`, `UOM`, `ACQUISITION PRICE`, and `DATE OUT` columns without overwriting existing data.
> 2. Create native **Google Sheets Dropdown Chips** in Column D with the 9 approved hospital units: `Tablet`, `Capsule`, `Vial`, `Ampule`, `Bottle`, `Tube`, `Nebule`, `Syringe`, `Sachet`.
> 3. Automatically backfill existing medicine rows with the appropriate UOM inferred from generic descriptions!

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
* **Master Formulary Dropdowns (`Medicine` Tab Sync)**:
  * Both **Item Code (DCI)** and **Generic Description** are searchable dropdowns dynamically populated from the **`Medicine`** tab.
  * Selecting or typing a **DCI Code** (e.g. `DMR000002`) immediately auto-fills the matching **Generic Description**.
  * Selecting or typing a **Generic Description** immediately auto-fills the matching **DCI Code**.
  * Click the **✕** button beside either input to quickly clear and pick another medicine.
* **Registering New Medicines into Master Catalog (When Code is Not on the List)**:
  * If a DCI code (e.g. `DMC000648`) or Generic Name is unlisted, the dropdown displays a green interactive card: **`➕ Add to Master Formulary`**.
  * You can also click the **`➕ Add to Master Catalog`** button directly in the header banner.
  * A modal opens allowing you to register the new DCI Code and Generic Name directly into the **`Medicine`** spreadsheet tab in 1 click!
  * **Automatic Dual-Save**: Even if you skip the modal and directly click **Save & Insert into Sheet**, the backend automatically checks if the medicine exists in the `Medicine` tab, and if missing, automatically enrolls it into the master formulary for all future searches!
* **Smart FEFO Batch Insertion (Near Expiry First)**:
  * **If the medicine already exists in storage** (e.g. `PNSS 1L` or `FLUCONAZOLE`): The system automatically detects existing batches and inserts the new batch in **FEFO order (First-Expired, First-Out)**. If the new batch has an earlier expiry (e.g., `Feb-28`), it is automatically placed **in the first row above later batches** (e.g., `Jan-30`)!
  * **If brand new**: Appends cleanly starting at **Row 2** without jumping down to row 75.
* **Unit of Measure (UOM) Dropdown & Auto-Detection**:
  * Selecting or typing a medicine automatically infers and selects the corresponding standard UOM (e.g. `FLUCONAZOLE 200MG CAPSULE` -> `Capsule`, `ZOLEDRONIC ACID 4MG/5ML VIAL` -> `Vial`).
  * Choose from the 9 standard hospital units: `Tablet`, `Capsule`, `Vial`, `Ampule`, `Bottle`, `Tube`, `Nebule`, `Syringe`, `Sachet`.
  * The storage quantity input label dynamically updates (e.g. `Storage Quantity (Vials)`).
* Input the **Brand Name** (required), **Storage Quantity**, **Expiry Date**, and **Acquisition Price** (₱ / unit).
* Click **Save & Insert into Sheet**.

### 2. ➖ Quick Stock-Out / Dispensing by Brand, UOM & Batch
* In the **Inventory & Stock-Out** tab, search by Generic Name, Item Code, Brand Name (e.g. `Diflucan`), UOM, or Expiry.
* Click **➖ Out** on the target batch.
* The modal clearly displays the **Brand Name**, **UOM Badge**, **Batch Expiry**, and **Current Storage Stock**.
* Enter how many units were dispensed.
* Select the **Date When Item Was Out** using the date picker (defaults to today's date).
* Click **Confirm Stock-Out**:
  * Updates `QTY STORAGE` in Google Sheets.
  * Records the selected date into the **`DATE OUT`** column (Column H) in Google Sheets.
  * Writes a complete audit record with UOM to **`STOCK_OUT_LOG`**.

### 3. 🔄 Restocking 0-Quantity Items
* Go to the **Restock 0-Qty Items** tab.
* Click **Restock & Update** on any out-of-stock medicine.
* Enter the incoming **Brand Name**, **UOM**, **New Expiry Date**, **Updated Price**, and **Quantity**.

### 4. 🧹 Clean Blank Rows & Consolidate Sheet
* If rows were previously cleared and entries appeared far down the sheet (e.g. at Row 75):
  * Click **`🏥 Pharmacy Expiry Monitoring` > `🧹 Clean Blank Rows & Consolidate Sheet`** in Google Sheets.
  * The system will automatically remove all ghost blank rows, shift stranded records up so data starts cleanly on **Row 2**, re-apply dropdowns and number formatting, and trim excess blank rows.
  * All future new medicine batches will automatically append contiguously (Row 2, Row 3, Row 4...) without skipping rows!

### 5. 📋 Audit Log Sheet (`STOCK_OUT_LOG`)
* Automatically maintains an audit trail with:
  * `TIMESTAMP`
  * `DATE OUT`
  * `ITEM CODE`
  * `GENERIC NAME`
  * `BRAND NAME`
  * `UOM`
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
