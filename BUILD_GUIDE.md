# Build guide: Global Health Supply Chain Performance (Power BI)

> **The finished project is `SCMS_Supply_Chain.pbip`** (see "How to open" in README.md). This guide explains how each part of it works, and walks you through rebuilding it yourself in Power BI Desktop. Rebuilding it once is the best way to be able to explain it in an interview.
>
> One difference from the finished file: it parses dates with explicit formats (`Date.FromText(_, [Format = "d-MMM-yy", Culture = "en-US"])`) instead of the locale-based Change Type in step 1b. Both give the same result.

This is a 4-page Power BI report on 10,324 real shipment line items (7,030 shipments, 2006–2015) from the USAID / PEPFAR SCMS programme.

**Time:** about 5–7 hours. **You need:** Power BI Desktop (free, Microsoft Store) and this folder.

**How to check your work:** after each part, compare your numbers with the answer key:

```
python scripts/check_answers.py
```

If a number differs, the bug is almost always one of three things: a missing filter, a wrong relationship, or counting line items where you meant shipments.

---

## 0. Understand the data first (15 min)

Open `data/SCMS_Delivery_History_Data_Dictionary.csv` and read it. Three facts shape the whole project:

| Fact | Why it matters |
|---|---|
| One row = one **line item** (a product on a shipment). One shipment (`ASN/DN #`) can have several line items. | 10,324 line items, but only **7,030 shipments**. |
| **Weight** and **Freight Cost** are stored once per shipment, on the row where `First Line Designation = Yes`. Other rows say `See ASN-93 (ID#:1281)`. | Summing freight over line items is wrong. You need a separate shipments table. |
| The dictionary says **Scheduled Delivery Date is not the client-promised date** and should not be used for on-time performance. | Call your measure **"% Late vs Schedule"**, never "On-Time Delivery", and say why on the report. |

Business questions your report answers:
1. How much was delivered, where, and how has volume changed over time?
2. Which shipment modes, vendors and fulfilment routes deliver later than scheduled?
3. What does freight cost per kg by mode, and how much of the value does it eat?
4. How reliable is the data itself?

---

## 1. Load the data (Power Query)

1. **Home → Get data → Text/CSV** → `data/SCMS_Delivery_History_Dataset.csv`.
2. In the preview window set:
   - **File Origin: 65001: Unicode (UTF-8)**. Check that "Côte d'Ivoire" displays correctly, not "CÃ´te".
   - **Data Type Detection: Do not detect data types.** You will set types yourself; auto-detection guesses wrong on columns that mix numbers and text.
3. Click **Transform Data**. Rename the query to **`Fact_LineItems`**.
4. If the headers are in the first row, use **Home → Use First Row as Headers**. Make sure the first column is called exactly `ID`.

### 1a. Capture text flags *before* converting types

Some date columns contain text such as `N/A - From RDC` or `Date Not Captured`. Converting them to dates turns that text into errors, so record what it said first.

**Add Column → Conditional Column**, name it `PO Date Status`:
- if `PO Sent to Vendor Date` equals `N/A - From RDC` → `Not applicable (RDC)`
- else if `PO Sent to Vendor Date` equals `Date Not Captured` → `Not captured`
- else → `Captured`

### 1b. Set data types

| Columns | Type | How |
|---|---|---|
| `ID`, `Line Item Quantity` | Whole Number | Column header icon |
| `Line Item Value`, `Pack Price`, `Unit Price`, `Line Item Insurance (USD)` | Decimal Number | Right-click → Change Type → **Using Locale…** → Decimal Number, **English (United States)** |
| `Scheduled Delivery Date`, `Delivered to Client Date`, `Delivery Recorded Date`, `PQ First Sent to Client Date`, `PO Sent to Vendor Date` | Date | Right-click → Change Type → **Using Locale…** → Date, **English (United States)** |

The file uses two date formats (`28-Sep-06` and `9/11/14`). The English (United States) locale reads both. A different locale can silently swap day and month.

### 1c. Clean the errors and blanks

1. Select the two PQ/PO date columns → **Transform → Replace Errors** → leave the value empty (null). The errors came from the text flags you already captured in 1a.
2. `Shipment Mode`: **Transform → Replace Values**: `N/A` → `Not recorded`.
3. Leave `Weight (Kilograms)` and `Freight Cost (USD)` as text in this table. You convert them in the shipments table.

✅ **Check:** 10,324 rows (shown at the bottom left of Power Query).

---

## 2. Build the shipments table

1. Right-click `Fact_LineItems` → **Reference**. Rename the new query to **`Fact_Shipments`**.
2. Filter `First Line Designation` = `Yes`.
3. **Choose Columns** and keep: `ASN/DN #`, `Country`, `Vendor`, `Fulfill Via`, `Vendor INCO Term`, `Shipment Mode`, `Managed By`, `PO Sent to Vendor Date`, `Scheduled Delivery Date`, `Delivered to Client Date`, `Weight (Kilograms)`, `Freight Cost (USD)`.

### 2a. Status columns (before converting freight and weight)

**Add Column → Conditional Column**, `Freight Status`:
- if `Freight Cost (USD)` **begins with** `Freight Included` → `Included in commodity cost`
- else if `Freight Cost (USD)` equals `Invoiced Separately` → `Invoiced separately`
- else → `Recorded`

**Conditional Column**, `Weight Status`:
- if `Weight (Kilograms)` equals `Weight Captured Separately` → `Captured separately`
- else → `Recorded`

### 2b. Convert to numbers

1. `Freight Cost (USD)` and `Weight (Kilograms)` → Change Type → Using Locale → Decimal Number, English (United States).
2. **Replace Errors** → null on both.
3. Rename them to `Freight Cost USD` and `Weight kg`.

### 2c. Calculated columns (Add Column → Custom Column)

| Name | Formula | Type |
|---|---|---|
| `Days vs Schedule` | `Duration.Days([Delivered to Client Date] - [Scheduled Delivery Date])` | Whole Number |
| `PO to Delivery Days` | `if [PO Sent to Vendor Date] = null then null else Duration.Days([Delivered to Client Date] - [PO Sent to Vendor Date])` | Whole Number |
| `Delivery Status` | `if [Days vs Schedule] > 0 then "Late" else if [Days vs Schedule] = 0 then "On schedule" else "Early"` | Text |
| `Has Freight and Weight` | `[Freight Cost USD] <> null and [Weight kg] <> null and [Weight kg] > 0` | True/False |

### 2d. Add each shipment's commodity value

1. Right-click `Fact_LineItems` → **Reference** → rename to `Shipment_Value`.
2. **Transform → Group By**: group by `ASN/DN #`, new column `Shipment Value USD` = **Sum** of `Line Item Value`.
3. Right-click `Shipment_Value` → untick **Enable load**. It's only a helper and isn't needed in the report.
4. In `Fact_Shipments`: **Home → Merge Queries** with `Shipment_Value` on `ASN/DN #`, **Left Outer** → expand only `Shipment Value USD`.

✅ **Check:** `Fact_Shipments` has **7,030** rows.

---

## 3. Dimension tables

Each one is made the same way: right-click `Fact_LineItems` → **Reference** → keep only the listed columns → **Home → Remove Rows → Remove Duplicates**.

| Query | Columns | Rows |
|---|---|---|
| `Dim_Country` | `Country` | 43 |
| `Dim_Vendor` | `Vendor` | 73 |
| `Dim_ShipmentMode` | `Shipment Mode` | 5 |
| `Dim_FulfillVia` | `Fulfill Via` | 2 |
| `Dim_Product` | `Item Description`, `Product Group`, `Sub Classification`, `Molecule/Test Type`, `Brand`, `Dosage`, `Dosage Form` | 184 |

Click **Close & Apply**.

### Date table (DAX)

**Modeling → New table**:

```DAX
Dim_Date =
ADDCOLUMNS (
    CALENDAR ( DATE ( 2006, 1, 1 ), DATE ( 2015, 12, 31 ) ),
    "Year", YEAR ( [Date] ),
    "Month Number", MONTH ( [Date] ),
    "Month", FORMAT ( [Date], "MMM" ),
    "Year-Month", FORMAT ( [Date], "YYYY-MM" )
)
```

Then **Table tools → Mark as date table** → `Date`. Set `Month` to **Sort by column → Month Number**.

---

## 4. Data model (Model view)

This is a **star schema with two fact tables at different grains**. Say that sentence in your interview.

```
                 Dim_Date
                    │ (Delivered to Client Date)
     ┌──────────────┼──────────────┐
     ▼                             ▼
Fact_LineItems  (10,324)      Fact_Shipments (7,030)
     ▲   ▲   ▲   ▲                 ▲   ▲   ▲   ▲
     │   │   │   └ Dim_Product      │   │   │   │
     └───┴───┴──── Dim_Country ─────┘   │   │   │
         │   └──── Dim_Vendor ──────────┘   │   │
         └──────── Dim_ShipmentMode ────────┘   │
                   Dim_FulfillVia ──────────────┘ (and to Fact_LineItems)
```

Create these relationships. Drag the dimension column onto the fact column; each is **many-to-one, single direction**:

| From (fact) | To (dimension) |
|---|---|
| `Fact_LineItems[Country]` and `Fact_Shipments[Country]` | `Dim_Country[Country]` |
| `Fact_LineItems[Vendor]` and `Fact_Shipments[Vendor]` | `Dim_Vendor[Vendor]` |
| `Fact_LineItems[Shipment Mode]` and `Fact_Shipments[Shipment Mode]` | `Dim_ShipmentMode[Shipment Mode]` |
| `Fact_LineItems[Fulfill Via]` and `Fact_Shipments[Fulfill Via]` | `Dim_FulfillVia[Fulfill Via]` |
| `Fact_LineItems[Item Description]` | `Dim_Product[Item Description]` |
| `Fact_LineItems[Delivered to Client Date]` and `Fact_Shipments[Delivered to Client Date]` | `Dim_Date[Date]` |

Then **hide** those key columns in both fact tables (right-click → Hide in report view). Always build slicers and axes from the **Dim_** tables; that's what lets one slicer filter both facts.

> Product slicers filter only line-item visuals, because a shipment can contain several products. That's correct behaviour, and a good interview point.

---

## 5. Measures (DAX)

Optional tidy-up: **Home → Enter data** → name the table `_Measures`, and create every measure in it.

```DAX
Line Items = COUNTROWS ( Fact_LineItems )

Commodity Value (USD) = SUM ( Fact_LineItems[Line Item Value] )

Packs Delivered = SUM ( Fact_LineItems[Line Item Quantity] )

Shipments = COUNTROWS ( Fact_Shipments )

Late Shipments =
CALCULATE ( [Shipments], Fact_Shipments[Days vs Schedule] > 0 )

% Late vs Schedule = DIVIDE ( [Late Shipments] + 0, [Shipments] )

Avg Days Late =
CALCULATE (
    AVERAGE ( Fact_Shipments[Days vs Schedule] ),
    Fact_Shipments[Days vs Schedule] > 0
)

Avg PO to Delivery Days = AVERAGE ( Fact_Shipments[PO to Delivery Days] )

Median PO to Delivery Days = MEDIAN ( Fact_Shipments[PO to Delivery Days] )

Freight Cost (USD) = SUM ( Fact_Shipments[Freight Cost USD] )

Freight per kg (USD) =
DIVIDE (
    CALCULATE ( SUM ( Fact_Shipments[Freight Cost USD] ), Fact_Shipments[Has Freight and Weight] = TRUE () ),
    CALCULATE ( SUM ( Fact_Shipments[Weight kg] ), Fact_Shipments[Has Freight and Weight] = TRUE () )
)

Freight % of Value =
DIVIDE (
    [Freight Cost (USD)],
    CALCULATE (
        SUM ( Fact_Shipments[Shipment Value USD] ),
        NOT ISBLANK ( Fact_Shipments[Freight Cost USD] )
    )
)

% Freight Not Recorded =
DIVIDE (
    CALCULATE ( [Shipments], Fact_Shipments[Freight Status] <> "Recorded" ),
    [Shipments]
)
```

Formats: percentages for the `%` measures, currency with 0 decimals for USD, 1 decimal for days.

Why each formula is written this way (you will be asked):
- `AVERAGE` and `MEDIAN` skip blanks, so the PO-to-delivery measures cover only the shipments that have a PO date. That means **Direct Drop shipments only**, because RDC shipments have no PO.
- `Freight per kg` uses only shipments where **both** values exist. Otherwise freight from one set of shipments gets divided by weight from another.
- `Freight % of Value` divides by the value of only those shipments whose freight was recorded, for the same reason.

✅ **Check against the answer key:** Commodity Value $1,627,584,457 · Shipments 7,030 · % Late vs Schedule 11.4% · Avg Days Late 20.0 · Avg PO to Delivery 113.0 days · Freight per kg $3.24 · Freight % of Value 4.5%.

---

## 6. Report pages

Use one colour for "late"/bad (e.g. dark orange) on every page and a neutral colour for everything else. Give every visual a title that states what it shows.

### Page 1 – Overview
- **Cards:** Commodity Value (USD), Shipments, % Late vs Schedule, Avg PO to Delivery Days, Freight Cost (USD).
- **Line and clustered column chart:** X = `Dim_Date[Year]`, columns = Shipments, line = % Late vs Schedule.
- **Bar chart:** Commodity Value by `Dim_Country[Country]`. In the Filters pane, set Country → **Top N**, show the top 10 by Commodity Value.
- **Slicers:** `Dim_Date[Year]`, `Dim_ShipmentMode[Shipment Mode]`, `Dim_FulfillVia[Fulfill Via]`.
- **Text box under the title:** "Late = delivered after the scheduled date. Per the data dictionary, the scheduled date is not the client-promised date."

### Page 2 – Delivery performance
- **Bar chart:** % Late vs Schedule by Shipment Mode, with Shipments in the tooltip.
- **Bar chart:** % Late vs Schedule by Vendor. In the visual-level filter, set **Shipments ≥ 50** so tiny vendors don't dominate.
- **Table:** Fulfill Via | Shipments | % Late vs Schedule | Avg PO to Delivery Days.
- **Matrix heat map:** rows = Country, columns = Year, values = % Late vs Schedule. Turn on Conditional formatting → Background colour, and filter to countries with Shipments ≥ 100.
- **Column chart:** Shipments by `Delivery Status` (Early / On schedule / Late).

### Page 3 – Freight and cost
- **Bar chart:** Freight per kg (USD) by Shipment Mode.
- **Bar chart:** Freight % of Value by Country (top 10 by Freight Cost).
- **Scatter chart:** X = `Weight kg`, Y = `Freight Cost USD`, Values = `ASN/DN #` (one dot per shipment), Legend = Shipment Mode. Set both axes to **logarithmic** scale.
- **Card:** Freight per kg (USD).

### Page 4 – Data quality
- **Bar chart:** Shipments by `Freight Status`.
- **Bar chart:** Shipments by `Weight Status`.
- **Card:** % Freight Not Recorded.
- **Table:** Shipment Mode | Shipments | Avg PO to Delivery Days. Look at the "Not recorded" row: an average of about 2 days from PO to delivery isn't believable. Write down what you think is going on.
- **Text box** listing the cleaning you did: encoding, two date formats, placeholder text → null with a status column, shipments vs line items.

---

## 7. Write your findings (your own words)

Answer these from **your** report, not from the answer key, and keep each to 1–2 sentences with a number:
1. Which shipment mode is cheapest per kg, and how does it compare on lateness? What trade-off would you raise with an operations manager?
2. How does Direct Drop compare with From RDC on % late? Give one possible reason, worded as a hypothesis, not a fact.
3. In which years did lateness rise, and which countries drove it (use the heat map)?
4. Which data-quality issue would you fix first, and why?
5. One recommendation, and the extra data you'd need to be sure of it (for example, client-promised dates).

---

## 8. Publish to GitHub

1. Save as `SCMS_Supply_Chain.pbix` in this folder.
2. **File → Export → Export to PDF** → `SCMS_Supply_Chain.pdf`.
3. Take one screenshot per page into `screenshots/` (`01_overview.png`, and so on).
4. Fill in every `TODO` in `README.md` with your own findings and screenshots.
5. Create a new GitHub repo (e.g. `supply-chain-powerbi`) and push this folder.

The online Power BI service needs a work or school account (Gmail isn't accepted), so the `.pbix`, PDF and screenshots on GitHub are your proof.

---

## 9. Then, and only then, update your resume

Once it's built and on GitHub, you can honestly add:

- **Skills:** Power BI (Power Query, DAX, data modeling)
- **Project:** *Supply Chain Performance Dashboard, Power BI (USAID SCMS data)*
  - Cleaned and modelled 10,324 shipment line items into a star schema with two fact tables (line items and 7,030 shipments) using Power Query.
  - Built DAX measures for % late vs schedule, PO-to-delivery lead time, freight cost per kg and freight as % of value across 43 countries and 73 vendors.
  - *A third bullet with your strongest finding, in your own words.*

Bring the screenshot to me and I'll add it to your HP resume.
