# Logistics Vessel Tracker

An operational analytics project that transforms item-level shipment records into a dynamic vessel-arrival dashboard using **Python in Excel**, **pandas**, **Matplotlib**, and Excel dynamic-array formulas.

> This portfolio project is based on a real logistics workflow. Every record in this repository is synthetic. No employer, customer, invoice, bill-of-lading, product, or shipment data is included.

## Business problem

Shipment information is commonly maintained at item level: one bill of lading can occupy several rows containing different products and quantities. Reading the raw table does not immediately answer:

- Which vessels are arriving soon?
- What commodities are included in each shipment?
- Which shipments remain active because delivery is incomplete?
- Which arrivals require attention today?

The project converts repeated operational records into a shipment-level summary and then plots active arrivals on a rolling timeline.

## Data pipeline

```mermaid
flowchart LR
    A["Raw item-level records"] --> B["Aggregate by shipment key"]
    B --> C["Apply delivery-date rules"]
    C --> D["Python in Excel"]
    D --> E["Vessel-status timeline"]
```

### 1. Raw data

The source worksheet contains one row per commodity or invoice line, including a shipment key, arrival date, delivery date, commodity, quantity, and unit.

### 2. Shipment-level transformation

Excel formulas use `UNIQUE`, `XLOOKUP`, and `HSTACK` to identify distinct shipments and retrieve their dates. `FILTER` and `TEXTJOIN` consolidate several commodity lines into one readable shipment summary.

### 3. Rolling operational window

A calendar centered on `TODAY()` covers the previous eight days and next eight days. Conditional filtering retains shipments whose delivery date is blank, zero, or later than today.

### 4. Python visualization

Python reads the transformed Excel range with `xl()`, validates dates and quantities, and draws:

- a daily arrival timeline;
- a ship icon for every scheduled shipment;
- combined commodity annotations;
- a red marker for today; and
- operational labels for arrived/customs-clearance and arriving shipments.

## Repository structure

```text
.
├── data/
│   └── sample_shipments.csv
├── docs/
│   └── excel_formulas.md
├── src/
│   └── vessel_status.py
├── .gitignore
├── requirements.txt
└── README.md
```

- [Synthetic shipment data](data/sample_shipments.csv)
- [Excel transformation formulas](docs/excel_formulas.md)
- [Python in Excel visualization](src/vessel_status.py)

## Technologies

- Python in Excel
- pandas
- Matplotlib
- Microsoft Excel dynamic arrays
- `UNIQUE`, `XLOOKUP`, `FILTER`, `TEXTJOIN`, and `HSTACK`

## Run in Python in Excel

1. Prepare a transformed range in cells `G1:I32` with the columns `Date`, `Commodity`, and `Qty`.
2. Insert a Python cell in Excel.
3. Copy the code from `src/vessel_status.py`.
4. Keep `fig` as the final expression so Excel returns the chart.

Change `SOURCE_RANGE` if the transformed table is stored elsewhere.

## Key analytical decisions

| Decision | Implementation |
|---|---|
| Analytical grain | Consolidate repeated item rows by shipment key |
| Date validation | Coerce invalid values to missing and exclude them |
| Quantity validation | Convert nonnumeric values to zero |
| Active shipment rule | Include blank, zero, or future delivery dates |
| Time reference | Anchor the dashboard dynamically to today |
| Confidentiality | Publish only synthetic records and generic product names |

## Skills demonstrated

- Translating a logistics process into data-transformation rules
- Cleaning and validating operational data
- Aggregating row-level records into business-level entities
- Combining spreadsheet functions with Python
- Building a time-based operational visualization
- Applying domain knowledge from international trade and logistics
- Protecting confidential business information through synthetic data

## Limitations and future improvements

- Replace full-column formulas with structured Excel Table references.
- Validate duplicated or inconsistent shipment keys before aggregation.
- Add explicit categories for in transit, customs clearance, delivered, and delayed.
- Separate extraction, transformation, and visualization for automated testing.
- Add unit tests and an interactive dashboard version.

## Author

**Khaothip**

Background in International Trade and Business Logistics, developing data and analytics skills for graduate study in Data Science.
