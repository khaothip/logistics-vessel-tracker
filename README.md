# Logistics Vessel Tracker

A vessel-arrival monitoring workflow built with **Excel dynamic-array formulas** and **Python in Excel**. It converts item-level shipment records into a shipment-level summary, filters shipments by operational status, and displays active arrivals on a rolling timeline.

![Vessel-status chart generated from sample data](assets/vessel_status_demo.svg)

## What the workflow does

A shipment can contain several commodities, so the same bill of lading appears across multiple source rows. Reading that table directly makes it difficult to see which vessels are approaching and what each shipment contains.

The workflow restructures the data in two stages:

1. **Shipment aggregation** — repeated item rows are grouped by shipment key. Arrival date and delivery date are retrieved, while commodity, quantity, and unit are combined into a multiline summary.
2. **Operational date view** — shipments are mapped to a rolling calendar from `TODAY()-8` to `TODAY()+8`. Records already delivered are excluded, while undelivered and upcoming deliveries remain visible.
3. **Vessel timeline** — Python reads the transformed range, cleans the date and quantity fields, and plots each active shipment on its arrival date.

## Data flow

```mermaid
flowchart LR
    A["Sheet1: item-level records"] --> B["Shipment-level summary"]
    B --> C["Rolling active-date table"]
    C --> D["Python data cleaning"]
    D --> E["Vessel timeline"]
```

## Processing stages

| Stage | Input | Processing | Result |
|---|---|---|---|
| Source data | One row per commodity line | Read arrival, delivery, item, quantity, unit, and shipment key | Repeated shipment records |
| Transformation 1 | Item-level records | `UNIQUE`, `XLOOKUP`, `HSTACK`, `FILTER`, and `TEXTJOIN` | One row per shipment |
| Transformation 2 | Shipment-level records | Match arrival dates and evaluate delivery conditions | One row per reporting date |
| Visualization | Date, commodity summary, and shipment count | pandas type cleaning and Matplotlib drawing | Vessel-status timeline |

The formulas and example tables for every stage are documented in [Excel transformation logic](docs/excel_formulas.md).

## Example data

The following files use the same set of sample shipments, allowing each stage to be followed from source to chart:

- [Item-level source table](data/sample_shipments.csv)
- [Shipment-level transformation](data/stage1_shipment_summary.csv)
- [Rolling date transformation](data/stage2_rolling_timeline.csv)

The example fixes the reporting date at **08-Sep-2026** so the tables and chart remain reproducible. In the working Excel file, the date window updates automatically through `TODAY()`.

All shipment identifiers, commodities, quantities, and dates in these files are fictional and are included only to demonstrate the processing logic.

## Python visualization

The code in [`src/vessel_status.py`](src/vessel_status.py):

- reads `G1:I32` through the Python in Excel `xl()` function;
- removes the text label `TODAY` before parsing dates;
- converts invalid dates to missing values;
- converts invalid quantities to zero;
- plots only rows with a positive shipment count;
- draws the current-date marker and status areas; and
- returns the Matplotlib figure to the Excel cell.

## Run in Python in Excel

1. Create the rolling date table with the columns `Date`, `Commodity`, and `Qty`.
2. Place it in `G1:I32`, or change `SOURCE_RANGE` in the Python code.
3. Insert a Python cell in Excel.
4. Copy the code from `src/vessel_status.py`.
5. Keep `fig` as the final expression so Excel displays the chart.

## Repository structure

```text
.
├── assets/
│   └── vessel_status_demo.svg
├── data/
│   ├── sample_shipments.csv
│   ├── stage1_shipment_summary.csv
│   └── stage2_rolling_timeline.csv
├── docs/
│   └── excel_formulas.md
├── src/
│   └── vessel_status.py
├── .gitignore
├── requirements.txt
└── README.md
```

## Design considerations

- Full-column Excel references make the formulas easy to read but can be replaced with structured Table references for larger workbooks.
- Each shipment key is expected to map to one arrival date and one delivery date.
- Blank, zero, and future delivery dates are treated as active according to the operational rule.
- The visualization is intentionally generated from the transformed table, keeping data preparation separate from chart construction.
