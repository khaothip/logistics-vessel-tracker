# Logistics Vessel Tracker

A vessel-arrival monitoring workflow built with **Excel dynamic-array formulas** and **Python in Excel**. It converts item-level shipment records into a shipment-level summary, filters records by operational status, and displays active arrivals on a rolling timeline.

![Vessel-status chart generated from sample data](assets/vessel_status_demo.svg)

## What the workflow does

The same shipment key can appear in several rows because one shipment contains multiple commodities. The workflow restructures these records in three steps:

1. **Shipment aggregation** — retrieve one arrival date and delivery date per shipment, then combine its commodity lines.
2. **Operational date view** — map active shipments to a calendar from `TODAY()-8` to `TODAY()+8`.
3. **Vessel visualization** — clean the transformed range with pandas and plot arrivals with Matplotlib.

```mermaid
flowchart LR
    A["Sheet1: item rows"] --> B["Shipment summary"]
    B --> C["Rolling date table"]
    C --> D["Python cleaning"]
    D --> E["Vessel timeline"]
```

## Source data — `Sheet1`

The table below represents the item-level source. Column references used by the workbook are documented in [Excel transformation logic](docs/excel_formulas.md).

| Shipment ID | ATA LCB | Delivery date | Commodity | Qty | Unit |
|---|---|---|---|---:|---|
| SYN-BL-001 | 04-Sep-2026 | 07-Sep-2026 | Inspection Camera | 80 | PCS |
| SYN-BL-001 | 04-Sep-2026 | 07-Sep-2026 | Mounting Bracket | 80 | PCS |
| SYN-BL-002 | 09-Sep-2026 | — | Navigation Unit | 400 | PCS |
| SYN-BL-002 | 09-Sep-2026 | — | Camera Module | 150 | PCS |
| SYN-BL-002 | 09-Sep-2026 | — | Charging Hub | 50 | PCS |
| SYN-BL-002 | 09-Sep-2026 | — | Battery Pack | 400 | PCS |
| SYN-BL-003 | 12-Sep-2026 | 16-Sep-2026 | Controller Unit | 100 | PCS |
| SYN-BL-003 | 12-Sep-2026 | 16-Sep-2026 | Sensor Module | 50 | PCS |
| SYN-BL-003 | 12-Sep-2026 | 16-Sep-2026 | Propeller Set | 550 | PCS |
| SYN-BL-003 | 12-Sep-2026 | 16-Sep-2026 | Protective Case | 250 | PCS |
| SYN-BL-004 | 15-Sep-2026 | 0 | Survey Kit | 60 | SET |
| SYN-BL-004 | 15-Sep-2026 | 0 | Remote Unit | 120 | PCS |
| SYN-BL-004 | 15-Sep-2026 | 0 | Spare Battery | 240 | PCS |

## Transformation 1 — shipment-level summary

`UNIQUE`, `XLOOKUP`, and `HSTACK` create one row per shipment. `FILTER` and `TEXTJOIN` combine all commodities belonging to that shipment.

| Shipment ID | ATA LCB | Delivery date | Commodity summary |
|---|---|---|---|
| SYN-BL-001 | 04-Sep-2026 | 07-Sep-2026 | Inspection Camera - 80 PCS<br>Mounting Bracket - 80 PCS |
| SYN-BL-002 | 09-Sep-2026 | — | Navigation Unit - 400 PCS<br>Camera Module - 150 PCS<br>Charging Hub - 50 PCS<br>Battery Pack - 400 PCS |
| SYN-BL-003 | 12-Sep-2026 | 16-Sep-2026 | Controller Unit - 100 PCS<br>Sensor Module - 50 PCS<br>Propeller Set - 550 PCS<br>Protective Case - 250 PCS |
| SYN-BL-004 | 15-Sep-2026 | 0 | Survey Kit - 60 SET<br>Remote Unit - 120 PCS<br>Spare Battery - 240 PCS |

## Transformation 2 — rolling active-date table

For a reproducible example, **08-Sep-2026** is treated as `TODAY`. The live workbook uses the actual `TODAY()` value. A shipment remains visible when its delivery date is blank, zero, or later than today.

| Offset | Date | Active commodity summary | Shipments |
|---:|---|---|---:|
| -8 | 31-Aug-2026 | — | 0 |
| -7 | 01-Sep-2026 | — | 0 |
| -6 | 02-Sep-2026 | — | 0 |
| -5 | 03-Sep-2026 | — | 0 |
| -4 | 04-Sep-2026 | — *(delivered 07-Sep)* | 0 |
| -3 | 05-Sep-2026 | — | 0 |
| -2 | 06-Sep-2026 | — | 0 |
| -1 | 07-Sep-2026 | — | 0 |
| 0 | **08-Sep-2026 TODAY** | — | 0 |
| +1 | 09-Sep-2026 | Navigation Unit - 400 PCS<br>Camera Module - 150 PCS<br>Charging Hub - 50 PCS<br>Battery Pack - 400 PCS | 1 |
| +2 | 10-Sep-2026 | — | 0 |
| +3 | 11-Sep-2026 | — | 0 |
| +4 | 12-Sep-2026 | Controller Unit - 100 PCS<br>Sensor Module - 50 PCS<br>Propeller Set - 550 PCS<br>Protective Case - 250 PCS | 1 |
| +5 | 13-Sep-2026 | — | 0 |
| +6 | 14-Sep-2026 | — | 0 |
| +7 | 15-Sep-2026 | Survey Kit - 60 SET<br>Remote Unit - 120 PCS<br>Spare Battery - 240 PCS | 1 |
| +8 | 16-Sep-2026 | — | 0 |

## Python visualization

The code in [`src/vessel_status.py`](src/vessel_status.py):

- reads `G1:I32` through Python in Excel's `xl()` function;
- removes the appended `TODAY` label before parsing dates;
- converts invalid dates to missing values and invalid quantities to zero;
- draws a ship only when the shipment count is positive;
- marks today's position with a red dashed line; and
- places the three footer statuses at fixed relative positions so they remain separated.

## Run in Python in Excel

1. Create the rolling table with the columns `Date`, `Commodity`, and `Qty`.
2. Place it in `G1:I32`, or change `SOURCE_RANGE`.
3. Insert a Python cell and copy the code from `src/vessel_status.py`.
4. Keep `fig` as the final expression so Excel displays the chart.

All identifiers, commodities, quantities, and dates shown above are fictional and demonstrate the processing logic only.
