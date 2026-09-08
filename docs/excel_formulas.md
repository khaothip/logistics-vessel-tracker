# Excel transformation logic

This walkthrough uses one consistent synthetic dataset from raw rows to the final chart. In the example workbook, the raw-data worksheet is named **`Sheet1`**.

## Column mapping in `Sheet1`

The original workbook stores the relevant fields in these columns:

| Excel column | Field | Example |
|---|---|---|
| C | ATA LCB / arrival date | 09-Sep-2026 |
| D | Delivery date | blank, 0, or 16-Sep-2026 |
| G | Commodity | Navigation Unit |
| H | Quantity | 400 |
| I | Unit | PCS |
| AH | Shipment key / B/L | SYN-BL-002 |

The downloadable [raw sample](../data/sample_shipments.csv) uses descriptive headers, but represents the same six fields.

## Raw table — `Sheet1`

| Shipment key | ATA LCB | Delivery date | Commodity | Qty | Unit |
|---|---:|---:|---|---:|---|
| SYN-BL-001 | 04-Sep-2026 | 07-Sep-2026 | Inspection Camera | 80 | PCS |
| SYN-BL-001 | 04-Sep-2026 | 07-Sep-2026 | Mounting Bracket | 80 | PCS |
| SYN-BL-002 | 09-Sep-2026 | blank | Navigation Unit | 400 | PCS |
| SYN-BL-002 | 09-Sep-2026 | blank | Camera Module | 150 | PCS |
| SYN-BL-002 | 09-Sep-2026 | blank | Charging Hub | 50 | PCS |
| SYN-BL-002 | 09-Sep-2026 | blank | Battery Pack | 400 | PCS |
| SYN-BL-003 | 12-Sep-2026 | 16-Sep-2026 | Controller Unit | 100 | PCS |
| SYN-BL-003 | 12-Sep-2026 | 16-Sep-2026 | Sensor Module | 50 | PCS |
| SYN-BL-003 | 12-Sep-2026 | 16-Sep-2026 | Propeller Set | 550 | PCS |
| SYN-BL-003 | 12-Sep-2026 | 16-Sep-2026 | Protective Case | 250 | PCS |
| SYN-BL-004 | 15-Sep-2026 | 0 | Survey Kit | 60 | SET |
| SYN-BL-004 | 15-Sep-2026 | 0 | Remote Unit | 120 | PCS |
| SYN-BL-004 | 15-Sep-2026 | 0 | Spare Battery | 240 | PCS |

## Transformation 1A — one row per shipment

```excel
=HSTACK(
    UNIQUE(Sheet1!AH:AH),
    XLOOKUP(UNIQUE(Sheet1!AH:AH),Sheet1!AH:AH,Sheet1!C:C),
    XLOOKUP(UNIQUE(Sheet1!AH:AH),Sheet1!AH:AH,Sheet1!D:D)
)
```

- `UNIQUE` extracts distinct shipment keys.
- The two `XLOOKUP` functions retrieve arrival and delivery dates.
- `HSTACK` creates a shipment-level table.

## Transformation 1B — combine commodity rows

Enter this formula beside the first shipment key and fill down:

```excel
=TEXTJOIN(
    CHAR(10),
    TRUE,
    FILTER(
        Sheet1!G:G&" - "&Sheet1!H:H&" "&Sheet1!I:I,
        Sheet1!AH:AH=A2
    )
)
```

`FILTER` selects all products belonging to the shipment in `A2`. `TEXTJOIN` combines them with line breaks.

### Result after Transformation 1

| Shipment key | ATA LCB | Delivery date | Commodity summary |
|---|---:|---:|---|
| SYN-BL-001 | 04-Sep-2026 | 07-Sep-2026 | Inspection Camera - 80 PCS<br>Mounting Bracket - 80 PCS |
| SYN-BL-002 | 09-Sep-2026 | blank | Navigation Unit - 400 PCS<br>Camera Module - 150 PCS<br>Charging Hub - 50 PCS<br>Battery Pack - 400 PCS |
| SYN-BL-003 | 12-Sep-2026 | 16-Sep-2026 | Controller Unit - 100 PCS<br>Sensor Module - 50 PCS<br>Propeller Set - 550 PCS<br>Protective Case - 250 PCS |
| SYN-BL-004 | 15-Sep-2026 | 0 | Survey Kit - 60 SET<br>Remote Unit - 120 PCS<br>Spare Battery - 240 PCS |

Download the complete [Transformation 1 output](../data/stage1_shipment_summary.csv).

## Transformation 2 — active shipments on a rolling calendar

Assume Transformation 1 contains:

- column B: arrival date;
- column C: delivery date;
- column D: commodity summary; and
- cell G2: one date from the rolling calendar.

```excel
=TEXTJOIN(
    CHAR(10),
    TRUE,
    FILTER(
        D:D,
        (B:B=G2)*((C:C="")+(C:C=0)+(C:C>TODAY())),
        ""
    )
)
```

The arrival date must match `G2`. The delivery date must be blank, zero, or later than today. Multiplication represents **AND**; addition represents **OR**.

For reproducibility, the demonstration below treats **08-Sep-2026 as TODAY**. The actual workbook remains dynamic and uses `TODAY()`.

### Result after Transformation 2

| Offset | Date | Active commodity summary | Shipments |
|---:|---:|---|---:|
| -4 | 04-Sep-2026 | *(blank — delivered on 07-Sep)* | 0 |
| 0 | 08-Sep-2026 TODAY | *(blank)* | 0 |
| +1 | 09-Sep-2026 | Navigation Unit - 400 PCS<br>Camera Module - 150 PCS<br>Charging Hub - 50 PCS<br>Battery Pack - 400 PCS | 1 |
| +4 | 12-Sep-2026 | Controller Unit - 100 PCS<br>Sensor Module - 50 PCS<br>Propeller Set - 550 PCS<br>Protective Case - 250 PCS | 1 |
| +7 | 15-Sep-2026 | Survey Kit - 60 SET<br>Remote Unit - 120 PCS<br>Spare Battery - 240 PCS | 1 |

The full 17-day table is available in [Transformation 2 output](../data/stage2_rolling_timeline.csv).

## Why SYN-BL-001 disappears

Its arrival is inside the rolling window, but its delivery date (07-Sep-2026) is not blank, not zero, and not later than the example TODAY (08-Sep-2026). It is therefore excluded from the active operational view.

## Production improvements

- Replace entire-column references with Excel Table structured references.
- Validate that every shipment key maps to one arrival and delivery date.
- Normalize missing and invalid dates before comparisons.
- Add explicit statuses for in transit, customs clearance, delivered, and delayed.
