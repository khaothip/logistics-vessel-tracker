# Excel transformation logic

This walkthrough uses one consistent synthetic dataset from raw rows to the final chart. In the example workbook, the raw-data worksheet is named **`Sheet1`**.

## Column mapping in `Sheet1`

The original workbook stores the relevant fields in these columns:

| Excel column | Field | Example |
|---|---|---|
| C | ATA LCB / actual arrival date used on the timeline | 09-Sep-2026 |
| D | Delivery date | blank, 0, or 16-Sep-2026 |
| G | Commodity | Navigation Unit |
| H | Quantity | 400 |
| I | Unit | PCS |
| AH | Shipment key / B/L | SYN-BL-002 |

`Shipment key / B/L` identifies one shipment. `ATA LCB` is the actual arrival date used to place that shipment on the vessel timeline.

The complete sample source table is displayed in the [project README](../README.md#source-data--sheet1).

## Raw table — `Sheet1`

| Shipment key | ATA LCB | Delivery date | Commodity | Qty | Unit |
|---|---:|---:|---|---:|---|
| SYN-BL-001 | 04-Sep-2026 | 07-Sep-2026 | Inspection Camera | 80 | PCS |
| SYN-BL-001 | 04-Sep-2026 | 07-Sep-2026 | Mounting Bracket | 80 | PCS |
| SYN-BL-005 | 06-Sep-2026 | blank | Power Distribution Unit | 75 | PCS |
| SYN-BL-005 | 06-Sep-2026 | blank | Communication Module | 75 | PCS |
| SYN-BL-002 | 09-Sep-2026 | 0 | Navigation Unit | 400 | PCS |
| SYN-BL-002 | 09-Sep-2026 | 0 | Camera Module | 150 | PCS |
| SYN-BL-002 | 09-Sep-2026 | 0 | Charging Hub | 50 | PCS |
| SYN-BL-002 | 09-Sep-2026 | 0 | Battery Pack | 400 | PCS |
| SYN-BL-003 | 12-Sep-2026 | 16-Sep-2026 | Controller Unit | 100 | PCS |
| SYN-BL-003 | 12-Sep-2026 | 16-Sep-2026 | Sensor Module | 50 | PCS |
| SYN-BL-003 | 12-Sep-2026 | 16-Sep-2026 | Propeller Set | 550 | PCS |
| SYN-BL-003 | 12-Sep-2026 | 16-Sep-2026 | Protective Case | 250 | PCS |

In the delivery-date field, `blank` means that no value has been entered. A value of `0` may be returned by the source system or an upstream formula when no valid date exists. The workflow treats both as **no confirmed delivery date**.

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

`XLOOKUP` returns the first matching arrival and delivery dates. The workflow therefore assumes that all rows belonging to the same shipment key contain consistent dates; conflicting dates should be validated before this step.

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
| SYN-BL-005 | 06-Sep-2026 | blank | Power Distribution Unit - 75 PCS<br>Communication Module - 75 PCS |
| SYN-BL-002 | 09-Sep-2026 | 0 | Navigation Unit - 400 PCS<br>Camera Module - 150 PCS<br>Charging Hub - 50 PCS<br>Battery Pack - 400 PCS |
| SYN-BL-003 | 12-Sep-2026 | 16-Sep-2026 | Controller Unit - 100 PCS<br>Sensor Module - 50 PCS<br>Propeller Set - 550 PCS<br>Protective Case - 250 PCS |

The complete output table is displayed in the [project README](../README.md#transformation-1--shipment-level-summary).

## Transformation 2 — active shipments on a dynamic calendar

Assume Transformation 1 contains:

- column B: arrival date;
- column C: delivery date;
- column D: commodity summary; and
- cell G2: one date from the dynamic calendar.

### Dynamic calendar range

The calendar ends eight days after today. Its start date is whichever is earlier: eight days before today or the oldest arrival date that is still active.

```excel
=LET(
    ActiveArrivals,FILTER(B2:B1000,(B2:B1000<>"")*((C2:C1000="")+(C2:C1000=0)+(C2:C1000>TODAY())),TODAY()-8),
    StartDate,MIN(TODAY()-8,MIN(ActiveArrivals)),
    SEQUENCE(TODAY()+8-StartDate+1,1,StartDate,1)
)
```

This prevents an undelivered shipment from disappearing merely because its arrival is more than eight days old. As the `TODAY` line moves forward, the ship stays on its original arrival date to the left of the line until its delivery date is completed.

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

The arrival date must match `G2`. The delivery date must be blank, zero, or later than today. Multiplication represents **AND**; addition represents **OR**. Because the comparison is `>TODAY()` rather than `>=TODAY()`, a delivery date equal to today is excluded under the current business rule.

For reproducibility, the demonstration below treats **08-Sep-2026 as TODAY**. The actual workbook remains dynamic and uses `TODAY()`.

### Result after Transformation 2

| Offset | Date | Active commodity summary | Plot flag |
|---:|---:|---|---:|
| -4 | 04-Sep-2026 | *(blank — delivered on 07-Sep)* | 0 |
| -2 | 06-Sep-2026 | Power Distribution Unit - 75 PCS<br>Communication Module - 75 PCS | 1 |
| 0 | 08-Sep-2026 TODAY | *(blank)* | 0 |
| +1 | 09-Sep-2026 | Navigation Unit - 400 PCS<br>Camera Module - 150 PCS<br>Charging Hub - 50 PCS<br>Battery Pack - 400 PCS | 1 |
| +4 | 12-Sep-2026 | Controller Unit - 100 PCS<br>Sensor Module - 50 PCS<br>Propeller Set - 550 PCS<br>Protective Case - 250 PCS | 1 |

The final numeric column is a plot flag rather than the product quantity from `Sheet1`: a positive value tells Python to draw a ship for that date, while `0` suppresses the ship icon.

The full 17-day output is displayed in the [project README](../README.md#transformation-2--dynamic-active-date-table).

## Why the two past arrivals behave differently

`SYN-BL-001` disappears because its delivery date (07-Sep-2026) is not blank, not zero, and not later than the example TODAY (08-Sep-2026). `SYN-BL-005` remains visible on 06-Sep because its delivery date is blank. This demonstrates how an undelivered ship can stay to the left of the moving `TODAY` line.

## Python input range

The Python cell reads `G1:I1000`. The larger range allows the calendar to expand backwards when older shipments remain undelivered. Blank or invalid date rows are removed by `prepare_data()`. If more than 999 data rows are required, `SOURCE_RANGE` can be increased. The script relies on Python in Excel's `xl()` function and will not run unchanged in a standard Python interpreter.

## Production improvements

- Replace entire-column references with Excel Table structured references.
- Validate that every shipment key maps to one arrival and delivery date.
- Normalize missing and invalid dates before comparisons.
- Add explicit statuses for in transit, customs clearance, delivered, and delayed.
