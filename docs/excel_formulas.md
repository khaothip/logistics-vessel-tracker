# Excel transformation logic

The workbook performs two transformations before Python creates the chart. The worksheet and columns below reflect the original design. All examples are documented without company data.

## Stage 1 — one row per shipment

```excel
=HSTACK(
    UNIQUE('Drone Lists'!AH:AH),
    XLOOKUP(UNIQUE('Drone Lists'!AH:AH),'Drone Lists'!AH:AH,'Drone Lists'!C:C),
    XLOOKUP(UNIQUE('Drone Lists'!AH:AH),'Drone Lists'!AH:AH,'Drone Lists'!D:D)
)
```

`UNIQUE` extracts shipment keys. The two `XLOOKUP` calls return the associated arrival and delivery dates. `HSTACK` produces a shipment-level table.

## Stage 1 — combine commodities in each shipment

```excel
=TEXTJOIN(
    CHAR(10),
    TRUE,
    FILTER(
        'Drone Lists'!G:G&" - "&'Drone Lists'!H:H&" "&'Drone Lists'!I:I,
        'Drone Lists'!AH:AH=A2
    )
)
```

`FILTER` selects all item lines belonging to the shipment key in `A2`. `TEXTJOIN` combines product, quantity, and unit using line breaks.

Synthetic example:

```text
Navigation Unit - 400 PCS
Camera Module - 150 PCS
Charging Hub - 50 PCS
```

## Stage 2 — map active shipments to the calendar

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

The arrival date must equal the calendar date in `G2`. The delivery date must be blank, zero, or later than today. Multiplication represents logical **AND**; addition represents logical **OR**.

## Rolling window

The dashboard displays `TODAY()-8` through `TODAY()+8`, keeping the view focused on recent and near-term vessel activity.

## Production improvements

- Replace whole-column references with structured Excel Table references.
- Validate that one shipment key maps to one arrival and delivery date.
- Normalize missing or invalid dates before comparisons.
- Add explicit statuses for in transit, customs clearance, delivered, and delayed.
