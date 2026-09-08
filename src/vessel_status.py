"""Vessel-arrival timeline designed for Python in Excel.

Expected Excel range G1:I32: Date | Commodity | Qty.
The chart uses the transformed worksheet table rather than confidential raw data.
"""

import matplotlib.dates as mdates
import matplotlib.patches as patches
import matplotlib.pyplot as plt
import pandas as pd

SOURCE_RANGE = "G1:I32"
SHIP_COLOR = "#4472C4"
ACCENT_COLOR = "#C00000"
GRID_COLOR = "#D9D9D9"


def prepare_data(raw: pd.DataFrame) -> pd.DataFrame:
    """Clean and validate the worksheet table."""
    data = raw.copy()
    data.columns = ["Date", "Commodity", "Qty"]
    data["Date"] = (
        data["Date"].astype(str)
        .str.replace(r"\s*TODAY\s*$", "", regex=True)
    )
    data["Date"] = pd.to_datetime(data["Date"], errors="coerce")
    data["Qty"] = pd.to_numeric(data["Qty"], errors="coerce").fillna(0)
    return data.dropna(subset=["Date"]).sort_values("Date")


def draw_ship(ax, date_value, y, x_size_days=0.7, y_size=0.18):
    """Draw a ship icon at a Matplotlib date coordinate."""
    x = mdates.date2num(date_value)
    hull = patches.Polygon(
        [
            (x - x_size_days, y),
            (x + x_size_days, y),
            (x + x_size_days * 0.7, y - y_size * 0.4),
            (x - x_size_days * 0.7, y - y_size * 0.4),
        ],
        closed=True,
        facecolor=SHIP_COLOR,
        edgecolor="black",
        linewidth=0.8,
        zorder=4,
    )
    ax.add_patch(hull)
    ax.plot([x, x], [y, y + y_size * 1.1], color="black", linewidth=1.2, zorder=4)
    flag = patches.Polygon(
        [
            (x, y + y_size * 1.1),
            (x + x_size_days * 0.9, y + y_size * 0.8),
            (x, y + y_size * 0.5),
        ],
        closed=True,
        facecolor=ACCENT_COLOR,
        zorder=4,
    )
    ax.add_patch(flag)


def build_vessel_chart(data: pd.DataFrame, today=None):
    """Return a vessel-status Matplotlib figure."""
    if data.empty:
        raise ValueError("No valid dates were found in the source range.")

    today = pd.Timestamp.today().normalize() if today is None else pd.Timestamp(today).normalize()
    ship_points = data.loc[data["Qty"] > 0]
    full_range = pd.date_range(data["Date"].min(), data["Date"].max(), freq="D")
    fig, ax = plt.subplots(figsize=(13, 4.5))
    water_level, label_y = 0.75, -0.52

    ax.set_xlim(mdates.date2num(full_range.min()), mdates.date2num(full_range.max()))
    ax.set_ylim(-0.55, 1.05)

    for day in full_range:
        ax.axvline(mdates.date2num(day), color=GRID_COLOR, linewidth=0.6, zorder=1)

    for _, row in ship_points.iterrows():
        draw_ship(ax, row["Date"], water_level)
        ax.annotate(
            str(row["Commodity"]),
            xy=(mdates.date2num(row["Date"]), water_level),
            xytext=(0, -22),
            textcoords="offset points",
            rotation=45,
            ha="right",
            va="top",
            fontsize=7,
            zorder=4,
        )

    today_num = mdates.date2num(today)
    ax.axvline(today_num, color=ACCENT_COLOR, linestyle="--", linewidth=1.2, zorder=3)
    ax.text(today_num, label_y, "TODAY", ha="center", va="bottom",
            fontsize=9, fontweight="bold", color=ACCENT_COLOR)
    ax.text(today_num - 2, label_y, "ARRIVED + DOING CUSTOMS CLEARANCE...",
            ha="right", va="bottom", fontsize=9, fontstyle="italic", color="gray")
    ax.text(today_num + 2, label_y, "ARRIVING...",
            ha="left", va="bottom", fontsize=9, fontstyle="italic", color="gray")

    ax.set_title("VESSEL STATUS", fontweight="bold", fontsize=14, pad=15)
    ax.xaxis.set_ticks_position("top")
    ax.spines["top"].set_position(("data", 1.0))
    ax.set_xticks(mdates.date2num(full_range))
    ax.set_xticklabels(full_range.strftime("%d-%b"), rotation=45, ha="left", fontsize=8)
    ax.set_yticks([])
    for side in ("left", "right", "bottom"):
        ax.spines[side].set_visible(False)

    plt.tight_layout()
    return fig


# Python in Excel entry point
df = prepare_data(xl(SOURCE_RANGE, headers=True))
fig = build_vessel_chart(df)

# Keep fig as the final expression so Excel displays the chart.
fig
