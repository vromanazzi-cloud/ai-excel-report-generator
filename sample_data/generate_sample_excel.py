"""Generate a realistic demo Excel file for the "AI Excel Report Generator" Streamlit app.

Fictional company: Nordhaven Home & Office (furniture and office supplies).
Period: January - October 2026.

The data contains deliberate patterns so the AI summary has something real to find:
- Seasonality: summer dip (Jun-Aug), strong growth from September
- Online channel growing steadily month over month
- East region dropping sharply from July (lost sales rep)
- One unusually large B2B order in August (outlier)
- Return spike for "ErgoChair Pro" in September-October (quality issue)
- Price increase on "StandDesk Lite" from June
And some "messy" data so the cleaning step is visible:
- duplicated rows, missing sales reps, missing regions, extra spaces in text
"""

import random
from pathlib import Path
from datetime import date, timedelta

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SEED = 42
OUTPUT = str(Path(__file__).with_name("nordhaven_sales_2026.xlsx"))

random.seed(SEED)
rng = np.random.default_rng(SEED)

PRODUCTS = [
    # name, category, unit price, unit cost
    ("ErgoChair Pro", "Seating", 349.0, 190.0),
    ("ErgoChair Basic", "Seating", 179.0, 95.0),
    ("StandDesk Lite", "Desks", 429.0, 250.0),
    ("StandDesk Max", "Desks", 799.0, 470.0),
    ("Oak Bookshelf", "Storage", 229.0, 120.0),
    ("Filing Cabinet 3D", "Storage", 159.0, 82.0),
    ("LED Desk Lamp", "Lighting", 59.0, 22.0),
    ("Floor Lamp Arc", "Lighting", 129.0, 58.0),
    ("Monitor Arm Dual", "Accessories", 89.0, 38.0),
    ("Cable Tray Kit", "Accessories", 29.0, 9.0),
    ("Acoustic Panel x4", "Accessories", 199.0, 96.0),
    ("Meeting Table 6p", "Desks", 1190.0, 690.0),
]

REGIONS = {
    "North": ["Erik Lindqvist", "Sara Holm"],
    "South": ["Johan Berg", "Maja Nyström"],
    "East": ["Anders Ek"],
    "West": ["Linnea Sjöberg", "Oskar Dahl"],
    "Online": ["Web Store"],
}

# Monthly demand multiplier (seasonality)
SEASON = {1: 0.95, 2: 0.9, 3: 1.0, 4: 1.05, 5: 1.0, 6: 0.85,
          7: 0.7, 8: 0.8, 9: 1.15, 10: 1.25}

# Popularity of each product (same order as PRODUCTS)
PRODUCT_WEIGHTS = np.array([14, 12, 10, 4, 7, 7, 13, 6, 10, 9, 5, 3], dtype=float)
PRODUCT_WEIGHTS /= PRODUCT_WEIGHTS.sum()

REGION_WEIGHT = {"North": 0.24, "South": 0.22, "East": 0.18, "West": 0.21, "Online": 0.15}


def region_weights(month: int) -> dict:
    w = dict(REGION_WEIGHT)
    w["Online"] *= 1 + 0.08 * (month - 1)      # online grows ~8% per month
    if month >= 7:
        w["East"] *= 0.35                       # East collapses after rep leaves
    total = sum(w.values())
    return {k: v / total for k, v in w.items()}


def unit_price(name: str, base: float, d: date) -> float:
    if name == "StandDesk Lite" and d.month >= 6:
        return round(base * 1.12, 2)            # price increase from June
    return base


def return_probability(name: str, d: date) -> float:
    if name == "ErgoChair Pro" and d.month >= 9:
        return 0.30                             # quality issue
    return 0.04


def build_sales() -> pd.DataFrame:
    rows = []
    order_id = 10001
    day = date(2026, 1, 1)
    end = date(2026, 10, 31)
    while day <= end:
        if day.weekday() < 6:  # closed on Sundays (online still sells)
            base_orders = 16 if day.weekday() < 5 else 6
        else:
            base_orders = 4
        n_orders = rng.poisson(base_orders * SEASON[day.month])
        weights = region_weights(day.month)
        for _ in range(n_orders):
            region = rng.choice(list(weights), p=list(weights.values()))
            if region != "Online" and day.weekday() == 6:
                region = "Online"
            rep = random.choice(REGIONS[region])
            if region == "East" and day.month >= 7:
                rep = "Unassigned"
            customer_type = "Business" if rng.random() < (0.25 if region == "Online" else 0.55) else "Retail"
            name, category, base_price, cost = PRODUCTS[rng.choice(len(PRODUCTS), p=PRODUCT_WEIGHTS)]
            if customer_type == "Retail":
                units = int(rng.integers(1, 3))
            else:
                units = int(rng.integers(2, 6 if base_price > 500 else 10))
            discount = 0.0
            if customer_type == "Business":
                discount = float(rng.choice([0.0, 0.05, 0.10, 0.15], p=[0.3, 0.35, 0.25, 0.1]))
            elif rng.random() < 0.15:
                discount = 0.10
            price = unit_price(name, base_price, day)
            status = "Returned" if rng.random() < return_probability(name, day) else "Completed"
            if rng.random() < 0.015:
                status = "Cancelled"
            rows.append([f"ORD-{order_id}", day, region, rep, customer_type, name, category,
                         units, price, discount, cost, status])
            order_id += 1
        day += timedelta(days=1)

    # Outlier: one very large B2B order in August
    rows.append([f"ORD-{order_id}", date(2026, 8, 19), "West", "Oskar Dahl", "Business",
                 "Meeting Table 6p", "Desks", 60, 1190.0, 0.18, 690.0, "Completed"])

    df = pd.DataFrame(rows, columns=["Order ID", "Order Date", "Region", "Sales Rep", "Customer Type",
                                     "Product", "Category", "Units", "Unit Price (EUR)", "Discount",
                                     "_unit_cost", "Status"])
    df["Revenue (EUR)"] = (df["Units"] * df["Unit Price (EUR)"] * (1 - df["Discount"])).round(2)
    df["Cost (EUR)"] = (df["Units"] * df["_unit_cost"]).round(2)
    df["Gross Profit (EUR)"] = (df["Revenue (EUR)"] - df["Cost (EUR)"]).round(2)
    df = df.drop(columns="_unit_cost").sort_values(["Order Date", "Order ID"]).reset_index(drop=True)

    # --- Messy data, as in real exports ---
    idx = rng.choice(df.index, size=25, replace=False)
    df.loc[idx, "Sales Rep"] = None                       # missing reps
    idx = rng.choice(df.index, size=8, replace=False)
    df.loc[idx, "Region"] = None                          # missing regions
    idx = rng.choice(df.index, size=30, replace=False)
    df.loc[idx, "Product"] = df.loc[idx, "Product"] + " "  # trailing spaces
    dupes = df.sample(12, random_state=SEED)              # duplicated rows
    df = pd.concat([df, dupes]).sort_values(["Order Date", "Order ID"]).reset_index(drop=True)
    return df


def build_targets(sales: pd.DataFrame) -> pd.DataFrame:
    """Targets = Q1 average per region + 3% planned growth per month (set at the start of the year)."""
    clean = sales.dropna(subset=["Region"])
    clean = clean[clean["Status"] == "Completed"]
    q1 = clean[pd.to_datetime(clean["Order Date"]).dt.month <= 3]
    base = q1.groupby("Region")["Revenue (EUR)"].sum() / 3
    months = pd.period_range("2026-01", "2026-12", freq="M")
    rows = []
    for m in months:
        for region in REGION_WEIGHT:
            target = base[region] * (1.0 + 0.03 * (m.month - 1)) * SEASON.get(m.month, 1.3)
            rows.append([m.strftime("%Y-%m"), region, round(target, -2)])
    return pd.DataFrame(rows, columns=["Month", "Region", "Revenue Target (EUR)"])


def build_products() -> pd.DataFrame:
    rows = []
    for name, category, price, cost in PRODUCTS:
        stock = int(rng.integers(5, 260))
        reorder = 40 if price < 300 else 15
        rows.append([name, category, price, cost, stock, reorder,
                     "Yes" if stock < reorder else "No"])
    return pd.DataFrame(rows, columns=["Product", "Category", "List Price (EUR)", "Unit Cost (EUR)",
                                       "Stock Units", "Reorder Level", "Below Reorder Level"])


def build_marketing() -> pd.DataFrame:
    rows = []
    for month in range(1, 11):
        for channel, base in [("Google Ads", 4200), ("Meta Ads", 2600), ("Email", 600), ("Trade Fairs", 0)]:
            spend = base * float(rng.uniform(0.85, 1.15))
            if channel == "Google Ads" and month >= 6:
                spend *= 1.6                      # budget shifted to online
            if channel == "Trade Fairs" and month in (3, 9):
                spend = 9500
            leads = int(spend / float(rng.uniform(18, 35))) if spend else 0
            rows.append([f"2026-{month:02d}", channel, round(spend, 2), leads])
    return pd.DataFrame(rows, columns=["Month", "Channel", "Spend (EUR)", "Leads"])


def style_workbook(path: str) -> None:
    wb = load_workbook(path)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")
    money_cols = ("EUR",)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 30
        for col_idx, cell in enumerate(ws[1], start=1):
            letter = get_column_letter(col_idx)
            header = str(cell.value)
            width = max(len(header), *(len(str(c.value)) for c in ws[letter][1:200] if c.value is not None)) + 3
            ws.column_dimensions[letter].width = min(width, 28)
            for c in ws[letter][1:]:
                if any(t in header for t in money_cols):
                    c.number_format = "#,##0.00"
                elif header == "Discount":
                    c.number_format = "0%"
                elif header == "Order Date":
                    c.number_format = "yyyy-mm-dd"
    wb.save(path)


def main() -> None:
    sales = build_sales()
    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        sales.to_excel(writer, sheet_name="Sales", index=False)
        build_targets(sales).to_excel(writer, sheet_name="Targets", index=False)
        build_products().to_excel(writer, sheet_name="Products", index=False)
        build_marketing().to_excel(writer, sheet_name="Marketing", index=False)
    style_workbook(OUTPUT)
    print(f"Saved {OUTPUT} with {len(sales)} sales rows")


if __name__ == "__main__":
    main()
