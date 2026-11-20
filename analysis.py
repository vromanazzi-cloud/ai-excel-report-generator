"""Loading, cleaning, profiling and insight extraction for Excel reports.

All numbers the AI writes about are computed here with pandas: the language model only
explains facts it receives, it never does the arithmetic itself.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

METRIC_HINTS = ("revenue", "sales", "turnover", "amount", "total", "income", "profit",
                "fatturato", "ricavi", "vendite", "importo", "totale", "value", "price", "cost")
STATUS_NAME = re.compile(r"status|stato|state|esito|outcome|result", re.I)
POSITIVE_VALUE = re.compile(
    r"complet|paid|closed|won|deliver|success|shipped|done|evas|pagat|consegn|chius|conclus", re.I)
ID_NAME = re.compile(r"(^|[\s_\-])(id|code|codice|nr|no|number|numero)([\s_\-]|$)", re.I)
DATE_NAME = re.compile(r"date|data|day|giorno|time|month|mese|period|periodo", re.I)

MAX_DIMENSION_MEMBERS = 40


# --------------------------------------------------------------------------- loading
def read_workbook(data: bytes, file_name: str) -> dict[str, pd.DataFrame]:
    """Read every sheet of an Excel file (or a single CSV) into DataFrames."""
    buffer = io.BytesIO(data)
    if file_name.lower().endswith(".csv"):
        return {"Data": pd.read_csv(buffer)}
    return pd.read_excel(buffer, sheet_name=None)


# --------------------------------------------------------------------------- cleaning
@dataclass
class CleaningReport:
    rows_before: int
    rows_after: int
    empty_rows_removed: int
    duplicates_removed: int
    text_cells_trimmed: int
    dates_parsed: list[str]
    missing_values: dict[str, int]

    def as_lines(self) -> list[str]:
        lines = [f"{self.rows_before:,} rows read, {self.rows_after:,} rows kept after cleaning"]
        if self.duplicates_removed:
            lines.append(f"{self.duplicates_removed} duplicated rows removed")
        if self.empty_rows_removed:
            lines.append(f"{self.empty_rows_removed} empty rows removed")
        if self.text_cells_trimmed:
            lines.append(f"{self.text_cells_trimmed} text cells with extra spaces fixed")
        if self.dates_parsed:
            lines.append("Date columns recognised: " + ", ".join(self.dates_parsed))
        for col, n in self.missing_values.items():
            lines.append(f"Column '{col}': {n} missing values")
        return lines

    def as_dict(self) -> dict:
        return {
            "rows_read": self.rows_before,
            "rows_after_cleaning": self.rows_after,
            "duplicates_removed": self.duplicates_removed,
            "empty_rows_removed": self.empty_rows_removed,
            "text_cells_trimmed": self.text_cells_trimmed,
            "missing_values_by_column": self.missing_values,
        }


def clean_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame, CleaningReport]:
    df = df.copy()
    rows_before = len(df)
    df = df.dropna(how="all").dropna(axis=1, how="all")
    empty_rows = rows_before - len(df)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.startswith("Unnamed:")]

    trimmed = 0
    for col in df.columns:
        if df[col].dtype == object:
            original = df[col]
            fixed = original.map(lambda v: re.sub(r"\s+", " ", v).strip() if isinstance(v, str) else v)
            trimmed += int(((fixed != original) & original.notna()).sum())
            df[col] = fixed.replace("", np.nan)

    dates = []
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            dates.append(col)
        elif df[col].dtype == object and DATE_NAME.search(col):
            parsed = pd.to_datetime(df[col], errors="coerce")
            if parsed.notna().mean() > 0.8:
                df[col] = parsed
                dates.append(col)

    before = len(df)
    df = df.drop_duplicates()
    duplicates = before - len(df)
    missing = {c: int(n) for c, n in df.isna().sum().items() if n > 0}
    report = CleaningReport(rows_before, len(df), empty_rows, duplicates, trimmed, dates, missing)
    return df.reset_index(drop=True), report


# --------------------------------------------------------------------------- profiling
@dataclass
class Profile:
    date_cols: list[str]
    numeric_cols: list[str]
    dimension_cols: list[str]
    status_cols: list[str]
    metric_col: str | None


def profile_dataframe(df: pd.DataFrame) -> Profile:
    date_cols = [c for c in df.columns if pd.api.types.is_datetime64_any_dtype(df[c])]
    numeric = [c for c in df.select_dtypes("number").columns if not ID_NAME.search(c)]
    dims, status = [], []
    for col in df.columns:
        if df[col].dtype != object or ID_NAME.search(col):
            continue
        n = df[col].nunique()
        if n < 2 or n > MAX_DIMENSION_MEMBERS:
            continue
        values = df[col].dropna().astype(str).unique()
        if n <= 6 and (STATUS_NAME.search(col) or any(POSITIVE_VALUE.search(v) for v in values)):
            status.append(col)
        else:
            dims.append(col)
    return Profile(date_cols, numeric, dims, status, pick_metric(df, numeric))


def pick_metric(df: pd.DataFrame, numeric: list[str]) -> str | None:
    for hint in METRIC_HINTS:
        for col in numeric:
            if hint in col.lower():
                return col
    if not numeric:
        return None
    return max(numeric, key=lambda c: df[c].abs().sum())


def suggest_keep_values(df: pd.DataFrame, status_col: str) -> list[str]:
    """Rows to count in the totals, e.g. 'Completed' orders only."""
    values = [str(v) for v in df[status_col].dropna().unique()]
    positive = [v for v in values if POSITIVE_VALUE.search(v)]
    return positive or values


# --------------------------------------------------------------------------- insights
@dataclass
class AnalysisResult:
    facts: dict
    filtered: pd.DataFrame
    monthly: pd.DataFrame | None
    by_dimension: dict[str, pd.DataFrame] = field(default_factory=dict)
    monthly_by_dimension: dict[str, pd.DataFrame] = field(default_factory=dict)


def _pct(new: float, old: float) -> float | None:
    if old is None or old == 0 or pd.isna(old):
        return None
    return round((new - old) / abs(old) * 100, 1)


def _num(v) -> float:
    return round(float(v), 2)


def _jsonable(v):
    if isinstance(v, pd.Timestamp):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (np.integer, np.floating)):
        return None if pd.isna(v) else round(float(v), 2)
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    return v


def compute_insights(df: pd.DataFrame, metric: str, date_col: str | None, dims: list[str],
                     status_col: str | None = None, keep_values: list[str] | None = None,
                     cleaning: CleaningReport | None = None) -> AnalysisResult:
    full = df
    data = df
    filter_text = "all rows"
    if status_col and keep_values:
        data = df[df[status_col].astype(str).isin(keep_values)]
        filter_text = f"only rows where {status_col} is {', '.join(keep_values)}"
    data = data[data[metric].notna()]

    facts: dict = {
        "metric": metric,
        "rows_in_file": int(len(full)),
        "rows_analysed": int(len(data)),
        "filter_applied": filter_text,
        "total": _num(data[metric].sum()),
        "average_per_row": _num(data[metric].mean()) if len(data) else 0,
    }
    if cleaning:
        facts["data_quality"] = cleaning.as_dict()

    monthly = None
    month_index = None
    if date_col:
        data = data[data[date_col].notna()]
        month_index = data[date_col].dt.to_period("M")
        facts["period"] = {"from": data[date_col].min().strftime("%Y-%m-%d"),
                           "to": data[date_col].max().strftime("%Y-%m-%d")}
        grouped = data.groupby(month_index)[metric].agg(["sum", "count"]).sort_index()
        monthly = pd.DataFrame({"Month": grouped.index.astype(str),
                                metric: grouped["sum"].round(2).values,
                                "Rows": grouped["count"].values})
        facts["monthly_totals"] = {m: _num(v) for m, v in zip(monthly["Month"], monthly[metric])}
        if len(monthly) >= 2:
            values = monthly[metric]
            trend = {
                "best_month": monthly.loc[values.idxmax(), "Month"],
                "worst_month": monthly.loc[values.idxmin(), "Month"],
                "last_month": monthly["Month"].iloc[-1],
                "last_month_vs_previous_pct": _pct(values.iloc[-1], values.iloc[-2]),
            }
            if len(monthly) >= 6:
                trend["avg_last_3_months_vs_first_3_months_pct"] = _pct(
                    values.tail(3).mean(), values.head(3).mean())
            facts["trend"] = trend

    # --- breakdown by dimension
    by_dimension, monthly_by_dimension, dim_facts = {}, {}, {}
    for dim in dims:
        sub = data[data[dim].notna()]
        totals = sub.groupby(dim)[metric].sum().sort_values(ascending=False)
        grand = totals.sum()
        table = pd.DataFrame({dim: totals.index.astype(str), metric: totals.round(2).values,
                              "Share %": (totals / grand * 100).round(1).values if grand else 0})
        by_dimension[dim] = table
        entry = {"members": int(len(totals)),
                 "ranking": [{"member": r[dim], "total": _num(r[metric]), "share_pct": _num(r["Share %"])}
                             for _, r in table.head(10).iterrows()]}
        missing = int(data[dim].isna().sum())
        if missing:
            entry["rows_without_value"] = missing

        if month_index is not None and month_index.nunique() >= 4:
            pivot = sub.pivot_table(index=month_index[sub.index], columns=dim, values=metric,
                                    aggfunc="sum", fill_value=0).sort_index()
            pivot.index = pivot.index.astype(str)
            if pivot.shape[1] <= 15:
                monthly_by_dimension[dim] = pivot.round(2)
            recent, earlier = pivot.tail(3).mean(), pivot.iloc[:-3].mean()
            changes = []
            for member in pivot.columns:
                if earlier[member] < 0.02 * earlier.sum():
                    continue
                changes.append({"member": str(member),
                                "avg_monthly_before_last_3_months": _num(earlier[member]),
                                "avg_monthly_last_3_months": _num(recent[member]),
                                "change_pct": _pct(recent[member], earlier[member])})
            changes.sort(key=lambda c: c["change_pct"] if c["change_pct"] is not None else 0)
            entry["sharp_drops"] = [c for c in changes if c["change_pct"] is not None and c["change_pct"] <= -35]
            entry["strong_growth"] = [c for c in changes if c["change_pct"] is not None and c["change_pct"] >= 35]
        dim_facts[dim] = entry
    facts["by_dimension"] = dim_facts

    # --- outliers: single rows far above normal
    if len(data) >= 50:
        values = data[metric]
        p99 = values.quantile(0.99)
        facts["distribution"] = {"median": _num(values.median()), "p99": _num(p99), "max": _num(values.max())}
        outliers = []
        if p99 > 0:
            for _, row in data.nlargest(5, metric).iterrows():
                if row[metric] >= 3 * p99:
                    record = {k: _jsonable(v) for k, v in row.items()}
                    record["times_the_99th_percentile"] = round(row[metric] / p99, 1)
                    outliers.append(record)
        facts["outliers"] = outliers

    # --- status rates (returns, cancellations...) computed on ALL rows
    if status_col:
        facts.update(_status_alerts(full, status_col, keep_values or [], date_col, dims))

    return AnalysisResult(facts, data, monthly, by_dimension, monthly_by_dimension)


def _status_alerts(full: pd.DataFrame, status_col: str, keep_values: list[str],
                   date_col: str | None, dims: list[str]) -> dict:
    s = full[full[status_col].notna()].copy()
    s[status_col] = s[status_col].astype(str)
    shares = s[status_col].value_counts(normalize=True)
    out: dict = {"status_breakdown_pct": {k: round(v * 100, 1) for k, v in shares.items()}}
    if not date_col:
        return out
    s = s[s[date_col].notna()]
    months = s[date_col].dt.to_period("M")
    last_two = sorted(months.unique())[-2:]
    recent = months.isin(last_two)
    rates, alerts = {}, []
    for value, share in shares.items():
        if share > 0.3 or value in keep_values:
            continue
        is_value = s[status_col] == value
        rates[value] = {str(m): round(r * 100, 1) for m, r in is_value.groupby(months).mean().items()}
        for dim in dims:
            frame = pd.DataFrame({"member": s[dim], "hit": is_value, "recent": recent})
            for member, grp in frame.groupby("member"):
                rec, ear = grp[grp["recent"]], grp[~grp["recent"]]
                if len(rec) < 15 or len(ear) < 30:
                    continue
                r_rate, e_rate = rec["hit"].mean(), ear["hit"].mean()
                if r_rate >= 0.10 and r_rate >= max(2 * e_rate, e_rate + 0.08):
                    alerts.append({"status": value, "dimension": dim, "member": str(member),
                                   "rate_last_2_months_pct": round(r_rate * 100, 1),
                                   "rate_before_pct": round(e_rate * 100, 1),
                                   "rows_last_2_months": int(len(rec))})
    out["status_rate_by_month_pct"] = rates
    out["status_alerts"] = sorted(alerts, key=lambda a: -a["rate_last_2_months_pct"])
    return out
