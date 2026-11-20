"""AI Excel Report Generator - upload an Excel file, get an AI-written report with charts."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from ai_summary import (DEFAULT_HOST, DEFAULT_MODEL, AISummaryError, build_context, demo_summary,
                        generate_ai_summary, list_local_models, unverified_numbers)
from analysis import (clean_dataframe, compute_insights, profile_dataframe, read_workbook,
                      suggest_keep_values)
from report_export import build_excel_report, build_word_report

SAMPLE_FILE = Path(__file__).parent / "sample_data" / "nordhaven_sales_2026.xlsx"
LANGUAGES = ["English", "Italiano", "Español", "Deutsch", "Français", "Svenska", "Português"]
NONE = "(none)"
COLORS = ["#1F4E78", "#2E86AB", "#F18F01", "#C73E1D", "#3B8B5A", "#7D5BA6", "#8C8C8C"]

st.set_page_config(page_title="AI Excel Report Generator", page_icon="📊", layout="wide")


@st.cache_data(show_spinner=False)
def load_sheets(data: bytes, name: str) -> dict[str, pd.DataFrame]:
    return read_workbook(data, name)


@st.cache_data(show_spinner=False)
def load_and_clean(data: bytes, name: str):
    sheets = load_sheets(data, name)
    return {sheet: clean_dataframe(df) for sheet, df in sheets.items()}


# ----------------------------------------------------------------------------- sidebar
with st.sidebar:
    st.header("⚙️ Settings")
    ollama_host = st.text_input("Ollama address", value=DEFAULT_HOST)
    local_models = list_local_models(ollama_host)
    if local_models:
        model = st.selectbox("Local model", local_models,
                             index=local_models.index(DEFAULT_MODEL) if DEFAULT_MODEL in local_models else 0,
                             help="Models installed in Ollama. Bigger models write better but are slower.")
        st.caption("🟢 Ollama connected")
    else:
        model = None
        st.warning("Ollama is not running or has no models. The app uses **demo mode**.\n\n"
                   f"Start Ollama, then run `ollama pull {DEFAULT_MODEL}`.")
    language = st.selectbox("Report language", LANGUAGES)
    st.divider()
    st.caption("🔒 100% local: your data never leaves this computer.")

# ----------------------------------------------------------------------------- header
st.title("📊 AI Excel Report Generator")
st.write("Upload an Excel file: the app cleans the data, computes the key figures and "
         "writes a management report with AI. Download it as Excel or Word.")

up_col, sample_col = st.columns([3, 1])
with up_col:
    uploaded = st.file_uploader("Excel or CSV file", type=["xlsx", "xls", "csv"])
with sample_col:
    st.write("")
    st.write("")
    if SAMPLE_FILE.exists() and st.button("Try with sample file", use_container_width=True):
        st.session_state["use_sample"] = True

if uploaded is not None:
    st.session_state["use_sample"] = False
    file_bytes, file_name = uploaded.getvalue(), uploaded.name
elif st.session_state.get("use_sample"):
    file_bytes, file_name = SAMPLE_FILE.read_bytes(), SAMPLE_FILE.name
else:
    st.info("Upload a file or click **Try with sample file** to see a demo.")
    st.stop()

try:
    cleaned = load_and_clean(file_bytes, file_name)
except Exception as exc:  # unreadable or corrupted file
    st.error(f"Could not read the file: {exc}")
    st.stop()

sheet_names = list(cleaned)
default_sheet = max(sheet_names, key=lambda s: len(cleaned[s][0]))
sheet = st.selectbox(f"Main sheet to analyse ({file_name})", sheet_names,
                     index=sheet_names.index(default_sheet))
df, cleaning = cleaned[sheet]
profile = profile_dataframe(df)
other_sheets = {name: cleaned[name][0] for name in sheet_names if name != sheet}

if df.empty or not profile.numeric_cols:
    st.warning("This sheet has no numeric columns to analyse. Choose another sheet.")
    st.stop()

# ----------------------------------------------------------------------------- 1. cleaning
with st.expander("① Data cleaning", expanded=False):
    for line in cleaning.as_lines():
        st.markdown(f"- {line}")
    st.dataframe(df.head(100), use_container_width=True, hide_index=True)
    if other_sheets:
        st.caption("Other sheets used as context for the AI: " + ", ".join(
            f"{n} ({len(d)} rows)" for n, d in other_sheets.items()))

# ----------------------------------------------------------------------------- 2. settings
st.subheader("② What to analyse")
c1, c2, c3 = st.columns(3)
metric = c1.selectbox("Main metric", profile.numeric_cols,
                      index=profile.numeric_cols.index(profile.metric_col) if profile.metric_col else 0)
date_options = [NONE] + profile.date_cols
date_choice = c2.selectbox("Date column", date_options, index=1 if profile.date_cols else 0)
date_col = None if date_choice == NONE else date_choice
dims = c3.multiselect("Break down by", profile.dimension_cols, default=profile.dimension_cols[:5])

c4, c5 = st.columns([1, 2])
status_options = [NONE] + profile.status_cols + [c for c in profile.dimension_cols if df[c].nunique() <= 8]
status_choice = c4.selectbox("Status column (optional)", status_options,
                             index=1 if profile.status_cols else 0,
                             help="E.g. order status. Totals count only the values you keep; "
                                  "the other values (returns, cancellations) are tracked as rates.")
status_col = None if status_choice == NONE else status_choice
keep_values = None
if status_col:
    all_values = sorted(str(v) for v in df[status_col].dropna().unique())
    keep_values = c5.multiselect("Count in totals only", all_values,
                                 default=suggest_keep_values(df, status_col))
    dims = [d for d in dims if d != status_col]

result = compute_insights(df, metric, date_col, dims, status_col, keep_values, cleaning)
facts = result.facts

# ----------------------------------------------------------------------------- KPIs + charts
k1, k2, k3, k4 = st.columns(4)
k1.metric(f"Total {metric}", f"{facts['total']:,.0f}")
k2.metric("Rows analysed", f"{facts['rows_analysed']:,}")
k3.metric("Average per row", f"{facts['average_per_row']:,.2f}")
trend = facts.get("trend", {})
if trend.get("last_month_vs_previous_pct") is not None:
    last_value = facts["monthly_totals"][trend["last_month"]]
    k4.metric(f"Last month ({trend['last_month']})", f"{last_value:,.0f}",
              f"{trend['last_month_vs_previous_pct']:+.1f}% vs previous")
else:
    k4.metric("Columns", len(df.columns))

if result.monthly is not None and len(result.monthly) > 1:
    fig = px.line(result.monthly, x="Month", y=metric, markers=True, title=f"{metric} by month",
                  color_discrete_sequence=COLORS)
    fig.update_layout(height=340, margin=dict(t=50, b=10), yaxis_tickformat=",.0f")
    st.plotly_chart(fig, use_container_width=True)

if dims:
    tabs = st.tabs([f"By {d}" for d in dims])
    for tab, dim in zip(tabs, dims):
        with tab:
            left, right = st.columns(2)
            table = result.by_dimension[dim].head(15)
            bar = px.bar(table.iloc[::-1], x=metric, y=dim, orientation="h", text="Share %",
                         title=f"Total {metric} by {dim}", color_discrete_sequence=COLORS)
            bar.update_traces(texttemplate="%{text}%", textposition="outside")
            bar.update_layout(height=380, margin=dict(t=50, b=10), xaxis_tickformat=",.0f")
            left.plotly_chart(bar, use_container_width=True)
            if dim in result.monthly_by_dimension:
                trend_df = result.monthly_by_dimension[dim].reset_index().melt(
                    id_vars=result.monthly_by_dimension[dim].index.name or "index",
                    var_name=dim, value_name=metric)
                x_col = trend_df.columns[0]
                line = px.line(trend_df, x=x_col, y=metric, color=dim, markers=True,
                               title=f"{metric} by month and {dim}", color_discrete_sequence=COLORS)
                line.update_layout(height=380, margin=dict(t=50, b=10), xaxis_title="Month",
                                   yaxis_tickformat=",.0f")
                right.plotly_chart(line, use_container_width=True)
            else:
                right.dataframe(result.by_dimension[dim], use_container_width=True, hide_index=True)

alerts = facts.get("status_alerts", [])
drops = [(d, c) for d, e in facts.get("by_dimension", {}).items() for c in e.get("sharp_drops", [])]
if alerts or drops or facts.get("outliers"):
    with st.container(border=True):
        st.markdown("**⚠️ Automatic alerts**")
        for d, c in drops:
            st.markdown(f"- {d} **{c['member']}**: {c['change_pct']:+.1f}% in the last 3 months vs before")
        for a in alerts:
            st.markdown(f"- {a['status']} rate for **{a['member']}** ({a['dimension']}): "
                        f"{a['rate_last_2_months_pct']}% in the last 2 months vs {a['rate_before_pct']}% before")
        for o in facts.get("outliers", []):
            st.markdown(f"- Unusually large row: {metric} **{o[metric]:,.0f}** "
                        f"({o['times_the_99th_percentile']}x the 99th percentile of all rows)")

# ----------------------------------------------------------------------------- 3. AI report
st.subheader("③ AI report")
signature = (file_name, sheet, metric, date_col, tuple(dims), status_col,
             tuple(keep_values or []), language, model)
title = f"Report – {Path(file_name).stem}"

if st.button("✨ Generate report", type="primary"):
    context = build_context(file_name, sheet, result, other_sheets)
    if model:
        with st.spinner(f"{model} is analysing your data on this computer… (usually 1-3 minutes)"):
            try:
                summary = generate_ai_summary(context, language, model, ollama_host)
            except AISummaryError as exc:
                st.error(str(exc))
                summary = None
    else:
        summary = demo_summary(result)
    if summary:
        st.session_state["report"] = {"signature": signature, "summary": summary,
                                      "unverified": unverified_numbers(summary, context) if model else []}

report = st.session_state.get("report")
if report:
    if report["signature"] != signature:
        st.info("Settings changed since this report was generated. Click **Generate report** to update it.")
    if report["unverified"]:
        st.warning("🔎 Number check: these values in the report do not match the computed data. "
                   "Check them before sending the report: " + ", ".join(f"`{n}`" for n in report["unverified"]))
    elif "unverified" in report and model:
        st.success("🔎 Number check passed: every figure in the report matches the computed data.")
    with st.container(border=True):
        st.markdown(report["summary"])
    d1, d2, d3 = st.columns(3)
    stem = Path(file_name).stem
    d1.download_button("⬇️ Excel report", build_excel_report(report["summary"], result, title),
                       file_name=f"{stem}_report.xlsx", use_container_width=True,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d2.download_button("⬇️ Word report", build_word_report(report["summary"], result, title),
                       file_name=f"{stem}_report.docx", use_container_width=True,
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    d3.download_button("⬇️ Markdown", report["summary"], file_name=f"{stem}_report.md",
                       use_container_width=True, mime="text/markdown")
