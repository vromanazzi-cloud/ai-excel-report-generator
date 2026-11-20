"""Export the report as Excel (with native charts), Word (with chart images) or Markdown."""

from __future__ import annotations

import io
import re
from datetime import date

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402
from docx import Document  # noqa: E402
from docx.shared import Inches, Pt, RGBColor  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from openpyxl.chart import BarChart, LineChart, Reference  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

from analysis import AnalysisResult  # noqa: E402

ACCENT = "1F4E78"
MAX_DIMENSION_SHEETS = 6


def _plain(text: str) -> str:
    return re.sub(r"\*\*|__|`", "", text)


def _sheet_title(name: str) -> str:
    return re.sub(r"[\[\]\*\?/\\:]", "", name)[:31]


def _style_header(ws, row: int, n_cols: int) -> None:
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=ACCENT)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _write_table(ws, df: pd.DataFrame) -> None:
    ws.append(list(map(str, df.columns)))
    for row in df.itertuples(index=False):
        ws.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in row])
    _style_header(ws, 1, len(df.columns))
    for i, col in enumerate(df.columns, start=1):
        sample = [len(str(col))] + [len(str(v)) for v in df[col].head(100)]
        ws.column_dimensions[get_column_letter(i)].width = min(max(sample) + 3, 40)
        if pd.api.types.is_float_dtype(df[col]):
            for cell in ws[get_column_letter(i)][1:]:
                cell.number_format = "#,##0.00"
        elif pd.api.types.is_datetime64_any_dtype(df[col]):
            for cell in ws[get_column_letter(i)][1:]:
                cell.number_format = "yyyy-mm-dd"
    ws.freeze_panes = "A2"


def build_excel_report(summary_md: str, result: AnalysisResult, title: str) -> bytes:
    wb = Workbook()
    metric = result.facts["metric"]

    ws = wb.active
    ws.title = "AI Summary"
    ws.column_dimensions["A"].width = 120
    ws.append([title])
    ws["A1"].font = Font(bold=True, size=16, color=ACCENT)
    ws.append([f"Generated on {date.today():%Y-%m-%d}"])
    ws.append([])
    for line in summary_md.splitlines():
        stripped = line.strip()
        if not stripped:
            ws.append([])
            continue
        heading = stripped.startswith("#")
        ws.append([_plain(stripped.lstrip("#").strip()) if heading else _plain(stripped)])
        cell = ws.cell(row=ws.max_row, column=1)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        if heading:
            cell.font = Font(bold=True, size=13, color=ACCENT)

    ws = wb.create_sheet("Key Figures")
    rows = [("Metric", metric), (f"Total {metric}", result.facts["total"]),
            ("Average per row", result.facts["average_per_row"]),
            ("Rows analysed", result.facts["rows_analysed"]), ("Filter", result.facts["filter_applied"])]
    if "period" in result.facts:
        rows.append(("Period", f"{result.facts['period']['from']} to {result.facts['period']['to']}"))
    ws.append(["Indicator", "Value"])
    for r in rows:
        ws.append(list(r))
    _style_header(ws, 1, 2)
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 30, 45
    for cell in ws["B"][1:]:
        if isinstance(cell.value, float):
            cell.number_format = "#,##0.00"

    if result.monthly is not None and len(result.monthly):
        ws = wb.create_sheet("Monthly Trend")
        _write_table(ws, result.monthly)
        chart = LineChart()
        chart.title, chart.y_axis.title, chart.height, chart.width = f"{metric} by month", metric, 8, 18
        n = len(result.monthly) + 1
        chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=n), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=n))
        ws.add_chart(chart, "E2")

    for dim, table in list(result.by_dimension.items())[:MAX_DIMENSION_SHEETS]:
        ws = wb.create_sheet(_sheet_title(f"By {dim}"))
        _write_table(ws, table)
        chart = BarChart()
        chart.type, chart.title, chart.height, chart.width = "bar", f"{metric} by {dim}", 9, 18
        n = min(len(table), 15) + 1
        chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=n), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=n))
        chart.y_axis.scaling.orientation = "minMax"
        chart.x_axis.scaling.orientation = "maxMin"
        chart.legend = None
        ws.add_chart(chart, "E2")

    ws = wb.create_sheet("Cleaned Data")
    _write_table(ws, result.filtered)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _chart_png(fig) -> io.BytesIO:
    buf = io.BytesIO()
    fig.tight_layout()
    fig.savefig(buf, format="png", dpi=160)
    plt.close(fig)
    buf.seek(0)
    return buf


def _add_markdown(doc: Document, md: str) -> None:
    for line in md.splitlines():
        s = line.strip()
        if not s:
            continue
        if s.startswith("#"):
            level = min(len(s) - len(s.lstrip("#")), 3)
            doc.add_heading(_plain(s.lstrip("#").strip()), level=level)
            continue
        style = None
        if re.match(r"^[-*] ", s):
            style, s = "List Bullet", s[2:]
        elif re.match(r"^\d+[.)] ", s):
            style, s = "List Number", re.sub(r"^\d+[.)] ", "", s)
        italic = len(s) > 2 and s.startswith("_") and s.endswith("_")
        if italic:
            s = s[1:-1]
        p = doc.add_paragraph(style=style)
        for i, part in enumerate(re.split(r"\*\*", s)):
            if part:
                run = p.add_run(part)
                run.bold = i % 2 == 1
                run.italic = italic


def build_word_report(summary_md: str, result: AnalysisResult, title: str) -> bytes:
    metric = result.facts["metric"]
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    heading = doc.add_heading(title, level=0)
    heading.runs[0].font.color.rgb = RGBColor.from_string(ACCENT)
    doc.add_paragraph(f"Generated on {date.today():%Y-%m-%d}")

    _add_markdown(doc, summary_md)

    doc.add_heading("Charts", level=1)
    if result.monthly is not None and len(result.monthly) > 1:
        fig, ax = plt.subplots(figsize=(8, 3.4))
        ax.plot(result.monthly["Month"], result.monthly[metric], marker="o", color="#" + ACCENT)
        ax.set_title(f"{metric} by month")
        ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
        ax.grid(axis="y", alpha=0.3)
        ax.spines[["top", "right"]].set_visible(False)
        plt.setp(ax.get_xticklabels(), rotation=45, ha="right")
        doc.add_picture(_chart_png(fig), width=Inches(6.3))

    for dim, table in list(result.by_dimension.items())[:3]:
        top = table.head(12).iloc[::-1]
        fig, ax = plt.subplots(figsize=(8, 0.35 * len(top) + 1.2))
        ax.barh(top[dim], top[metric], color="#" + ACCENT)
        ax.set_title(f"{metric} by {dim}")
        ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
        ax.spines[["top", "right"]].set_visible(False)
        doc.add_picture(_chart_png(fig), width=Inches(6.3))

    if result.monthly is not None and len(result.monthly):
        doc.add_heading("Monthly figures", level=1)
        table = doc.add_table(rows=1, cols=3)
        table.style = "Light Grid Accent 1"
        for cell, text in zip(table.rows[0].cells, ["Month", metric, "Rows"]):
            cell.text = text
        for _, r in result.monthly.iterrows():
            cells = table.add_row().cells
            cells[0].text, cells[1].text, cells[2].text = r["Month"], f"{r[metric]:,.2f}", str(r["Rows"])

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
