"""AI-written report with a local model (Ollama), plus a rule-based fallback for demo mode.

Small local models (7-8B) make mistakes when they have to read raw JSON or tables and do
arithmetic. So Python turns every computed fact into a short, ready-made sentence, and the
model only has to prioritise, explain, translate and recommend actions.
"""

from __future__ import annotations

import os
import re

import httpx
import ollama
import pandas as pd
from dotenv import load_dotenv

from analysis import AnalysisResult

load_dotenv()
DEFAULT_MODEL = os.getenv("OLLAMA_MODEL", "qwen3:8b")
DEFAULT_HOST = os.getenv("OLLAMA_HOST", "http://localhost:11434")
EMBEDDING_TAGS = ("embed", "bge", "minilm")
MODEL_OPTIONS = {
    "num_ctx": 8192,         # Ollama's default window (2-4k tokens) would silently cut the input
    "num_predict": 1500,     # keeps the report short and the wait reasonable on a laptop
    "temperature": 0.3,
    "repeat_penalty": 1.15,  # small models tend to loop on the same sentence
}

SECTION_TITLES = {
    "English": ["Executive summary", "Key figures", "Risks and anomalies", "Opportunities",
                "Recommended actions", "Data quality notes"],
    "Italiano": ["Sintesi", "Dati chiave", "Rischi e anomalie", "Opportunità",
                 "Azioni consigliate", "Qualità dei dati"],
    "Español": ["Resumen ejecutivo", "Cifras clave", "Riesgos y anomalías", "Oportunidades",
                "Acciones recomendadas", "Calidad de los datos"],
    "Deutsch": ["Zusammenfassung", "Kennzahlen", "Risiken und Auffälligkeiten", "Chancen",
                "Empfohlene Maßnahmen", "Datenqualität"],
    "Français": ["Synthèse", "Chiffres clés", "Risques et anomalies", "Opportunités",
                 "Actions recommandées", "Qualité des données"],
    "Svenska": ["Sammanfattning", "Nyckeltal", "Risker och avvikelser", "Möjligheter",
                "Rekommenderade åtgärder", "Datakvalitet"],
    "Português": ["Resumo executivo", "Números-chave", "Riscos e anomalias", "Oportunidades",
                  "Ações recomendadas", "Qualidade dos dados"],
}

SYSTEM_PROMPT = """You are a business analyst. You turn a list of verified findings into a short report for managers.

Strict rules:
- Use ONLY the findings provided. Do not add numbers, names or facts that are not in them.
- Copy every number exactly as written in the findings. Never calculate or change numbers.
- Each finding appears at most once in the whole report. Never repeat a sentence.
- Keep names of regions, products, people and categories exactly as written (do not translate them).
- Write EVERYTHING in {language}: translate all the findings you use, including the key figures
  and data quality notes. Only numbers and names stay unchanged.

Use exactly these six section titles, in this order:
## {t0}
3 to 5 bullet points: the most important findings and why they matter for the business.
## {t1}
The KEY FIGURES block as bullet points, translated.
## {t2}
Up to 4 bullet points from the RISKS block, each with its likely business impact.
## {t3}
Up to 3 bullet points from the GROWTH block.
## {t4}
Up to 5 numbered, concrete actions, each linked to a finding.
## {t5}
The DATA QUALITY block as bullet points, translated.

Maximum 450 words. Plain Markdown, no tables."""


def system_prompt(language: str) -> str:
    titles = SECTION_TITLES.get(language, SECTION_TITLES["English"])
    return SYSTEM_PROMPT.format(language=language, **{f"t{i}": t for i, t in enumerate(titles)})


class AISummaryError(Exception):
    """Raised with a user-friendly message when the AI summary cannot be generated."""


# ----------------------------------------------------------------------------- findings
def _fmt(v: float) -> str:
    return f"{v:,.0f}"


def findings(result: AnalysisResult) -> dict[str, list[str]]:
    """Every computed fact as a ready-to-use English sentence, grouped by topic."""
    f = result.facts
    metric = f["metric"]
    key, growth, risks, quality = [], [], [], []

    if "period" in f:
        key.append(f"Total {metric}: {_fmt(f['total'])}, from {f['period']['from']} to "
                   f"{f['period']['to']} ({f['rows_analysed']:,} rows, {f['filter_applied']}).")
    else:
        key.append(f"Total {metric}: {_fmt(f['total'])} ({f['rows_analysed']:,} rows).")
    key.append(f"Average {metric} per row: {f['average_per_row']:,.2f}.")
    t = f.get("trend")
    if t:
        months = f["monthly_totals"]
        key.append(f"Best month: {t['best_month']} ({_fmt(months[t['best_month']])}). "
                   f"Worst month: {t['worst_month']} ({_fmt(months[t['worst_month']])}).")
        if t.get("last_month_vs_previous_pct") is not None:
            key.append(f"Last month {t['last_month']}: {t['last_month_vs_previous_pct']:+.1f}% "
                       f"vs the previous month.")
        if t.get("avg_last_3_months_vs_first_3_months_pct") is not None:
            growth.append(f"Average monthly {metric} of the last 3 months is "
                          f"{t['avg_last_3_months_vs_first_3_months_pct']:+.1f}% vs the first 3 months.")
    for dim, d in f.get("by_dimension", {}).items():
        if d["ranking"]:
            top = d["ranking"][0]
            key.append(f"Top {dim}: {top['member']} with {_fmt(top['total'])} "
                       f"({top['share_pct']}% of the total).")
        for c in d.get("strong_growth", []):
            growth.append(f"{dim} '{c['member']}' grew {c['change_pct']:+.1f}%: from "
                          f"{_fmt(c['avg_monthly_before_last_3_months'])} to "
                          f"{_fmt(c['avg_monthly_last_3_months'])} per month (last 3 months vs before).")
        for c in d.get("sharp_drops", []):
            risks.append(f"{dim} '{c['member']}' dropped {c['change_pct']:+.1f}%: from "
                         f"{_fmt(c['avg_monthly_before_last_3_months'])} to "
                         f"{_fmt(c['avg_monthly_last_3_months'])} per month (last 3 months vs before).")
    for a in f.get("status_alerts", []):
        risks.append(f"'{a['status']}' rate for {a['dimension']} '{a['member']}' is "
                     f"{a['rate_last_2_months_pct']}% in the last 2 months, vs {a['rate_before_pct']}% before.")
    for o in f.get("outliers", []):
        detail = ", ".join(f"{k}: {v}" for k, v in o.items()
                           if k not in (metric, "times_the_99th_percentile") and v is not None)
        risks.append(f"One single row has {metric} {_fmt(o[metric])}, {o['times_the_99th_percentile']} times "
                     f"the 99th percentile of all rows ({detail}). Check whether it is correct and "
                     f"whether it inflates the totals.")

    dq = f.get("data_quality")
    if dq:
        quality.append(f"{dq['duplicates_removed']} duplicated rows and {dq['empty_rows_removed']} "
                       f"empty rows were removed; {dq['text_cells_trimmed']} text cells were cleaned.")
        for col, n in dq["missing_values_by_column"].items():
            quality.append(f"Column '{col}' has {n} missing values.")
    return {"KEY FIGURES": key, "GROWTH": growth, "RISKS": risks, "DATA QUALITY": quality}


def build_context(file_name: str, sheet_name: str, result: AnalysisResult,
                  other_sheets: dict[str, pd.DataFrame]) -> str:
    parts = [f"Source: file '{file_name}', sheet '{sheet_name}'."]
    if other_sheets:
        parts.append("Other sheets in the file (not analysed): " + ", ".join(other_sheets) + ".")
    for section, items in findings(result).items():
        parts += ["", f"{section}:"] + [f"- {i}" for i in items or ["(none)"]]
    return "\n".join(parts)


# ----------------------------------------------------------------------------- Ollama
def list_local_models(host: str) -> list[str]:
    """Chat models installed in Ollama (embedding models are left out). Empty if Ollama is off."""
    try:
        models = ollama.Client(host=host, timeout=3).list().models
    except Exception:  # Ollama not running or not reachable
        return []
    return sorted(m.model for m in models if not any(tag in m.model for tag in EMBEDDING_TAGS))


def _chat(client: ollama.Client, model: str, messages: list[dict], think: bool | None):
    extra = {} if think is None else {"think": think}
    return client.chat(model=model, messages=messages, options=MODEL_OPTIONS, **extra)


def generate_ai_summary(context: str, language: str, model: str = DEFAULT_MODEL,
                        host: str = DEFAULT_HOST) -> str:
    client = ollama.Client(host=host, timeout=900)
    messages = [{"role": "system", "content": system_prompt(language)},
                {"role": "user", "content": context}]
    try:
        try:
            # think=False skips the slow reasoning phase of models like qwen3
            response = _chat(client, model, messages, think=False)
        except ollama.ResponseError as e:
            if "think" not in str(e).lower():
                raise
            response = _chat(client, model, messages, think=None)  # model has no thinking switch
    except ollama.ResponseError as e:
        if e.status_code == 404:
            raise AISummaryError(f"The model '{model}' is not installed. Run: ollama pull {model}") from e
        raise AISummaryError(f"Ollama returned an error: {e.error}") from e
    except (ConnectionError, httpx.ConnectError) as e:
        raise AISummaryError("Ollama is not running. Open the Ollama app (or run 'ollama serve') "
                             "and try again.") from e
    except httpx.TimeoutException as e:
        raise AISummaryError("The local model took too long. Try a smaller model.") from e

    text = re.sub(r"<think>.*?</think>", "", response.message.content or "", flags=re.S).strip()
    text = _drop_repeated_lines(text)
    if not text:
        raise AISummaryError("The model returned an empty answer. Try again.")
    if response.done_reason == "length":
        text += "\n\n_(The report was cut because it reached the length limit.)_"
    return text


NUMBER = re.compile(r"(?<![\w.,])\d[\d.,]*\d|(?<![\w.,])\d")


def _number_readings(token: str) -> set[float]:
    """All plausible values of a number written in English (1,234.5) or European (1.234,5) style."""
    token = token.strip(".,")
    readings = set()
    for thousands, decimal in ((",", "."), (".", ",")):
        candidate = token.replace(thousands, "").replace(decimal, ".")
        try:
            readings.add(round(float(candidate), 2))
        except ValueError:
            pass
    return readings


def unverified_numbers(report: str, context: str) -> list[str]:
    """Numbers in the AI report that do not appear in the findings it received.

    Dates and small integers (1-31, e.g. "3 months") are ignored.
    """
    report = re.sub(r"\d{4}-\d{2}(-\d{2})?", " ", report)
    known: set[float] = set()
    for token in NUMBER.findall(context):
        known |= _number_readings(token)
    suspicious = []
    for token in NUMBER.findall(report):
        readings = _number_readings(token)
        if not readings or (token.isdigit() and int(token) <= 31):
            continue
        if not readings & known and token not in suspicious:
            suspicious.append(token)
    return suspicious


def _drop_repeated_lines(text: str) -> str:
    """Safety net against small models looping on the same sentence."""
    seen, out = set(), []
    for line in text.splitlines():
        key = re.sub(r"^\s*(?:[-*]|\d+[.)])\s*", "", line).strip().lower()
        if key and len(key) > 20 and key in seen:
            continue
        seen.add(key)
        out.append(line)
    return "\n".join(out)


# ----------------------------------------------------------------------------- demo mode
def demo_summary(result: AnalysisResult) -> str:
    """Rule-based report used when Ollama is not available: same findings, no AI wording."""
    titles = {"KEY FIGURES": "Key figures", "GROWTH": "Opportunities",
              "RISKS": "Risks and anomalies", "DATA QUALITY": "Data quality notes"}
    lines = ["_Demo mode: this summary is rule-based. Start Ollama to get the AI-written analysis._"]
    for section, items in findings(result).items():
        if items:
            lines += ["", f"## {titles[section]}"] + [f"- {i}" for i in items]
    return "\n".join(lines)
