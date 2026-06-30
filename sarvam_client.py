"""
sarvam_client.py
------------------
Thin wrapper around the Sarvam AI Chat Completions API used to generate
narrative business insights from computed GST metrics.

Degrades gracefully: if no API key is configured, or the API call fails
for any reason (network, auth, rate limit, malformed response), the app
falls back to rule-based template insights generated locally so the
dashboard always has something useful to show.
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import requests
from dotenv import load_dotenv

from data_loader import format_indian_rupees, format_currency_in_narrative

load_dotenv(Path(__file__).resolve().parent / ".env")

SARVAM_CHAT_URL = "https://api.sarvam.ai/v1/chat/completions"
DEFAULT_MODEL = os.environ.get("SARVAM_MODEL", "sarvam-30b")
REQUEST_TIMEOUT = 60


@dataclass
class InsightResult:
    text: str
    source: str  # "sarvam_ai" or "fallback_template"
    error: Optional[str] = None


def _get_api_key() -> Optional[str]:
    key = os.environ.get("SARVAM_API_KEY", "").strip()
    return key or None


def is_sarvam_configured() -> bool:
    return _get_api_key() is not None


def _extract_message_text(data: dict) -> str:
    """Pull assistant text from Sarvam chat completion response."""
    message = data["choices"][0]["message"]
    text = (message.get("content") or "").strip()
    if text:
        return text
    reasoning = (message.get("reasoning_content") or "").strip()
    if reasoning:
        return reasoning
    raise KeyError("No content or reasoning_content in API response")


def _format_request_error(resp: requests.Response) -> str:
    try:
        body = resp.json()
        err = body.get("error", body)
        if isinstance(err, dict):
            return err.get("message") or json.dumps(err)
        return str(err)
    except (json.JSONDecodeError, ValueError):
        return resp.text[:500] or f"HTTP {resp.status_code}"


def _call_sarvam(prompt: str, api_key: str, system_prompt: str, max_tokens: int = 700) -> str:
    headers = {
        "api-subscription-key": api_key,
        "Content-Type": "application/json",
    }
    payload = {
        "model": DEFAULT_MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.4,
        "max_tokens": max_tokens,
        "reasoning_effort": None,
    }
    resp = requests.post(SARVAM_CHAT_URL, headers=headers, json=payload, timeout=REQUEST_TIMEOUT)
    if not resp.ok:
        raise requests.exceptions.HTTPError(
            f"{resp.status_code} {resp.reason}: {_format_request_error(resp)}",
            response=resp,
        )
    data = resp.json()
    return _extract_message_text(data)


SYSTEM_PROMPT = (
    "You are a senior Indian GST compliance and tax analyst. You write clear, "
    "specific, numbers-grounded narrative insights for a small/medium business "
    "owner reading their own monthly GST return data. Be direct and practical. "
    "Avoid generic filler. Use Indian Rupee formatting with Indian comma grouping (e.g. ₹ 12,34,567.89). "
    "Keep paragraphs short. Do not give legal advice or definitive tax-saving "
    "instructions; frame suggestions as observations worth discussing with a CA."
)


def _fmt_date(value) -> str:
    if value is None:
        return "Unknown"
    if hasattr(value, "strftime"):
        return value.strftime("%b %Y")
    return str(value)


def build_dashboard_prompt(
    stats: dict, yearly_df_text: str, monthly_df_text: str, trade_name: str
) -> str:
    date_start = _fmt_date(stats.get("date_range", (None, None))[0])
    date_end = _fmt_date(stats.get("date_range", (None, None))[1])
    
    total_sales = stats.get('total_sales', stats.get('total_turnover', 0))
    total_gst = stats.get('total_gst_liability', stats.get('total_outward_tax', 0))
    avg_sales = stats.get('avg_monthly_sales', stats.get('avg_monthly_turnover', 0))
    itc_avail = stats.get('total_itc_available', 0)
    net_itc = stats.get('net_itc', 0)
    itc_reversal = stats.get('itc_reversal', 0)
    total_tax_paid = stats.get('total_tax_paid', stats.get('total_cash_paid', 0) + stats.get('total_itc_utilized', 0))
    itc_eff = stats.get('itc_efficiency', 0.0)
    cash_dep = stats.get('cash_dependency', 0.0)
    rcm_tax = stats.get('total_rcm_tax', 0.0)
    rcm_pct = stats.get('rcm_ratio_pct', 0.0)
    delay = stats.get('avg_filing_delay_days', 0.0)
    late = stats.get('late_filings_count', 0)

    return f"""
Business: {trade_name}

Summary metrics across {stats.get('n_periods')} monthly GST return periods
({date_start} to {date_end}) — all figures taken directly from the uploaded Excel file:

- Total Sales / Turnover: {format_indian_rupees(total_sales, prefix="₹ ")}
- Average Monthly Sales: {format_indian_rupees(avg_sales, prefix="₹ ")}
- Total Outward GST Liability: {format_indian_rupees(total_gst, prefix="₹ ")}
- Average Effective Tax Rate: {stats.get('avg_effective_tax_rate'):.2f}%
- Total ITC (input tax credit) available: {format_indian_rupees(itc_avail, prefix="₹ ")}
- Net ITC (after reversals): {format_indian_rupees(net_itc, prefix="₹ ")}
- ITC Reversals: {format_indian_rupees(itc_reversal, prefix="₹ ")}
- Total tax liability paid: {format_indian_rupees(total_tax_paid, prefix="₹ ")}
- Total tax paid via cash ledger: {format_indian_rupees(stats.get('total_cash_paid', 0), prefix="₹ ")}
- Total ITC utilized: {format_indian_rupees(stats.get('total_itc_utilized', 0), prefix="₹ ")}
- ITC Efficiency (% of available utilized): {itc_eff:.2f}%
- Cash Dependency Ratio: {cash_dep:.2f}%
- Total Reverse Charge (RCM) tax liability: {format_indian_rupees(rcm_tax, prefix="₹ ")}
- Reverse Charge % (RCM tax as % of total GST liability): {rcm_pct:.2f}%
- Average filing delay vs due date: {delay:.1f} days
- Number of late filings: {late}
- Best month by turnover: {stats.get('best_month')}
- Weakest month by turnover: {stats.get('weakest_month')}

Month-by-month data from the uploaded file:
{monthly_df_text}

Year-wise breakdown:
{yearly_df_text}

Write a structured analysis containing EXACTLY these 6 sections, with each section name on its own line followed by a paragraph of narrative analysis:

Monthly Sales Trend
[Analyze the monthly sales trend, average sales, best/weakest months, and year-over-year change.]

GST Liability Trend
[Analyze the GST liability trend, effective tax rate, and how output tax tracks with sales.]

ITC Efficiency
[Analyze ITC available, net ITC, ITC reversals, and how efficiently input credit is utilized to offset tax.]

Cash Flow Analysis
[Analyze total tax paid, cash ledger payments, ITC offsets, cash dependency, and its working capital impact.]

Reverse Charge Analysis
[Analyze total reverse charge (RCM) tax paid, RCM %, and potential cash flow implications.]

Filing Compliance
[Analyze compliance timeliness, late filings count, average filing delay, and potential compliance risk.]

Rules:
- Format each section with the exact title above as a header, followed by its analysis paragraph.
- Use the actual numbers and month names from the data. Do not invent figures.
- Do not use markdown hash headers (# or ##). Just print the section title on its own line.
""".strip()


def generate_dashboard_insight(
    stats: dict, yearly_df_text: str, monthly_df_text: str, trade_name: str
) -> InsightResult:
    key = _get_api_key()
    if not key:
        raw_text = _fallback_dashboard_insight(stats)
        return InsightResult(text=format_currency_in_narrative(raw_text), source="fallback_template",
                               error="No Sarvam API key configured.")
    try:
        prompt = build_dashboard_prompt(stats, yearly_df_text, monthly_df_text, trade_name)
        text = _call_sarvam(prompt, key, SYSTEM_PROMPT, max_tokens=1000)
        return InsightResult(text=format_currency_in_narrative(text), source="sarvam_ai")
    except requests.exceptions.RequestException as e:
        raw_text = _fallback_dashboard_insight(stats)
        return InsightResult(text=format_currency_in_narrative(raw_text), source="fallback_template", error=str(e))
    except (KeyError, IndexError, json.JSONDecodeError) as e:
        raw_text = _fallback_dashboard_insight(stats)
        return InsightResult(text=format_currency_in_narrative(raw_text), source="fallback_template",
                               error=f"Unexpected API response format: {e}")


def generate_custom_insight(
    question: str, stats: dict, monthly_df_text: str, trade_name: str
) -> InsightResult:
    """Answer an ad-hoc question the user types about their own data."""
    key = _get_api_key()
    date_start = _fmt_date(stats.get("date_range", (None, None))[0])
    date_end = _fmt_date(stats.get("date_range", (None, None))[1])
    context = f"""
Business: {trade_name}
Periods covered: {stats.get('n_periods')} months ({date_start} to {date_end})
Total sales: {format_indian_rupees(stats.get('total_sales', 0), prefix="₹ ")}
Total GST liability: {format_indian_rupees(stats.get('total_gst_liability', 0), prefix="₹ ")}
Total ITC available: {format_indian_rupees(stats.get('total_itc_available', 0), prefix="₹ ")}
Total ITC utilized: {format_indian_rupees(stats.get('total_itc_utilized', 0), prefix="₹ ")}
Total cash paid: {format_indian_rupees(stats.get('total_cash_paid', 0), prefix="₹ ")}
Average effective tax rate: {stats.get('avg_effective_tax_rate'):.2f}%
Late filings: {stats.get('late_filings_count')} out of {stats.get('n_periods')}

Month-by-month data from the uploaded Excel file:
{monthly_df_text}

Question from the business owner: {question}

Answer using only the numbers and months shown above. Do not invent figures.
""".strip()
    if not key:
        return InsightResult(
            text="AI-powered Q&A needs a Sarvam API key. Add SARVAM_API_KEY to your .env file; "
                 "for now, here are the relevant numbers above to help you answer that yourself.",
            source="fallback_template", error="No Sarvam API key configured.",
        )
    try:
        text = _call_sarvam(context, key, SYSTEM_PROMPT, max_tokens=500)
        return InsightResult(text=format_currency_in_narrative(text), source="sarvam_ai")
    except requests.exceptions.RequestException as e:
        return InsightResult(text="Could not reach Sarvam AI right now. Please try again in a moment.",
                               source="fallback_template", error=str(e))
    except (KeyError, IndexError, json.JSONDecodeError) as e:
        return InsightResult(text="Received an unexpected response from Sarvam AI.",
                               source="fallback_template", error=str(e))


def _fallback_dashboard_insight(stats: dict) -> str:
    """Rule-based narrative used when the Sarvam API is unavailable."""
    rate = stats.get("avg_effective_tax_rate") or 0.0
    late = stats.get("late_filings_count") or 0
    n = stats.get("n_periods") or 1
    cash = stats.get("total_cash_paid") or 0.0
    itc_used = stats.get("total_itc_utilized") or 0.0
    itc_avail = stats.get("total_itc_available") or 0.0
    itc_net = stats.get("net_itc") or 0.0
    itc_reversal = stats.get("itc_reversal") or 0.0
    total_sales = stats.get("total_sales") or 0.0
    avg_sales = stats.get("avg_monthly_sales") or 0.0
    total_liability = stats.get("total_gst_liability") or 0.0
    total_tax_paid = stats.get("total_tax_paid") or 0.0
    total_rcm = stats.get("total_rcm_tax") or 0.0
    rcm_pct = stats.get("rcm_ratio_pct") or 0.0
    itc_eff = stats.get("itc_efficiency") or 0.0
    cash_dep = stats.get("cash_dependency") or 0.0
    delay = stats.get("avg_filing_delay_days") or 0.0

    sections = []

    # 1. Monthly Sales Trend
    sections.append(
        "Monthly Sales Trend\n"
        f"Across {n} return periods, total sales stood at {format_indian_rupees(total_sales, prefix='₹ ')} with an average "
        f"monthly sales of {format_indian_rupees(avg_sales, prefix='₹ ')}. The strongest month was {stats.get('best_month', 'N/A')} "
        f"and the weakest month was {stats.get('weakest_month', 'N/A')}, showing directional changes in sales trajectory."
    )

    # 2. GST Liability Trend
    sections.append(
        "GST Liability Trend\n"
        f"Total GST liability was {format_indian_rupees(total_liability, prefix='₹ ')}, representing an average effective "
        f"tax rate of {rate:.2f}%. Liability trends tracked closely with sales fluctuations, showing "
        f"normal tax incidence across the periods."
    )

    # 3. ITC Efficiency
    sections.append(
        "ITC Efficiency\n"
        f"Total Input Tax Credit (ITC) available was {format_indian_rupees(itc_avail, prefix='₹ ')}. Net ITC stood at {format_indian_rupees(itc_net, prefix='₹ ')} "
        f"after accounting for {format_indian_rupees(itc_reversal, prefix='₹ ')} in ITC Reversals. The overall ITC efficiency ratio "
        f"reached {itc_eff:.2f}%, indicating the percentage of available credits offset against liability."
    )

    # 4. Cash Flow Analysis
    sections.append(
        "Cash Flow Analysis\n"
        f"The total tax paid was {format_indian_rupees(total_tax_paid, prefix='₹ ')}, discharged through {format_indian_rupees(cash, prefix='₹ ')} via the cash ledger "
        f"and {format_indian_rupees(itc_used, prefix='₹ ')} via ITC offset. This results in a Cash Dependency ratio of {cash_dep:.2f}%, "
        f"highlighting the proportion of tax liability requiring direct cash outflow and its working capital impact."
    )

    # 5. Reverse Charge Analysis
    sections.append(
        "Reverse Charge Analysis\n"
        f"Total reverse charge (RCM) tax liability was {format_indian_rupees(total_rcm, prefix='₹ ')}, representing {rcm_pct:.2f}% "
        "of the total GST liability. Keeping track of RCM is critical as it requires upfront cash outflow "
        "before it can be claimed as ITC."
    )

    # 6. Filing Compliance
    compliance_msg = (
        f"{late} of {n} returns were filed after the statutory due date, with an average delay of {delay:.1f} days. "
        "Consistent delay can attract late fees and interest."
        if late > 0
        else "All returns in this dataset were filed on or before the due date, reflecting strong compliance discipline with zero delays."
    )
    sections.append(
        "Filing Compliance\n" + compliance_msg
    )

    return "\n\n".join(sections)

