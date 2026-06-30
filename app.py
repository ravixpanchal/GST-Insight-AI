"""
app.py
--------
GST Insight AI — Streamlit dashboard.

Upload a GST returns Excel workbook -> see KPIs, charts, year-wise
breakdown, AI-generated narrative insights (via Sarvam AI, with graceful
fallback), and download a polished Word report.

Run with:
    streamlit run app.py
"""

import hashlib

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

from data_loader import (
    load_gst_workbook,
    get_summary_stats,
    get_yearly_summary,
    get_monthly_summary_text,
    format_indian_rupees,
)
from sarvam_client import generate_dashboard_insight, generate_custom_insight, is_sarvam_configured
from report_generator import build_word_report

# ---------------------------------------------------------------------------
# Page config & theme
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="GST Insight AI",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="auto",
)

NAVY = "#0A1F44"
NAVY_LIGHT = "#16305E"
GOLD = "#C9A227"
GOLD_LIGHT = "#E0C158"
BG = "#F7F8FA"

CUSTOM_CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&display=swap');

/* Global Font Override */
.stApp, .stApp [data-testid="stMarkdownContainer"] p, .stApp label, .stApp button, .stApp select, .stApp input, .stApp textarea, .stApp [data-baseweb="tab"] {{
    font-family: 'Outfit', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif !important;
}}

.stApp {{ background-color: {BG}; }}

/* Main content container with responsive padding */
.block-container {{
    padding-top: 2rem !important;
    padding-bottom: 3rem !important;
    padding-left: 3rem !important;
    padding-right: 3rem !important;
    max-width: 100% !important;
}}

@media (max-width: 1024px) {{
    .block-container {{
        padding-left: 2rem !important;
        padding-right: 2rem !important;
        padding-top: 1.5rem !important;
    }}
}}

@media (max-width: 768px) {{
    .block-container {{
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        padding-top: 1rem !important;
    }}
}}

/* Sidebar Styling */
[data-testid="stSidebar"] {{
    background-color: {NAVY} !important;
    border-right: 1px solid rgba(255, 255, 255, 0.05);
}}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h1,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] li,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stCaption {{
    color: #E2E8F0 !important;
}}

[data-testid="stSidebar"] .stButton button {{
    background: linear-gradient(135deg, {GOLD}, {GOLD_LIGHT});
    color: {NAVY} !important;
    font-weight: 600;
    border: none;
    border-radius: 8px;
    width: 100%;
    transition: all 0.25s ease;
    box-shadow: 0 4px 10px rgba(201, 162, 39, 0.15);
}}
[data-testid="stSidebar"] .stButton button:hover {{
    transform: translateY(-1px);
    box-shadow: 0 6px 14px rgba(201, 162, 39, 0.25);
    background: linear-gradient(135deg, {GOLD_LIGHT}, {GOLD});
    color: {NAVY} !important;
}}

h1, h2, h3 {{
    color: {NAVY};
    font-weight: 700;
    word-wrap: break-word;
}}

/* Responsive KPI Grid using auto-fit */
.kpi-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(min(240px, 100%), 1fr));
    gap: 16px;
    margin-bottom: 8px;
}}

.kpi-card {{
    background: white;
    border-radius: 12px;
    padding: 20px 24px;
    border-left: 6px solid {GOLD};
    box-shadow: 0 4px 12px rgba(10, 31, 68, 0.04);
    transition: transform 0.25s cubic-bezier(0.4, 0, 0.2, 1), box-shadow 0.25s ease;
    min-width: 0;
}}
.kpi-card:hover {{
    transform: translateY(-2px);
    box-shadow: 0 8px 20px rgba(10, 31, 68, 0.08);
}}

.kpi-label {{
    font-size: 12px;
    color: #64748B;
    text-transform: uppercase;
    font-weight: 600;
    letter-spacing: 0.06em;
    margin-bottom: 6px;
}}
.kpi-value {{
    font-size: clamp(20px, 2.2vw, 26px);
    font-weight: 700;
    color: {NAVY};
    line-height: 1.2;
    word-break: break-word;
}}
.kpi-sub {{
    font-size: 12px;
    color: #94A3B8;
    margin-top: 4px;
    word-break: break-word;
}}

/* Header meta chips */
.meta-row {{
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    margin: 12px 0 8px 0;
}}
.meta-chip {{
    display: inline-flex;
    align-items: center;
    background: white;
    border: 1px solid #E2E8F0;
    border-radius: 20px;
    padding: 6px 14px;
    font-size: 13px;
    color: #475569;
    max-width: 100%;
    word-break: break-word;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02);
}}
.meta-chip strong {{
    color: {NAVY};
    margin-right: 6px;
    font-weight: 600;
}}

/* AI Insights / Narrative box */
.insight-box {{
    background: white;
    border-radius: 12px;
    padding: 24px 30px;
    border-top: 5px solid {GOLD};
    box-shadow: 0 4px 16px rgba(10, 31, 68, 0.04);
    line-height: 1.7;
    font-size: 16px;
    color: #1E293B;
    word-wrap: break-word;
    overflow-wrap: anywhere;
    white-space: pre-wrap;
}}

.section-divider {{
    height: 4px;
    background: linear-gradient(90deg, {NAVY}, {GOLD}, {BG});
    border-radius: 4px;
    margin: 20px 0 24px 0;
}}

.badge {{
    display: inline-block;
    padding: 4px 12px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 600;
    margin-bottom: 10px;
    letter-spacing: 0.04em;
    text-transform: uppercase;
}}
.badge-ai {{ background: #E8F5E9; color: #1B5E20; border: 1px solid #C8E6C9; }}
.badge-fallback {{ background: #FFF3E0; color: #E65100; border: 1px solid #FFE0B2; }}

/* Scrollable tabs on narrow screens */
.stTabs [data-baseweb="tab-list"] {{
    gap: 8px;
    flex-wrap: nowrap;
    overflow-x: auto;
    overflow-y: hidden;
    -webkit-overflow-scrolling: touch;
    scrollbar-width: none; /* Hide scrollbar for clean look */
}}
.stTabs [data-baseweb="tab-list"]::-webkit-scrollbar {{
    display: none;
}}
.stTabs [data-baseweb="tab"] {{
    white-space: nowrap;
    flex-shrink: 0;
    font-weight: 500;
    padding: 10px 16px;
    border-radius: 8px 8px 0 0;
}}

/* Plotly charts fill container */
[data-testid="stPlotlyChart"] {{
    width: 100% !important;
    overflow-x: auto;
}}

/* Touch-friendly buttons and forms */
.stButton button {{
    min-height: 2.8rem;
    border-radius: 8px;
    font-weight: 500;
    transition: all 0.2s ease;
}}

/* Tablet columns responsiveness override */
@media (max-width: 992px) {{
    .insight-box {{
        padding: 20px;
    }}
    
    /* Stack layout columns on tablet/mobile screens */
    [data-testid="stHorizontalBlock"] {{
        flex-wrap: wrap !important;
        gap: 1.5rem !important;
    }}
    [data-testid="column"], [data-testid="stColumn"] {{
        width: 100% !important;
        flex: 1 1 100% !important;
        min-width: 100% !important;
    }}
}}

/* Mobile adjustments */
@media (max-width: 768px) {{
    h1 {{
        font-size: 1.65rem !important;
        line-height: 1.3 !important;
    }}
    h2 {{
        font-size: 1.35rem !important;
    }}
    h3, h4 {{
        font-size: 1.15rem !important;
    }}

    .meta-chip {{
        font-size: 12px;
        padding: 5px 10px;
    }}

    .insight-box {{
        padding: 16px 18px;
        font-size: 15px;
    }}

    .stButton button, [data-testid="stDownloadButton"] button {{
        width: 100% !important;
    }}
}}

/* Small phones */
@media (max-width: 480px) {{
    .kpi-card {{
        padding: 16px;
    }}
    .kpi-value {{
        font-size: 20px;
    }}
}}

/* Report Preview styling */
.report-preview-container {{
    background-color: white;
    padding: 40px;
    border-radius: 12px;
    box-shadow: 0 10px 30px rgba(10, 31, 68, 0.05);
    border: 1px solid #E2E8F0;
    max-width: 850px;
    margin: 20px auto;
    font-family: 'Outfit', sans-serif;
    color: #334155;
}}
.report-preview-header {{
    text-align: center;
    border-bottom: 2px solid {NAVY};
    padding-bottom: 24px;
    margin-bottom: 30px;
}}
.report-preview-title {{
    font-size: 32px;
    font-weight: 800;
    color: {NAVY};
    margin: 0 0 8px 0;
    letter-spacing: -0.02em;
}}
.report-preview-subtitle {{
    font-size: 18px;
    font-weight: 500;
    color: {GOLD};
    margin: 0 0 20px 0;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}}
.report-preview-meta {{
    font-size: 14px;
    color: #64748B;
    margin: 4px 0;
}}
.report-preview-section-title {{
    font-size: 20px;
    font-weight: 700;
    color: {NAVY};
    margin-top: 36px;
    margin-bottom: 16px;
    padding-bottom: 8px;
    border-bottom: 1px solid {GOLD};
}}
.report-preview-table {{
    width: 100%;
    border-collapse: collapse;
    margin: 16px 0 24px 0;
    font-size: 14px;
}}
.report-preview-table th {{
    background-color: {NAVY};
    color: white;
    font-weight: 600;
    text-align: left;
    padding: 10px 14px;
    border: 1px solid #E2E8F0;
}}
.report-preview-table td {{
    padding: 10px 14px;
    border: 1px solid #E2E8F0;
    color: #334155;
}}
.report-preview-table tr:nth-child(even) td {{
    background-color: #F8FAFC;
}}
.report-preview-conclusion {{
    font-size: 13px;
    font-style: italic;
    color: #64748B;
    margin-top: 40px;
    border-top: 1px solid #E2E8F0;
    padding-top: 16px;
    line-height: 1.6;
}}
@media (max-width: 768px) {{
    .report-preview-container {{
        padding: 24px 16px;
        margin: 10px 0;
    }}
    .report-preview-title {{
        font-size: 24px;
    }}
    .report-preview-subtitle {{
        font-size: 15px;
    }}
    .report-preview-table th, .report-preview-table td {{
        padding: 8px 10px;
        font-size: 12px;
    }}
}}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

PLOTLY_TEMPLATE = dict(
    layout=dict(
        font=dict(family="'Outfit', sans-serif", color=NAVY, size=11),
        plot_bgcolor="white",
        paper_bgcolor="white",
        colorway=[NAVY, GOLD, "#5B7FB5", "#D9B66A", "#8FA3C7"],
        title_font=dict(size=14, color=NAVY, family="'Outfit', sans-serif"),
        margin=dict(t=80, l=50, r=30, b=50),
        autosize=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, xanchor="left"),
    )
)


def style_dataframe_indian_rupees(dataframe: pd.DataFrame):
    """Format numeric columns in the dataframe to Indian money representation (Lakhs/Crores grouping)."""
    format_dict = {}
    for col in dataframe.columns:
        col_lower = str(col).lower()
        is_msme_money = any(term in col_lower for term in [
            "written down value", "wdv", "net investment", "total turnover", "net turnover", "pollution control", "export turnover"
        ])
        
        if is_msme_money:
            format_dict[col] = lambda x: format_indian_rupees(x) if pd.notna(x) and str(x).strip() != "" else ""
        elif pd.api.types.is_numeric_dtype(dataframe[col]):
            if any(term in col_lower for term in [
                "year", "delay", "month", "period", "pct", "ratio", "rate", 
                "status", "date", "arn", "gstin", "trade_name", "count", "filed", "number"
            ]) or col_lower == "no" or " no " in col_lower or col_lower.endswith(" no") or col_lower.startswith("no "):
                if any(term in col_lower for term in ["pct", "ratio", "rate"]):
                    format_dict[col] = lambda x: f"{x:.2f}%" if pd.notna(x) else ""
                elif "delay" in col_lower:
                    format_dict[col] = lambda x: f"{int(x)}" if pd.notna(x) else ""
                continue
            format_dict[col] = lambda x: format_indian_rupees(x) if pd.notna(x) else ""
        elif pd.api.types.is_datetime64_any_dtype(dataframe[col]):
            format_dict[col] = lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else ""
    return dataframe.style.format(format_dict)


def format_indian_rupees_short(val: float | int, prefix: str = "₹ ") -> str:
    if " " in prefix:
        prefix = prefix.replace(" ", "\u00A0")
    if pd.isna(val) or val is None or val == 0:
        return f"{prefix}0"
    is_neg = val < 0
    val = abs(val)
    if val >= 10000000: # 1 Crore
        formatted = f"{val / 10000000:.2f}".rstrip('0').rstrip('.') + " Cr"
    elif val >= 100000: # 1 Lakh
        formatted = f"{val / 100000:.2f}".rstrip('0').rstrip('.') + " L"
    elif val >= 1000: # Thousand
        formatted = f"{val / 1000:.2f}".rstrip('0').rstrip('.') + " K"
    else:
        formatted = f"{val:.2f}".rstrip('0').rstrip('.')
    
    res = f"{prefix}{formatted}"
    if is_neg:
        res = f"-{res}"
    return res


def get_indian_ticks(max_val: float):
    if pd.isna(max_val) or max_val <= 0:
        return [0], ["₹ 0"]
    
    steps = [
        1000, 2500, 5000, 10000, 25000, 50000, 100000, 250000, 500000, 
        1000000, 2500000, 5000000, 10000000, 25000000, 50000000, 100000000
    ]
    raw_step = max_val / 5
    step = steps[-1]
    for s in steps:
        if s >= raw_step:
            step = s
            break
            
    tick_max = (int(max_val / step) + 1) * step
    tickvals = list(range(0, int(tick_max) + 1, int(step)))
    ticktext = [format_indian_rupees_short(v) for v in tickvals]
    return tickvals, ticktext


def style_chart(fig, height: int = 400, n_labels: int = 0):
    """Apply shared layout and mobile-friendly axis labels."""
    fig.update_layout(**PLOTLY_TEMPLATE["layout"], height=height)
    if n_labels > 6:
        fig.update_xaxes(tickangle=-45)
        fig.update_layout(margin=dict(t=80, l=50, r=30, b=80))
    else:
        fig.update_layout(margin=dict(t=80, l=50, r=30, b=50))
    fig.update_layout(hovermode="x unified")
    return fig


def kpi_grid(cards: list[tuple[str, str, str]]):
    """Render KPI cards in a responsive CSS grid."""
    items = "".join(
        f"""<div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-sub">{sub}</div>
        </div>"""
        for label, value, sub in cards
    )
    st.markdown(f'<div class="kpi-grid">{items}</div>', unsafe_allow_html=True)


def meta_chips(chips: list[tuple[str, str]]):
    """Render wrapping header metadata chips."""
    items = "".join(
        f'<span class="meta-chip"><strong>{label}</strong>{value}</span>'
        for label, value in chips
    )
    st.markdown(f'<div class="meta-row">{items}</div>', unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## 📊 GST Insight AI")
    st.markdown("Upload GST return data and get instant analysis, charts, and AI narrative insights.")
    st.markdown("---")

    uploaded_file = st.file_uploader("Upload GST Excel file", type=["xlsx", "xls"])

    if uploaded_file is not None:
        st.markdown("---")
        st.markdown("**Current file**")
        st.markdown(f"📄 `{uploaded_file.name}`")
        st.caption(f"{uploaded_file.size / 1024:.1f} KB")

    st.markdown("---")
    st.caption("Made with ♥ by GSVians")

# ---------------------------------------------------------------------------
# Empty state
# ---------------------------------------------------------------------------
if uploaded_file is None:
    st.title("GST Insight AI")
    st.markdown("#### Upload a GST returns Excel file to generate your dashboard")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(
            """
            **What you get**
            - KPI cards for turnover, tax, ITC, and cash payments
            - Interactive charts across overview, tax, ITC, and compliance tabs
            - AI narrative insights and a downloadable Word report
            """
        )
    with col2:
        st.markdown(
            """
            **Getting started**
            1. Use the sidebar to upload your `.xlsx` or `.xls` file
            2. Review the dashboard tabs
            3. Download the Word report when ready
            """
        )

    if is_sarvam_configured():
        st.success("Sarvam AI is configured via `.env` — AI Insights will use the live API.")
    else:
        st.info(
            "No `SARVAM_API_KEY` found in `.env`. The app will use built-in fallback insights. "
            "Copy `.env.example` to `.env` and add your key to enable Sarvam AI."
        )
    st.stop()

# ---------------------------------------------------------------------------
# Load & process data — always from the uploaded Excel file
# ---------------------------------------------------------------------------
file_bytes = uploaded_file.getvalue()
data_fingerprint = hashlib.sha256(file_bytes).hexdigest()

if st.session_state.get("data_fingerprint") != data_fingerprint:
    st.session_state["data_fingerprint"] = data_fingerprint
    st.session_state.pop("ai_insight", None)
    st.session_state.pop("ai_insight_fingerprint", None)
    st.session_state.pop("custom_q_answer", None)
    st.session_state.pop("custom_q_question", None)

try:
    with st.spinner("Reading and analyzing your GST data..."):
        result = load_gst_workbook(file_bytes)
        df = result.df
        stats = get_summary_stats(df)
        yearly_df = get_yearly_summary(df)
        monthly_text = get_monthly_summary_text(df)
except ValueError as e:
    st.error(f"Could not process this file: {e}")
    st.stop()
except Exception as e:
    st.error(f"Unexpected error while reading the file: {e}")
    st.stop()

trade_name = df["trade_name"].iloc[0] if "trade_name" in df.columns and df["trade_name"].notna().any() else "Your Business"
gstin = df["gstin"].iloc[0] if "gstin" in df.columns and df["gstin"].notna().any() else None

if result.warnings:
    with st.expander("⚠️ Data notes", expanded=False):
        for w in result.warnings:
            st.write(f"- {w}")

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.markdown(f"# {trade_name}")
date_start, date_end = stats["date_range"]
chips = [
    ("File", uploaded_file.name),
    ("Sheet", result.sheet_used),
    ("Periods", str(stats["n_periods"])),
    ("Range", f"{date_start.strftime('%b %Y')} – {date_end.strftime('%b %Y')}"),
]
if gstin and str(gstin).lower() != "nan":
    chips.append(("GSTIN", str(gstin)))
meta_chips(chips)
st.caption(
    f"{len(result.mapped_columns)} Excel columns mapped to the dashboard. "
    "All KPIs, charts, and AI insights below are computed from this uploaded file."
)
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# KPI cards
# ---------------------------------------------------------------------------
kpi_grid([
    ("Total Sales", format_indian_rupees(stats['total_sales']), f"Avg/Month: {format_indian_rupees(stats['avg_monthly_sales'])}"),
    ("Total GST Liability", format_indian_rupees(stats['total_gst_liability']), f"Eff. Rate: {stats['avg_effective_tax_rate']:.2f}%"),
    ("Total ITC Available", format_indian_rupees(stats['total_itc_available']), f"Net ITC: {format_indian_rupees(stats['net_itc'])}"),
    ("Total Tax Paid", format_indian_rupees(stats['total_tax_paid']), f"Cash Dep: {stats['cash_dependency']:.2f}%"),
])

st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tab_dashboard, tab_insights, tab_data, tab_report = st.tabs(
    ["📊 Dashboard & Charts", "🤖 AI Insights", "📋 Raw Data", "📄 View Report"]
)

n_months = len(df)

# --- Dashboard & Charts -------------------------------------------------
with tab_dashboard:
    col1, col2 = st.columns(2)
    
    with col1:
        # Calculate ticks for ITC Available (left y-axis)
        max_val_itc = df["total_itc_available"].fillna(0).max()
        tickvals_itc, ticktext_itc = get_indian_ticks(max_val_itc)

        # Calculate ticks for ITC Reversal (right y-axis)
        max_val_rev = df["total_itc_reversed"].fillna(0).max()
        tickvals_rev, ticktext_rev = get_indian_ticks(max_val_rev)

        fig_itc_combined = make_subplots(specs=[[{"secondary_y": True}]])
        
        # Add bar chart on primary axis
        fig_itc_combined.add_trace(
            go.Bar(
                x=df["month_label"],
                y=df["total_itc_available"],
                name="Total ITC Available",
                marker_color=NAVY,
                customdata=[format_indian_rupees(v) for v in df["total_itc_available"]],
                hovertemplate="Total ITC Available: %{customdata}<extra></extra>"
            ),
            secondary_y=False
        )
        
        # Add line chart on secondary axis
        fig_itc_combined.add_trace(
            go.Scatter(
                x=df["month_label"],
                y=df["total_itc_reversed"],
                name="ITC Reversal",
                mode="lines+markers",
                marker=dict(color="#C0392B", size=6),
                line=dict(color="#C0392B", width=2),
                customdata=[format_indian_rupees(v) for v in df["total_itc_reversed"]],
                hovertemplate="ITC Reversal: %{customdata}<extra></extra>"
            ),
            secondary_y=True
        )

        fig_itc_combined.update_layout(
            title="ITC Available vs Reversal Trend (₹)"
        )
        
        # Update primary y-axis
        fig_itc_combined.update_yaxes(
            title_text="Amount (₹)",
            tickvals=tickvals_itc,
            ticktext=ticktext_itc,
            secondary_y=False
        )
        
        # Update secondary y-axis (disable grid lines to prevent overlapping lines)
        fig_itc_combined.update_yaxes(
            title_text="Reversal Amount (₹)",
            tickvals=tickvals_rev,
            ticktext=ticktext_rev,
            showgrid=False,
            secondary_y=True
        )
        
        style_chart(fig_itc_combined, height=380, n_labels=n_months)
        st.plotly_chart(fig_itc_combined, use_container_width=True)
        
    with col2:
        # Calculate ticks for Total Tax Paid (stacked)
        max_val_tax = (df["total_itc_utilized"].fillna(0) + df["total_cash_paid"].fillna(0)).max()
        tickvals_tax, ticktext_tax = get_indian_ticks(max_val_tax)

        fig_tax_paid = go.Figure()
        fig_tax_paid.add_trace(
            go.Bar(
                x=df["month_label"],
                y=df["total_itc_utilized"],
                name="ITC Offset",
                marker_color=GOLD,
                customdata=[format_indian_rupees(v) for v in df["total_itc_utilized"]],
                hovertemplate="ITC Offset: %{customdata}<extra></extra>"
            )
        )
        fig_tax_paid.add_trace(
            go.Bar(
                x=df["month_label"],
                y=df["total_cash_paid"],
                name="Cash Ledger",
                marker_color=NAVY,
                customdata=[format_indian_rupees(v) for v in df["total_cash_paid"]],
                hovertemplate="Cash Ledger: %{customdata}<extra></extra>"
            )
        )
        fig_tax_paid.update_layout(
            barmode="stack",
            title="Total Tax Paid (Cash vs ITC Offset) Trend (₹)",
            yaxis_title="Tax Paid (₹)"
        )
        fig_tax_paid.update_yaxes(tickvals=tickvals_tax, ticktext=ticktext_tax)
        style_chart(fig_tax_paid, height=380, n_labels=n_months)
        st.plotly_chart(fig_tax_paid, use_container_width=True)

# --- AI Insights -----------------------------------------------------------
with tab_insights:
    yearly_text = yearly_df.to_string(index=False)

    if is_sarvam_configured():
        st.caption("Sarvam AI connected via `.env` — insights use your uploaded Excel data")
    else:
        st.warning(
            "Sarvam AI is not configured. Add `SARVAM_API_KEY` to your `.env` file "
            "to enable live AI insights; showing rule-based fallback for now.",
            icon="⚠️",
        )

    if (
        "ai_insight" not in st.session_state
        or st.session_state.get("ai_insight_fingerprint") != data_fingerprint
    ):
        with st.spinner("Generating narrative insights from your uploaded data..."):
            insight = generate_dashboard_insight(stats, yearly_text, monthly_text, trade_name)
        st.session_state["ai_insight"] = insight
        st.session_state["ai_insight_fingerprint"] = data_fingerprint

    insight = st.session_state["ai_insight"]
    badge_class = "badge-ai" if insight.source == "sarvam_ai" else "badge-fallback"
    badge_text = "Sarvam AI" if insight.source == "sarvam_ai" else "Fallback (local)"
    st.markdown(f'<span class="badge {badge_class}">{badge_text}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="insight-box">{insight.text}</div>', unsafe_allow_html=True)

    if insight.error and insight.source == "fallback_template":
        with st.expander("Why fallback was used"):
            st.code(insight.error)

    if st.button("🔄 Regenerate insight"):
        with st.spinner("Regenerating..."):
            st.session_state["ai_insight"] = generate_dashboard_insight(
                stats, yearly_text, monthly_text, trade_name
            )
            st.session_state["ai_insight_fingerprint"] = data_fingerprint
        st.rerun()

    st.markdown("---")
    st.markdown("#### Ask a question about your GST data")

    if "custom_q_answer" not in st.session_state:
        st.session_state["custom_q_answer"] = None
        st.session_state["custom_q_question"] = ""

    with st.form(key="ask_form", clear_on_submit=False):
        question = st.text_input("e.g. Why did my tax liability spike in March 2025?", key="custom_q_input")
        submit_button = st.form_submit_button(label="Ask AI")

    if submit_button and question.strip():
        with st.spinner("Thinking..."):
            answer = generate_custom_insight(question, stats, monthly_text, trade_name)
            st.session_state["custom_q_answer"] = answer
            st.session_state["custom_q_question"] = question

    if st.session_state["custom_q_answer"] is not None:
        answer = st.session_state["custom_q_answer"]
        q_asked = st.session_state["custom_q_question"]
        st.markdown(f"**Question Asked:** {q_asked}")
        answer_badge = "Sarvam AI" if answer.source == "sarvam_ai" else "Fallback"
        st.markdown(f'<span class="badge {"badge-ai" if answer.source == "sarvam_ai" else "badge-fallback"}">{answer_badge}</span>', unsafe_allow_html=True)
        st.markdown(f'<div class="insight-box">{answer.text}</div>', unsafe_allow_html=True)
        if answer.error and answer.source == "fallback_template":
            with st.expander("Why fallback was used"):
                st.code(answer.error)

# --- Raw Data ---------------------------------------------------------------
with tab_data:
    st.markdown(f"#### Data from `{uploaded_file.name}` → sheet `{result.sheet_used}`")
    
    # Summary KPI cards for Raw Data
    total_months_count = stats.get("n_periods", 0)
    total_penalty_value = stats.get("total_late_fee", 0.0)
    kpi_grid([
        ("Total Duration", f"{total_months_count} Months", "Total periods in dataset"),
        ("Total GSTR-3B Penalty", format_indian_rupees(total_penalty_value), "Cumulative filing penalties"),
    ])
    st.markdown("<br/>", unsafe_allow_html=True)

    st.markdown("##### Normalized monthly data (computed from your Excel rows)")
    display_cols = [c for c in [
        "month_label", "financial_year", "taxable_value", "total_outward_tax",
        "total_itc_available", "total_itc_utilized", "total_cash_paid",
        "effective_tax_rate_pct", "filing_delay_days", "late_fee",
    ] if c in df.columns]
    
    monthly_display_df = df[display_cols].copy()
    monthly_rename_dict = {
        "month_label": "Month",
        "financial_year": "Financial Year",
        "taxable_value": "Taxable Value (Turnover)",
        "total_outward_tax": "Outward Tax (Output)",
        "total_itc_available": "ITC Available",
        "total_itc_utilized": "ITC Utilized",
        "total_cash_paid": "Cash Paid (Cash Ledger)",
        "effective_tax_rate_pct": "Eff. Tax Rate",
        "filing_delay_days": "Filing Delay (Days)",
        "late_fee": "GSTR-3B Penalty"
    }
    monthly_display_df = monthly_display_df.rename(columns={k: v for k, v in monthly_rename_dict.items() if k in monthly_display_df.columns})
    st.dataframe(style_dataframe_indian_rupees(monthly_display_df), use_container_width=True, height=400)

    st.markdown("##### All parsed columns from the uploaded file")
    st.dataframe(style_dataframe_indian_rupees(df), use_container_width=True, height=300)

    with st.expander("Excel column mapping (source header → internal field)"):
        mapping_df = pd.DataFrame([
            {"Excel column": raw, "Mapped to": canon}
            for raw, canon in sorted(result.mapped_columns.items(), key=lambda x: x[1])
        ])
        st.dataframe(mapping_df, use_container_width=True, hide_index=True)
        if result.unmapped_columns:
            st.caption(f"Unmapped columns (ignored): {', '.join(result.unmapped_columns[:20])}"
                       + (" …" if len(result.unmapped_columns) > 20 else ""))

    st.markdown("##### Year-wise summary")
    display_yearly_df = yearly_df.copy()
    if not display_yearly_df.empty:
        total_row = pd.Series({
            "financial_year": "Total",
            "months_filed": display_yearly_df["months_filed"].sum(),
            "total_turnover": display_yearly_df["total_turnover"].sum(),
            "total_outward_tax": display_yearly_df["total_outward_tax"].sum(),
            "total_itc_available": display_yearly_df["total_itc_available"].sum(),
            "total_itc_utilized": display_yearly_df["total_itc_utilized"].sum(),
            "total_cash_paid": display_yearly_df["total_cash_paid"].sum(),
        })
        if "total_gstr3b_filing_penalty" in display_yearly_df.columns:
            total_row["total_gstr3b_filing_penalty"] = display_yearly_df["total_gstr3b_filing_penalty"].sum()
        
        total_row["effective_tax_rate_pct"] = (total_row["total_outward_tax"] / total_row["total_turnover"] * 100) if total_row["total_turnover"] > 0 else np.nan
        display_yearly_df = pd.concat([display_yearly_df, pd.DataFrame([total_row])], ignore_index=True)
        
    yearly_rename_dict = {
        "financial_year": "Financial Year",
        "months_filed": "Months Filed",
        "total_turnover": "Turnover",
        "total_outward_tax": "Outward Tax (Output)",
        "total_itc_available": "ITC Available",
        "total_itc_utilized": "ITC Utilized",
        "total_cash_paid": "Cash Paid",
        "effective_tax_rate_pct": "Eff. Tax Rate",
        "total_gstr3b_filing_penalty": "GSTR-3B Penalty"
    }
    display_yearly_df = display_yearly_df.rename(columns={k: v for k, v in yearly_rename_dict.items() if k in display_yearly_df.columns})
    st.dataframe(style_dataframe_indian_rupees(display_yearly_df), use_container_width=True)

    if result.employment_info:
        st.markdown("#### Employment details (from MSME/Udyam sheet)")
        st.json(result.employment_info)

    if result.msme_info is not None:
        st.markdown("#### MSME / Udyam investment & turnover history")
        st.dataframe(style_dataframe_indian_rupees(result.msme_info), use_container_width=True)

# --- View Report -------------------------------------------------------------
with tab_report:
    from datetime import datetime

    insight = st.session_state.get("ai_insight")
    if insight is None or st.session_state.get("ai_insight_fingerprint") != data_fingerprint:
        with st.spinner("Generating insights..."):
            insight = generate_dashboard_insight(
                stats, yearly_df.to_string(index=False), monthly_text, trade_name
            )
            st.session_state["ai_insight"] = insight
            st.session_state["ai_insight_fingerprint"] = data_fingerprint

    # Build HTML for report
    date_start, date_end = stats.get("date_range", (None, None))
    period_text = f"{date_start.strftime('%b %Y')} to {date_end.strftime('%b %Y')}" if (date_start and date_end) else ""
    gstin_text = f"GSTIN: {gstin}" if (gstin and str(gstin).lower() != "nan") else ""
    
    # KPI rows
    kpi_rows_data = [
        ("Total Sales", format_indian_rupees(stats.get('total_sales', 0))),
        ("Total GST Liability", format_indian_rupees(stats.get('total_gst_liability', 0))),
        ("Average Monthly Sales", format_indian_rupees(stats.get('avg_monthly_sales', 0))),
        ("Total ITC Available", format_indian_rupees(stats.get('total_itc_available', 0))),
        ("Net ITC", format_indian_rupees(stats.get('net_itc', 0))),
        ("ITC Efficiency", f"{stats.get('itc_efficiency', 0):.2f}%"),
        ("Cash Dependency", f"{stats.get('cash_dependency', 0):.2f}%"),
        ("Reverse Charge %", f"{stats.get('rcm_ratio_pct', 0):.2f}%"),
    ]
    if "filing_date" in df.columns and df["filing_date"].notna().any():
        kpi_rows_data.append(("Total Late Fees / Penalty", format_indian_rupees(stats.get('total_late_fee', 0))))
        
    kpi_rows_html = ""
    for label, val in kpi_rows_data:
        kpi_rows_html += f"""
        <tr>
            <td style="font-weight: 500;">{label}</td>
            <td style="text-align: right; font-family: monospace; white-space: nowrap;">{val}</td>
        </tr>
        """

    # Yearly summary rows
    yearly_rows_html = ""
    for _, row in yearly_df.iterrows():
        outward_tax_pct = (row['total_outward_tax'] / row['total_turnover'] * 100) if row['total_turnover'] > 0 else 0.0
        itc_utilized_pct = (row['total_itc_utilized'] / row['total_outward_tax'] * 100) if row['total_outward_tax'] > 0 else 0.0

        yearly_rows_html += f"""
        <tr>
            <td style="text-align: center;">{row['financial_year']}</td>
            <td style="text-align: center;">{int(row['months_filed'])}</td>
            <td style="text-align: right; font-family: monospace; white-space: nowrap;">{format_indian_rupees(row['total_turnover'])}</td>
            <td style="text-align: right; font-family: monospace; white-space: nowrap;">
                {format_indian_rupees(row['total_outward_tax'])}<br/>
                <span style="font-size: 0.8em; color: #64748B;">({outward_tax_pct:.2f}%)</span>
            </td>
            <td style="text-align: right; font-family: monospace; white-space: nowrap;">
                {format_indian_rupees(row['total_itc_utilized'])}<br/>
                <span style="font-size: 0.8em; color: #64748B;">({itc_utilized_pct:.2f}%)</span>
            </td>
            <td style="text-align: right; font-family: monospace; white-space: nowrap;">{format_indian_rupees(row['total_cash_paid'])}</td>
        </tr>
        """

    compliance_rows_html = ""
    if "filing_date" in df.columns and df["filing_date"].notna().any():
        for _, row in df.iterrows():
            due_dt = row['due_date'].strftime('%d %b %Y') if pd.notna(row['due_date']) else "N/A"
            fil_dt = row['filing_date'].strftime('%d %b %Y') if pd.notna(row['filing_date']) else "N/A"
            delay = int(row['filing_delay_days']) if pd.notna(row['filing_delay_days']) else 0
            delay_str = f"{delay} days" if delay > 0 else "On Time"
            penalty = format_indian_rupees(row['late_fee']) if pd.notna(row['late_fee']) else "₹ 0.00"
            compliance_rows_html += f"""
            <tr>
                <td style="text-align: center;">{row['month_label']}</td>
                <td style="text-align: center; white-space: nowrap;">{due_dt}</td>
                <td style="text-align: center; white-space: nowrap;">{fil_dt}</td>
                <td style="text-align: center; white-space: nowrap;">{delay_str}</td>
                <td style="text-align: right; font-family: monospace; white-space: nowrap;">{penalty}</td>
            </tr>
            """
        
        total_months = len(df)
        total_penalty = stats.get('total_late_fee', 0)
        compliance_rows_html += f"""
        <tr style="font-weight: bold; background-color: #F8FAFC; border-top: 2px solid #E2E8F0;">
            <td style="text-align: center;">Total ({total_months} Months)</td>
            <td style="text-align: center;"></td>
            <td style="text-align: center;"></td>
            <td style="text-align: center;"></td>
            <td style="text-align: right; font-family: monospace; white-space: nowrap;">{format_indian_rupees(total_penalty)}</td>
        </tr>
        """
    else:
        compliance_rows_html = """
        <tr>
            <td colspan="5" style="text-align: center; color: #64748B;">No filing dates available in the uploaded data.</td>
        </tr>
        """

    # Format AI Insights Narrative
    paragraphs = [p.strip() for p in insight.text.split("\n\n") if p.strip()]
    narrative_html = ""
    sections_list = [
        "Monthly Sales Trend",
        "GST Liability Trend",
        "ITC Efficiency",
        "Cash Flow Analysis",
        "Reverse Charge Analysis",
        "Filing Compliance",
    ]
    for para in paragraphs:
        matched_section = None
        for sec in sections_list:
            if para.startswith(sec) or para.lower().startswith(sec.lower()):
                matched_section = sec
                break
        
        if matched_section:
            lines = para.split("\n")
            sec_heading = lines[0].strip()
            content = "\n".join(lines[1:]).strip() if len(lines) > 1 else ""
            narrative_html += f"""
            <h4 style="color: {GOLD}; margin-top: 20px; margin-bottom: 8px; font-size: 16px; font-weight: 600;">{sec_heading}</h4>
            """
            if content:
                formatted_content = content.replace('\n', '<br/>')
                narrative_html += f"""
                <p style="margin-bottom: 12px; line-height: 1.6;">{formatted_content}</p>
                """
        else:
            formatted_para = para.replace('\n', '<br/>')
            narrative_html += f"""
            <p style="margin-bottom: 12px; line-height: 1.6;">{formatted_para}</p>
            """

    source_badge = "Generated by Sarvam AI" if insight.source == "sarvam_ai" else "Generated locally (Sarvam AI unavailable)"
    
    report_html = f"""
    <div class="report-preview-container">
        <div class="report-preview-header">
            <h1 class="report-preview-title">GST Insight AI</h1>
            <h2 class="report-preview-subtitle">GST Performance & Compliance Report</h2>
            <div class="report-preview-meta" style="font-weight: 600; color: {NAVY}; font-size: 16px; margin-bottom: 12px;">{trade_name}</div>
            {"<div class='report-preview-meta'><strong>GSTIN:</strong> " + gstin_text.split(": ")[1] + "</div>" if gstin_text else ""}
            {"<div class='report-preview-meta'><strong>Period Covered:</strong> " + period_text + "</div>" if period_text else ""}
            <div class="report-preview-meta"><strong>Report Date:</strong> {datetime.now().strftime('%d %B %Y')}</div>
        </div>
        
        <h3 class="report-preview-section-title">Executive Summary</h3>
        <table class="report-preview-table">
            <thead>
                <tr>
                    <th>Metric</th>
                    <th style="text-align: right;">Value</th>
                </tr>
            </thead>
            <tbody>
                {kpi_rows_html}
            </tbody>
        </table>
        
        <h3 class="report-preview-section-title">Year-wise Performance</h3>
        <div style="overflow-x: auto;">
            <table class="report-preview-table" style="min-width: 600px;">
                <thead>
                    <tr>
                        <th style="text-align: center;">Financial Year</th>
                        <th style="text-align: center;">Months Filed</th>
                        <th style="text-align: right;">Turnover</th>
                        <th style="text-align: right;">
                            Output Tax<br/>
                            <span style="font-size: 0.75em; font-weight: normal; opacity: 0.85; display: block; margin-top: 4px;">(Output Tax ÷ Turnover × 100)</span>
                        </th>
                        <th style="text-align: right;">
                            ITC Utilized<br/>
                            <span style="font-size: 0.75em; font-weight: normal; opacity: 0.85; display: block; margin-top: 4px;">(ITC Utilized ÷ Output Tax × 100)</span>
                        </th>
                        <th style="text-align: right;">Cash Paid</th>
                    </tr>
                </thead>
                <tbody>
                    {yearly_rows_html}
                </tbody>
            </table>
        </div>

        <h3 class="report-preview-section-title">GSTR-3B Filing Compliance & Penalty Details</h3>
        <div style="overflow-x: auto;">
            <table class="report-preview-table" style="min-width: 600px;">
                <thead>
                    <tr>
                        <th style="text-align: center;">Return Period</th>
                        <th style="text-align: center;">Due Date</th>
                        <th style="text-align: center;">Filing Date</th>
                        <th style="text-align: center;">Delay</th>
                        <th style="text-align: right;">Penalty (Late Fee)</th>
                    </tr>
                </thead>
                <tbody>
                    {compliance_rows_html}
                </tbody>
            </table>
        </div>
        
        <h3 class="report-preview-section-title">AI-Generated Narrative Insights</h3>
        <div style="font-size: 12px; font-style: italic; color: #64748B; margin-bottom: 16px;">
            {source_badge}
        </div>
        <div class="report-preview-insights">
            {narrative_html}
        </div>
        
        <div class="report-preview-conclusion">
            This report summarizes {stats.get('n_periods')} GST return periods for {trade_name}. 
            All figures are derived directly from the uploaded return data. This report is intended 
            for internal review and does not constitute tax or legal advice; please consult a qualified 
            Chartered Accountant before acting on any observation in this report.
        </div>
    </div>
    """
    
    st.html(report_html)

# ---------------------------------------------------------------------------
# Word report download
# ---------------------------------------------------------------------------
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
st.markdown("### 📄 Download Full Report")
st.write("Generate a polished Word document with the executive summary, year-wise table, "
         "and AI narrative insights — ready to share or archive.")

if st.button("Generate Word Report", type="primary"):
    with st.spinner("Building report..."):
        insight = st.session_state.get("ai_insight")
        if insight is None or st.session_state.get("ai_insight_fingerprint") != data_fingerprint:
            insight = generate_dashboard_insight(
                stats, yearly_df.to_string(index=False), monthly_text, trade_name
            )
        buffer = build_word_report(
            trade_name=trade_name,
            gstin=gstin,
            stats=stats,
            yearly_df=yearly_df,
            insight_text=insight.text,
            insight_source=insight.source,
            df=df,
        )
    st.download_button(
        label="⬇️ Download GST_Insight_Report.docx",
        data=buffer,
        file_name=f"GST_Insight_Report_{trade_name.replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
