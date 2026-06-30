"""
data_loader.py
----------------
Loads a GST returns Excel workbook and normalizes it into a clean,
analysis-ready DataFrame.

Design goal: DO NOT hardcode to one exact spreadsheet layout. GST data
exports (from GSTN portals, CAs, or ERP systems) use inconsistent column
names ("sup_details_osup_det_txval" vs "Taxable Value" vs "Outward Taxable
Supply Value", etc). This module maps many possible aliases onto a fixed
set of canonical internal field names, so the rest of the app can work off
one stable schema regardless of the exact source file.
"""

from __future__ import annotations

import io
import re
import warnings
from dataclasses import dataclass, field
from typing import Optional

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# ---------------------------------------------------------------------------
# Canonical schema
# ---------------------------------------------------------------------------
# Each canonical field maps to a list of normalized alias patterns that may
# appear in the source file. Matching is done on a "slug" of the header:
# lowercased, with all non-alphanumeric characters stripped.

CANONICAL_ALIASES: dict[str, list[str]] = {
    # Identity / period
    "gstin": ["gstin", "gstno", "gstnumber"],
    "trade_name": ["tradename", "legalname", "businessname", "name", "companyname"],
    "financial_year": ["financialyear", "fy", "fyear"],
    "ret_period": ["retperiod", "returnperiod", "period", "taxperiod", "month"],
    "filing_date": ["fildt", "filingdate", "filedate", "datefiled", "submissiondate"],

    # Outward supplies (sales side)
    "taxable_value": [
        "supdetailsosupdettxval", "outwardtaxablevalue", "taxablevalue",
        "totaltaxablevalue", "outwardtaxablesupplyvalue",
    ],
    "igst_outward": ["supdetailsosupdetiamt", "outwardigst", "igstoutward", "igstonoutward"],
    "cgst_outward": ["supdetailsosupdetcamt", "outwardcgst", "cgstoutward"],
    "sgst_outward": ["supdetailsosupdetsamt", "outwardsgst", "sgstoutward"],
    "cess_outward": ["supdetailsosupdetcsamt", "outwardcess", "cessoutward"],

    "zero_rated_value": ["supdetailsosupzerotxval", "zeroratedvalue", "exportvalue"],
    "zero_rated_igst": ["supdetailsosupzeroiamt"],
    "zero_rated_cess": ["supdetailsosupzerocsamt"],
    "nil_exempt_value": ["supdetailsosupnilexmptxval", "nilratedvalue", "exemptvalue", "nilexemptvalue"],
    "non_gst_value": ["supdetailsosupnongsttxval", "nongstvalue", "nongstsupplyvalue"],

    "rcm_inward_value": ["supdetailsisuprevtxval", "reversechargevalue", "rcmvalue", "inwardrcmvalue"],
    "rcm_igst": ["supdetailsisupreviamt", "supdetailsisuprevamt", "rcmigst"],
    "rcm_cgst": ["supdetailsisuprevcamt"],
    "rcm_sgst": ["supdetailsisuprevsamt"],
    "rcm_cess": ["supdetailsisuprevcsamt"],

    # ITC available (input tax credit)
    "itc_import_goods_igst": ["itcelgitcavlimpgiamt", "itcimportgoodsigst"],
    "itc_import_services_igst": ["itcelgitcavlimpsiamt", "itcimportservicesigst"],
    "itc_inward_supply_igst": ["itcelgitcavlisrciamt", "itcisrciamt"],
    "itc_inward_supply_cgst": ["itcelgitcavlisrccamt", "itcisrccamt"],
    "itc_inward_supply_sgst": ["itcelgitcavlisrcsamt", "itcisrcsamt"],
    "itc_isd_igst": ["itcelgitcavlisdiamt"],
    "itc_other_igst": ["itcelgitcavlothiamt"],
    "itc_other_cgst": ["itcelgitcavlothcamt"],
    "itc_other_sgst": ["itcelgitcavlothsamt"],

    "itc_net_igst": ["itcelgitcnetiamt", "netitcigst", "itcnetigst"],
    "itc_net_cgst": ["itcelgitcnetcamt", "netitccgst", "itcnetcgst"],
    "itc_net_sgst": ["itcelgitcnetsamt", "netitcsgst", "itcnetsgst"],
    "itc_net_cess": ["itcelgitcnetcsamt", "netitccess", "itcnetcess"],

    "itc_reversed_rule": ["itcelgitcrevruliamt", "itcelgitcrevrulcamt", "itcelgitcrevrulsamt"],
    "itc_reversed_other": ["itcelgitcrevothiamt", "itcelgitcrevothcamt", "itcelgitcrevothsamt"],
    "itc_reclaimed": ["itcelgitcreclaim4d1iamt", "itcelgitcreclaim4d1camt", "itcelgitcreclaim4d1samt", "itcelgitcreclaim4d1csamt"],
    "itc_ineligible": ["itcelgitcinelg4d2iamt", "itcelgitcinelg4d2camt", "itcelgitcinelg4d2samt", "ineligibleitc"],

    # Tax paid - cash ledger
    "igst_paid_cash": ["igstwithcash", "igstcashpaid", "igstpaidincash"],
    "cgst_paid_cash": ["cgstwithcash", "cgstcashpaid", "cgstpaidincash"],
    "sgst_paid_cash": ["sgstwithcash", "sgstcashpaid", "sgstpaidincash"],
    "cess_paid_cash": ["cesswithcash", "cesscashpaid", "cesspaidincash"],

    # Tax paid - ITC offset (credit ledger utilization)
    "igst_offset_igst": ["igstwithigst"],
    "igst_offset_cgst": ["igstwithcgst"],
    "igst_offset_sgst": ["igstwithsgst"],
    "cgst_offset_igst": ["cgstwithigst"],
    "cgst_offset_cgst": ["cgstwithcgst"],
    "sgst_offset_igst": ["sgstwithigst"],
    "sgst_offset_sgst": ["sgstwithsgst"],
    "cess_offset_cess": ["cesswithcess"],
}

REQUIRED_MIN_FIELDS = {"ret_period", "taxable_value"}


def _slug(text: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def _build_reverse_lookup() -> dict[str, str]:
    lookup = {}
    for canon, aliases in CANONICAL_ALIASES.items():
        lookup[_slug(canon)] = canon
        for alias in aliases:
            lookup[_slug(alias)] = canon
    return lookup


_REVERSE_LOOKUP = _build_reverse_lookup()


@dataclass
class LoadResult:
    df: pd.DataFrame
    mapped_columns: dict[str, str] = field(default_factory=dict)
    unmapped_columns: list[str] = field(default_factory=list)
    sheet_used: str = ""
    warnings: list[str] = field(default_factory=list)
    employment_info: Optional[dict] = None
    msme_info: Optional[pd.DataFrame] = None


def _map_columns(raw_columns: list[str]) -> dict[str, str]:
    """Map raw spreadsheet headers -> canonical field names."""
    mapping = {}
    for col in raw_columns:
        slug = _slug(col)
        if slug in _REVERSE_LOOKUP:
            mapping[col] = _REVERSE_LOOKUP[slug]
    return mapping


def _score_header_row(header_row: list[str]) -> tuple[int, set[str]]:
    """Score a candidate header row; higher = more likely to be the real header."""
    mapping = _map_columns(header_row)
    if not mapping:
        return -1, set()
    canon_fields = set(mapping.values())
    score = len(mapping)
    if REQUIRED_MIN_FIELDS.issubset(canon_fields):
        score += 20
    return score, canon_fields


def _find_header_row(raw_df: pd.DataFrame, max_scan_rows: int = 20) -> int:
    """Scan the first rows to find the one that looks like column headers.
    Real GST exports often have title/metadata rows before the actual header."""
    best_idx, best_score = 0, -1
    for i in range(min(max_scan_rows, len(raw_df))):
        header_row = raw_df.iloc[i].astype(str).tolist()
        score, _ = _score_header_row(header_row)
        if score > best_score:
            best_idx, best_score = i, score
    return best_idx


def _find_best_sheet(sheets: dict[str, pd.DataFrame]) -> tuple[str, pd.DataFrame, int]:
    """Pick the sheet that looks most like the monthly return-level data
    (i.e. has the most mappable columns and more than a couple of rows)."""
    best_name, best_df, best_header_idx, best_score = None, None, 0, -1
    for name, df in sheets.items():
        if df.empty:
            continue
        header_idx = _find_header_row(df)
        header_row = df.iloc[header_idx].astype(str).tolist()
        score, _ = _score_header_row(header_row)
        score += 1 if df.shape[0] > header_idx + 3 else 0
        if score > best_score:
            best_name, best_df, best_header_idx, best_score = name, df, header_idx, score
    return best_name, best_df, best_header_idx


def _extract_employment_and_msme(sheets: dict[str, pd.DataFrame], exclude_sheet: str) -> tuple[Optional[dict], Optional[pd.DataFrame]]:
    """Best-effort extraction of an Udyam/MSME-style supplementary sheet
    (employment headcount + investment/turnover history), if present.
    Purely additive — failures here never break the main pipeline."""
    employment_info, msme_df = None, None
    for name, df in sheets.items():
        if name == exclude_sheet or df.empty:
            continue
        flat = df.astype(str).apply(lambda col: col.str.lower())
        text_blob = " ".join(str(v) for v in flat.values.flatten().tolist())

        if employment_info is None and "employment" in text_blob:
            for i in range(len(df) - 1):
                row_vals = [str(v).strip().lower() for v in df.iloc[i].tolist()]
                if "male" in row_vals and "female" in row_vals:
                    try:
                        male_idx = row_vals.index("male")
                        female_idx = row_vals.index("female")
                        total_idx = row_vals.index("total") if "total" in row_vals else None
                        next_row = df.iloc[i + 1]
                        employment_info = {
                            "male": pd.to_numeric(next_row.iloc[male_idx], errors="coerce"),
                            "female": pd.to_numeric(next_row.iloc[female_idx], errors="coerce"),
                            "total": pd.to_numeric(next_row.iloc[total_idx], errors="coerce") if total_idx is not None else None,
                        }
                    except Exception:
                        pass
                    break

        if msme_df is None and "turnover" in text_blob and "financial year" in text_blob:
            for i in range(len(df)):
                row_vals = [str(v).strip().lower() for v in df.iloc[i].tolist()]
                if "financial year" in row_vals:
                    header_idx = i
                    headers = df.iloc[header_idx].tolist()
                    data_rows = df.iloc[header_idx + 1:].copy()
                    data_rows.columns = headers
                    data_rows = data_rows.dropna(how="all")
                    data_rows = data_rows[data_rows.iloc[:, 0].apply(
                        lambda v: str(v).strip().replace(".0", "").isdigit()
                    )]
                    if not data_rows.empty:
                        msme_df = data_rows.reset_index(drop=True)
                        for col in msme_df.columns:
                            col_str = str(col).lower()
                            if any(term in col_str for term in [
                                "written down value", "wdv", "net investment", "total turnover", "net turnover",
                                "pollution control", "export turnover"
                            ]):
                                cleaned = msme_df[col].astype(str).str.replace("₹", "", regex=False) \
                                                                  .str.replace(",", "", regex=False) \
                                                                  .str.replace(" ", "", regex=False) \
                                                                  .str.replace("$", "", regex=False)
                                msme_df[col] = pd.to_numeric(cleaned, errors="coerce")
                    break
    return employment_info, msme_df


def _coerce_period(series: pd.Series) -> pd.Series:
    """ret_period in GSTN exports is typically the first day of the month
    (e.g. 2025-01-01 meaning Jan 2025). Coerce robustly to datetime."""
    return pd.to_datetime(series, errors="coerce", dayfirst=False)


def load_gst_workbook(file_like) -> LoadResult:
    """
    Main entry point. Accepts a file path or a file-like object (e.g. an
    uploaded Streamlit file). Returns a LoadResult with a normalized,
    analysis-ready DataFrame plus metadata about what was found.
    """
    if isinstance(file_like, (bytes, bytearray)):
        excel_obj = io.BytesIO(file_like)
    elif hasattr(file_like, "getvalue"):
        excel_obj = io.BytesIO(file_like.getvalue())
    elif hasattr(file_like, "read"):
        excel_obj = io.BytesIO(file_like.read())
    else:
        excel_obj = file_like

    sheets = pd.read_excel(excel_obj, sheet_name=None, header=None)
    sheet_name, raw_df, header_idx = _find_best_sheet(sheets)
    warn_list = []

    if raw_df is None:
        raise ValueError("No usable sheet found in the uploaded workbook.")

    header_row = raw_df.iloc[header_idx].astype(str).tolist()
    mapping = _map_columns(header_row)
    body = raw_df.iloc[header_idx + 1:].copy()
    if header_idx > 0:
        warn_list.append(
            f"Detected column headers on row {header_idx + 1} (skipped {header_idx} title/metadata row(s))."
        )
    body.columns = header_row
    body = body.dropna(how="all")

    mapped_cols_present = [c for c in body.columns if c in mapping]
    unmapped = [c for c in body.columns if c not in mapping]

    df = pd.DataFrame(index=body.index)
    for raw_col in mapped_cols_present:
        canon = mapping[raw_col]
        col_data = body[raw_col]
        if canon in ("gstin", "trade_name", "financial_year"):
            df[canon] = col_data.astype(str).str.strip()
        elif canon in ("ret_period", "filing_date"):
            df[canon] = _coerce_period(col_data)
        else:
            df[canon] = pd.to_numeric(col_data, errors="coerce")

    missing_required = REQUIRED_MIN_FIELDS - set(df.columns)
    if missing_required:
        raise ValueError(
            f"Could not find required columns in the file: {sorted(missing_required)}. "
            f"Recognized columns: {sorted(mapped_cols_present)}"
        )

    df = df.dropna(subset=["ret_period"]).sort_values("ret_period").reset_index(drop=True)

    numeric_cols = [c for c in df.columns if c not in ("gstin", "trade_name", "financial_year", "ret_period", "filing_date")]
    for c in numeric_cols:
        df[c] = df[c].fillna(0.0)

    df = _derive_metrics(df)

    if "financial_year" not in df.columns or df["financial_year"].isna().all():
        df["financial_year"] = df["ret_period"].apply(_infer_fy)
        warn_list.append("financial_year column was missing or empty; inferred from ret_period.")

    if "filing_date" in df.columns and df["filing_date"].notna().any():
        due_date = (df["ret_period"] + pd.DateOffset(months=1)).apply(lambda d: d.replace(day=20))
        df["due_date"] = due_date
        df["filing_delay_days"] = (df["filing_date"] - due_date).dt.days
        
        # Calculate aggregate turnover per financial year to determine late fee caps
        fy_turnover = df.groupby("financial_year")["total_turnover"].sum().to_dict()
        
        def calculate_row_late_fee(row):
            delay = row["filing_delay_days"]
            if pd.isna(delay) or delay <= 0:
                return 0.0
            
            # Check if Nil return (no turnover and no RCM tax)
            is_nil = (row.get("total_turnover", 0.0) == 0.0) and (row.get("total_rcm_tax", 0.0) == 0.0)
            
            if is_nil:
                rate = 20.0
                cap = 500.0
            else:
                rate = 50.0
                ret_period = row.get("ret_period")
                if pd.notna(ret_period) and ret_period < pd.Timestamp("2021-06-01"):
                    cap = 10000.0
                else:
                    fy = row.get("financial_year", "Unknown")
                    # Find preceding FY
                    prev_fy = "Unknown"
                    match = re.match(r"(\d{4})[-/](\d{2,4})", str(fy).strip())
                    if match:
                        start_year = int(match.group(1))
                        prev_start = start_year - 1
                        prev_end = start_year
                        if len(match.group(2)) == 2:
                            prev_fy = f"{prev_start}-{str(prev_end)[-2:]}"
                        else:
                            prev_fy = f"{prev_start}-{prev_end}"
                    
                    turnover = fy_turnover.get(prev_fy)
                    if turnover is None:
                        turnover = fy_turnover.get(fy, 0.0)
                    
                    if turnover <= 1.5 * 10**7: # 1.5 Crores
                        cap = 2000.0
                    elif turnover <= 5.0 * 10**7: # 5 Crores
                        cap = 5000.0
                    else:
                        cap = 10000.0
            
            return min(delay * rate, cap)
            
        df["late_fee"] = df.apply(calculate_row_late_fee, axis=1)
    else:
        df["filing_delay_days"] = np.nan
        df["late_fee"] = np.nan
        warn_list.append("filing_date not found; on-time filing analysis will be skipped.")

    employment_info, msme_df = _extract_employment_and_msme(sheets, sheet_name)

    return LoadResult(
        df=df,
        mapped_columns=mapping,
        unmapped_columns=unmapped,
        sheet_used=sheet_name,
        warnings=warn_list,
        employment_info=employment_info,
        msme_info=msme_df,
    )


def _infer_fy(ts: pd.Timestamp) -> str:
    if pd.isna(ts):
        return "Unknown"
    year = ts.year
    if ts.month >= 4:
        return f"{year}-{str(year + 1)[-2:]}"
    return f"{year - 1}-{str(year)[-2:]}"


def _safe_sum(df: pd.DataFrame, cols: list[str]) -> pd.Series:
    present = [c for c in cols if c in df.columns]
    if not present:
        return pd.Series(0.0, index=df.index)
    return df[present].sum(axis=1)


def _derive_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Compute the derived business metrics every part of the app relies on."""
    df = df.copy()

    df["total_outward_tax"] = _safe_sum(df, ["igst_outward", "cgst_outward", "sgst_outward", "cess_outward"])
    df["total_turnover"] = _safe_sum(df, ["taxable_value", "zero_rated_value", "nil_exempt_value", "non_gst_value"])

    df["total_itc_available"] = _safe_sum(df, [
        "itc_import_goods_igst", "itc_import_services_igst",
        "itc_inward_supply_igst", "itc_inward_supply_cgst", "itc_inward_supply_sgst",
        "itc_isd_igst", "itc_other_igst", "itc_other_cgst", "itc_other_sgst",
    ])
    df["total_itc_net"] = _safe_sum(df, ["itc_net_igst", "itc_net_cgst", "itc_net_sgst", "itc_net_cess"])
    df["total_itc_reversed"] = _safe_sum(df, ["itc_reversed_rule", "itc_reversed_other", "itc_ineligible"])

    df["total_cash_paid"] = _safe_sum(df, ["igst_paid_cash", "cgst_paid_cash", "sgst_paid_cash", "cess_paid_cash"])
    df["total_itc_utilized"] = _safe_sum(df, [
        "igst_offset_igst", "igst_offset_cgst", "igst_offset_sgst",
        "cgst_offset_igst", "cgst_offset_cgst",
        "sgst_offset_igst", "sgst_offset_sgst",
        "cess_offset_cess",
    ])
    df["total_tax_liability_paid"] = df["total_cash_paid"] + df["total_itc_utilized"]

    # Reverse Charge Mechanism (RCM) metrics
    df["total_rcm_tax"] = _safe_sum(df, ["rcm_igst", "rcm_cgst", "rcm_sgst", "rcm_cess"])
    df["total_gst_liability"] = df["total_outward_tax"] + df["total_rcm_tax"]
    df["rcm_ratio_pct"] = np.where(
        df["total_gst_liability"] > 0,
        (df["total_rcm_tax"] / df["total_gst_liability"]) * 100,
        0.0,
    )

    df["effective_tax_rate_pct"] = np.where(
        df["taxable_value"] > 0,
        (df["total_outward_tax"] / df["taxable_value"]) * 100,
        np.nan,
    )
    df["itc_utilization_ratio_pct"] = np.where(
        df["total_itc_net"] > 0,
        (df["total_itc_utilized"] / df["total_itc_net"]) * 100,
        np.nan,
    )
    df["cash_dependency_ratio_pct"] = np.where(
        df["total_tax_liability_paid"] > 0,
        (df["total_cash_paid"] / df["total_tax_liability_paid"]) * 100,
        np.nan,
    )

    df["mom_turnover_growth_pct"] = df["taxable_value"].pct_change() * 100
    df["mom_tax_growth_pct"] = df["total_outward_tax"].pct_change() * 100

    df["month_label"] = df["ret_period"].dt.strftime("%b %Y")

    return df


def get_summary_stats(df: pd.DataFrame) -> dict:
    """High-level KPI numbers used across the dashboard header and report."""
    latest = df.iloc[-1] if not df.empty else None
    
    total_sales = df["taxable_value"].sum() if not df.empty else 0.0
    total_outward_tax = df["total_outward_tax"].sum() if not df.empty else 0.0
    total_itc_available = df["total_itc_available"].sum() if not df.empty else 0.0
    total_itc_utilized = df["total_itc_utilized"].sum() if not df.empty else 0.0
    total_cash_paid = df["total_cash_paid"].sum() if not df.empty else 0.0
    total_tax_liability_paid = df["total_tax_liability_paid"].sum() if not df.empty else 0.0
    total_rcm_tax = df["total_rcm_tax"].sum() if not df.empty else 0.0
    
    itc_efficiency = (total_itc_utilized / total_itc_available * 100) if total_itc_available > 0 else 0.0
    cash_dependency = (total_cash_paid / total_tax_liability_paid * 100) if total_tax_liability_paid > 0 else 0.0
    total_gst_liability = total_outward_tax + total_rcm_tax
    rcm_ratio_pct = (total_rcm_tax / total_gst_liability * 100) if total_gst_liability > 0 else 0.0

    stats = {
        "n_periods": len(df),
        "date_range": (df["ret_period"].min(), df["ret_period"].max()) if not df.empty else (None, None),
        "total_turnover": total_sales,
        "total_sales": total_sales,
        "total_outward_tax": total_outward_tax,
        "total_gst_liability": total_gst_liability,
        "total_itc_available": total_itc_available,
        "total_itc_utilized": total_itc_utilized,
        "net_itc": df["total_itc_net"].sum() if not df.empty else 0.0,
        "itc_reversal": df["total_itc_reversed"].sum() if not df.empty else 0.0,
        "total_cash_paid": total_cash_paid,
        "total_tax_paid": total_tax_liability_paid,
        "total_rcm_tax": total_rcm_tax,
        "itc_efficiency": itc_efficiency,
        "cash_dependency": cash_dependency,
        "rcm_ratio_pct": rcm_ratio_pct,
        "avg_effective_tax_rate": df["effective_tax_rate_pct"].mean() if not df.empty else 0.0,
        "avg_monthly_turnover": df["taxable_value"].mean() if not df.empty else 0.0,
        "avg_monthly_sales": df["taxable_value"].mean() if not df.empty else 0.0,
        "best_month": df.loc[df["taxable_value"].idxmax(), "month_label"] if not df.empty else None,
        "weakest_month": df.loc[df["taxable_value"].idxmin(), "month_label"] if not df.empty else None,
        "avg_filing_delay_days": df["filing_delay_days"].mean() if "filing_delay_days" in df.columns else np.nan,
        "late_filings_count": int((df["filing_delay_days"] > 0).sum()) if "filing_delay_days" in df.columns else None,
        "total_late_fee": df["late_fee"].sum() if "late_fee" in df.columns else 0.0,
        "latest_month": latest["month_label"] if latest is not None else None,
        "latest_turnover": latest["taxable_value"] if latest is not None else None,
    }
    return stats


def get_monthly_summary_text(df: pd.DataFrame) -> str:
    """Compact month-by-month table from the uploaded file for AI prompts."""
    cols = [
        c for c in [
            "month_label", "financial_year", "taxable_value", "total_outward_tax",
            "total_itc_available", "total_itc_utilized", "total_cash_paid",
            "effective_tax_rate_pct", "filing_delay_days",
        ]
        if c in df.columns
    ]
    if not cols:
        return "(no monthly rows parsed)"
    return df[cols].to_string(index=False)


def get_yearly_summary(df: pd.DataFrame) -> pd.DataFrame:
    agg = df.groupby("financial_year").agg(
        months_filed=("ret_period", "count"),
        total_turnover=("taxable_value", "sum"),
        total_outward_tax=("total_outward_tax", "sum"),
        total_itc_available=("total_itc_available", "sum"),
        total_itc_utilized=("total_itc_utilized", "sum"),
        total_cash_paid=("total_cash_paid", "sum"),
        total_gstr3b_filing_penalty=("late_fee", "sum") if "late_fee" in df.columns else ("ret_period", lambda x: 0.0),
    ).reset_index()
    agg["effective_tax_rate_pct"] = np.where(
        agg["total_turnover"] > 0, (agg["total_outward_tax"] / agg["total_turnover"]) * 100, np.nan
    )
    return agg.sort_values("financial_year")


def format_indian_rupees(amount: float | int, prefix: str = "₹ ") -> str:
    """Format a numeric value into the Indian numbering format (Lakhs/Crores grouping)."""
    if " " in prefix:
        prefix = prefix.replace(" ", "\u00A0")
    if pd.isna(amount) or amount is None:
        return f"{prefix}0.00"
    try:
        val = float(amount)
    except (ValueError, TypeError):
        return str(amount)
        
    is_neg = val < 0
    val = abs(val)
    
    s = f"{val:.2f}"
    parts = s.split('.')
    integer_part = parts[0]
    decimal_part = parts[1] if len(parts) > 1 else "00"
    
    if len(integer_part) <= 3:
        formatted_integer = integer_part
    else:
        last_three = integer_part[-3:]
        remaining = integer_part[:-3]
        rev_remaining = remaining[::-1]
        chunks = [rev_remaining[i:i+2] for i in range(0, len(rev_remaining), 2)]
        formatted_remaining = ",".join(chunks)[::-1]
        
        formatted_integer = f"{formatted_remaining},{last_three}"
        
    res = f"{prefix}{formatted_integer}.{decimal_part}"
    if is_neg:
        res = f"-{res}"
    return res


def format_currency_in_narrative(text: str) -> str:
    """Find numeric currency values in a narrative text and format them to Indian Rupees with commas."""
    if not text:
        return text

    # 1. Replace numbers with currency prefix (Rs., Rs, INR, ₹)
    def replace_prefix(match):
        num_str = match.group(1).replace(',', '')
        try:
            val = float(num_str)
            return format_indian_rupees(val, prefix="₹ ")
        except ValueError:
            return match.group(0)

    # Note: \b before Rs and INR, but not before ₹ (since ₹ is a non-word character)
    text = re.sub(r'(?i)(?:\bRs\.?|\bINR|₹)\s*([0-9]+(?:,[0-9]+)*(?:\.[0-9]+)?)', replace_prefix, text)

    # 2. Replace numbers with currency suffix (rupees, rupee, INR)
    def replace_suffix(match):
        num_str = match.group(1).replace(',', '')
        try:
            val = float(num_str)
            return format_indian_rupees(val, prefix="₹ ")
        except ValueError:
            return match.group(0)

    text = re.sub(r'(?i)\b([0-9]+(?:,[0-9]+)*(?:\.[0-9]+)?)\s*(?:rupees|rupee|\bINR)\b', replace_suffix, text)

    # 3. Replace standalone large numbers that likely represent currency
    # Must be 4 or more digits and not year / percent / compliance metrics
    def replace_standalone(match):
        num_str = match.group(1).replace(',', '')
        if len(num_str) == 4 and (num_str.startswith('19') or num_str.startswith('20')):
            return match.group(0)
        try:
            val = float(num_str)
            return format_indian_rupees(val, prefix="₹ ")
        except ValueError:
            return match.group(0)

    # Match numbers >= 1000 not followed by %, percent, day, filing, month, period, time, year
    pattern = r'(?i)\b([1-9]\d{3,}(?:,\d{3})*(?:\.[0-9]+)?)\b(?!\s*(?:%|percent|percentage|day|days|filing|filings|month|months|period|periods|times|year|years|due))'
    text = re.sub(pattern, replace_standalone, text)

    return text

