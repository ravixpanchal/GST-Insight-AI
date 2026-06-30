# GST Insight AI — Streamlit Dashboard

Upload a GST returns Excel workbook (GSTR-3B style monthly export) and get
an instant dashboard: KPIs, charts, year-wise breakdown, AI-generated
narrative insights via **Sarvam AI**, and a downloadable Word report.

## Features

- **Flexible column mapping** — doesn't require an exact spreadsheet template.
  Recognizes common GSTN/CA-export column naming variants (e.g.
  `sup_details_osup_det_txval`, `Taxable Value`, `Outward Taxable Value`, etc.)
  and maps them onto a stable internal schema.
- **KPI dashboard** — turnover, output tax, ITC available/utilized, cash tax
  paid, effective tax rate, filing punctuality.
- **Year-wise Performance** — shows financial year-wise summaries of turnover, output tax, ITC utilized, and cash paid. The table headers explicitly state the percentage formulas used:
  - Output Tax Rate: `(Output Tax ÷ Turnover × 100)`
  - ITC Utilization Rate: `(ITC Utilized ÷ Output Tax × 100)`
- **Charts** — turnover & tax trends, tax composition, ITC utilization,
  payment mix (cash vs ITC), filing delay timeline.
- **AI narrative insights** — calls Sarvam AI's chat completion API to
  generate a plain-English analysis of the business's GST performance.
  Falls back to a rule-based template automatically if no API key is set or
  the API call fails, so the dashboard never breaks.
- **Ask-a-question** — free-text Q&A against your own summarized GST metrics.
- **Word report export** — polished `.docx` report (navy/gold themed) with
  executive summary table, year-wise table (with formulas), and the AI narrative, ready to
  download and share.
- **MSME/Udyam panel** — if the workbook includes an Udyam-style supplementary
  sheet (employment headcount, investment & turnover history), it's parsed
  and shown in the Raw Data tab.

## Setup

```bash
pip install -r requirements.txt
```

### Sarvam AI API key (optional but recommended)

Copy the example env file and add your key:

```bash
cp .env.example .env
# Edit .env and set SARVAM_API_KEY=your_key_here
streamlit run app.py
```

The key is read from `.env` at startup — it is never shown in the UI or written to disk by the app.

Without a key, the **AI Insights** tab and the Word report still work —
they use a rule-based fallback narrative generated from your actual
computed metrics instead of a model-generated one.

## Running

```bash
streamlit run app.py
```

Then open the local URL shown in the terminal (typically
`http://localhost:8501`) and upload your GST Excel file from the sidebar.

Every KPI, chart, table, and AI insight is computed live from the file you
upload. Upload a different Excel file and the entire dashboard updates
automatically.

## Expected input format

A monthly return-level sheet, one row per filing period, with columns for:

- Period identifiers: GSTIN, trade name, financial year, return period, filing date
- Outward supply values & tax (IGST/CGST/SGST/Cess)
- Input tax credit (ITC) available, net, and reversed
- Tax paid — cash ledger and ITC-offset breakdown

Column names don't need to match exactly — see `data_loader.py`'s
`CANONICAL_ALIASES` dict to add more aliases if your export uses different
header names.

## Project structure

```
app.py                 Streamlit UI — dashboard, tabs, charts, report button
data_loader.py         Excel parsing, column-alias mapping, metric calculations
sarvam_client.py       Sarvam AI API wrapper with graceful fallback
report_generator.py    python-docx Word report builder
requirements.txt
.streamlit/config.toml Theme (navy/gold)
```

## Notes on assumptions

- GSTR-3B due date is estimated as the **20th of the month following** the
  return period (the standard statutory due date). Actual due dates can
  shift with government notifications/extensions, so filing-delay figures
  are directional, not authoritative.
- This tool is for internal analysis only and does not constitute tax or
  legal advice. Always confirm interpretations with a qualified CA.
