# Gain Estonia Financials Extractor

Public Streamlit app for retrieving structured Estonian annual-report data from
RIK's SOAP/XML service and filling the Gain EUR financials template.

## Local setup

1. Create a virtual environment and install `requirements.txt`.
2. Set `RIK_USERNAME` and `RIK_PASSWORD` as environment variables, or copy
   `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and fill in
   the values locally.
3. Run `streamlit run app.py`.

Credentials and raw SOAP responses must never be committed or logged. RIK may
echo authentication fields in its XML response, so saved fixtures must be
redacted.

## Data flow

- Consolidated RIK structured statement rows are the primary source whenever
  RIK exposes them.
- If RIK exposes only explicitly unconsolidated statement types but the complete
  annual report contains consolidated statements, the consolidated document
  replaces the whole financial block. The app does not mix consolidated and
  unconsolidated figures within a fiscal year.
- Unconsolidated XML is used only when no consolidated source is available.
- Each historical year prefers the comparative column in the following year's
  annual report (for example, FY2024 comes from AR2025) so later corrections
  and reclassifications are retained. The same-year report is the fallback,
  and still supplies the target year's reporting-period dates for any required
  annualisation.
- PDF or BDOC annual reports supplement missing FTE, segment, and disclosure
  data through the existing Estonia parser.
- The generation report lists the exact mapped items added by document
  fallback and the number of added segment records by dimension.
- A single ZIP download contains the complete RIK annual-report PDF for each
  fiscal year selected for the Excel workbook, when RIK provides one.
- Only directly sourceable raw inputs are written. Template formulas remain
  formulas.
- The template's extended Excel validations are restored after workbook save.
