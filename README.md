# Gain Estonia Financials Extractor

Public Streamlit app for retrieving Estonian annual-report XBRL and statement
data from RIK and filling the Gain EUR financials template.

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

- The full annual-report XBRL package is the primary structured source. The
  parser chooses consolidated facts when the package contains a complete group
  statement block and otherwise uses the standalone facts.
- RIK's statement-line XML API is the second structured source. Explicitly
  unconsolidated statement XML is never mixed into a consolidated XBRL block.
- PDF or BDOC annual reports fill only remaining gaps and narrative disclosures.
- Each historical year prefers the comparative column in the following year's
  annual report (for example, FY2024 comes from AR2025) so later corrections
  and reclassifications are retained. The same-year report is the fallback,
  and still supplies the target year's reporting-period dates for any required
  annualisation.
- XBRL note facts supply FTE, goodwill amortisation, trade balances, CAPEX, and
  revenue segmentations where disclosed. PDF or BDOC parsing remains available
  for missing or non-tabular disclosures.
- The generation report lists exact XBRL, statement XML, and PDF/BDOC items plus
  the segment-record counts by source and dimension.
- A single ZIP download contains the complete RIK annual-report PDF for each
  fiscal year selected for the Excel workbook, when RIK provides one.
- Only directly sourceable raw inputs are written. Template formulas remain
  formulas.
- The template's extended Excel validations are restored after workbook save.
