# Gain Estonia Financials Extractor

Public Streamlit app for retrieving structured Estonian annual-report data from RIK's SOAP/XML service and filling the Gain EUR financials template.

## Local setup

1. Create a virtual environment and install `requirements.txt`.
2. Set `RIK_USERNAME` and `RIK_PASSWORD` as environment variables, or copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and fill in the values locally.
3. Run `streamlit run app.py`.

Credentials and raw SOAP responses must never be committed or logged. RIK may echo authentication fields in its XML response, so saved fixtures must be redacted.

## Data flow

- RIK structured statement rows are the primary source.
- PDF or BDOC annual reports supplement missing FTE, segment, and disclosure data through the existing Estonia parser.
- Only directly sourceable raw inputs are written. Template formulas remain formulas.
- The template's extended Excel validations are restored after workbook save.
