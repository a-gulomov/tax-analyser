# TradeFlow Analyzer — local starter

## Run
1. Install Python 3.10+
2. Extract this ZIP.
3. Open terminal/cmd in this folder.
4. Run:
   ```bash
   python -m venv .venv
   ```
   Windows:
   ```bash
   .venv\Scripts\activate
   ```
   macOS/Linux:
   ```bash
   source .venv/bin/activate
   ```
5. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
6. Launch:
   ```bash
   streamlit run app.py
   ```
The app opens in your browser. SQLite database (`tradeflow.db`) is created locally.

## Current capabilities
- Upload Excel (`.xlsx`) with multiple sheets or CSV.
- Attempts to recognize invoice columns by Uzbek/Russian/English header aliases.
- Saves imported line items locally in SQLite.
- Filters by year, STIR (seller/buyer), product and direction.
- Tables and basic grouped summaries.
- Export filtered records to CSV/XLSX.
- Duplicate prevention using a fingerprint.

## Notes / limitations
- This is a local prototype, not a production deployment.
- Excel imports are best suited to detailed invoice line sheets. Summary/pivot sheets may not be recognized.
- The workbook's `ТАХЛИЛ` summary sheet has multi-row headers and is not imported as invoice detail; the detailed `УМУМИЙ РУЙХАТ`, `4-илова кирим`, and `5-илова чиким` sheets are the likely import targets.
- Header matching is heuristic; inspect imports and compare against source.
- ONKM/terminal feeds are not connected; only files can be imported for now.
- Data access, sharing, and processing must be authorized. Dashboard results are for informational review, not automatic conclusions.
# tax-analyser
