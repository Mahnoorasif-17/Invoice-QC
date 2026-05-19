# Invoice QC Tool — Streamlit Edition

A Python/Streamlit rebuild of the original HTML Invoice QC Tool, with **no API
dependency**. PDF data is extracted locally using `pdfplumber`, so the app runs
fully offline once dependencies are installed.

## What changed vs the original

| Component | Original (HTML) | This version (Python) |
| --- | --- | --- |
| UI | Vanilla JS in the browser | Streamlit |
| Excel parsing | SheetJS (`xlsx`) | pandas + openpyxl |
| PDF data extraction | **Claude API** (paid, requires key) | **pdfplumber** (local, free) |
| QC engine | `runQC()` in JS | `qc_engine.run_qc()` in Python — same 6 rules, same date windows |

## File layout

```
invoice_qc/
├── app.py              # Streamlit UI + orchestration
├── qc_engine.py        # Pure QC logic — direct port of the JS rules
├── pdf_extractor.py    # PDF → JSON (replaces the Claude API call)
├── requirements.txt
└── README.md
```

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app opens at `http://localhost:8501` and works in any modern browser.

## Deploy free on Streamlit Community Cloud

1. Push this folder to a GitHub repo (public or private, both work).
2. Go to https://share.streamlit.io → **New app**.
3. Pick your repo, branch, and set the main file path to `app.py`.
4. Click **Deploy**. Streamlit Cloud reads `requirements.txt` and installs everything.

No secrets, no API keys, no environment variables — there's nothing to configure.

## Other deployment options

| Platform | Notes |
| --- | --- |
| **Hugging Face Spaces** | Choose "Streamlit" SDK → upload these files. Free tier works. |
| **Railway / Render / Fly.io** | Use `streamlit run app.py --server.port $PORT --server.address 0.0.0.0` as the start command. |
| **Docker** | `FROM python:3.11-slim` → `pip install -r requirements.txt` → `CMD ["streamlit","run","app.py","--server.address","0.0.0.0"]` |
| **Internal server** | `streamlit run app.py --server.port 80 --server.address 0.0.0.0` (then expose via reverse proxy). |

## ⚠️ Tuning the PDF extractor

PDF layouts vary, and `pdfplumber` is doing pattern matching on extracted text and
tables — it doesn't have the layout-understanding ability the LLM had. The defaults
in `pdf_extractor.py` cover common Redirect Health invoice formats, but **you'll
likely need to tune them once** for your specific PDFs.

Here's the workflow:

1. Run the app on one of your invoices.
2. Expand the result group → click **🔎 Debug — raw extracted text from this PDF**.
3. Compare what pdfplumber extracted against what the QC engine got. The four
   places to tune:
   - `_find_group_id()` — regex patterns for the Group ID field
   - `_find_group_name()` — regex patterns for the group/company name
   - `_find_amount()` — keywords that precede dollar totals (e.g. "Invoice Total")
   - `_classify_table()` — how tables get bucketed into current / charges / credits

Every function is short and commented. Adjusting one or two regex patterns is
usually enough to get a new layout working.

## QC rules (unchanged from the original)

1. **Terminated but billed** — `EmploymentStatus = Terminated` AND `Charge > 0` → ERROR
2. **Active but not billed** — `EmploymentStatus = Active`, `StartDate ≤ invoice month`, valid plan, no charge → ERROR
3. **Future start but billed** — `StartDate > invoice month` AND `Charge > 0` → ERROR
4. **SKU mismatch / PEPM** — EN plan code ≠ invoice plan code = ERROR; same code, different amount = PEPM warning
5. **Missing adjustment charge** — Enrolled in the data window but no adjustment charge in the PDF
6. **Missing adjustment credit** — Coverage ended in the data window but no adjustment credit in the PDF
7. **Late adjustments** — Anything dated before the 60-day boundary needs approval

## Excel format expected

The Excel file should have these columns (matching the original tool):

`Identifier`, `FirstName`, `LastName`, `EmploymentStatus`, `Charge`, `PlanCost`,
`CarrierPlanCode`, `PlanSKUCSV`, `StartDate`, `EnrolledOn`, `EndDate`, `EndedOn`

Rows are grouped by `Identifier`, which is matched against `groupId` extracted from each PDF.
