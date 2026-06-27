# Serials Holdings Cataloging Tool

A free, locally-run web application that helps library catalogers automate the most repetitive parts of serial holdings cataloging for Ex Libris Alma.

**Live app:** [serials-catalog-app.streamlit.app](https://serials-catalog-app.streamlit.app/)

Built for Hood College Library (Frederick, MD).

---

## What it does

Given a journal ISSN, the tool automatically retrieves all known volumes and issues from the CrossRef API, then guides the cataloger through building a correctly formatted Excel import file for Alma's **Items Creator** bulk import tool — without inventing or guessing any data.

### Phase 1 — Item Records Generator (current)

1. **Enter an ISSN or journal URL** — the tool extracts the ISSN automatically from any URL
2. **CrossRef lookup** — retrieves all registered volumes, issues, and dates automatically
3. **Holdings range** — select the first and last volume Hood College actually owns; records outside that range are excluded
4. **Fill gaps** — paste manually extracted records for issues not in CrossRef (e.g. older print-only volumes); the tool accepts months, combined months (Jan/Feb, Jan-Mar), and seasons (Spring, Fall/Winter)
5. **Add special items** — add index volumes, supplements, parts, or other non-standard items separately
6. **Review and edit** — full editable table; mark individual issues as "Missing / not owned" to exclude them from the export
7. **Download** — generates a ready-to-import `.xlsx` file with all columns pre-filled: barcodes, enumeration, chronology, descriptions, mms_id, holding_id

### Phase 2 — Pattern Recognition & MARC Generation (planned)

Analyzes the approved item records to detect publication frequency and generates MARC 21 Holdings fields (853/863/866) with plain-English reasoning for every decision.

---

## Key design principles

- **Never fabricates data** — missing months, unknown issues, and coverage gaps are flagged, never guessed
- **Human in the loop** — every output is a proposal; the cataloger reviews and approves before export
- **Free tools only** — no paid APIs, no subscriptions required
- **Standards-compliant** — follows MARC 21 Holdings (MFHD), ANSI/NISO Z39.71, CONSER, and Alma's item import format

---

## Running locally

**Requirements:** Python 3.9+

```bash
# Clone the repo
git clone https://github.com/kfabo1990/serials-cataloging-tool.git
cd serials-cataloging-tool

# Create a virtual environment and install dependencies
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # Mac/Linux

pip install -r requirements.txt

# Run the app
streamlit run app.py
```

The app opens automatically in your browser at `http://localhost:8501`.

---

## Tech stack

| Tool | Purpose |
|---|---|
| Python 3.x | Core language |
| Streamlit | Web UI (runs locally in the browser) |
| requests | CrossRef API calls |
| pandas | Data handling |
| openpyxl | Excel file generation |

All free and open source. No Claude API or any paid service is called by the app itself.

---

## Data sources

- **CrossRef REST API** (`api.crossref.org`) — free, no account required
- Manual entry by the cataloger for records not in CrossRef

---

## Standards

| Standard | Role |
|---|---|
| MARC 21 Holdings Format (MFHD) | Structure of 853/863/866 fields (Phase 2) |
| ANSI/NISO Z39.71 | Holdings statements |
| CONSER Cataloging Manual | Serials-specific cataloging practices |
| Ex Libris Alma item import format | Column names and rules for Excel upload |

---

## License

MIT — see [LICENSE](LICENSE)
