# RFQ compliance automation - step 1 (extract + load)

PDF quotations -> structured rows -> the compliance workbook. AI extracts; the sheet rules classify; humans decide.

## Setup (uses your Claude subscription, no API credits)
    pip install -r requirements.txt
    # Windows PowerShell: install Claude Code, then open a NEW window
    irm https://claude.ai/install.ps1 | iex
    claude --version
    claude            # first run only: log in with your Claude account (Pro/Max), then type /exit

Extraction runs `claude -p` in the background with your login, so it counts against your plan's usage limits.
If ANTHROPIC_API_KEY is set on your machine, the pipeline ignores it for this mode.
Optional pay-per-use mode: add `--backend api` and set ANTHROPIC_API_KEY.

## Run
    # 1. Extract (one Claude Code call per quotation, pages from the manifest)
    python run.py extract --workbook P01113_template.xlsx --manifest manifest_p01113.json \
        --pdf "Binder_aluminium_33.pdf" --out extracted/
    #    add --dry-run to check page ranges and request sizes without calling the API

    # 2. Load into the workbook (append-only; a returning supplier gets the next revision)
    python run.py load --workbook P01113_template.xlsx --manifest manifest_p01113.json \
        --json "extracted/*.json" --out P01113_result.xlsx

    # 3. Open in Excel. Review P01113_result_review_queue.csv, then work the Compliance Matrix.

    python run.py check --workbook P01113_result.xlsx   # sheet formulas vs Python rules (recalculated file)
    python test_pipeline.py                              # RECALC=<path to recalc.py> adds the full recalculation test

## How review works
- Every extracted value carries confidence, source page and quote (Submissions columns R-S).
- Confidence below the threshold (Baseline!B9, default 80%) turns a GREEN line YELLOW with the flag "Low AI confidence - verify".
- After checking the value against the PDF, fix column G if needed and clear column R. The line then follows the rules.
- Exceptions still need an Accepted decision with an approver on the Exceptions tab. Nothing is auto-accepted.

## For a new RFQ
1. Copy the template; replace the Baseline header, BOQ items and requirement lines.
2. Write a manifest (supplier -> pages). Separate PDFs per supplier also work: give each entry `"file"` instead of using `--pdf`.
3. Extract, load, review.

## Limits
- 8 suppliers, 30 requirement lines, 200 submission rows, unit rates for BOQ items 1-3 on the Evaluation tab.
- Bank details and IBANs on quotation pages are sent to the API as part of the PDF. Redaction is not built yet.
- Extraction quality on the 7 binder PDFs has not been measured yet; the test suite uses fixtures and a mocked API response.

## Streamlit app (upload up to 8 PDFs, compare, download Excel)

Run on your PC (uses your Claude subscription through Claude Code, no API credits):

    pip install -r requirements.txt
    streamlit run app.py

Open http://localhost:8501. Steps: 1 check baseline, 2 upload PDFs and name suppliers, 3 Run comparison,
4 correct extracted values (tick Reviewed), 5 exceptions (Accepted + approver clears a line), 6 rates, then Build Excel and Download.

Online (a link others can open): put these files on GitHub and deploy on Streamlit Community Cloud
(share.streamlit.io, main file `app.py`). Add secrets `ANTHROPIC_API_KEY` and `APP_PASSWORD`
(see `.streamlit/secrets.toml.example`). Online use needs an API key: a personal subscription cannot be used by a hosted app.
Quotations (incl. bank details) are sent to Claude, so always set APP_PASSWORD.
Limits: 8 suppliers, 3 BOQ items, 30 lines, one PDF per supplier, 100 pages per PDF on the API.
Tests: `python test_engine.py`, `python test_app.py`, `python test_pipeline.py`.
