# Controlled Document Formatter

A Streamlit application that converts an unformatted `.docx` into a consistently formatted SOP, Working Guideline, Policy, or IT Policy.

## Main features
- Upload an unformatted DOCX.
- Optionally upload the approved DOCX template.
- Apply Times New Roman 12 pt, justified alignment, 1.5 line spacing, heading hierarchy, bullets, table styling, header/footer, and live `Page X of Y` fields.
- Recognize common controlled-document sections.
- Add an updateable Word table of contents.
- Add `END OF DOCUMENT` when missing.
- Produce a validation report for missing sections, dates, and document-number pattern.
- Keep processing in memory; the app does not intentionally write uploaded files to persistent storage.

## Important validation boundary
This is a formatting aid, not an approval or document-control system. It does not electronically sign, approve, stamp, issue, obsolete, or assign controlled status. For regulated use, validate the application and review every output.

## Run locally
```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy from GitHub to Streamlit Community Cloud
1. Create a GitHub repository.
2. Upload all files from this project, including `.streamlit/config.toml`.
3. Commit and push to the main branch.
4. In Streamlit Community Cloud, choose **Create app**.
5. Select your repository and branch.
6. Set the entrypoint to `app.py`.
7. Select a supported Python version, preferably Python 3.12.
8. Deploy.

## Exact template fidelity
For exact logo, margins, header table, signature area, approved numbering, and annexure design, upload the approved editable Word template. PDF layouts are not suitable as the editable template.

## Known limitations
- Input must be `.docx`; legacy `.doc` and PDF are not accepted.
- Heading recognition is rule-based; ambiguous paragraphs may require manual review.
- Word may require **Update Field** for the table of contents and page-total field.
- The generic header is a fallback and is not a substitute for the approved template.
- First-page approval/signature tables should be embedded in the approved template for exact positioning.

## Repository structure
- `app.py`: Streamlit interface
- `formatter.py`: DOCX transformation logic
- `validator.py`: post-format checks
- `requirements.txt`: deployment dependencies
- `.streamlit/config.toml`: upload and theme configuration
- `tests/test_smoke.py`: basic test
