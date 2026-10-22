import io
from pathlib import Path
import streamlit as st
from formatter import FormatOptions, format_document
from validator import validate_document

st.set_page_config(page_title="Controlled Document Formatter", page_icon="📄", layout="wide")
st.title("Controlled Document Formatter")
st.caption("Convert an unformatted DOCX into a consistently formatted SOP, WG, Policy, or IT Policy.")

with st.sidebar:
    st.header("Document control")
    document_type = st.selectbox("Document type", ["SOP", "WG", "Policy", "IT Policy"])
    company_name = st.text_input("Company name", value="Company Name")
    department = st.text_input("Department", value="Pharmacovigilance")
    document_number = st.text_input("Document number", placeholder="XXX/SOP/PVG/01-00")
    effective_date = st.text_input("Effective date", placeholder="DD-MMM-YYYY")
    review_date = st.text_input("Review date", placeholder="DD-MMM-YYYY")
    st.divider()
    add_toc = st.checkbox("Insert/updateable table of contents", value=(document_type == "IT Policy"))
    normalize_sections = st.checkbox("Normalize recognized section headings", value=True)
    add_end_marker = st.checkbox("Add END OF DOCUMENT if missing", value=True)
    preserve_header = st.checkbox("Preserve uploaded template header/footer", value=True)

col1, col2 = st.columns(2)
with col1:
    source = st.file_uploader("1. Upload unformatted Word document", type=["docx"], help="Only .docx is supported.")
with col2:
    template = st.file_uploader("2. Optional approved Word template", type=["docx"], help="Recommended for exact logo, margins, numbering, header, footer, and controlled layout.")

st.info("The app formats content but does not approve, sign, stamp, or assign controlled-document status. Review the output before use.")

if st.button("Format document", type="primary", disabled=source is None, use_container_width=True):
    options = FormatOptions(
        document_type=document_type,
        company_name=company_name.strip(), department=department.strip(),
        document_number=document_number.strip(), effective_date=effective_date.strip(), review_date=review_date.strip(),
        add_toc=add_toc, normalize_sections=normalize_sections, add_end_marker=add_end_marker,
        preserve_template_header_footer=preserve_header,
    )
    try:
        output, report = format_document(source.getvalue(), template.getvalue() if template else None, options)
        validation = validate_document(output, options)
        st.session_state["formatted"] = output
        st.session_state["report"] = report + validation
        st.session_state["name"] = f"FORMATTED_{Path(source.name).stem}.docx"
        st.success("Formatting completed. Review all warnings before controlled use.")
    except Exception as exc:
        st.exception(exc)

if "formatted" in st.session_state:
    report = st.session_state.get("report", [])
    errors = [x for x in report if x[0] == "ERROR"]
    warnings = [x for x in report if x[0] == "WARNING"]
    passed = [x for x in report if x[0] == "PASS"]
    a,b,c=st.columns(3)
    a.metric("Passed checks", len(passed)); b.metric("Warnings", len(warnings)); c.metric("Errors", len(errors))
    with st.expander("Formatting and validation report", expanded=True):
        for level, message in report:
            {"PASS": st.success, "WARNING": st.warning, "ERROR": st.error, "INFO": st.info}.get(level, st.write)(message)
    st.download_button("Download formatted DOCX", data=st.session_state["formatted"], file_name=st.session_state["name"], mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
