"""
Streamlit web UI for the vCard conversion tool.

Run with:
  streamlit run src/web/app.py

Three tabs share the same core library:
  1. Upload File  — drag-drop CSV or VCF
  2. Paste Text   — textarea for messy / free-form data
  3. Google Contacts — upload Google Takeout CSV
"""
from __future__ import annotations

import io
import sys
from pathlib import Path

# Ensure project root is on sys.path when run directly
_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

import pandas as pd
import streamlit as st

from vcard_convert.core.models import Contact, ParseResult, SourceFormat, VCardVersion
from vcard_convert.core.parsers.csv_parser import CSVParser
from vcard_convert.core.parsers.text_parser import FreeTextParser
from vcard_convert.core.parsers.vcf_parser import VCFParser
from vcard_convert.core.writers.vcf_writer import VCFWriter

# ── Page config ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="vCard Converter",
    page_icon="📇",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Sidebar settings ──────────────────────────────────────────────────────────

with st.sidebar:
    st.title("⚙️ Settings")
    country_code = st.text_input("Default country code", value="+91", max_chars=5)
    vcf_version_label = st.radio("vCard version", options=["3.0 (recommended)", "2.1"])
    vcf_version = VCardVersion.V3_0 if vcf_version_label.startswith("3") else VCardVersion.V2_1
    dedupe = st.checkbox("Remove duplicate phone numbers", value=True)

    st.markdown("---")
    st.caption("Low-confidence rows (⚠) are shown in yellow. Review them before downloading.")


# ── Helper functions ──────────────────────────────────────────────────────────

def contacts_to_df(contacts: list[Contact]) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "First Name": [c.first_name for c in contacts],
            "Last Name": [c.last_name for c in contacts],
            "Phones": [", ".join(c.phones) for c in contacts],
            "Emails": [", ".join(c.emails) for c in contacts],
            "Organization": [c.organization for c in contacts],
            "Notes": [c.notes for c in contacts],
            "⚠ Low confidence": [c.confidence < 1.0 for c in contacts],
        }
    )


def df_to_contacts(df: pd.DataFrame) -> list[Contact]:
    contacts = []
    for _, row in df.iterrows():
        phones = [p.strip() for p in str(row.get("Phones", "")).split(",") if p.strip()]
        emails = [e.strip() for e in str(row.get("Emails", "")).split(",") if e.strip()]
        contacts.append(
            Contact(
                first_name=str(row.get("First Name", "")).strip(),
                last_name=str(row.get("Last Name", "")).strip(),
                phones=phones,
                emails=emails,
                organization=str(row.get("Organization", "")).strip(),
                notes=str(row.get("Notes", "")).strip(),
            )
        )
    return [c for c in contacts if c.first_name or c.last_name]


def show_parse_result(result: ParseResult) -> None:
    col1, col2, col3 = st.columns(3)
    col1.metric("✅ Valid contacts", len(result.contacts))
    col2.metric("❌ Errors / skipped", len(result.errors))
    col3.metric("Success rate", f"{result.success_rate:.0%}")

    if result.errors:
        with st.expander(f"⚠️ {len(result.errors)} rows were skipped — click to review"):
            err_data = [
                {"Line": e.line_number, "Reason": e.reason, "Raw": e.raw[:100]}
                for e in result.errors
            ]
            st.dataframe(pd.DataFrame(err_data), use_container_width=True)


def editor_and_download(result: ParseResult, tab_key: str) -> None:
    if not result.contacts:
        st.warning("No valid contacts to show.")
        return

    show_parse_result(result)

    st.markdown("### ✏️ Review & Edit")
    st.caption(
        "You can edit names, phones, and other fields directly in this table "
        "before downloading. Rows marked ⚠ were extracted with low confidence."
    )

    df_base = contacts_to_df(result.contacts)

    # Highlight low-confidence rows
    def highlight_confidence(row):
        if row["⚠ Low confidence"]:
            return ["background-color: #fff3cd"] * len(row)
        return [""] * len(row)

    edited_df = st.data_editor(
        df_base.drop(columns=["⚠ Low confidence"]),
        use_container_width=True,
        num_rows="dynamic",
        key=f"editor_{tab_key}",
    )

    final_contacts = df_to_contacts(edited_df)

    writer = VCFWriter(version=vcf_version, dedupe=dedupe)
    vcf_text = writer.write_text(final_contacts)

    st.download_button(
        label=f"⬇️  Download VCF  ({len(final_contacts)} contacts)",
        data=vcf_text.encode("utf-8"),
        file_name="contacts.vcf",
        mime="text/vcard",
        type="primary",
        key=f"download_{tab_key}",
    )


# ── Tabs ──────────────────────────────────────────────────────────────────────

st.title("📇 vCard Converter")
st.caption("Convert contacts from CSV, VCF, or pasted text into a .vcf file ready to import into any phone.")

tab_upload, tab_paste, tab_google = st.tabs(
    ["📂 Upload File", "📋 Paste Text", "🔵 Google Contacts"]
)

# ── Tab 1: Upload File ────────────────────────────────────────────────────────

with tab_upload:
    st.subheader("Upload a CSV or VCF file")
    uploaded = st.file_uploader(
        "Drag & drop or browse",
        type=["csv", "vcf", "vcard", "txt"],
        help="Supported: standard CSV, VCF 2.1/3.0, plain text",
    )

    if uploaded:
        text = uploaded.read().decode("utf-8-sig")
        ext = Path(uploaded.name).suffix.lower()

        if ext in (".vcf", ".vcard"):
            parser = VCFParser(country_code=country_code)
            result = parser.parse_text(text, source_label=uploaded.name)
        elif ext == ".txt":
            parser = FreeTextParser(country_code=country_code)
            result = parser.parse_text(text, source_label=uploaded.name)
        else:
            parser = CSVParser(country_code=country_code)
            result = parser.parse_text(text, source_label=uploaded.name)

        editor_and_download(result, tab_key="upload")


# ── Tab 2: Paste Text ─────────────────────────────────────────────────────────

with tab_paste:
    st.subheader("Paste any contact data")
    st.caption(
        "Supported formats: WhatsApp group export (`Name ~ +91XXXXXXXXXX`), "
        "tab-separated rows, comma-separated rows, or loose prose like "
        '"Call Priya at 98765 43210".'
    )

    sample_placeholder = (
        "Raj Kumar ~ +91 98765 43210\n"
        "Priya Sharma\t8765432109\n"
        "Arun, 9123456789\n"
        "Call Neha at 9000011112"
    )

    pasted = st.text_area(
        "Paste here",
        height=200,
        placeholder=sample_placeholder,
        key="paste_input",
    )

    if st.button("Parse", key="parse_paste"):
        if pasted.strip():
            parser = FreeTextParser(country_code=country_code)
            result = parser.parse_text(pasted)
            st.session_state["paste_result"] = result
        else:
            st.warning("Nothing to parse.")

    if "paste_result" in st.session_state:
        editor_and_download(st.session_state["paste_result"], tab_key="paste")


# ── Tab 3: Google Contacts ────────────────────────────────────────────────────

with tab_google:
    st.subheader("Google Contacts CSV")
    st.caption(
        "Export from [Google Contacts](https://contacts.google.com) → "
        "**Export → Google CSV** and upload the file here."
    )

    google_file = st.file_uploader(
        "Upload Google Contacts CSV",
        type=["csv"],
        key="google_upload",
    )

    if google_file:
        text = google_file.read().decode("utf-8-sig")
        parser = CSVParser(country_code=country_code, force_google_mode=True)
        result = parser.parse_text(text, source_label=google_file.name)
        editor_and_download(result, tab_key="google")
