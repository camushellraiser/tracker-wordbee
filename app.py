"""Streamlit Tracker

Current scope:
- Upload an Excel workbook
- Select a sheet
- Extract a clean `reference`
- Read `person_name`
- Filter by one or more people
- Parse a status block from an Outlook email body

Status rule:
- Take everything after `Message:` and before `Click to access job online`

Run locally:
    streamlit run app.py
"""

from __future__ import annotations

from io import BytesIO
from typing import List

import pandas as pd
import streamlit as st


APP_TITLE = "Excel Tracker"
APP_ICON = "📊"


st.set_page_config(
    page_title=APP_TITLE,
    page_icon=APP_ICON,
    layout="wide",
    initial_sidebar_state="expanded",
)


CUSTOM_CSS = """
<style>
    .main {
        background: linear-gradient(180deg, #f8fafc 0%, #eef2ff 100%);
    }
    .block-container {
        padding-top: 1.4rem;
        padding-bottom: 2rem;
        max-width: 1400px;
    }
    .hero {
        padding: 1.5rem 1.6rem;
        border-radius: 24px;
        background: linear-gradient(135deg, #111827 0%, #312e81 50%, #4f46e5 100%);
        color: white;
        box-shadow: 0 18px 45px rgba(15, 23, 42, 0.18);
        margin-bottom: 1.2rem;
    }
    .hero h1 {
        margin: 0;
        font-size: 2rem;
        line-height: 1.1;
    }
    .hero p {
        margin: 0.35rem 0 0 0;
        opacity: 0.92;
        font-size: 1rem;
    }
    .card {
        padding: 1rem 1.1rem;
        background: rgba(255, 255, 255, 0.78);
        border: 1px solid rgba(148, 163, 184, 0.25);
        border-radius: 18px;
        box-shadow: 0 10px 30px rgba(15, 23, 42, 0.06);
    }
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    return pd.read_excel(BytesIO(file_bytes), sheet_name=None)


def clean_reference(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    return text.split("_")[0].strip()


def guess_reference_column(columns: List[str]) -> int:
    lowered = [str(c).strip().lower() for c in columns]
    for key in ("reference", "ref", "reference id", "reference_id"):
        if key in lowered:
            return lowered.index(key)
    return 0


def guess_person_column(columns: List[str]) -> int:
    lowered = [str(c).strip().lower() for c in columns]
    for key in ("person name", "person", "name"):
        if key in lowered:
            return lowered.index(key)
    return 3 if len(columns) > 3 else max(0, len(columns) - 1)


def extract_status_from_email(email_body: str) -> str:
    """Extract everything between `Message:` and `Click to access job online`."""
    if not email_body:
        return ""

    text = email_body.replace("\r\n", "\n").replace("\r", "\n")
    lower = text.lower()

    start_token = "message:"
    end_token = "click to access job online"

    start_idx = lower.find(start_token)
    if start_idx == -1:
        return ""

    start_idx += len(start_token)
    end_idx = lower.find(end_token, start_idx)
    if end_idx == -1:
        extracted = text[start_idx:]
    else:
        extracted = text[start_idx:end_idx]

    lines = [line.strip() for line in extracted.splitlines()]
    cleaned = [line for line in lines if line]
    return " ".join(cleaned).strip()


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Processed")
    return buffer.getvalue()


def main() -> None:
    st.markdown(
        """
        <div class="hero">
            <h1>📊 Excel Tracker</h1>
            <p>Upload your workbook, extract a clean reference, filter by person name, and parse Outlook status text.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.header("Configuration")
        uploaded_file = st.file_uploader(
            "Upload Excel file",
            type=["xlsx", "xls"],
            help="Your tracker workbook.",
        )
        st.caption("Ready for GitHub and Streamlit Cloud.")

    if not uploaded_file:
        st.info("Upload an Excel file to begin.")
        st.stop()

    try:
        workbook = load_excel(uploaded_file.getvalue())
    except Exception as exc:
        st.error(f"Could not read the Excel file: {exc}")
        st.stop()

    sheet_names = list(workbook.keys())
    if not sheet_names:
        st.error("The file does not contain any valid sheets.")
        st.stop()

    left, right = st.columns([1, 2])
    with left:
        selected_sheet = st.selectbox("Sheet", sheet_names)
    df = workbook[selected_sheet].copy()

    if df.empty:
        st.warning("The selected sheet is empty.")
        st.stop()

    columns = df.columns.tolist()
    if len(columns) < 2:
        st.error("The selected sheet must have at least two columns.")
        st.stop()

    reference_default = guess_reference_column(columns)
    person_default = guess_person_column(columns)

    st.markdown("<div class='card'>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)

    with c1:
        reference_col = st.selectbox(
            "Reference column",
            columns,
            index=reference_default,
            help="Select the column that contains values like GTS260050_Web_SWu_Incubators Models_AEM.",
        )

    with c2:
        person_col = st.selectbox(
            "Person name column",
            columns,
            index=person_default,
            help="Select the column that contains the person's name.",
        )

    with c3:
        download_excel = st.checkbox("Also download as Excel", value=True)

    st.markdown("</div>", unsafe_allow_html=True)

    result = pd.DataFrame(
        {
            "reference": df[reference_col].apply(clean_reference),
            "person_name": df[person_col].astype(str).str.strip(),
        }
    )

    result = result.replace({"": pd.NA}).dropna(how="all").fillna("")

    person_options = sorted([p for p in result["person_name"].unique().tolist() if p])
    default_people = person_options[:2] if len(person_options) >= 2 else person_options

    f1, f2 = st.columns(2)
    with f1:
        selected_people = st.multiselect(
            "Filter by person name",
            options=person_options,
            default=default_people,
            help="Pick one or more people to show.",
        )
    with f2:
        reference_search = st.text_input(
            "Reference search",
            value="",
            help="Type part of a reference, for example GTS260050.",
        ).strip()

    filtered = result.copy()
    if selected_people:
        filtered = filtered[filtered["person_name"].isin(selected_people)]
    if reference_search:
        filtered = filtered[filtered["reference"].astype(str).str.contains(reference_search, case=False, na=False)]

    filtered = filtered.sort_values(["person_name", "reference"], kind="stable").reset_index(drop=True)

    m1, m2, m3 = st.columns(3)
    m1.metric("Visible rows", len(filtered))
    m2.metric("Unique people", filtered["person_name"].nunique())
    m3.metric("Unique references", filtered["reference"].nunique())

    st.subheader("Processed data")
    st.dataframe(filtered, use_container_width=True, hide_index=True)

    csv_bytes = filtered.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="Download CSV",
        data=csv_bytes,
        file_name="tracker_processed.csv",
        mime="text/csv",
        use_container_width=True,
    )

    if download_excel:
        xlsx_bytes = to_excel_bytes(filtered)
        st.download_button(
            label="Download Excel",
            data=xlsx_bytes,
            file_name="tracker_processed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.divider()
    st.subheader("Outlook status parser")
    st.write("Paste the full email body below to extract the status block.")

    email_body = st.text_area(
        "Email body",
        height=250,
        placeholder="Paste the email text here...",
    )
    extracted_status = extract_status_from_email(email_body)

    status_cols = st.columns([1, 2])
    with status_cols[0]:
        st.metric("Status characters", len(extracted_status))
    with status_cols[1]:
        st.text_area("Extracted status", value=extracted_status, height=160)

    with st.expander("How it works"):
        st.write(
            """
            - `reference`: the text before the first underscore.
            - `person_name`: the selected person column, trimmed.
            - Filters let you show one or more people, such as just two names.
            - The status parser extracts everything after `Message:` and before `Click to access job online`.
            """
        )


if __name__ == "__main__":
    main()
