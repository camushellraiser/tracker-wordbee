from __future__ import annotations

import html
import re
from io import BytesIO
from typing import List, Optional, Tuple

import msal
import pandas as pd
import requests
import streamlit as st


APP_TITLE = "Excel Tracker"
APP_ICON = "📊"
SCOPES = ["User.Read", "Mail.Read"]


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


# -----------------------------
# Excel helpers
# -----------------------------
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


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Processed")
    return buffer.getvalue()


# -----------------------------
# HTML / status helpers
# -----------------------------
def html_to_text(raw_html: str) -> str:
    if not raw_html:
        return ""

    text = html.unescape(raw_html)

    # Make common HTML line breaks visible.
    text = re.sub(r"(?is)<br\s*/?>", "\n", text)
    text = re.sub(r"(?is)</p>", "\n", text)
    text = re.sub(r"(?is)</div>", "\n", text)
    text = re.sub(r"(?is)</tr>", "\n", text)

    # Remove remaining tags.
    text = re.sub(r"(?is)<script.*?>.*?</script>", " ", text)
    text = re.sub(r"(?is)<style.*?>.*?</style>", " ", text)
    text = re.sub(r"(?s)<.*?>", " ", text)

    # Clean whitespace.
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s+", "\n", text)
    return text.strip()


def extract_status_from_email(email_body: str) -> str:
    """
    Extract everything between:
      Message:
    and
      Click to access job online
    """
    if not email_body:
        return ""

    text = html_to_text(email_body)
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


# -----------------------------
# Microsoft auth helpers
# -----------------------------
def get_secrets() -> Tuple[str, str]:
    client_id = st.secrets.get("CLIENT_ID", "")
    tenant_id = st.secrets.get("TENANT_ID", "")
    if not client_id or not tenant_id:
        raise RuntimeError("Missing CLIENT_ID or TENANT_ID in Streamlit secrets.")
    return client_id, tenant_id


def load_token_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    cache_state = st.session_state.get("msal_cache_state")
    if cache_state:
        cache.deserialize(cache_state)
    return cache


def save_token_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        st.session_state["msal_cache_state"] = cache.serialize()


def build_msal_app(cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
    client_id, tenant_id = get_secrets()
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    return msal.PublicClientApplication(
        client_id=client_id,
        authority=authority,
        token_cache=cache,
    )


def try_silent_token(app: msal.PublicClientApplication, cache: msal.SerializableTokenCache) -> Optional[str]:
    accounts = app.get_accounts()
    if not accounts:
        return None

    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if result and "access_token" in result:
        save_token_cache(cache)
        return result["access_token"]
    return None


def start_device_flow(app: msal.PublicClientApplication) -> dict:
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(flow.get("error_description", "Could not start device flow."))
    return flow


def complete_device_flow(app: msal.PublicClientApplication, flow: dict) -> str:
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(result.get("error_description", "Login failed."))
    return result["access_token"]


def authenticate_user() -> Optional[str]:
    cache = load_token_cache()
    app = build_msal_app(cache)

    # 1) Try silent login first.
    token = try_silent_token(app, cache)
    if token:
        st.session_state["access_token"] = token
        return token

    # 2) Device flow UI.
    st.markdown("### Microsoft sign-in")

    flow = st.session_state.get("device_flow")

    if not flow:
        if st.button("Start Microsoft sign-in", type="primary"):
            flow = start_device_flow(app)
            st.session_state["device_flow"] = flow
            st.rerun()

        st.info("Start sign-in to get a code for Microsoft login.")
        return None

    st.code(
        f"Go to: {flow['verification_uri']}\n"
        f"Code: {flow['user_code']}",
        language="text",
    )

    c1, c2 = st.columns([1, 1])
    with c1:
        if st.button("I already signed in"):
            try:
                token = complete_device_flow(app, flow)
                save_token_cache(cache)
                st.session_state["access_token"] = token
                st.session_state["device_flow"] = None
                st.success("Signed in successfully.")
                st.rerun()
            except Exception as exc:
                st.error(f"Login failed: {exc}")

    with c2:
        if st.button("Reset sign-in"):
            st.session_state["device_flow"] = None
            st.rerun()

    st.info("Finish the Microsoft login in the browser, then click **I already signed in**.")
    return None


# -----------------------------
# Graph mail helpers
# -----------------------------
def graph_get(url: str, token: str, params: dict) -> dict:
    headers = {
        "Authorization": f"Bearer {token}",
        "ConsistencyLevel": "eventual",
        "Accept": "application/json",
    }
    response = requests.get(url, headers=headers, params=params, timeout=30)
    if response.status_code >= 400:
        raise RuntimeError(f"Graph error {response.status_code}: {response.text}")
    return response.json()


def get_latest_email(reference: str, token: str) -> Optional[dict]:
    if not reference:
        return None

    url = "https://graph.microsoft.com/v1.0/me/messages"

    # Search both quoted and unquoted to improve recall.
    search_queries = [f'"{reference}"', reference]

    best_item = None

    for q in search_queries:
        params = {
            "$search": q,
            "$top": 25,
            "$select": "subject,receivedDateTime,body,from",
            "$orderby": "receivedDateTime DESC",
        }

        data = graph_get(url, token, params)
        items = data.get("value", [])

        if not items:
            continue

        # Pick the newest item in case Graph does not honor order perfectly.
        items = sorted(
            items,
            key=lambda x: x.get("receivedDateTime", ""),
            reverse=True,
        )
        best_item = items[0]
        break

    return best_item


def get_status_from_reference(reference: str, token: str) -> str:
    email = get_latest_email(reference, token)
    if not email:
        return ""

    body = email.get("body", {})
    content = body.get("content", "")
    return extract_status_from_email(content)


# -----------------------------
# Main app
# -----------------------------
def main() -> None:
    st.markdown(
        """
        <div class="hero">
            <h1>📊 Excel Tracker</h1>
            <p>Upload your workbook, extract a clean reference, filter by person name, and fetch Outlook status text.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Microsoft sign-in first.
    try:
        token = authenticate_user()
    except Exception as exc:
        st.error(f"Authentication setup error: {exc}")
        st.stop()

    if not token:
        st.stop()

    st.success("Connected to Microsoft Graph ✅")

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
    result = result[(result["reference"] != "") | (result["person_name"] != "")].copy()

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
        filtered = filtered[
            filtered["reference"].astype(str).str.contains(reference_search, case=False, na=False)
        ]

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
    st.subheader("Outlook status fetch")
    st.write("Click the button below to fetch the latest email status for each visible reference.")

    if st.button("Get Outlook Status", type="primary"):
        updated = filtered.copy()
        statuses = []

        with st.spinner("Fetching emails from Outlook..."):
            for ref in updated["reference"].tolist():
                try:
                    status = get_status_from_reference(ref, token)
                except Exception as exc:
                    status = f"ERROR: {exc}"
                statuses.append(status)

        updated["latest_status"] = statuses
        st.session_state["last_result"] = updated
        st.success("Done ✅")

    if "last_result" in st.session_state:
        st.subheader("Result with Outlook status")
        st.dataframe(st.session_state["last_result"], use_container_width=True, hide_index=True)

        final_csv = st.session_state["last_result"].to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="Download final CSV",
            data=final_csv,
            file_name="tracker_with_status.csv",
            mime="text/csv",
            use_container_width=True,
        )

        if download_excel:
            final_xlsx = to_excel_bytes(st.session_state["last_result"])
            st.download_button(
                label="Download final Excel",
                data=final_xlsx,
                file_name="tracker_with_status.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    with st.expander("How it works"):
        st.write(
            """
            - `reference` is the text before the first underscore.
            - `person_name` is the selected person column, trimmed.
            - The app signs in with Microsoft device login and uses `/me/messages`.
            - The status parser extracts everything after `Message:` and before `Click to access job online`.
            """
        )


if __name__ == "__main__":
    main()
