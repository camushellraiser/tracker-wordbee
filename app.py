import streamlit as st
import pandas as pd
import msal
import requests
from io import BytesIO


# -----------------------------
# 🔐 MICROSOFT AUTH
# -----------------------------
def get_access_token():
    client_id = st.secrets["CLIENT_ID"]
    tenant_id = st.secrets["TENANT_ID"]
    client_secret = st.secrets["CLIENT_SECRET"]

    authority = f"https://login.microsoftonline.com/{tenant_id}"

    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret,
    )

    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    if "access_token" not in result:
        raise Exception(result.get("error_description", "Could not get access token"))

    return result["access_token"]


# -----------------------------
# 📬 OUTLOOK FUNCTIONS
# -----------------------------
def get_latest_email(reference):
    token = get_access_token()

    headers = {
        "Authorization": f"Bearer {token}",
        "ConsistencyLevel": "eventual"
    }

    url = "https://graph.microsoft.com/v1.0/me/messages"

    params = {
        "$search": f'"{reference}"',
        "$top": 5,
        "$orderby": "receivedDateTime DESC"
    }

    response = requests.get(url, headers=headers, params=params)

    if response.status_code != 200:
        return None

    data = response.json()

    if "value" not in data or len(data["value"]) == 0:
        return None

    return data["value"][0]


# -----------------------------
# 🧠 STATUS PARSER
# -----------------------------
def extract_status_from_email(email_body):
    if not email_body:
        return ""

    text = email_body.replace("\r\n", "\n")
    lower = text.lower()

    start = lower.find("message:")
    end = lower.find("click to access job online")

    if start == -1:
        return ""

    start += len("message:")

    if end == -1:
        content = text[start:]
    else:
        content = text[start:end]

    lines = [line.strip() for line in content.splitlines() if line.strip()]
    return " ".join(lines)


def get_status_from_reference(reference):
    email = get_latest_email(reference)

    if not email:
        return ""

    body = email["body"]["content"]

    return extract_status_from_email(body)


# -----------------------------
# 📊 EXCEL FUNCTIONS
# -----------------------------
def load_excel(file):
    return pd.read_excel(file, sheet_name=None)


def clean_reference(value):
    if pd.isna(value):
        return ""
    return str(value).split("_")[0]


# -----------------------------
# 🚀 MAIN APP
# -----------------------------
def main():

    st.title("📊 Excel Tracker + Outlook")

    # 🔥 TEST CONNECTION
    st.write("Testing connection...")
    try:
        token = get_access_token()
        st.success("Connected to Microsoft Graph ✅")
    except Exception as e:
        st.error(f"Connection error: {e}")
        st.stop()

    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

    if not uploaded_file:
        return

    workbook = load_excel(uploaded_file)
    sheet = st.selectbox("Select sheet", list(workbook.keys()))
    df = workbook[sheet]

    st.write("Preview:")
    st.dataframe(df.head())

    columns = df.columns.tolist()

    ref_col = st.selectbox("Reference column", columns)
    person_col = st.selectbox("Person column", columns, index=3 if len(columns) > 3 else 0)

    result = pd.DataFrame()
    result["reference"] = df[ref_col].apply(clean_reference)
    result["person_name"] = df[person_col]

    result = result.dropna()

    # 🔍 FILTER
    people = result["person_name"].unique().tolist()
    selected_people = st.multiselect("Filter people", people, default=people[:2])

    filtered = result[result["person_name"].isin(selected_people)]

    st.dataframe(filtered)

    # 🚀 GET OUTLOOK STATUS
    if st.button("Get Outlook Status"):

        statuses = []

        with st.spinner("Fetching emails..."):
            for ref in filtered["reference"]:
                status = get_status_from_reference(ref)
                statuses.append(status)

        filtered["latest_status"] = statuses

        st.success("Done ✅")
        st.dataframe(filtered)


# -----------------------------
if __name__ == "__main__":
    main()
