"""Streamlit Tracker - Excel loader

Primera versión del tracker:
- Carga un Excel
- Selecciona hoja
- Extrae `reference` limpio
- Lee `person name`
- Muestra la tabla procesada
- Permite descargar el resultado

Run locally:
    streamlit run app.py
"""

from __future__ import annotations

from io import BytesIO
import pandas as pd
import streamlit as st


APP_TITLE = "Tracker de Excel"
APP_ICON = "📊"


st.set_page_config(
    page_title=APP_TITLE,
    page_icon=APP_ICON,
    layout="wide",
    initial_sidebar_state="expanded",
)


@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes):
    return pd.read_excel(BytesIO(file_bytes), sheet_name=None)


def clean_reference(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    return text.split("_")[0].strip()


def main():
    st.title("📊 Tracker de Excel")

    uploaded_file = st.file_uploader("Sube tu Excel", type=["xlsx", "xls"])

    if not uploaded_file:
        st.info("Sube un archivo Excel para comenzar.")
        st.stop()

    workbook = load_excel(uploaded_file.getvalue())
    sheet_names = list(workbook.keys())

    selected_sheet = st.selectbox("Hoja", sheet_names)
    df = workbook[selected_sheet].copy()

    st.dataframe(df.head(20), use_container_width=True)

    columns = df.columns.tolist()

    reference_col = st.selectbox("Columna reference", columns, index=0)
    person_col = st.selectbox("Columna person", columns, index=3 if len(columns)>3 else 1)

    result = pd.DataFrame()
    result["reference_raw"] = df[reference_col]
    result["reference"] = result["reference_raw"].apply(clean_reference)
    result["person_name"] = df[person_col]

    st.dataframe(result, use_container_width=True)


if __name__ == "__main__":
    main()
