import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Filter Tool", layout="wide")
st.title(" Excel Filtering Tool")

# Upload Excel file
uploaded_file = st.file_uploader(" Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file)

        st.subheader(" Original Data Preview")
        st.dataframe(df.head())

        #  fixed columns
        fixed_columns = [
            'counterpartyName',
            'caseId',
            'completedDate',
            'rubix_Score',
            'counterparty_panNo',
            'MCA_CIN'
        ]

        # Filter by case type (Domestic, International, Both)
        if "case type" in df.columns:
            case_type = st.selectbox("🔍 Select Case Type", ["Domestic", "International", "Both"])
            if case_type != "Both":
                df = df[df["case type"].str.lower() == case_type.lower()]
        else:
            st.warning("⚠️ 'case type' column not found. Skipping filter.")

        # Select additional columns
        st.subheader("➕ Select Additional Columns to Include")
        selectable_columns = [col for col in df.columns if col not in fixed_columns]
        selected_columns = st.multiselect("Pick extra columns", selectable_columns)

        # Combine and filter only existing columns
        requested_columns = fixed_columns + selected_columns
        existing_columns_set = set(df.columns)
        final_columns = [col for col in requested_columns if col in existing_columns_set]

        # Identify any missing fixed columns
        missing_fixed_columns = list(set(fixed_columns) - existing_columns_set)
        if missing_fixed_columns:
            st.error(f"⚠️ Missing fixed columns: {missing_fixed_columns}")

        #column slicing
        filtered_df = df.loc[:, final_columns]

        st.subheader(" Filtered Result")
        st.dataframe(filtered_df)

        # Excel download function
        def to_excel(dataframe):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                dataframe.to_excel(writer, index=False, sheet_name='Filtered')
            return output.getvalue()

        excel_data = to_excel(filtered_df)

        # Download button
        st.download_button(
            label=" Download Filtered Excel",
            data=excel_data,
            file_name="filtered_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f" Error reading file: {e}")
