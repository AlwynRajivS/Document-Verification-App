import streamlit as st
import pandas as pd
import pdfplumber
import io
import re

st.set_page_config(page_title="Dynamic Excel vs PDF Comparator", layout="wide")
st.title("Excel vs PDF/Text Comparator")
st.markdown("""
Upload any **Excel file** (master data) and **PDF/Text file** (to verify).  
This tool automatically converts the PDF to text and compares **only the columns you select**.
""")

# ---------- File Upload ----------
excel_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
pdf_file = st.file_uploader("Upload PDF or Text File", type=["pdf", "txt"])

if excel_file:
    try:
        # ---------- STEP 1: Load Excel ----------
        df_excel = pd.read_excel(excel_file)
        df_excel.columns = [str(c).strip().upper().replace("\n", " ").replace("  ", " ") for c in df_excel.columns]
        df_excel = df_excel.astype(str).apply(lambda x: x.str.strip())

        st.success("Excel file loaded successfully.")

        # ---------- STEP 2: Select Headers ----------
        st.markdown("### Select Columns to Compare")
        selected_cols = st.multiselect(
            "Select one or more columns from Excel to check in PDF/Text file:",
            options=list(df_excel.columns),
            default=list(df_excel.columns)
        )

        if not selected_cols:
            st.warning("Please select at least one column to compare.")
            st.stop()

        # ---------- STEP 3: Upload and Extract PDF ----------
        if pdf_file:
            if pdf_file.name.lower().endswith(".pdf"):
                pdf_text = ""
                with pdfplumber.open(pdf_file) as pdf:
                    for page in pdf.pages:
                        pdf_text += (page.extract_text() or "") + "\n"
                if not pdf_text.strip():
                    st.error("No text extracted from PDF. Try uploading a text version of the file.")
                    st.stop()
            else:
                pdf_text = pdf_file.read().decode("utf-8", errors="ignore")

            # Clean and standardize PDF text
            pdf_text = re.sub(r'\s+', ' ', pdf_text.strip()).upper()

            # ---------- STEP 4: Compare Dynamically ----------
            st.info("Comparing selected Excel columns with PDF/Text content...")

            mismatches = []
            total_checked = 0

            for idx, row in df_excel.iterrows():
                missing_fields = {}
                found_all = True

                for col_name in selected_cols:
                    value_str = str(row[col_name]).strip().upper()
                    if not value_str or value_str in ["NAN", "NONE"]:
                        continue

                    total_checked += 1
                    if value_str not in pdf_text:
                        found_all = False
                        missing_fields[col_name] = value_str

                if not found_all:
                    mismatch_info = {"ROW_INDEX": idx + 1}
                    mismatch_info.update(missing_fields)
                    mismatches.append(mismatch_info)

            # ---------- STEP 5: Display Results ----------
            if not mismatches:
                st.success(f"All {total_checked} selected data values were found in the PDF/Text. No mismatches detected.")
            else:
                mismatch_df = pd.DataFrame(mismatches)
                st.warning(f"Found {len(mismatch_df)} rows with missing or mismatched data in the PDF/Text.")
                st.dataframe(mismatch_df)

                # Allow download
                csv_buffer = io.StringIO()
                mismatch_df.to_csv(csv_buffer, index=False)
                st.download_button(
                    label="Download Mismatch Report (CSV)",
                    data=csv_buffer.getvalue(),
                    file_name="mismatched_data.csv",
                    mime="text/csv"
                )

        else:
            st.info("Please upload a PDF or Text file to start comparison.")

    except Exception as e:
        st.error(f"Error: {e}")

else:
    st.info("Please upload both Excel and PDF/Text files to start.")
