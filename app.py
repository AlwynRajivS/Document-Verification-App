import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

# ---------------- HEADER MAPPING ----------------
HEADER_MAP = {
    "EXAM": "EXAM",
    "PROGRAMME": "PROGRAMME",
    "REGISTER NO": "REGISTER_NO",
    "STUDENT NAME": "NAME",
    "SEM": "SEM_NO",
    "SUBJECT ORDER": "SUB_ORDER",
    "SUB CODE": "SUB_CODE",
    "SUBJECT NAME": "SUBJECT_NAME",
    "INT": "INT",
    "EXT": "EXT",
    "TOT": "TOTAL",
    "RESULT": "RESULT",
    "GRADE": "GRADE",
    "GRADE POINT": "GRADE_POINT"
}

# ---------------- PDF PARSER ----------------
def extract_pdf_data(pdf_bytes):
    """Extracts structured student data from PDF text."""
    text_data = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text_data += (page.extract_text() or "") + "\n"

    text_data = re.sub(r"\s+", " ", text_data.strip())
    records = []

    # Split PDF by student blocks
    for block in re.split(r"Credits\s+Registered", text_data, flags=re.I):
        reg_match = re.search(r"REGISTER\s*NO\.?\s*:?[\s]*([0-9]+)", block, flags=re.I)
        if not reg_match:
            continue
        regno = reg_match.group(1).strip()

        # Extract subject info lines
        for line in re.split(r"(?=\d{1,2}\s+[A-Z]{2,3}\d{4})", block):
            m = re.search(
                r"^\s*\d+\s+([A-Z]{2,3}\d{4})\s+(.+?)\s+(\d+(?:\.\d+)?)\s+([A-Z\+OU]{1,2})\s+(\d+)\s+(PASS|RA)",
                line.strip(), flags=re.I
            )
            if m:
                records.append({
                    "REGISTER_NO": regno,
                    "SUB_CODE": m.group(1).strip().upper(),
                    "SUBJECT_NAME": m.group(2).strip(),
                    "COURSE_CREDIT": m.group(3).strip(),
                    "GRADE": m.group(4).strip().upper(),
                    "GRADE_POINT": m.group(5).strip(),
                    "RESULT": m.group(6).strip().upper(),
                })
    return pd.DataFrame(records)


# ---------------- RESULT NORMALIZATION ----------------
def normalize_result(value):
    """Convert result abbreviations to a common form for comparison."""
    v = str(value).strip().upper()
    if v == "F":
        return "RA"
    elif v == "P":
        return "PASS"
    return v


# ---------------- STREAMLIT APP ----------------
st.set_page_config(page_title="Excel vs PDF Comparator", layout="wide")
st.title("Excel vs PDF Comparator (with Missing Record Detection)")

st.markdown("""
This tool compares **Excel master data** and **multi-student PDF marksheets**,  
maps results and detects missing or extra records.
""")

# ---------------- FILE UPLOAD ----------------
excel_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
pdf_file = st.file_uploader("Upload PDF File", type=["pdf"])

if excel_file and pdf_file:
    try:
        # ---------- STEP 1: Load Excel ----------
        df_excel = pd.read_excel(excel_file)
        df_excel.columns = [str(c).strip().upper() for c in df_excel.columns]
        st.success("Excel file loaded successfully.")

        # Map headers
        mapped_cols = {}
        for col in df_excel.columns:
            clean_col = col.strip().upper()
            if clean_col in HEADER_MAP:
                mapped_cols[col] = HEADER_MAP[clean_col]
        df_excel.rename(columns=mapped_cols, inplace=True)
        df_excel = df_excel.astype(str).apply(lambda x: x.str.strip())

        # Normalize Excel 'RESULT' column
        if "RESULT" in df_excel.columns:
            df_excel["RESULT"] = df_excel["RESULT"].apply(normalize_result)

        st.subheader("Excel Data")
        st.dataframe(df_excel.head())

        # Check required columns
        required_cols = ["REGISTER_NO", "SUB_CODE", "SUBJECT_NAME", "GRADE", "GRADE_POINT", "RESULT"]
        missing = [c for c in required_cols if c not in df_excel.columns]
        if missing:
            st.warning(f"Missing columns in Excel after mapping: {missing}")

        # ---------- STEP 2: Parse PDF ----------
        st.info("Extracting structured data from PDF...")
        pdf_bytes = pdf_file.read()
        df_pdf = extract_pdf_data(pdf_bytes)

        if df_pdf.empty:
            st.error("No valid course data extracted from PDF. Please verify the PDF format.")
            st.stop()

        st.success(f"Extracted {len(df_pdf)} subject records from PDF.")
        st.dataframe(df_pdf.head())

        # Normalize PDF 'RESULT' column
        df_pdf["RESULT"] = df_pdf["RESULT"].apply(normalize_result)

        # ---------- STEP 3: Merge and Compare ----------
        merged = pd.merge(
            df_excel, df_pdf,
            on=["REGISTER_NO", "SUB_CODE"],
            how="inner",
            suffixes=("_EXCEL", "_PDF")
        )

        compare_cols = ["SUBJECT_NAME", "GRADE", "GRADE_POINT", "RESULT"]

        mismatches = merged[
            merged.apply(
                lambda row: any(
                    str(row.get(f"{col}_EXCEL", "")).strip().upper() !=
                    str(row.get(f"{col}_PDF", "")).strip().upper()
                    for col in compare_cols if f"{col}_PDF" in merged.columns
                ),
                axis=1
            )
        ]

        # ---------- STEP 4: Missing Record Detection ----------
        excel_keys = df_excel[["REGISTER_NO", "SUB_CODE"]].drop_duplicates()
        pdf_keys = df_pdf[["REGISTER_NO", "SUB_CODE"]].drop_duplicates()

        missing_in_pdf = pd.merge(excel_keys, pdf_keys, on=["REGISTER_NO", "SUB_CODE"], how="left", indicator=True)
        missing_in_pdf = missing_in_pdf[missing_in_pdf["_merge"] == "left_only"].drop(columns=["_merge"])

        missing_in_excel = pd.merge(pdf_keys, excel_keys, on=["REGISTER_NO", "SUB_CODE"], how="left", indicator=True)
        missing_in_excel = missing_in_excel[missing_in_excel["_merge"] == "left_only"].drop(columns=["_merge"])

        # ---------- STEP 5: Display Results ----------
        st.subheader("Comparison Result")

        if mismatches.empty and missing_in_pdf.empty and missing_in_excel.empty:
            st.success("All records match perfectly! No mismatches or missing records found.")
        else:
            if not mismatches.empty:
                st.error(f"{len(mismatches)} mismatched rows detected!")
                st.dataframe(mismatches)
                csv_buf = io.StringIO()
                mismatches.to_csv(csv_buf, index=False)
                st.download_button("Download Mismatch Report (CSV)", csv_buf.getvalue(),
                                   file_name="mismatch_report.csv", mime="text/csv")

            if not missing_in_pdf.empty:
                st.warning(f"{len(missing_in_pdf)} records missing in PDF.")
                st.dataframe(missing_in_pdf)
                csv_buf2 = io.StringIO()
                missing_in_pdf.to_csv(csv_buf2, index=False)
                st.download_button("Download Missing in PDF (CSV)", csv_buf2.getvalue(),
                                   file_name="missing_in_pdf.csv", mime="text/csv")

            if not missing_in_excel.empty:
                st.warning(f"{len(missing_in_excel)} extra records found in PDF (not in Excel).")
                st.dataframe(missing_in_excel)
                csv_buf3 = io.StringIO()
                missing_in_excel.to_csv(csv_buf3, index=False)
                st.download_button("Download Extra in PDF (CSV)", csv_buf3.getvalue(),
                                   file_name="extra_in_pdf.csv", mime="text/csv")

    except Exception as e:
        st.error(f"Error during processing: {e}")

else:
    st.info("Please upload both Excel and PDF files to start comparison.")

