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

# ---------------- NORMALIZATION HELPERS ----------------
def normalize_result(value):
    """Convert result abbreviations to a common form."""
    v = str(value).strip().upper()
    if v == "F":
        return "RA"
    elif v == "P":
        return "PASS"
    return v

def normalize_subject_name_after_extraction(name):
    """
    Clean subject names AFTER extraction for comparison:
    - remove various star characters and weird unicode variants,
    - remove other noise, collapse spaces, uppercase.
    """
    name = str(name)
    # remove unicode variants of star/asterisk and bullets
    name = re.sub(r"[\*\u2217\u2731\u204E\uFE61\uFF0A\u002A]+", "", name)
    # remove common noise characters but keep letters, digits, parentheses, hyphen, comma, colon
    name = re.sub(r"[^A-Za-z0-9\-\(\)\:\,\/\s&]", "", name)
    # collapse spaces
    name = re.sub(r"\s+", " ", name)
    return name.strip().upper()

def normalize_register(v):
    """
    Normalize register number:
    - If numeric (int/float or scientific), convert to integer string (no .0)
    - Else extract long digit sequence if present
    - Fallback: return stripped string
    """
    s = str(v).strip()
    if s == "" or s.lower() == "nan":
        return ""
    # Try numeric parse first (handles 9.20423E+11, floats that become 920423xxxxxx, and plain ints)
    try:
        # Some register IDs may be large; float may lose precision for very large numbers.
        # But typical register numbers fit within float precision for conversion to int; attempt it.
        f = float(s)
        i = int(f)
        return str(i)
    except Exception:
        pass
    # If scientific-like (contains E), try Decimal? fallback to extracting digits
    digits = "".join(re.findall(r"\d+", s))
    if len(digits) >= 5:
        return digits
    # fallback: remove trailing .0
    s2 = re.sub(r"\.0+$", "", s)
    return s2

# ---------------- PDF PARSER ----------------
def extract_pdf_data(pdf_bytes):
    """
    Robust extraction:
    - Build full text
    - Locate each 'REGISTER NO' and treat the following portion as that student's block
    - Within block, find subject occurrences using flexible regexes
    """
    text_data = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            # some pdfs have NBSP; keep them consistent
            page_txt = page.extract_text() or ""
            page_txt = page_txt.replace("\xa0", " ").replace("\u200b", "")
            text_data += page_txt + "\n"

    # normalize whitespace
    text_data = re.sub(r"\s+", " ", text_data).strip()
    records = []

    # Find all REGISTER NO occurrences with their span so we can slice blocks reliably
    reg_iter = list(re.finditer(r"(REGISTER\s*NO\.?\s*:?\s*([0-9A-Za-z\.\-E\+]+))", text_data, flags=re.I))
    if not reg_iter:
        # try a looser pattern
        reg_iter = list(re.finditer(r"(REGISTER\s*NO\.?\s*:?\s*([0-9]+))", text_data, flags=re.I))

    if not reg_iter:
        return pd.DataFrame(records)  # empty

    for idx, m in enumerate(reg_iter):
        reg_raw = m.group(2)
        # block starts at end of REGISTER NO match
        block_start = m.end()
        block_end = reg_iter[idx + 1].start() if idx + 1 < len(reg_iter) else len(text_data)
        block = text_data[block_start:block_end].strip()
        regno = normalize_register(reg_raw)

        # Now find subject lines inside block.
        # We'll use two patterns:
        # 1) full line with credits & gradepoint & PASS/RA: e.g. "04 AUD101 *Constitution of India 0 S 0 PASS"
        # 2) shorter audit-style line: "AUD101 *Constitution of India PASS" or "AUD101 *Constitution of India 0 S PASS"
        # We'll run finditer with a fairly flexible pattern.

        # Pattern A (more columns present): code + name + credit (number) + grade-letter + grade-point + PASS/RA
        pat_a = re.compile(
            r"(?:\d{1,3}\s+)?([A-Z]{2,3}\d{3,4})\s+([^\d]{2,160}?)\s+(\d+(?:\.\d+)?)\s+([A-Z\+OU]{1,2})\s+(\d+)\s+(PASS|RA|P)",
            flags=re.I,
        )
        # Pattern B (shorter): code + name + (maybe credit) + maybe grade letter + PASS/RA or just PASS/RA
        pat_b = re.compile(
            r"(?:\d{1,3}\s+)?([A-Z]{2,3}\d{3,4})\s+([A-Za-z\*\u2217\u2731\u204E\uFE61\uFF0A\u002A0-9\(\)\-\:\,\/&\s]{1,160}?)\s+(?:([\d\.]+)\s+)?(?:([A-Z\+OU]{1,2})\s+)?(?:([\d]+)\s+)?(PASS|RA|P)\b",
            flags=re.I,
        )

        # First capture all using pat_a (more specific), then pat_b for the remainder
        used_spans = []

        for msub in pat_a.finditer(block):
            sub_code = msub.group(1).strip().upper()
            sub_name = msub.group(2).strip()
            credit = msub.group(3).strip()
            grade = msub.group(4).strip().upper()
            grade_point = msub.group(5).strip()
            result = msub.group(6).strip().upper()
            records.append({
                "REGISTER_NO": regno,
                "SUB_CODE": sub_code,
                "SUBJECT_NAME": sub_name,
                "COURSE_CREDIT": credit,
                "GRADE": grade,
                "GRADE_POINT": grade_point,
                "RESULT": "PASS" if result == "P" else ("PASS" if result == "PASS" else result),
            })
            used_spans.append((msub.start(), msub.end()))

        # For pat_b, avoid duplicates by checking overlap with used_spans
        for msub in pat_b.finditer(block):
            s, e = msub.start(), msub.end()
            overlap = any(not (e <= us or s >= ue) for us, ue in used_spans)
            if overlap:
                continue
            sub_code = msub.group(1).strip().upper()
            sub_name = msub.group(2).strip()
            credit = msub.group(3).strip() if msub.group(3) else "0"
            grade = (msub.group(4).strip().upper() if msub.group(4) else "S")
            grade_point = (msub.group(5).strip() if msub.group(5) else "0")
            result = msub.group(6).strip().upper()
            records.append({
                "REGISTER_NO": regno,
                "SUB_CODE": sub_code,
                "SUBJECT_NAME": sub_name,
                "COURSE_CREDIT": credit,
                "GRADE": grade,
                "GRADE_POINT": grade_point,
                "RESULT": "PASS" if result == "P" else ("PASS" if result == "PASS" else result),
            })
            used_spans.append((s, e))

    return pd.DataFrame(records)

# ---------------- STREAMLIT APP ----------------
st.set_page_config(page_title="Excel vs PDF Comparator", layout="wide")
st.title("Excel vs PDF Comparator (with Missing Record Detection)")

st.markdown("""
This tool compares **Excel master data** and **multi-student PDF marksheets**,  
detects mismatches, and identifies missing or extra records.
""")

# ---------------- FILE UPLOAD ----------------
excel_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
pdf_file = st.file_uploader("Upload PDF File", type=["pdf"])

if excel_file and pdf_file:
    try:
        # ---------- STEP 1: Load Excel ----------
        df_excel = pd.read_excel(excel_file, dtype=str)  # read everything as string initially
        df_excel.columns = [str(c).strip().upper() for c in df_excel.columns]
        st.success("Excel file loaded successfully.")

        # Map headers
        mapped_cols = {col: HEADER_MAP[col.strip().upper()] for col in df_excel.columns if col.strip().upper() in HEADER_MAP}
        df_excel.rename(columns=mapped_cols, inplace=True)

        # Trim whitespace in all cells
        df_excel = df_excel.fillna("").astype(str).apply(lambda col: col.str.strip())

        # Normalize excel register numbers (convert scientific/float strings to plain digits)
        if "REGISTER_NO" in df_excel.columns:
            df_excel["REGISTER_NO"] = df_excel["REGISTER_NO"].apply(normalize_register)

        # Normalize subject name in excel AFTER extraction (remove leading *)
        if "SUBJECT_NAME" in df_excel.columns:
            df_excel["SUBJECT_NAME"] = df_excel["SUBJECT_NAME"].apply(normalize_subject_name_after_extraction)

        if "RESULT" in df_excel.columns:
            df_excel["RESULT"] = df_excel["RESULT"].apply(normalize_result)

        st.subheader("Excel Data Preview")
        st.dataframe(df_excel.head())

        # ---------- STEP 2: Parse PDF ----------
        st.info("Extracting structured data from PDF...")
        pdf_bytes = pdf_file.read()
        df_pdf = extract_pdf_data(pdf_bytes)

        if df_pdf.empty:
            st.error("Could not extract any course data from PDF/Text file.")
            st.stop()

        st.success(f"Extracted {len(df_pdf)} subject records from PDF.")
        st.dataframe(df_pdf.head())

        # Normalize PDF registers and subject names AFTER extraction
        df_pdf["REGISTER_NO"] = df_pdf["REGISTER_NO"].apply(normalize_register)
        df_pdf["SUBJECT_NAME"] = df_pdf["SUBJECT_NAME"].apply(normalize_subject_name_after_extraction)
        df_pdf["SUB_CODE"] = df_pdf["SUB_CODE"].astype(str).str.strip().str.upper()
        df_pdf["RESULT"] = df_pdf["RESULT"].apply(normalize_result)

        # Ensure Excel SUB_CODE normalized
        if "SUB_CODE" in df_excel.columns:
            df_excel["SUB_CODE"] = df_excel["SUB_CODE"].astype(str).str.strip().str.upper()
        else:
            st.error("Excel missing SUB_CODE after mapping. Check your header mapping.")
            st.stop()

        # ---------- STEP 3: Compare ----------
        # build keys for comparison
        excel_keys = df_excel[["REGISTER_NO", "SUB_CODE"]].copy()
        pdf_keys = df_pdf[["REGISTER_NO", "SUB_CODE"]].copy()

        # find missing in PDF (excel key not in pdf)
        missing_in_pdf = pd.merge(excel_keys.drop_duplicates(), pdf_keys.drop_duplicates(),
                                  on=["REGISTER_NO", "SUB_CODE"], how="left", indicator=True)
        missing_in_pdf = missing_in_pdf[missing_in_pdf["_merge"] == "left_only"].drop(columns=["_merge"])

        # extra in PDF (pdf key not in excel)
        missing_in_excel = pd.merge(pdf_keys.drop_duplicates(), excel_keys.drop_duplicates(),
                                   on=["REGISTER_NO", "SUB_CODE"], how="left", indicator=True)
        missing_in_excel = missing_in_excel[missing_in_excel["_merge"] == "left_only"].drop(columns=["_merge"])

        # For mismatches, merge the two fully and compare chosen columns
        merged = pd.merge(df_excel, df_pdf, on=["REGISTER_NO", "SUB_CODE"], how="inner", suffixes=("_EXCEL", "_PDF"))
        compare_cols = ["SUBJECT_NAME", "GRADE", "GRADE_POINT", "RESULT"]

        # make sure comparison columns exist (fill if missing)
        for c in compare_cols:
            if f"{c}_EXCEL" not in merged.columns:
                merged[f"{c}_EXCEL"] = ""
            if f"{c}_PDF" not in merged.columns:
                merged[f"{c}_PDF"] = ""

        # normalize text columns for comparison
        merged["SUBJECT_NAME_EXCEL"] = merged["SUBJECT_NAME_EXCEL"].apply(normalize_subject_name_after_extraction)
        merged["SUBJECT_NAME_PDF"] = merged["SUBJECT_NAME_PDF"].apply(normalize_subject_name_after_extraction)

        mismatches = merged[
            merged.apply(
                lambda row: any(
                    str(row.get(f"{col}_EXCEL", "")).strip().upper() != str(row.get(f"{col}_PDF", "")).strip().upper()
                    for col in compare_cols
                ),
                axis=1
            )
        ]

        # ---------- STEP 4: Display ----------
        st.subheader("Comparison Results")

        if mismatches.empty and missing_in_pdf.empty and missing_in_excel.empty:
            st.success("All records match perfectly! No mismatches or missing records found.")
        else:
            if not mismatches.empty:
                st.error(f"{len(mismatches)} mismatched rows found!")
                st.dataframe(mismatches)
                buf = io.StringIO()
                mismatches.to_csv(buf, index=False)
                st.download_button("Download Mismatch Report (CSV)", buf.getvalue(), "mismatch_report.csv", "text/csv")

            if not missing_in_pdf.empty:
                st.warning(f"{len(missing_in_pdf)} records missing in PDF.")
                st.dataframe(missing_in_pdf)
                buf2 = io.StringIO()
                missing_in_pdf.to_csv(buf2, index=False)
                st.download_button("Download Missing in PDF (CSV)", buf2.getvalue(), "missing_in_pdf.csv", "text/csv")

            if not missing_in_excel.empty:
                st.warning(f"{len(missing_in_excel)} extra records found in PDF (not in Excel).")
                st.dataframe(missing_in_excel)
                buf3 = io.StringIO()
                missing_in_excel.to_csv(buf3, index=False)
                st.download_button("Download Extra in PDF (CSV)", buf3.getvalue(), "extra_in_pdf.csv", "text/csv")

    except Exception as e:
        st.error(f"Error: {e}")

else:
    st.info("Please upload both Excel and PDF files to begin comparison.")
