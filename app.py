import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
from datetime import datetime

# ------------------------------- PHASE SELECTION -------------------------------
st.set_page_config(page_title="Excel vs PDF Comparator", layout="wide")
st.title("Excel vs PDF Comparator Tool")

phase = st.radio("Select Phase", ["Phase 1: Student Info Comparison", "Phase 2: Marks/Grades Comparison"])

# ------------------------------- COMMON HELPERS -------------------------------

def normalize_register(v):
    s = str(v).strip()
    if s in ["", "nan", "None"]:
        return ""
    try:
        f = float(s)
        i = int(f)
        return str(i)
    except Exception:
        digits = "".join(re.findall(r"\d+", s))
        return digits.lstrip("0") if digits else re.sub(r"\.0+$", "", s)

def normalize_name(name):
    return re.sub(r"\s+", " ", str(name)).strip().upper()

def normalize_text(txt):
    return re.sub(r"\s+", " ", str(txt)).strip().upper()

def normalize_gender(g):
    g = str(g).strip().upper()
    if g.startswith("M"): return "MALE"
    if g.startswith("F"): return "FEMALE"
    return g

def normalize_dob(dob):
    dob = str(dob).strip()
    if dob == "" or dob.lower() == "nan":
        return ""
    for fmt in ("%Y-%m-%d", "%d-%b-%Y", "%d/%b/%Y", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            d = datetime.strptime(dob, fmt)
            return d.strftime("%d-%b-%Y")
        except Exception:
            continue
    return dob

def normalize_subject_name_after_extraction(name):
    name = str(name)
    name = re.sub(r"[\*\u2217\u2731\u204E\uFE61\uFF0A\u002A]+", "", name)
    name = re.sub(r"[^A-Za-z0-9\-\(\)\:\,\/\s&]", "", name)
    return re.sub(r"\s+", " ", name).strip().upper()

def normalize_result(value):
    v = str(value).strip().upper()
    if v == "F": return "RA"
    elif v == "P": return "PASS"
    return v

# ------------------------------- PHASE 1: STUDENT INFO -------------------------------
def extract_pdf_students(pdf_bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        all_text = ""
        for page in pdf.pages:
            text = page.extract_text() or ""
            text = text.replace("\xa0", " ")
            all_text += text + "\n"

    clean_text = re.sub(r"\s+", " ", all_text.strip())
    blocks = re.split(r"UMIS\s*No\.?\s*[:\-]?\s*", clean_text, flags=re.I)

    students = []
    for blk in blocks[1:]:
        umis_match = re.match(r"(\d{10,20})", blk)
        umis_no = umis_match.group(1) if umis_match else ""
        rest = blk[len(umis_no):].strip()

        name_reg = re.search(
            r"NAME\s+OF\s+THE\s+(?:CANDIDATE\s+)?([A-Z\.\s]+?)\s+REGISTER\s*NO\.?\s*[:\-]?\s*([0-9]{9,15})",
            rest, flags=re.I)
        if not name_reg:
            name_reg = re.search(
                r"NAME\s+OF\s+THE\s+([A-Z\.\s]+?)\s+CANDIDATE\s+REGISTER\s*NO\.?\s*[:\-]?\s*([0-9]{9,15})",
                rest, flags=re.I)

        name = name_reg.group(1) if name_reg else ""
        register_no = name_reg.group(2) if name_reg else ""

        dob_match = re.search(r"DATE\s+OF\s+BIRTH\s*[:\-]?\s*([0-9]{1,2}[-/][A-Za-z]{3,}[-/][0-9]{2,4})", rest, flags=re.I)
        gender_match = re.search(r"GENDER\s*[:\-]?\s*([A-Za-z]+)", rest, flags=re.I)
        prog_match = re.search(r"PROGRAMME\s*&?\s*B\.?(?:E|TECH)\.?\s*([A-Z\s]+?)\s+REGULATIONS", rest, flags=re.I)

        students.append({
            "UMIS_NO": umis_no,
            "REGISTER_NO": normalize_register(register_no),
            "NAME": normalize_name(name),
            "DOB": normalize_dob(dob_match.group(1)) if dob_match else "",
            "GENDER": normalize_gender(gender_match.group(1)) if gender_match else "",
            "PROGRAMME": normalize_text(prog_match.group(1)) if prog_match else "",
        })

    return pd.DataFrame(students)

# ------------------------------- PHASE 2: MARKS/GRADES -------------------------------
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

def extract_pdf_data(pdf_bytes):
    text_data = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            page_txt = page.extract_text() or ""
            page_txt = page_txt.replace("\xa0", " ").replace("\u200b", "")
            text_data += page_txt + "\n"

    text_data = re.sub(r"\s+", " ", text_data).strip()
    records = []

    reg_iter = list(re.finditer(r"(REGISTER\s*NO\.?\s*:?\s*([0-9A-Za-z\.\-E\+]+))", text_data, flags=re.I))
    if not reg_iter:
        reg_iter = list(re.finditer(r"(REGISTER\s*NO\.?\s*:?\s*([0-9]+))", text_data, flags=re.I))
    if not reg_iter:
        return pd.DataFrame(records)

    for idx, m in enumerate(reg_iter):
        reg_raw = m.group(2)
        block_start = m.end()
        block_end = reg_iter[idx + 1].start() if idx + 1 < len(reg_iter) else len(text_data)
        block = text_data[block_start:block_end].strip()
        regno = normalize_register(reg_raw)

        pat_a = re.compile(
            r"(?:\d{1,3}\s+)?([A-Z]{2,3}\d{3,4})\s+([^\d]{2,160}?)\s+(\d+(?:\.\d+)?)\s+([A-Z\+OU]{1,2})\s+(\d+)\s+(PASS|RA|P)",
            flags=re.I,
        )
        pat_b = re.compile(
            r"(?:\d{1,3}\s+)?([A-Z]{2,3}\d{3,4})\s+([A-Za-z\*\u2217\u2731\u204E\uFE61\uFF0A\u002A0-9\(\)\-\:\,\/&\s]{1,160}?)\s+(?:([\d\.]+)\s+)?(?:([A-Z\+OU]{1,2})\s+)?(?:([\d]+)\s+)?(PASS|RA|P)\b",
            flags=re.I,
        )
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

# ------------------------------- FILE UPLOAD -------------------------------
excel_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
pdf_file = st.file_uploader("Upload PDF File", type=["pdf"])

if excel_file and pdf_file:
    try:
        if phase == "Phase 1: Student Info Comparison":
            # ---------- Phase 1 Excel ----------
            df_excel = pd.read_excel(excel_file, dtype=str)
            df_excel.columns = [c.strip().upper() for c in df_excel.columns]
            expected = ["REGISTER NO", "STUDENT NAME", "DATE OF BIRTH", "GENDER", "PROGRAMME", "UMIS NO"]
            missing_cols = [c for c in expected if c not in df_excel.columns]
            if missing_cols:
                st.error(f"Missing columns in Excel: {missing_cols}")
                st.stop()
            df_excel.rename(columns={
                "REGISTER NO": "REGISTER_NO",
                "STUDENT NAME": "NAME",
                "DATE OF BIRTH": "DOB",
                "GENDER": "GENDER",
                "PROGRAMME": "PROGRAMME",
                "UMIS NO": "UMIS_NO"
            }, inplace=True)

            for col in df_excel.columns:
                df_excel[col] = df_excel[col].astype(str).fillna("").str.strip()

            df_excel["REGISTER_NO"] = df_excel["REGISTER_NO"].apply(normalize_register)
            df_excel["NAME"] = df_excel["NAME"].apply(normalize_name)
            df_excel["DOB"] = df_excel["DOB"].apply(normalize_dob)
            df_excel["GENDER"] = df_excel["GENDER"].apply(normalize_gender)
            df_excel["PROGRAMME"] = df_excel["PROGRAMME"].apply(normalize_text)
            df_excel["UMIS_NO"] = df_excel["UMIS_NO"].astype(str).str.strip()

            st.success("Excel file loaded successfully.")
            st.dataframe(df_excel.head())

            # ---------- Phase 1 PDF ----------
            df_pdf = extract_pdf_students(pdf_file.read())
            if df_pdf.empty:
                st.error("Could not extract any student data from PDF.")
                st.stop()
            st.success(f"Extracted {len(df_pdf)} students from PDF.")
            st.dataframe(df_pdf.head())

            # ---------- Compare Phase 1 ----------
            merged = pd.merge(
                df_excel, df_pdf,
                on=["REGISTER_NO", "UMIS_NO"],
                how="outer",
                suffixes=("_EXCEL", "_PDF"),
                indicator=True
            )

            missing_in_pdf = merged[merged["_merge"] == "left_only"]
            missing_in_excel = merged[merged["_merge"] == "right_only"]

            compare_cols = ["NAME", "DOB", "GENDER", "PROGRAMME"]
            mismatches = []
            for _, row in merged[merged["_merge"] == "both"].iterrows():
                diffs = [col for col in compare_cols
                         if str(row[f"{col}_EXCEL"]).strip().upper() != str(row[f"{col}_PDF"]).strip().upper()]
                if diffs:
                    mismatches.append({**row, "MISMATCH_FIELDS": ", ".join(diffs)})
            mismatches = pd.DataFrame(mismatches)

            st.subheader("Comparison Results")
            if mismatches.empty and missing_in_pdf.empty and missing_in_excel.empty:
                st.success("All student info matches perfectly!")
            else:
                if not mismatches.empty:
                    st.error(f"{len(mismatches)} mismatched student records found!")
                    st.dataframe(mismatches)
                if not missing_in_pdf.empty:
                    st.warning(f"{len(missing_in_pdf)} students missing in PDF.")
                    st.dataframe(missing_in_pdf)
                if not missing_in_excel.empty:
                    st.warning(f"{len(missing_in_excel)} students missing in Excel.")
                    st.dataframe(missing_in_excel)

        elif phase == "Phase 2: Marks/Grades Comparison":
            # ---------- Phase 2 Excel ----------
            df_excel = pd.read_excel(excel_file, dtype=str)
            df_excel.columns = [str(c).strip().upper() for c in df_excel.columns]
            mapped_cols = {col: HEADER_MAP[col.strip().upper()] for col in df_excel.columns if col.strip().upper() in HEADER_MAP}
            df_excel.rename(columns=mapped_cols, inplace=True)
            df_excel = df_excel.fillna("").astype(str).apply(lambda col: col.str.strip())

            if "REGISTER_NO" in df_excel.columns:
                df_excel["REGISTER_NO"] = df_excel["REGISTER_NO"].apply(normalize_register)
            if "SUBJECT_NAME" in df_excel.columns:
                df_excel["SUBJECT_NAME"] = df_excel["SUBJECT_NAME"].apply(normalize_subject_name_after_extraction)
            if "RESULT" in df_excel.columns:
                df_excel["RESULT"] = df_excel["RESULT"].apply(normalize_result)

            st.subheader("Excel Data Preview")
            st.dataframe(df_excel.head())

            # ---------- Phase 2 PDF ----------
            pdf_bytes = pdf_file.read()
            df_pdf = extract_pdf_data(pdf_bytes)
            if df_pdf.empty:
                st.error("Could not extract any course data from PDF.")
                st.stop()

            st.success(f"Extracted {len(df_pdf)} subject records from PDF.")
            st.dataframe(df_pdf.head())

            df_pdf["REGISTER_NO"] = df_pdf["REGISTER_NO"].apply(normalize_register)
            df_pdf["SUBJECT_NAME"] = df_pdf["SUBJECT_NAME"].apply(normalize_subject_name_after_extraction)
            df_pdf["SUB_CODE"] = df_pdf["SUB_CODE"].astype(str).str.strip().str.upper()
            df_pdf["RESULT"] = df_pdf["RESULT"].apply(normalize_result)

            if "SUB_CODE" in df_excel.columns:
                df_excel["SUB_CODE"] = df_excel["SUB_CODE"].astype(str).str.strip().str.upper()
            else:
                st.error("Excel missing SUB_CODE after mapping.")
                st.stop()

            excel_keys = df_excel[["REGISTER_NO", "SUB_CODE"]].copy()
            pdf_keys = df_pdf[["REGISTER_NO", "SUB_CODE"]].copy()

            missing_in_pdf = pd.merge(excel_keys.drop_duplicates(), pdf_keys.drop_duplicates(),
                                      on=["REGISTER_NO", "SUB_CODE"], how="left", indicator=True)
            missing_in_pdf = missing_in_pdf[missing_in_pdf["_merge"] == "left_only"].drop(columns=["_merge"])

            missing_in_excel = pd.merge(pdf_keys.drop_duplicates(), excel_keys.drop_duplicates(),
                                       on=["REGISTER_NO", "SUB_CODE"], how="left", indicator=True)
            missing_in_excel = missing_in_excel[missing_in_excel["_merge"] == "left_only"].drop(columns=["_merge"])

            merged = pd.merge(df_excel, df_pdf, on=["REGISTER_NO", "SUB_CODE"], how="inner", suffixes=("_EXCEL", "_PDF"))
            compare_cols = ["SUBJECT_NAME", "GRADE", "GRADE_POINT", "RESULT"]
            for c in compare_cols:
                if f"{c}_EXCEL" not in merged.columns:
                    merged[f"{c}_EXCEL"] = ""
                if f"{c}_PDF" not in merged.columns:
                    merged[f"{c}_PDF"] = ""

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
