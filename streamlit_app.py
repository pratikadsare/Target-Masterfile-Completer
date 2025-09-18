
import io
from io import BytesIO
import pandas as pd
import streamlit as st
from typing import List, Dict, Tuple, Optional
from openpyxl import load_workbook

st.set_page_config(page_title="Masterfile Filler", layout="wide")

# --- Header (as requested) ---
st.markdown("<h1 style='text-align: center;'>🧩 Masterfile Filler</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; font-style: italic;'>Innovating with AI Today ⏐ Leading Automation Tomorrow</h4>", unsafe_allow_html=True)
st.caption("Only the FIRST sheet is modified. Row 1 = headers, Row 2 is preserved, data is written starting from Row 3. Other sheets remain unchanged (names, styles, formulas, merges).")

# ---------- Helpers ----------
def read_raw(uploaded_file, sheet: Optional[str]) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame()
    name = uploaded_file.name.lower()
    try:
        if name.endswith((".csv", ".txt")):
            return pd.read_csv(uploaded_file)
        # Excel
        if sheet is not None:
            return pd.read_excel(uploaded_file, sheet_name=sheet, engine="openpyxl")
        # default first
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
        first = xls.sheet_names[0]
        return pd.read_excel(xls, sheet_name=first, engine="openpyxl")
    except Exception as e:
        st.error(f"Failed to read RAW: {e}")
        return pd.DataFrame()

def list_excel_sheets(uploaded_file) -> List[str]:
    if uploaded_file is None:
        return []
    try:
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
        return xls.sheet_names
    except Exception:
        return []

def read_two_col_mapping(uploaded_file) -> pd.DataFrame:
    """Expect exactly 2 columns:
       - 'header of row sheet' (or 'header of raw sheet')
       - 'header of masterfile template'
    """
    if uploaded_file is None:
        return pd.DataFrame()
    name = uploaded_file.name.lower()
    try:
        if name.endswith((".csv", ".txt")):
            m = pd.read_csv(uploaded_file)
        else:
            m = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Failed to read mapping: {e}")
        return pd.DataFrame()
    if m.empty:
        return m

    # Detect columns by flexible names
    cols_norm = {c.strip().lower(): c for c in m.columns}
    raw_keys = ["header of row sheet", "header of raw sheet", "raw header", "raw", "source"]
    tpl_keys = ["header of masterfile template", "template header", "masterfile header", "template", "target"]

    raw_col = next((cols_norm[k] for k in raw_keys if k in cols_norm), None)
    tpl_col = next((cols_norm[k] for k in tpl_keys if k in cols_norm), None)
    if not raw_col or not tpl_col:
        st.error("Mapping must contain two columns: 'header of row sheet' and 'header of masterfile template'.")
        return pd.DataFrame()

    m = m[[raw_col, tpl_col]].copy()
    m.columns = ["raw_header", "template_header"]
    m["raw_header"] = m["raw_header"].astype(str).str.strip()
    m["template_header"] = m["template_header"].astype(str).str.strip()
    m = m[(m["raw_header"] != "") & (m["template_header"] != "")].drop_duplicates(subset=["template_header"])
    return m.reset_index(drop=True)

def build_header_index_first_sheet(ws, header_row: int = 1) -> Dict[str, int]:
    """Return a dict normalized header -> column index for the FIRST sheet row1 only."""
    headers = {}
    max_col = ws.max_column or 1
    for c in range(1, max_col + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            continue
        key = str(v).strip().lower()
        if key and key not in headers:
            headers[key] = c
    return headers

def fill_first_sheet_by_headers(template_bytes: BytesIO, mapping_df: pd.DataFrame, raw_df: pd.DataFrame, template_filename: str) -> BytesIO:
    """Write values ONLY on the first sheet using headers in row 1. Start writing from row 3.
       Other sheets remain untouched (names, formatting, formulas, merges).
    """
    keep_vba = template_filename.lower().endswith(".xlsm")
    wb = load_workbook(filename=template_bytes, data_only=False, keep_vba=keep_vba)
    ws = wb.worksheets[0]  # FIRST sheet only

    # Build header index from row 1
    tpl_header_to_col = build_header_index_first_sheet(ws, header_row=1)
    if not tpl_header_to_col:
        raise ValueError("No headers found in row 1 of the first sheet. Please ensure row 1 contains headers.")

    # Map raw columns by lowercase
    raw_norm = {c.strip().lower(): c for c in raw_df.columns}

    # Build mapping pairs: (raw_col_name, target_col_idx)
    pairs = []
    missing_raw, missing_tpl = [], []
    for _, r in mapping_df.iterrows():
        raw_hdr = r["raw_header"].strip().lower()
        tpl_hdr = r["template_header"].strip().lower()
        raw_col_name = raw_norm.get(raw_hdr)
        col_idx = tpl_header_to_col.get(tpl_hdr)
        if raw_col_name is None:
            missing_raw.append(r["raw_header"])
            continue
        if col_idx is None:
            missing_tpl.append(r["template_header"])
            continue
        pairs.append((raw_col_name, col_idx))

    if missing_tpl:
        st.warning("Template headers not found in row 1 (skipped): " + ", ".join(sorted(set(missing_tpl))))
    if missing_raw:
        st.warning("RAW columns missing (skipped): " + ", ".join(sorted(set(missing_raw))))

    # Write starting from row 3
    start_row = 3
    for i, (_, raw_row) in enumerate(raw_df.iterrows()):
        out_row = start_row + i
        for raw_col_name, col_idx in pairs:
            val = raw_row[raw_col_name]
            ws.cell(row=out_row, column=col_idx, value=("" if pd.isna(val) else val))

    # Save to bytes
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ---------- UI (3 tabs) ----------
with st.tabs(["1) Upload Raw Data", "2) Upload Masterfile (XLSX; FIRST sheet will be filled)", "3) Mapping (2 columns) & Download"]):
    # Tab 1: Raw
    st.subheader("Upload Raw Data (CSV/XLSX)")
    raw_file = st.file_uploader("Raw file", type=["csv","xlsx"], key="raw_file")
    raw_sheet = None
    if raw_file is not None and raw_file.name.lower().endswith(".xlsx"):
        try:
            sheets = pd.ExcelFile(raw_file, engine="openpyxl").sheet_names
            if sheets:
                raw_sheet = st.selectbox("Select RAW sheet", options=sheets, index=0)
        except Exception as e:
            st.error(f"Could not read RAW Excel: {e}")

    raw_df = read_raw(raw_file, raw_sheet)
    if not raw_df.empty:
        st.success(f"Loaded RAW: {len(raw_df):,} rows × {len(raw_df.columns):,} columns.")
        st.dataframe(raw_df.head(50), use_container_width=True)

    # Tab 2: Template
    st.subheader("Upload Masterfile (XLSX only)")
    template_file = st.file_uploader("Masterfile (Excel .xlsx or .xlsm)", type=["xlsx","xlsm"], key="template_file")
    tpl_preview = None
    first_sheet_name = None
    if template_file is not None:
        try:
            xls = pd.ExcelFile(template_file, engine="openpyxl")
            first_sheet_name = xls.sheet_names[0]
            tpl_preview = pd.read_excel(xls, sheet_name=first_sheet_name, nrows=5, engine="openpyxl")
            st.info(f"Only the FIRST sheet will be modified: **{first_sheet_name}**")
            st.dataframe(tpl_preview, use_container_width=True)
        except Exception as e:
            st.error(f"Failed to read template for preview: {e}")

    # Tab 3: Mapping + Process
    st.subheader("Upload 2-column mapping (XLSX/CSV)")
    st.markdown("Columns required: **header of row sheet** (RAW) → **header of masterfile template** (FIRST sheet, row 1).")
    mapping_file = st.file_uploader("Mapping file", type=["xlsx","csv"], key="mapping_file")

    # Download mapping template
    def mapping_template_bytes():
        df = pd.DataFrame({
            "header of row sheet": ["sku","name","description","price","qty","category"],
            "header of masterfile template": ["SKU","Title","Description","Price","Quantity","Category"]
        })
        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            df.to_excel(w, index=False, sheet_name="mapping")
        out.seek(0)
        return out
    if st.button("⬇️ Download mapping template"):
        out = mapping_template_bytes()
        st.download_button("Download mapping_template.xlsx", data=out.getvalue(), file_name="mapping_template.xlsx")

    # Process & Download
    can_process = (not raw_df.empty) and (template_file is not None) and (mapping_file is not None)
    if not can_process:
        st.info("Upload RAW, Template, and Mapping to enable processing.")

    if st.button("⚙️ Process & Download", type="primary", disabled=not can_process):
        try:
            # Read mapping
            mapping_name = mapping_file.name
            if mapping_name.lower().endswith(".csv"):
                mapping_df = pd.read_csv(mapping_file)
            else:
                mapping_df = pd.read_excel(mapping_file, engine="openpyxl")
            # Normalize mapping
            def normalize_mapping(df):
                cols_norm = {c.strip().lower(): c for c in df.columns}
                raw_keys = ["header of row sheet", "header of raw sheet", "raw header", "raw", "source"]
                tpl_keys = ["header of masterfile template", "template header", "masterfile header", "template", "target"]
                raw_col = next((cols_norm[k] for k in raw_keys if k in cols_norm), None)
                tpl_col = next((cols_norm[k] for k in tpl_keys if k in cols_norm), None)
                if not raw_col or not tpl_col:
                    raise ValueError("Mapping must have 'header of row sheet' and 'header of masterfile template'.")
                out = df[[raw_col, tpl_col]].copy()
                out.columns = ["raw_header", "template_header"]
                out["raw_header"] = out["raw_header"].astype(str).str.strip()
                out["template_header"] = out["template_header"].astype(str).str.strip()
                out = out[(out["raw_header"] != "") & (out["template_header"] != "")].drop_duplicates(subset=["template_header"])
                return out.reset_index(drop=True)
            mapping_df = normalize_mapping(mapping_df)

            # Read template bytes
            tpl_bytes = BytesIO(template_file.getbuffer())

            # Fill
            out_bytes = fill_first_sheet_by_headers(
                template_bytes=tpl_bytes,
                mapping_df=mapping_df,
                raw_df=raw_df,
                template_filename=template_file.name
            )

            st.success("Done! Your updated masterfile is ready.")
            st.download_button(
                label="⬇️ Download Updated Masterfile (Excel)",
                data=out_bytes.getvalue(),
                file_name="filled_masterfile.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.exception(e)
