import io
import os
import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(page_title="FL Dashboard", layout="wide")

@st.cache_data
def load_data(file_bytes_or_path):
    # Accept both uploaded bytes and file path
    if isinstance(file_bytes_or_path, (bytes, bytearray)):
        xls = pd.ExcelFile(io.BytesIO(file_bytes_or_path))
    else:
        xls = pd.ExcelFile(file_bytes_or_path)
    # Assume the first sheet
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    # Basic cleaning
    # Strip column names
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    # Forward-fill identity columns to deal with sparsely filled tables
    identity_cols = [c for c in df.columns if any(k in str(c).lower() for k in ["employee", "fl code", "currency"])]
    if identity_cols:
        df[identity_cols] = df[identity_cols].ffill()
    # Standardize language columns
    src_col = next((c for c in df.columns if "source language" in str(c).lower()), None)
    tgt_col = next((c for c in df.columns if "target language" in str(c).lower()), None)
    # Subject columns (union across levels if present)
    subj_cols = [c for c in df.columns if "subject areas level" in str(c).lower()]
    # Step name
    step_col = next((c for c in df.columns if "step name" in str(c).lower()), None)
    # Rate columns
    flat_rate_col = next((c for c in df.columns if "flat rate" in str(c).lower()), None)
    range1_col = next((c for c in df.columns if "range-1" in str(c).lower()), None)
    range2_col = next((c for c in df.columns if "range-2" in str(c).lower()), None)
    currency_col = next((c for c in df.columns if "currency" in str(c).lower()), None)

    # Clean up whitespace in key columns
    for c in [src_col, tgt_col, step_col, currency_col]:
        if c and c in df:
            df[c] = df[c].astype(str).str.strip().replace({"nan": np.nan})

    # Build language pair
    if src_col and tgt_col:
        df["Language Pair"] = df[src_col].fillna("").astype(str).str.strip() + " \u2192 " + df[tgt_col].fillna("").astype(str).str.strip()
# Remove the 'r' prefix so Python processes the Unicode escape
df["Language Pair"] = df["Language Pair"].str.replace("^\s*\u...", ...)    
    else:
        df["Language Pair"] = np.nan

    # Build unified Subject Area column (union across levels, explode to 1 per row)
    if subj_cols:
        subj_df = df[subj_cols].astype(str).replace({"nan": np.nan})
        df["__SubjectsMerged"] = subj_df.apply(lambda r: [x for x in r.dropna().tolist() if str(x).strip() != ""], axis=1)
    else:
        df["__SubjectsMerged"] = [[] for _ in range(len(df))]

    # Attempt to tidy ranges to numeric when possible
    for c in [flat_rate_col, range1_col, range2_col]:
        if c and c in df:
            # Coerce to numeric safely
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # Standardize employee columns for display
    emp_name_col = next((c for c in df.columns if "employee name" in str(c).lower()), None)
    emp_code_col = next((c for c in df.columns if "employee code" in str(c).lower()), None)
    fl_code_col  = next((c for c in df.columns if "fl code" in str(c).lower()), None)

    return df, {
        "src_col": src_col,
        "tgt_col": tgt_col,
        "step_col": step_col,
        "flat_rate_col": flat_rate_col,
        "range1_col": range1_col,
        "range2_col": range2_col,
        "currency_col": currency_col,
        "emp_name_col": emp_name_col,
        "emp_code_col": emp_code_col,
        "fl_code_col": fl_code_col,
        "subj_cols": subj_cols,
    }

def kpi_card(label, value, help_text=None):
    st.metric(label, value, help=help_text)

st.title("📊 FL Dashboard by Language Pair & Subject Area")

st.sidebar.header("Filters")

# File input
uploaded = st.sidebar.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

default_path = "/mnt/data/hr.employee (6).xlsx"
use_default = False
if uploaded is None and os.path.exists(default_path):
    use_default = True

if uploaded is None and not use_default:
    st.info("Upload your Excel file from the sidebar to get started.")
    st.stop()

# Load data
if use_default:
    df, meta = load_data(default_path)
else:
    df, meta = load_data(uploaded.getvalue())

# Build filter controls
src_col = meta["src_col"]
tgt_col = meta["tgt_col"]
step_col = meta["step_col"]
currency_col = meta["currency_col"]

# Distinct language values (dropna + sorted)
src_values = sorted([x for x in df[src_col].dropna().unique().tolist()]) if src_col else []
tgt_values = sorted([x for x in df[tgt_col].dropna().unique().tolist()]) if tgt_col else []
pair_values = sorted([x for x in df["Language Pair"].dropna().unique().tolist() if "→" in str(x)])

subj_pool = sorted({s for row in df["__SubjectsMerged"] for s in (row if isinstance(row, list) else [])})

st.sidebar.subheader("Language Pair")
pair = st.sidebar.selectbox("Select Language Pair", options=["(All)"] + pair_values, index=0)

# Optional: Allow independent source/target selection if (All)
with st.sidebar.expander("Advanced language filter", expanded=False):
    src_sel = st.multiselect("Source Language", options=src_values, default=[])
    tgt_sel = st.multiselect("Target Language", options=tgt_values, default=[])

st.sidebar.subheader("Subject Area")
subj_sel = st.sidebar.multiselect("Choose one or more", options=subj_pool, default=[])

with st.sidebar.expander("More filters", expanded=False):
    step_sel = st.multiselect("Step Name", options=sorted([x for x in df[step_col].dropna().unique().tolist()]) if step_col else [], default=[])
    currency_sel = st.multiselect("Currency", options=sorted([x for x in df[currency_col].dropna().unique().tolist()]) if currency_col else [], default=[])

# Apply filters
flt = df.copy()

if pair and pair != "(All)":
    flt = flt[flt["Language Pair"] == pair]

# Advanced independent source/target filters if user set them
if src_sel and src_col:
    flt = flt[flt[src_col].isin(src_sel)]
if tgt_sel and tgt_col:
    flt = flt[flt[tgt_col].isin(tgt_sel)]

# Subjects: include rows that have ANY of the selected subjects
if subj_sel:
    flt = flt[flt["__SubjectsMerged"].apply(lambda lst: any(s in (lst or []) for s in subj_sel))]

if step_sel and step_col:
    flt = flt[flt[step_col].isin(step_sel)]
if currency_sel and currency_col:
    flt = flt[flt[currency_col].isin(currency_sel)]

# KPI Row
st.markdown("### Overview")
col1, col2, col3, col4 = st.columns(4)

# Unique FL count
emp_name_col = meta["emp_name_col"]
emp_code_col = meta["emp_code_col"]
fl_code_col = meta["fl_code_col"]

def unique_fl_count(df_):
    keys = []
    if fl_code_col and fl_code_col in df_:
        keys.append(fl_code_col)
    if emp_code_col and emp_code_col in df_:
        keys.append(emp_code_col)
    if emp_name_col and emp_name_col in df_:
        keys.append(emp_name_col)
    if keys:
        return df_[keys].dropna(how="all").drop_duplicates().shape[0]
    return df_.drop_duplicates().shape[0]

with col1:
    kpi_card("Matching FLs", unique_fl_count(flt))

# Unique Steps
with col2:
    if step_col and step_col in flt:
        kpi_card("Unique Steps", int(flt[step_col].dropna().nunique()))
    else:
        kpi_card("Unique Steps", "—")

# Currency info
with col3:
    if currency_col and currency_col in flt:
        currs = sorted([str(x) for x in flt[currency_col].dropna().unique().tolist()])
        kpi_card("Currencies", ", ".join(currs) if currs else "—")
    else:
        kpi_card("Currencies", "—")

# Rate summary (by currency)
with col4:
    flat_rate_col = meta["flat_rate_col"]
    if flat_rate_col and flat_rate_col in flt and not flt[flat_rate_col].dropna().empty:
        kpi_card("Median Flat Rate", f'{flt[flat_rate_col].median():,.2f}')
    else:
        kpi_card("Median Flat Rate", "—")

st.divider()

# Grouped stats by Step & Currency
flat_rate_col = meta["flat_rate_col"]
range1_col = meta["range1_col"]
range2_col = meta["range2_col"]

st.markdown("### Rates by Step & Currency")
if step_col and currency_col and flat_rate_col and flat_rate_col in flt:
    grp_cols = [c for c in [step_col, currency_col] if c]
    summary = (
        flt.groupby(grp_cols, dropna=False)
           .agg(
               records=("Language Pair", "count"),
               flat_min=(flat_rate_col, "min"),
               flat_med=(flat_rate_col, "median"),
               flat_max=(flat_rate_col, "max"),
               r1_min=(range1_col, "min") if range1_col in flt else ("Language Pair", "count"),
               r1_max=(range1_col, "max") if range1_col in flt else ("Language Pair", "count"),
               r2_min=(range2_col, "min") if range2_col in flt else ("Language Pair", "count"),
               r2_max=(range2_col, "max") if range2_col in flt else ("Language Pair", "count"),
           )
           .reset_index()
    )
    # Clean columns if range1/2 missing
    to_drop = [c for c in ["r1_min","r1_max","r2_min","r2_max"] if c in summary and summary[c].dtype == "int64" and summary[c].max() == summary[c].min()]
    if to_drop:
        summary = summary.drop(columns=to_drop)
    st.dataframe(summary, use_container_width=True)
else:
    st.info("Rate columns not found or no data after filters.")

st.divider()

# Detailed table
st.markdown("### Matching FLs — Details")

display_cols = []
for c in [fl_code_col, emp_code_col, emp_name_col, src_col, tgt_col, "Language Pair", step_col, currency_col, flat_rate_col, range1_col, range2_col]:
    if c and c in flt.columns and c not in display_cols:
        display_cols.append(c)

# Add a readable Subject list
flt = flt.copy()
flt["Subjects"] = flt["__SubjectsMerged"].apply(lambda lst: ", ".join(lst) if isinstance(lst, list) else "")

if "Subjects" not in display_cols:
    display_cols.append("Subjects")

if display_cols:
    st.dataframe(flt[display_cols], use_container_width=True)
else:
    st.dataframe(flt, use_container_width=True)

# Export
csv_bytes = flt[display_cols].to_csv(index=False).encode("utf-8") if display_cols else flt.to_csv(index=False).encode("utf-8")
st.download_button("Download filtered data (CSV)", data=csv_bytes, file_name="fl_filtered_export.csv", mime="text/csv")

st.caption("Tip: Use the sidebar to filter by language pair and subject area. Expand 'Advanced' sections for more controls.")
