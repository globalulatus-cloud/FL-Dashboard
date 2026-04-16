import io
import os
import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(page_title="FL Dashboard", layout="wide")


# ---------------------------
# Utility Functions
# ---------------------------

def find_column(df, keywords):
    """Robust column finder using keyword matching."""
    for col in df.columns:
        col_lower = str(col).lower()
        if all(k in col_lower for k in keywords):
            return col
    return None


def safe_strip(series):
    """Strip safely without breaking NaNs."""
    return series.astype(str).str.strip().replace("nan", np.nan)


# ---------------------------
# Data Loading
# ---------------------------

@st.cache_data
def load_data(file_bytes_or_path):

    if isinstance(file_bytes_or_path, (bytes, bytearray)):
        df = pd.read_excel(io.BytesIO(file_bytes_or_path))
    else:
        df = pd.read_excel(file_bytes_or_path)

    # Clean column names
    df.columns = [str(c).strip() for c in df.columns]

    # Detect columns
    src_col = find_column(df, ["source", "language"])
    tgt_col = find_column(df, ["target", "language"])
    step_col = find_column(df, ["step"])
    currency_col = find_column(df, ["currency"])

    flat_rate_col = find_column(df, ["flat", "rate"])
    range1_col = find_column(df, ["range-1"])
    range2_col = find_column(df, ["range-2"])

    emp_name_col = find_column(df, ["employee", "name"])
    emp_code_col = find_column(df, ["employee", "code"])
    fl_code_col = find_column(df, ["fl", "code"])

    subj_cols = [c for c in df.columns if "subject areas level" in c.lower()]

    # Clean key columns
    for col in [src_col, tgt_col, step_col, currency_col]:
        if col:
            df[col] = safe_strip(df[col])

    # Forward fill identity columns
    identity_cols = [c for c in [emp_name_col, emp_code_col, fl_code_col, currency_col] if c]
    if identity_cols:
        df[identity_cols] = df[identity_cols].ffill()

    # Build language pair
    if src_col and tgt_col:
        df["Language Pair"] = df[src_col].fillna("") + " → " + df[tgt_col].fillna("")
        df["Language Pair"] = df["Language Pair"].str.strip()
    else:
        df["Language Pair"] = np.nan

    # Subjects (optimized)
    if subj_cols:
        df[subj_cols] = df[subj_cols].replace({np.nan: None})
        df["__SubjectsMerged"] = df[subj_cols].values.tolist()
        df["__SubjectsMerged"] = df["__SubjectsMerged"].apply(
            lambda x: [i for i in x if i]
        )
    else:
        df["__SubjectsMerged"] = [[] for _ in range(len(df))]

    # Convert numeric columns safely
    for col in [flat_rate_col, range1_col, range2_col]:
        if col:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df, {
        "src_col": src_col,
        "tgt_col": tgt_col,
        "step_col": step_col,
        "currency_col": currency_col,
        "flat_rate_col": flat_rate_col,
        "range1_col": range1_col,
        "range2_col": range2_col,
        "emp_name_col": emp_name_col,
        "emp_code_col": emp_code_col,
        "fl_code_col": fl_code_col,
        "subj_cols": subj_cols,
    }


# ---------------------------
# KPI Logic
# ---------------------------

def unique_fl_count(df, meta):
    keys = [meta["fl_code_col"], meta["emp_code_col"], meta["emp_name_col"]]
    keys = [k for k in keys if k and k in df]

    if not keys:
        return len(df)

    return df[keys].drop_duplicates().shape[0]


# ---------------------------
# UI
# ---------------------------

st.title("FL Dashboard")

uploaded = st.sidebar.file_uploader("Upload Excel", type=["xlsx"])

if not uploaded:
    st.warning("Upload a file to proceed")
    st.stop()

df, meta = load_data(uploaded.getvalue())

src_col = meta["src_col"]
tgt_col = meta["tgt_col"]
step_col = meta["step_col"]
currency_col = meta["currency_col"]

# ---------------------------
# Filters
# ---------------------------

pair_values = sorted(df["Language Pair"].dropna().unique())

pair = st.sidebar.selectbox("Language Pair", ["(All)"] + pair_values)

subj_pool = sorted({s for lst in df["__SubjectsMerged"] for s in lst})
subj_sel = st.sidebar.multiselect("Subjects", subj_pool)

flt = df.copy()

if pair != "(All)":
    flt = flt[flt["Language Pair"] == pair]

if subj_sel:
    flt = flt[
        flt["__SubjectsMerged"].apply(lambda x: bool(set(x) & set(subj_sel)))
    ]

# ---------------------------
# KPIs
# ---------------------------

col1, col2, col3, col4 = st.columns(4)

col1.metric("Matching FLs", unique_fl_count(flt, meta))

if step_col:
    col2.metric("Steps", flt[step_col].nunique())
else:
    col2.metric("Steps", "-")

if currency_col:
    col3.metric("Currencies", ", ".join(flt[currency_col].dropna().unique()))
else:
    col3.metric("Currencies", "-")

if meta["flat_rate_col"]:
    col4.metric("Median Rate", round(flt[meta["flat_rate_col"]].median(), 2))
else:
    col4.metric("Median Rate", "-")

# ---------------------------
# Aggregation (Fixed)
# ---------------------------

st.subheader("Rates Summary")

if step_col and currency_col and meta["flat_rate_col"]:

    summary = (
        flt.groupby([step_col, currency_col])
        .agg(
            records=("Language Pair", "count"),
            flat_min=(meta["flat_rate_col"], "min"),
            flat_med=(meta["flat_rate_col"], "median"),
            flat_max=(meta["flat_rate_col"], "max"),
        )
        .reset_index()
    )

    st.dataframe(summary, use_container_width=True)

else:
    st.info("Missing required columns for summary")

# ---------------------------
# Table
# ---------------------------

st.subheader("Details")

flt["Subjects"] = flt["__SubjectsMerged"].apply(lambda x: ", ".join(x))

display_cols = [c for c in [
    meta["fl_code_col"],
    meta["emp_code_col"],
    meta["emp_name_col"],
    src_col,
    tgt_col,
    "Language Pair",
    step_col,
    currency_col,
    meta["flat_rate_col"],
    "Subjects"
] if c and c in flt.columns]

st.dataframe(flt[display_cols], use_container_width=True)

# ---------------------------
# Export
# ---------------------------

csv = flt.to_csv(index=False).encode("utf-8")
st.download_button("Download CSV", csv, "filtered.csv", "text/csv")
