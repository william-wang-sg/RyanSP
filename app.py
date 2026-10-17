"""Update WC Contract & Location files – Databricks App (Streamlit).

Converted from the tkinter desktop version.

Contract file : in every column whose name contains "Peril",
                PES is replaced with PEA (whole-code match).
Location file :
  1. Rows whose OccupancyCodeType contains "AIR" are left alone.
  2. Other rows are matched to the lookup
     (Markel EDC Parameters.xlsx, sheet tConvertEDCOccupancy)
     on OccupancyCodeType = OCCSCHEME and OccupancyCode = OCCTYPE
     and get AIROccupancyType / AIROccupancyCode.  No match -> both blank.
  3. In LocPerils-1, PES is replaced with PEA.
"""

import re
import io
import os

import streamlit as st
import pandas as pd

# -------------------- Settings --------------------
# Contract
CONTRACT_COLUMN_KEY = "Peril"
CONTRACT_OLD, CONTRACT_NEW = "PES", "PEA"
WHOLE_CODE_ONLY = True

# Location
DEFAULT_LOOKUP_SHEET = "tConvertEDCOccupancy"
SKIP_TYPE_TEXT = "AIR"
PERIL_FIELD, PERIL_OLD, PERIL_NEW = "LocPerils-1", "PES", "PEA"
LOOKUP_FILE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Markel EDC Parameters.xlsx")


# ==================== Processing ====================

def read_csv_text(file):
    """Read a CSV as text, trying common encodings."""
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            file.seek(0)
            return pd.read_csv(file, dtype=str, keep_default_na=False, encoding=enc)
        except UnicodeDecodeError:
            continue
    raise ValueError("Could not decode the CSV with any supported encoding.")


def process_contract(contract_file, log):
    log(f"--- Contract: {contract_file.name} ---")
    contract = read_csv_text(contract_file)
    log(f"Rows: {len(contract):,}   columns: {len(contract.columns)}")

    target_cols = [c for c in contract.columns if CONTRACT_COLUMN_KEY.lower() in c.lower()]
    log(f"Columns to update: {target_cols}")

    pattern = (
        rf"(?<![A-Za-z0-9]){re.escape(CONTRACT_OLD)}(?![A-Za-z0-9])"
        if WHOLE_CODE_ONLY
        else re.escape(CONTRACT_OLD)
    )
    for col in target_cols:
        before = contract[col]
        after = before.str.replace(pattern, CONTRACT_NEW, regex=True)
        log(f"  {col}: {int((before != after).sum())} cells changed")
        contract[col] = after

    log(f"Contract processing complete — {len(contract):,} rows")
    return contract


def read_lookup(lookup_path, sheet, log):
    log(f"Loading lookup: {os.path.basename(lookup_path)}")
    name = lookup_path.lower()
    if name.endswith((".xlsx", ".xlsm", ".xls")):
        xls = pd.ExcelFile(lookup_path, engine="openpyxl")
        if sheet not in xls.sheet_names:
            raise ValueError(
                f"Sheet '{sheet}' not found in {os.path.basename(lookup_path)}.\n"
                f"Available sheets: {xls.sheet_names}"
            )
        lookup = pd.read_excel(xls, sheet_name=sheet, dtype=str).fillna("")
    else:
        lookup = pd.read_csv(lookup_path, dtype=str, keep_default_na=False)

    lookup = lookup[["OCCSCHEME", "OCCTYPE", "AIROccupancyCode", "AIROccupancyType"]].copy()
    for c in lookup.columns:
        lookup[c] = lookup[c].str.strip()

    dups = lookup.duplicated(["OCCSCHEME", "OCCTYPE"]).sum()
    if dups:
        log(f"WARNING: {dups} duplicate lookup keys — keeping the first of each")
        lookup = lookup.drop_duplicates(["OCCSCHEME", "OCCTYPE"], keep="first")
    log(f"Lookup rows: {len(lookup):,}")
    return lookup


def process_location(location_file, reference_file, log):
    log(f"--- Location: {location_file.name} ---")
    loc = read_csv_text(location_file)
    log(f"Rows: {len(loc):,}   columns: {len(loc.columns)}")
    lookup = read_lookup(LOOKUP_FILE_PATH, DEFAULT_LOOKUP_SHEET, log)

    is_air = loc["OccupancyCodeType"].str.contains(
        SKIP_TYPE_TEXT, case=False, regex=False
    )
    already_air = loc[is_air].copy()
    to_convert = loc[~is_air].copy()

    converted = to_convert.merge(
        lookup,
        how="left",
        left_on=["OccupancyCodeType", "OccupancyCode"],
        right_on=["OCCSCHEME", "OCCTYPE"],
    )
    converted["OccupancyCodeType"] = converted["AIROccupancyType"].fillna("")
    converted["OccupancyCode"] = converted["AIROccupancyCode"].fillna("")

    out = pd.concat(
        [converted[loc.columns], already_air[loc.columns]], ignore_index=True
    )
    log(f"Already AIR (unchanged): {len(already_air):,}")
    log(f"Converted:               {len(to_convert):,}")
    log(f"  - matched in lookup:   {converted['AIROccupancyCode'].notna().sum():,}")
    log(f"  - no match (blanked):  {converted['AIROccupancyCode'].isna().sum():,}")

    out[PERIL_FIELD] = out[PERIL_FIELD].str.replace(PERIL_OLD, PERIL_NEW, regex=False)
    log(f"{PERIL_FIELD}: {PERIL_OLD} -> {PERIL_NEW} done")

    if reference_file:
        log(f"--- Compare with Alteryx output: {reference_file.name} ---")
        ref = read_csv_text(reference_file)
        log(f"Same columns/order: {list(ref.columns) == list(out.columns)}")
        log(f"Same row count:     {len(ref) == len(out)}")
        cmp = out.merge(
            ref, on="LocationID", suffixes=("_py", "_alt"), how="outer", indicator=True
        )
        log(f"LocationIDs only in one file: {(cmp['_merge'] != 'both').sum()}")
        bad = {
            c: int((cmp[c + "_py"] != cmp[c + "_alt"]).sum())
            for c in ref.columns
            if c != "LocationID"
        }
        bad = {c: n for c, n in bad.items() if n}
        log(
            f"Columns with differences: "
            f"{bad if bad else 'none — output matches Alteryx'}"
        )

    log(f"Location processing complete — {len(out):,} rows")
    return out


def df_to_csv_bytes(df):
    """Convert DataFrame to CSV bytes for download."""
    buf = io.BytesIO()
    df.to_csv(buf, index=False, encoding="utf-8")
    return buf.getvalue()


# ==================== Streamlit UI ====================

st.set_page_config(page_title="Update WC Contract & Location", layout="wide")
st.title("Update WC Contract & Location")
st.caption(
    "Upload your Contract CSV and/or Location CSV to process. "
    "Leave either section blank to skip it."
)

# ---------- File inputs ----------
col_contract, col_location = st.columns(2)

with col_contract:
    st.subheader("Contract")
    contract_file = st.file_uploader(
        "Contract CSV", type=["csv"], key="contract"
    )

with col_location:
    st.subheader("Location")
    location_file = st.file_uploader(
        "Location CSV", type=["csv"], key="location"
    )
    st.caption("Using bundled lookup: Markel EDC Parameters.xlsx")
    reference_file = st.file_uploader(
        "Alteryx output for comparison (optional)",
        type=["csv"],
        key="reference",
    )

st.divider()

# ---------- Session state defaults ----------
for _key in ("contract_csv", "contract_name", "location_csv", "location_name", "logs", "processed"):
    if _key not in st.session_state:
        st.session_state[_key] = None
if "processed" not in st.session_state:
    st.session_state["processed"] = False

# ---------- Run ----------
btn_run, btn_reset = st.columns(2)
with btn_run:
    run_clicked = st.button("Run", type="primary", use_container_width=True)
with btn_reset:
    reset_clicked = st.button("Reset", use_container_width=True)

if reset_clicked:
    for _key in ("contract_csv", "contract_name", "location_csv", "location_name", "logs"):
        st.session_state[_key] = None
    st.session_state["processed"] = False
    st.rerun()

if run_clicked:
    if not contract_file and not location_file:
        st.error("Upload a Contract CSV and/or a Location CSV.")
    else:
        logs: list[str] = []

        def log(msg: str):
            logs.append(msg)

        try:
            with st.spinner("Processing..."):
                if contract_file:
                    result = process_contract(contract_file, log)
                    stem = contract_file.name.rsplit(".", 1)[0]
                    st.session_state["contract_csv"] = df_to_csv_bytes(result)
                    st.session_state["contract_name"] = f"{stem}_updated.csv"
                if location_file:
                    result = process_location(
                        location_file, reference_file, log
                    )
                    stem = location_file.name.rsplit(".", 1)[0]
                    st.session_state["location_csv"] = df_to_csv_bytes(result)
                    st.session_state["location_name"] = f"{stem}_Output.csv"
            logs.append("=== Done ===")
            st.session_state["logs"] = "\n".join(logs)
            st.session_state["processed"] = True

        except Exception as e:
            logs.append(f"ERROR: {e}")
            st.session_state["logs"] = "\n".join(logs)
            st.session_state["processed"] = True
            st.error(str(e))

# ---------- Persistent results (survive reruns) ----------
if st.session_state["processed"]:
    if st.session_state["logs"]:
        st.text_area("Processing Log", st.session_state["logs"], height=300)

    st.subheader("Download Results")
    dl1, dl2 = st.columns(2)
    if st.session_state["contract_csv"] is not None:
        with dl1:
            st.download_button(
                label="Download Contract Result",
                data=st.session_state["contract_csv"],
                file_name=st.session_state["contract_name"],
                mime="text/csv",
            )
    if st.session_state["location_csv"] is not None:
        with dl2:
            st.download_button(
                label="Download Location Result",
                data=st.session_state["location_csv"],
                file_name=st.session_state["location_name"],
                mime="text/csv",
            )

    st.success("Processing finished! Click Reset to start over.")
