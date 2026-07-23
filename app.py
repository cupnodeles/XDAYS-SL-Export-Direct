import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="XDAYS SL Export Direct", layout="wide")
st.title("XDAYS SL Export Direct")

# ── Raw file headers (53 columns) ────────────────────────────────────────────
RAW_HEADERS = [
    "S.No", "Date", "Time", "Debtor", "Account No.", "Card No.", "Service No.",
    "DPD", "Call Status", "Status", "Remark", "Remark By", "Remark Type",
    "Field Visit Date", "Collector", "Client", "Product Description", "Product Type",
    "Batch No", "Account Type", "Relation", "PTP Amount", "Next Call", "PTP Date",
    "Claim Paid Amount", "Claim Paid Date", "Dialed Number", "Days Past Write Off",
    "Balance", "Contact Type", "Black Case No.", "Red Case No.", "Court Name",
    "Lawyer", "Legal Stage", "Legal Status", "Next Legal Follow up", "Call Duration",
    "Talk Time Duration", "Cycle", "Old IC", "I.C Issue Date", "Bank Code",
    "Over Limit Amount", "Min Payment", "Due Date", "Monthly Installment",
    "30 Days", "MIA", "Area", "Debtor ID", "Last Pay Date", "Last Pay Amount"
]

# ── Expected headers after restructuring (51 columns) ────────────────────────
EXPECTED_HEADERS = [
    "S.No", "Date", "Time", "Debtor", "Account No.", "Card No.", "Service No.",
    "DPD", "Call Status", "Status", "Remark", "Remark By", "Remark Type",
    "Field Visit Date", "Collector", "Client", "Product Description", "Product Type",
    "Batch No", "Account Type", "Relation", "PTP Amount", "PTP Date", "Next Call",
    "Claim Paid Amount", "Claim Paid Date", "Dialed Number", "Days Past Write Off",
    "Balance", "Contact Type", "Cycle", "Old IC", "I.C Issue Date", "Bank Code",
    "Over Limit Amount", "Min Payment", "Due Date", "Monthly Installment",
    "30 Days", "MIA", "Area", "Call Duration", "Talk Time Duration", "Debtor ID",
    "Black Case No.", "Red Case No.", "Court Name", "Lawyer", "Legal Stage",
    "Legal Status", "Next Legal Follow up"
]

REMARK_PREFIXES_TO_REMOVE = [
    "new assignment",
    "new contact",
    "system auto update"
]

STATUS_TO_REMOVE = [
    "", "abort", "locked", "new", "reactive",
    "sms failed", "sms replied", "sms reply", "sms sent", "unlocked"
]


# ── File Loader ───────────────────────────────────────────────────────────────
def load_file(uploaded_file):
    file_name = uploaded_file.name.lower()

    if file_name.endswith(".csv"):
        try:
            df = pd.read_csv(uploaded_file, encoding="utf-8")
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding="latin-1")

    elif file_name.endswith(".xlsx"):
        df = pd.read_excel(uploaded_file, engine="openpyxl")

    elif file_name.endswith(".xls"):
        df = pd.read_excel(uploaded_file, engine="xlrd")

    else:
        return None

    return df


# ── Validate Raw Headers ──────────────────────────────────────────────────────
def validate_raw_headers(df_columns):
    actual = list(df_columns)
    mismatches = []

    if len(actual) != len(RAW_HEADERS):
        return False, [
            f"Column count mismatch — Expected {len(RAW_HEADERS)} raw columns, got {len(actual)}"
        ]

    for i, (a, e) in enumerate(zip(actual, RAW_HEADERS)):
        if str(a).strip() != e:
            mismatches.append(f"Col {i+1}: expected **'{e}'** → got **'{a}'**")

    return len(mismatches) == 0, mismatches


# ── Validate Final Headers ────────────────────────────────────────────────────
def validate_final_headers(df_columns):
    actual = list(df_columns)
    mismatches = []

    if len(actual) != len(EXPECTED_HEADERS):
        return False, [
            f"Column count mismatch — Expected {len(EXPECTED_HEADERS)} columns after restructuring, got {len(actual)}"
        ]

    for i, (a, e) in enumerate(zip(actual, EXPECTED_HEADERS)):
        if str(a).strip() != e:
            mismatches.append(f"Col {i+1}: expected **'{e}'** → got **'{a}'**")

    return len(mismatches) == 0, mismatches


# ── Restructure Columns ───────────────────────────────────────────────────────
def restructure(df):
    """
    Transform raw file (53 cols) into expected structure (51 cols).

    Step 1 — Swap Next Call (index 22) ↔ PTP Date (index 23)
    Step 2 — Insert copy of Old IC (index 40) at index 30, beside Contact Type (index 29)
    Step 3 — After Step 2 shift, insert copy of Call Duration at index 51,
              between Area (index 50) and Debtor ID (index 51)
    Step 4 — Drop Last Pay Date and Last Pay Amount
    Step 5 — Drop original Old IC and original Call Duration
    Step 6 — Rename inserted copies back to original names
    """

    df = df.copy()

    # ── Step 1: Swap Next Call ↔ PTP Date ────────────────────────────────────
    cols = list(df.columns)
    cols[22], cols[23] = cols[23], cols[22]
    df = df[cols]

    # ── Step 2: Insert Old IC copy beside Contact Type (index 29) → index 30 ─
    # Old IC is at index 40 before this insert
    old_ic_col_name = df.columns[40]
    old_ic_data = df.iloc[:, 40].copy()
    df.insert(30, "__OLD_IC__", old_ic_data)
    # After insert at 30: everything from index 30 onward shifts +1
    # Old IC original → index 41
    # Call Duration (was 37) → index 38
    # Area (was 49) → index 50
    # Debtor ID (was 50) → index 51

    # ── Step 3: Insert Call Duration copy between Area and Debtor ID ─────────
    # Call Duration is now at index 38 (after Step 2 shift)
    # Area is at index 50, Debtor ID at index 51 → insert at 51
    call_dur_data = df.iloc[:, 38].copy()
    df.insert(51, "__CALL_DUR__", call_dur_data)
    # After insert at 51: everything from index 51 onward shifts +1
    # Old IC original → index 41 (unchanged)
    # Call Duration original → index 38 (unchanged)

    # ── Step 4: Drop Last Pay Date and Last Pay Amount ────────────────────────
    df = df.drop(columns=["Last Pay Date", "Last Pay Amount"])

    # ── Step 5: Drop original Old IC and original Call Duration ───────────────
    df = df.drop(columns=[old_ic_col_name])
    df = df.drop(columns=["Call Duration"])

    # ── Step 6: Rename inserted placeholders back to proper names ─────────────
    df = df.rename(columns={
        "__OLD_IC__": old_ic_col_name,
        "__CALL_DUR__": "Call Duration"
    })

    return df


# ── Remark Filter ─────────────────────────────────────────────────────────────
def remark_should_remove(val):
    if pd.isna(val):
        return False
    return any(
        str(val).strip().lower().startswith(prefix)
        for prefix in REMARK_PREFIXES_TO_REMOVE
    )


# ── Status Filter ─────────────────────────────────────────────────────────────
def status_should_remove(val):
    if pd.isna(val) or str(val).strip() == "":
        return True
    return str(val).strip().lower() in STATUS_TO_REMOVE


# ── Main App ──────────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "Upload Raw Export File (.csv, .xlsx, .xls)",
    type=["csv", "xlsx", "xls"]
)

if uploaded_file:
    try:
        # ── 1. Load File ──────────────────────────────────────────────────────
        df_raw = load_file(uploaded_file)

        if df_raw is None:
            st.error("❌ Unsupported file format. Please upload a .csv, .xlsx, or .xls file.")
            st.stop()

        st.subheader("Preview — Raw File")
        st.dataframe(df_raw.head(10), use_container_width=True)
        st.caption(f"Raw file: {len(df_raw):,} rows | {len(df_raw.columns)} columns")

        # ── 2. Validate Raw Headers ───────────────────────────────────────────
        st.subheader("Step 1 — Raw File Header Check")
        raw_valid, raw_mismatches = validate_raw_headers(df_raw.columns)

        if raw_valid:
            st.success(f"✅ Raw file headers validated — {len(RAW_HEADERS)} columns confirmed.")
        else:
            st.error("❌ Raw file headers do not match the expected raw structure:")
            for m in raw_mismatches:
                st.markdown(f"- {m}")
            st.warning("⚠️ Processing halted. Please check the uploaded file and try again.")
            st.stop()

        # ── 3. Restructure Columns ────────────────────────────────────────────
        st.subheader("Step 2 — Column Restructuring")
        with st.spinner("Restructuring columns..."):
            df_restructured = restructure(df_raw)
        st.success(
            "✅ Columns restructured — "
            "Next Call ↔ PTP Date swapped | "
            "Old IC copied to AE | "
            "Call Duration copied between Area and Debtor ID | "
            "Last Pay Date & Last Pay Amount removed."
        )

        # ── 4. Validate Final Headers ─────────────────────────────────────────
        st.subheader("Step 3 — Post-Restructure Header Validation")
        final_valid, final_mismatches = validate_final_headers(df_restructured.columns)

        if final_valid:
            st.success(f"✅ All {len(EXPECTED_HEADERS)} headers are in the correct position after restructuring.")
        else:
            st.error("❌ Header validation failed after restructuring:")
            for m in final_mismatches:
                st.markdown(f"- {m}")
            st.warning("⚠️ Processing halted. Please report this issue.")
            st.stop()

        # ── 5. Clean — Remark Filter ──────────────────────────────────────────
        st.subheader("Step 4 — Data Cleaning")
        with st.spinner("Cleaning data..."):
            before_remark = len(df_restructured)
            df_cleaned = df_restructured[
                ~df_restructured["Remark"].apply(remark_should_remove)
            ].copy()
            removed_remark = before_remark - len(df_cleaned)

            # ── 6. Clean — Status Filter ──────────────────────────────────────
            before_status = len(df_cleaned)
            df_cleaned = df_cleaned[
                ~df_cleaned["Status"].apply(status_should_remove)
            ].copy()
            removed_status = before_status - len(df_cleaned)

            df_cleaned = df_cleaned.reset_index(drop=True)

        st.success("✅ Data cleaning complete.")

        # ── 7. Summary ────────────────────────────────────────────────────────
        st.subheader("Summary")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Raw Rows", f"{len(df_raw):,}")
        c2.metric("Final Rows", f"{len(df_cleaned):,}")
        c3.metric("Removed by Remark Filter", f"{removed_remark:,}")
        c4.metric("Removed by Status Filter", f"{removed_status:,}")

        # ── 8. Preview ────────────────────────────────────────────────────────
        st.subheader("Preview — Processed File")
        st.dataframe(df_cleaned.head(20), use_container_width=True)

        # ── 9. Download ───────────────────────────────────────────────────────
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_cleaned.to_excel(writer, index=False, sheet_name="Processed")
        output.seek(0)

        st.download_button(
            label="⬇️ Download Processed File",
            data=output,
            file_name="XDAYS_SL_Export_Direct_Processed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
        st.exception(e)