import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="XDAYS SL Export Direct", layout="wide")
st.title("XDAYS SL Export Direct")

# ── Raw file headers (53 columns) ─────────────────────────────────────────────
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

# ── Expected headers after restructuring (51 columns) ─────────────────────────
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


# ── File Loader ────────────────────────────────────────────────────────────────
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


# ── Validate Raw Headers ───────────────────────────────────────────────────────
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


# ── Validate Final Headers ─────────────────────────────────────────────────────
def validate_final_headers(df_columns):
    actual = list(df_columns)
    mismatches = []

    if len(actual) != len(EXPECTED_HEADERS):
        return False, [
            f"Column count mismatch — Expected {len(EXPECTED_HEADERS)} columns "
            f"after restructuring, got {len(actual)}"
        ]

    for i, (a, e) in enumerate(zip(actual, EXPECTED_HEADERS)):
        if str(a).strip() != e:
            mismatches.append(f"Col {i+1}: expected **'{e}'** → got **'{a}'**")

    return len(mismatches) == 0, mismatches


# ── Restructure Columns ────────────────────────────────────────────────────────
def restructure(df):
    """
    Transform raw file (53 cols) into expected structure (51 cols)
    by explicitly reordering columns by name.

    Changes from raw → expected:
    1. Next Call (W) and PTP Date (X) are swapped
    2. Cycle moves from after Talk Time Duration to right after Contact Type
    3. Old IC moves from AO to right after Cycle
    4. Call Duration is copied to sit between Area and Talk Time Duration
    5. Black Case No. and the legal columns move to the end (after Debtor ID)
    6. Last Pay Date and Last Pay Amount are dropped
    """

    df = df.copy()

    # Step 1: Drop columns not needed in output
    df = df.drop(columns=["Last Pay Date", "Last Pay Amount"])

    # Step 2: Reorder all columns explicitly by name to match EXPECTED_HEADERS
    # This handles all swaps, moves, and copies in one clean operation
    df = df[[
        "S.No",
        "Date",
        "Time",
        "Debtor",
        "Account No.",
        "Card No.",
        "Service No.",
        "DPD",
        "Call Status",
        "Status",
        "Remark",
        "Remark By",
        "Remark Type",
        "Field Visit Date",
        "Collector",
        "Client",
        "Product Description",
        "Product Type",
        "Batch No",
        "Account Type",
        "Relation",
        "PTP Amount",
        "PTP Date",           # was at X (index 23) — swapped with Next Call
        "Next Call",          # was at W (index 22) — swapped with PTP Date
        "Claim Paid Amount",
        "Claim Paid Date",
        "Dialed Number",
        "Days Past Write Off",
        "Balance",
        "Contact Type",
        "Cycle",              # moved from index 39 to here (index 30)
        "Old IC",             # moved from index 40 to here (index 31)
        "I.C Issue Date",
        "Bank Code",
        "Over Limit Amount",
        "Min Payment",
        "Due Date",
        "Monthly Installment",
        "30 Days",
        "MIA",
        "Area",
        "Call Duration",      # copied from index 37, placed after Area
        "Talk Time Duration",
        "Debtor ID",
        "Black Case No.",     # moved from index 30 to here (index 44)
        "Red Case No.",
        "Court Name",
        "Lawyer",
        "Legal Stage",
        "Legal Status",
        "Next Legal Follow up"
    ]]

    return df


# ── Remark Filter ──────────────────────────────────────────────────────────────
def remark_should_remove(val):
    if pd.isna(val):
        return False
    return any(
        str(val).strip().lower().startswith(prefix)
        for prefix in REMARK_PREFIXES_TO_REMOVE
    )


# ── Status Filter ──────────────────────────────────────────────────────────────
def status_should_remove(val):
    if pd.isna(val) or str(val).strip() == "":
        return True
    return str(val).strip().lower() in STATUS_TO_REMOVE


# ── Main App ───────────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "Upload Raw Export File (.csv, .xlsx, .xls)",
    type=["csv", "xlsx", "xls"]
)

if uploaded_file:
    try:
        # ── 1. Load File ───────────────────────────────────────────────────────
        df_raw = load_file(uploaded_file)

        if df_raw is None:
            st.error(
                "❌ Unsupported file format. "
                "Please upload a .csv, .xlsx, or .xls file."
            )
            st.stop()

        st.subheader("Preview — Raw File")
        st.dataframe(df_raw.head(10), use_container_width=True)
        st.caption(
            f"Raw file: {len(df_raw):,} rows | {len(df_raw.columns)} columns"
        )

        # ── 2. Validate Raw Headers ────────────────────────────────────────────
        st.subheader("Step 1 — Raw File Header Check")
        raw_valid, raw_mismatches = validate_raw_headers(df_raw.columns)

        if raw_valid:
            st.success(
                f"✅ Raw file headers validated — "
                f"{len(RAW_HEADERS)} columns confirmed."
            )
        else:
            st.error(
                "❌ Raw file headers do not match the expected raw structure:"
            )
            for m in raw_mismatches:
                st.markdown(f"- {m}")
            st.warning(
                "⚠️ Processing halted. "
                "Please check the uploaded file and try again."
            )
            st.stop()

        # ── 3. Restructure Columns ─────────────────────────────────────────────
        st.subheader("Step 2 — Column Restructuring")
        with st.spinner("Restructuring columns..."):
            df_restructured = restructure(df_raw)
        st.success(
            "✅ Columns restructured — "
            "PTP Date & Next Call swapped | "
            "Cycle & Old IC moved after Contact Type | "
            "Call Duration placed after Area | "
            "Black Case No. & Legal columns moved to end | "
            "Last Pay Date & Last Pay Amount removed."
        )

        # ── 4. Validate Final Headers ──────────────────────────────────────────
        st.subheader("Step 3 — Post-Restructure Header Validation")
        final_valid, final_mismatches = validate_final_headers(
            df_restructured.columns
        )

        if final_valid:
            st.success(
                f"✅ All {len(EXPECTED_HEADERS)} headers are in the correct "
                f"position after restructuring."
            )
        else:
            st.error(
                "❌ Header validation failed after restructuring:"
            )
            for m in final_mismatches:
                st.markdown(f"- {m}")
            st.warning("⚠️ Processing halted. Please report this issue.")
            st.stop()

        # ── 5. Clean — Remark Filter ───────────────────────────────────────────
        st.subheader("Step 4 — Data Cleaning")
        with st.spinner("Cleaning data..."):

            before_remark = len(df_restructured)
            df_cleaned = df_restructured[
                ~df_restructured["Remark"].apply(remark_should_remove)
            ].copy()
            removed_remark = before_remark - len(df_cleaned)

            # ── 6. Clean — Status Filter ───────────────────────────────────────
            before_status = len(df_cleaned)
            df_cleaned = df_cleaned[
                ~df_cleaned["Status"].apply(status_should_remove)
            ].copy()
            removed_status = before_status - len(df_cleaned)

            df_cleaned = df_cleaned.reset_index(drop=True)

        st.success("✅ Data cleaning complete.")

        # ── 7. Summary ─────────────────────────────────────────────────────────
        st.subheader("Summary")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Raw Rows", f"{len(df_raw):,}")
        c2.metric("Final Rows", f"{len(df_cleaned):,}")
        c3.metric("Removed by Remark Filter", f"{removed_remark:,}")
        c4.metric("Removed by Status Filter", f"{removed_status:,}")

        # ── 8. Preview ─────────────────────────────────────────────────────────
        st.subheader("Preview — Processed File")
        st.dataframe(df_cleaned.head(20), use_container_width=True)

        # ── 9. Download ────────────────────────────────────────────────────────
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_cleaned.to_excel(writer, index=False, sheet_name="Processed")
        output.seek(0)

        st.download_button(
            label="⬇️ Download Processed File",
            data=output,
            file_name="XDAYS_SL_Export_Direct_Processed.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            )
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
        st.exception(e)