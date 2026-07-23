import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="XDAYS SL Export Direct", layout="wide")
st.title("XDAYS SL Export Direct")

uploaded_file = st.file_uploader("Upload Raw Export File", type=["xlsx", "xls", "csv"])

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


def validate_headers(df_columns, expected):
    """Compare restructured columns vs expected. Returns (is_valid, mismatches)."""
    actual = list(df_columns)
    mismatches = []

    if len(actual) != len(expected):
        return False, [
            f"Column count mismatch — Expected {len(expected)}, got {len(actual)}"
        ]

    for i, (a, e) in enumerate(zip(actual, expected)):
        if str(a).strip() != e:
            mismatches.append(
                f"Col {i+1} ('{chr(64 + i+1) if i < 26 else 'A' + chr(64 + i-25)}'): "
                f"expected **'{e}'** → got **'{a}'**"
            )

    return len(mismatches) == 0, mismatches


if uploaded_file:
    try:
        # ── 1. Load file ──────────────────────────────────────────────────────
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        st.subheader("Preview — Raw File")
        st.dataframe(df.head(10), use_container_width=True)
        st.caption(f"Raw file: {len(df)} rows | {len(df.columns)} columns")

        # ── 2. Restructure Columns ────────────────────────────────────────────
        with st.spinner("Restructuring columns..."):

            # Step 1: Swap Next Call (W = index 22) and PTP Date (X = index 23)
            # Raw file: W = Next Call, X = PTP Date
            # After swap: W = PTP Date, X = Next Call
            cols = list(df.columns)
            cols[22], cols[23] = cols[23], cols[22]
            df = df[cols]

            # Step 2: Copy Old IC (AO = index 40) and insert at index 30
            # This places Old IC at AE, right beside Contact Type (AD = index 29)
            old_ic_data = df.iloc[:, 40].copy()
            df.insert(30, old_ic_data.name + "_COPY", old_ic_data)

            # Step 3: Copy Call Duration — after Step 2 insert, original AW
            # shifted from index 48 to index 49. Insert between Area and
            # Black Case No. — Area is now at index 41, so insert at index 42.
            call_duration_data = df.iloc[:, 49].copy()
            df.insert(42, call_duration_data.name + "_COPY", call_duration_data)

        # ── 3. Validate Headers After Restructuring ───────────────────────────
        st.subheader("Header Validation")
        st.caption(
            "Validation is performed **after** column swapping and insertions, "
            "not on the raw file."
        )

        is_valid, mismatches = validate_headers(df.columns, EXPECTED_HEADERS)

        if is_valid:
            st.success("✅ All 51 headers are in the correct position and naming after restructuring.")
        else:
            st.error("❌ Header validation failed after restructuring. See mismatches below:")
            for m in mismatches:
                st.markdown(f"- {m}")
            st.warning(
                "⚠️ Processing has been halted. "
                "The raw file's column structure may not match what is expected. "
                "Please verify the raw file before uploading again."
            )
            st.stop()

        # ── 4. Filter Remark (Column K) — remove by first-character prefix ────
        with st.spinner("Cleaning data..."):

            def remark_should_remove(val):
                if pd.isna(val):
                    return False
                val_lower = str(val).strip().lower()
                return any(
                    val_lower.startswith(prefix)
                    for prefix in REMARK_PREFIXES_TO_REMOVE
                )

            before_remark = len(df)
            df = df[~df.iloc[:, 10].apply(remark_should_remove)]
            removed_remark = before_remark - len(df)

            # ── 5. Filter Status (Column J) — remove blanks + listed values ───
            def status_should_remove(val):
                if pd.isna(val) or str(val).strip() == "":
                    return True
                return str(val).strip().lower() in STATUS_TO_REMOVE

            before_status = len(df)
            df = df[~df.iloc[:, 9].apply(status_should_remove)]
            removed_status = before_status - len(df)

            df = df.reset_index(drop=True)

        # ── 6. Summary ────────────────────────────────────────────────────────
        st.success("✅ Processing complete!")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Final Rows", len(df))
        col2.metric("Final Columns", len(df.columns))
        col3.metric("Removed by Remark Filter", removed_remark)
        col4.metric("Removed by Status Filter", removed_status)

        st.subheader("Preview — Processed File")
        st.dataframe(df.head(20), use_container_width=True)

        # ── 7. Download ───────────────────────────────────────────────────────
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Processed")
        output.seek(0)

        st.download_button(
            label="⬇️ Download Processed File",
            data=output,
            file_name="XDAYS_SL_Export_Direct_Processed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")