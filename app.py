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
    actual = list(df_columns)
    mismatches = []

    if len(actual) != len(expected):
        return False, [
            f"Column count mismatch — Expected {len(expected)}, got {len(actual)}"
        ]

    for i, (a, e) in enumerate(zip(actual, expected)):
        if str(a).strip() != e:
            mismatches.append(f"Col {i+1}: expected **'{e}'** → got **'{a}'**")

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
        st.caption(f"Rows: {len(df)} | Columns: {len(df.columns)}")

        with st.spinner("Restructuring columns..."):

            # ── 2. Swap Next Call (W=col 22) and PTP Date (X=col 23) ─────────
            cols = list(df.columns)
            cols[22], cols[23] = cols[23], cols[22]
            df = df[cols]

            # ── 3. Copy Old IC (AO=col 40) and insert beside Contact Type (AD=col 29) → AE=col 30 ──
            old_ic_data = df.iloc[:, 40].copy()
            df.insert(30, "__OLD_IC_COPY__", old_ic_data)

            # ── 4. Copy Call Duration (now shifted) and insert between Area and Black Case No. ──
            # After insert at 30, original AW (col 48) is now col 49
            call_duration_data = df.iloc[:, 49].copy()
            # Area is now at col 41, Black Case No. at col 42 → insert at 42
            df.insert(42, "__CALL_DUR_COPY__", call_duration_data)

        # ── 5. Validate Headers ───────────────────────────────────────────────
        st.subheader("Header Validation")
        is_valid, mismatches = validate_headers(df.columns, EXPECTED_HEADERS)

        if is_valid:
            st.success("✅ All headers are in the correct order and naming.")
        else:
            st.error("❌ Header validation failed. Please review the mismatches below:")
            for m in mismatches:
                st.markdown(f"- {m}")
            st.warning(
                "Processing has been halted. "
                "Please ensure the raw file matches the expected structure before proceeding."
            )
            st.stop()

        # ── 6. Filter Remark (Column K) — remove matching prefixes ───────────
        with st.spinner("Cleaning data..."):

            def remark_should_remove(val):
                if pd.isna(val):
                    return False
                val_lower = str(val).strip().lower()
                return any(val_lower.startswith(prefix) for prefix in REMARK_PREFIXES_TO_REMOVE)

            before_remark = len(df)
            df = df[~df["Remark"].apply(remark_should_remove)]
            removed_remark = before_remark - len(df)

            # ── 7. Filter Status (Column J) — remove blanks + listed values ──
            def status_should_remove(val):
                if pd.isna(val) or str(val).strip() == "":
                    return True
                return str(val).strip().lower() in STATUS_TO_REMOVE

            before_status = len(df)
            df = df[~df["Status"].apply(status_should_remove)]
            removed_status = before_status - len(df)

            df = df.reset_index(drop=True)

        # ── 8. Summary ────────────────────────────────────────────────────────
        st.success("✅ Processing complete!")
        col1, col2, col3 = st.columns(3)
        col1.metric("Final Rows", len(df))
        col2.metric("Removed (Remark filter)", removed_remark)
        col3.metric("Removed (Status filter)", removed_status)

        st.subheader("Preview — Processed File")
        st.dataframe(df.head(20), use_container_width=True)

        # ── 9. Download ───────────────────────────────────────────────────────
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