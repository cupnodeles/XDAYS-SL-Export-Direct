import io
import math
import re
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st
from openpyxl import Workbook

st.set_page_config(page_title="XDAYS SL Export Direct", layout="wide")
st.title("XDAYS SL Export Direct")

# ── Canonical standardized DRR headers (51, exact order) ─────────────────────
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

# ── Date-typed standard headers (closed list per spec) ───────────────────────
DATE_COLUMNS = [
    "Date",
    "Field Visit Date",
    "PTP Date",
    "Next Call",
    "Claim Paid Date",
    "I.C Issue Date",
    "Due Date",
]
DATE_SET = set(DATE_COLUMNS)

# ── Amount/balance columns where comma-grouped text becomes numbers ──────────
AMOUNT_COLUMNS = [
    "PTP Amount",
    "Claim Paid Amount",
    "Balance",
    "Over Limit Amount",
    "Min Payment",
    "Monthly Installment",
]
AMOUNT_SET = set(AMOUNT_COLUMNS)

STATUS_TO_REMOVE = [
    "",
]

EXCEL_EPOCH = date(1899, 12, 30)


# ── Header auto-match ────────────────────────────────────────────────────────
def normalize_header(header):
    """Lowercase + drop separators so variants map to one standard header."""
    return re.sub(r"[.\s_\-/()#]+", "", str(header if header is not None else "").strip().lower())


def build_header_map(input_columns):
    """Map any input headers to the 51 canonical headers.

    Returns (standard_to_input, dropped, mapped):
    - standard_to_input: canonical header -> first matching input column (or missing).
    - dropped: input columns with no canonical match.
    - mapped: canonical headers matched, in canonical order.
    """
    by_normalized = {}
    for std in EXPECTED_HEADERS:
        by_normalized.setdefault(normalize_header(std), std)

    standard_to_input = {}
    dropped = []
    for col in input_columns:
        name = "" if col is None else str(col)
        if name.strip() == "":
            continue
        std = by_normalized.get(normalize_header(name))
        if std is None:
            if name not in dropped:
                dropped.append(name)
        elif std not in standard_to_input:
            standard_to_input[std] = name
        elif standard_to_input[std] != name and name not in dropped:
            dropped.append(name)

    mapped = [h for h in EXPECTED_HEADERS if h in standard_to_input]
    return standard_to_input, dropped, mapped


# ── Formula defuse ───────────────────────────────────────────────────────────
def defuse_cell_value(raw):
    """Turn formula-looking cells into static values. Never returns a formula.

    Returns (value, was_formula).
    """
    if raw is None:
        return None, False
    try:
        if pd.isna(raw):
            return None, False
    except (TypeError, ValueError):
        pass
    if isinstance(raw, bool):
        return raw, False
    if isinstance(raw, (datetime, date)):
        return raw, False
    if type(raw).__name__ == "Timestamp":
        try:
            return raw.date(), False
        except (ValueError, TypeError, AttributeError):
            return None, False
    if isinstance(raw, (int, float)):
        return raw, False

    text = str(raw).strip()
    if text == "":
        return "", False
    double_quoted = re.match(r'^="([\s\S]*)"$', text)
    if double_quoted:
        return double_quoted.group(1).replace('""', '"'), True
    single_quoted = re.match(r"^='([\s\S]*)'$", text)
    if single_quoted:
        return single_quoted.group(1).replace('""', '"'), True
    if text.startswith("=") and len(text) > 1:
        inner = text[1:].strip()
        if (
            len(inner) >= 2
            and ((inner.startswith('"') and inner.endswith('"'))
                 or (inner.startswith("'") and inner.endswith("'")))
        ):
            inner = inner[1:-1].replace('""', '"')
        return inner, True
    return text, False


# ── Account No. cleaning ─────────────────────────────────────────────────────
def clean_account_no(raw):
    """Digits-only exact-8 rule: 12345678 -> 00000012345678 (text).

    All other lengths pass through untouched (never double-pads).
    Returns (text, padded).
    """
    value, _ = defuse_cell_value(raw)
    if value is None:
        return "", False
    text = str(value).strip()
    if text == "":
        return "", False
    text = re.sub(r"[\s-]+", "", text)
    dot_zero = re.match(r"^(\d+)\.0+$", text)
    if dot_zero:
        text = dot_zero.group(1)
    if re.fullmatch(r"\d{8}", text):
        return "000000" + text, True
    return text, False


# ── Comma amounts ────────────────────────────────────────────────────────────
def parse_comma_amount(text):
    """Parse proper thousands-grouped text ("21,482.51", "-1,000") to float.

    Returns None when the text is not grouped-amount shaped, so IDs and
    free text are never coerced.
    """
    t = str(text).strip()
    if not re.fullmatch(r"-?\d{1,3}(,\d{3})+(\.\d+)?", t):
        return None
    try:
        number = float(t.replace(",", ""))
    except ValueError:
        return None
    return number if math.isfinite(number) else None


# ── Explicit-order date parsing ──────────────────────────────────────────────
def _to_date_or_none(year, month, day):
    try:
        if not (1900 <= year <= 2100 and 1 <= month <= 12 and 1 <= day <= 31):
            return None
        return date(year, month, day)
    except ValueError:
        return None


def parse_date_with_order(raw, order="dmy"):
    """Parse a date explicitly per order — never month-first by accident.

    `10/9/2026` with order="dmy" -> Sep 10; with "mdy" -> Oct 9.
    Excel serials and real dates are order-independent. `0000-00-00…`
    maps to blank. Time suffixes are dropped (date-only output).

    Returns (date|None, ambiguous: bool, blank: bool).
    """
    if raw is None:
        return None, False, True
    try:
        if pd.isna(raw):
            return None, False, True
    except (TypeError, ValueError):
        pass
    if isinstance(raw, bool):
        return None, False, False
    if isinstance(raw, datetime):
        try:
            return raw.date(), False, False
        except (ValueError, TypeError, AttributeError):
            return None, False, False
    if isinstance(raw, date):
        return raw, False, False
    if type(raw).__name__ == "Timestamp":
        try:
            return raw.date(), False, False
        except (ValueError, TypeError, AttributeError):
            return None, False, False
    if isinstance(raw, (int, float)):
        if isinstance(raw, float):
            if math.isnan(raw) or math.isinf(raw):
                return None, False, True
            if not raw.is_integer():
                return None, False, False
            serial = int(raw)
        else:
            serial = raw
        if serial < 1 or serial > 2958465:
            return None, False, False
        return EXCEL_EPOCH + timedelta(days=serial), False, False

    text = str(raw).strip()
    if text == "" or text.lower() in ("nan", "nat", "none"):
        return None, False, True
    if re.match(r"^0{4}-0{2}-0{2}", text):
        return None, False, True

    iso = re.match(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})(?:[T ].*)?$", text)
    if iso:
        parsed = _to_date_or_none(int(iso.group(1)), int(iso.group(2)), int(iso.group(3)))
        return parsed, False, False

    mdy = re.match(
        r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})"
        r"(?:[ T]\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AP]M)?)?$",
        text,
        re.IGNORECASE,
    )
    if not mdy:
        return None, False, False
    first, second = int(mdy.group(1)), int(mdy.group(2))
    year = int(mdy.group(3))
    if len(mdy.group(3)) == 2:
        year += 2000
    ambiguous = 1 <= first <= 12 and 1 <= second <= 12
    month, day = (second, first) if order == "dmy" else (first, second)
    return _to_date_or_none(year, month, day), ambiguous, False


def format_mmddyyyy(value):
    if isinstance(value, datetime):
        value = value.date()
    if not isinstance(value, date):
        return ""
    return f"{value.month:02d}/{value.day:02d}/{value.year:04d}"


# ── File loader (raw text preserved — no month-first coercion) ───────────────
def load_file(uploaded_file):
    file_name = uploaded_file.name.lower()

    if file_name.endswith(".csv"):
        try:
            df = pd.read_csv(
                uploaded_file, dtype=str, keep_default_na=False, na_filter=False,
                encoding="utf-8",
            )
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            df = pd.read_csv(
                uploaded_file, dtype=str, keep_default_na=False, na_filter=False,
                encoding="latin-1",
            )
    elif file_name.endswith(".xlsx"):
        df = pd.read_excel(uploaded_file, engine="openpyxl", dtype=str, na_filter=False)
    elif file_name.endswith(".xls"):
        df = pd.read_excel(uploaded_file, engine="xlrd", dtype=str, na_filter=False)
    else:
        return None

    # Drop blank-named columns; keep the first of any duplicated headers.
    df = df.loc[:, [str(c).strip() != "" for c in df.columns]]
    df = df.loc[:, ~pd.Index([str(c) for c in df.columns]).duplicated(keep="first")]
    return df


# ── Status filter (kept per spec) ────────────────────────────────────────────
def status_should_remove(val):
    if val is None:
        return True
    try:
        if pd.isna(val):
            return True
    except (TypeError, ValueError):
        pass
    return str(val).strip().lower() in STATUS_TO_REMOVE


# ── Core conversion ──────────────────────────────────────────────────────────
def convert_export_direct(df_raw, date_order="dmy"):
    """Standardize any-header input into the 51-header DRR + Status clean.

    Returns (df_out, report) where report holds mapped/dropped lists,
    padded/ambiguous counts and samples.
    """
    standard_to_input, dropped, mapped = build_header_map(list(df_raw.columns))

    out_rows = []
    padded_count = 0
    padded_samples = []
    ambiguous_count = 0
    ambiguous_samples = []

    for i, (_, row) in enumerate(df_raw.iterrows()):
        row_num = i + 2
        out = {}
        for std in EXPECTED_HEADERS:
            input_col = standard_to_input.get(std)
            raw = row[input_col] if input_col is not None else None

            if std == "Account No.":
                before = "" if raw is None else str(raw).strip()
                cleaned, was_padded = clean_account_no(raw)
                out[std] = cleaned
                if was_padded:
                    padded_count += 1
                    if len(padded_samples) < 5:
                        padded_samples.append(
                            {"row": row_num, "before": before, "after": cleaned}
                        )
                continue

            if std in DATE_SET:
                value, _ = defuse_cell_value(raw)
                parsed, ambiguous, _ = parse_date_with_order(value, date_order)
                out[std] = parsed
                if ambiguous and parsed is not None:
                    ambiguous_count += 1
                    if len(ambiguous_samples) < 20:
                        ambiguous_samples.append({
                            "row": row_num,
                            "column": std,
                            "raw": "" if value is None else str(value).strip(),
                            "interpreted": format_mmddyyyy(parsed),
                        })
                continue

            value, _ = defuse_cell_value(raw)
            if value is None:
                out[std] = ""
            elif isinstance(value, bool):
                out[std] = str(value)
            elif isinstance(value, (int, float)):
                out[std] = value
            elif isinstance(value, (datetime, date)):
                out[std] = value
            else:
                text = str(value)
                if std in AMOUNT_SET:
                    number = parse_comma_amount(text)
                    if number is not None:
                        out[std] = number
                    elif re.fullmatch(r"-?\d+(\.\d+)?", text.strip()):
                        out[std] = float(text.strip())
                    else:
                        out[std] = text
                else:
                    out[std] = text
        out_rows.append(out)

    df_out = pd.DataFrame(out_rows, columns=EXPECTED_HEADERS)

    removed_status = 0
    if "Status" in df_out.columns and len(df_out) > 0:
        before_status = len(df_out)
        df_out = df_out[~df_out["Status"].apply(status_should_remove)].copy()
        removed_status = before_status - len(df_out)

    df_out = df_out.reset_index(drop=True)

    report = {
        "rows_in": len(df_raw),
        "rows_out": len(df_out),
        "mapped": mapped,
        "dropped": dropped,
        "padded_count": padded_count,
        "padded_samples": padded_samples,
        "ambiguous_count": ambiguous_count,
        "ambiguous_samples": ambiguous_samples,
    }
    return df_out, report


def preview_dates(df_raw, standard_to_input, date_order="dmy", max_per_column=3, include_invalid=True):
    """Sample raw -> interpreted date previews for the confirm screen.

    Realtime: pure function of `date_order`, so the Step-2 radio refreshes
    it instantly. Invalid/unparseable dates are included as
    "— (blank)" with status instead of being skipped, so picking the
    wrong order (e.g. mdy on 15/09/2026) is visible rather than empty.
    """
    samples = []
    seen = {}
    for _, row in df_raw.iterrows():
        for std in DATE_COLUMNS:
            input_col = standard_to_input.get(std)
            if input_col is None or seen.get(std, 0) >= max_per_column:
                continue
            value, _ = defuse_cell_value(row[input_col])
            if value is None or str(value).strip() == "":
                continue
            if re.match(r"^0{4}-0{2}-0{2}", str(value).strip()):
                continue
            if re.match(r"^0+$", str(value).strip()):
                continue  # placeholder zero, not a real date attempt
            parsed, ambiguous, _ = parse_date_with_order(value, date_order)
            if parsed is None:
                if not include_invalid:
                    continue
                samples.append({
                    "column": std,
                    "raw": str(value).strip(),
                    "interpreted": "— (blank)",
                    "ambiguous": False,
                    "status": "blank/invalid",
                })
            else:
                samples.append({
                    "column": std,
                    "raw": str(value).strip(),
                    "interpreted": format_mmddyyyy(parsed),
                    "ambiguous": ambiguous,
                    "status": "ok",
                })
            seen[std] = seen.get(std, 0) + 1
        if all(
            standard_to_input.get(std) is None or seen.get(std, 0) >= max_per_column
            for std in DATE_COLUMNS
        ):
            break
    return samples


def find_ambiguous_example(df_raw, standard_to_input, fallback="10/9/2026"):
    """First ambiguous raw date in the file for the live Step-2 labels.

    Scans date columns in file order; returns the raw string of the first
    value ambiguous under both orders (e.g. "11/9/2026"). Falls back to
    `fallback` when the file has no ambiguous date.
    """
    for _, row in df_raw.iterrows():
        for std in DATE_COLUMNS:
            input_col = standard_to_input.get(std)
            if input_col is None:
                continue
            value, _ = defuse_cell_value(row[input_col])
            if value is None or str(value).strip() == "":
                continue
            text = str(value).strip()
            if re.match(r"^0{4}-0{2}-0{2}", text):
                continue
            if re.match(r"^0+$", text):
                continue
            _, ambiguous, _ = parse_date_with_order(value, "dmy")
            if ambiguous:
                return text
    return fallback


def short_date_label(raw):
    """Strip year for radio labels: "11/9/2026" -> "11/9"."""
    m = re.match(r"^\s*(\d{1,2})[/\-.](\d{1,2})[/\-.]\d{2,4}", str(raw))
    if m:
        return f"{int(m.group(1))}/{int(m.group(2))}"
    return str(raw).strip()


def output_name(uploaded_name):
    """Keep the uploaded file's name; only the extension becomes .xlsx."""
    clean = (uploaded_name or "").strip() or "XDAYS_SL_Export_Direct"
    lowered = clean.lower()
    for ext in (".xlsx", ".xls", ".csv"):
        if lowered.endswith(ext):
            clean = clean[: -len(ext)]
            break
    return clean + ".xlsx"


def write_output_workbook(df):
    """Build the .xlsx bytes: dates as mm/dd/yyyy, accounts as text.

    Strings are written as plain text (never formulas); date cells carry
    the mm/dd/yyyy number format; amounts stay numeric.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Processed"
    ws.append(list(EXPECTED_HEADERS))

    for _, row in df.iterrows():
        values = []
        for header in EXPECTED_HEADERS:
            value = row[header]
            if header in DATE_SET:
                values.append(value if isinstance(value, (datetime, date)) else "")
            else:
                values.append(value)
        ws.append(values)

    date_positions = {EXPECTED_HEADERS.index(h) + 1 for h in DATE_COLUMNS}
    for excel_row in ws.iter_rows(min_row=2):
        for cell in excel_row:
            if cell.column in date_positions:
                if isinstance(cell.value, (datetime, date)):
                    cell.number_format = "mm/dd/yyyy"
                else:
                    cell.value = ""
            elif isinstance(cell.value, str) and cell.value.startswith("="):
                # Plain-text guarantee: openpyxl stores str as text, but
                # pin it explicitly so it can never become a live formula.
                cell.value = str(cell.value)
                cell.data_type = "s"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ── Main App ─────────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "Upload Raw Export File (.csv, .xlsx, .xls — any headers)",
    type=["csv", "xlsx", "xls"],
)

if uploaded_file:
    try:
        # ── 1. Load file (raw text preserved) ────────────────────────────────
        df_raw = load_file(uploaded_file)

        if df_raw is None:
            st.error(
                "❌ Unsupported file format. "
                "Please upload a .csv, .xlsx, or .xls file."
            )
            st.stop()

        st.subheader("Preview — Raw File")
        st.dataframe(df_raw.head(10), use_container_width=True)
        st.caption(f"Raw file: {len(df_raw):,} rows | {len(df_raw.columns)} columns")

        # ── 2. Header coverage (auto-match, extras listed) ───────────────────
        standard_to_input, dropped, mapped = build_header_map(list(df_raw.columns))

        st.subheader("Step 1 — Header Coverage")
        c1, c2 = st.columns(2)
        c1.metric("Matched", f"{len(mapped)} / 51")
        c2.metric("Dropped (extras)", len(dropped))
        if dropped:
            st.caption("Dropped: " + ", ".join(f"`{d}`" for d in dropped))

        # ── 3. Date order confirm (labels live from the file) ──────────────────
        st.subheader("Step 2 — Date Order (single confirm)")
        example_raw = find_ambiguous_example(df_raw, standard_to_input)
        ex_dmy, _, _ = parse_date_with_order(example_raw, "dmy")
        ex_mdy, _, _ = parse_date_with_order(example_raw, "mdy")
        ex_short = short_date_label(example_raw)
        date_order = st.radio(
            f"How should ambiguous dates like {example_raw} be read?",
            options=["dmy", "mdy"],
            format_func=lambda o, _s=ex_short, _d=format_mmddyyyy(ex_dmy), _m=format_mmddyyyy(ex_mdy): (
                f"DD/MM — {_s} → {_d}" if o == "dmy"
                else f"MM/DD — {_s} → {_m}"
            ),
            index=0,
        )

        date_samples = preview_dates(df_raw, standard_to_input, date_order)
        if date_samples:
            st.caption("Date preview (raw → interpreted)")
            st.dataframe(
                pd.DataFrame(date_samples)[["column", "raw", "interpreted", "ambiguous", "status"]],
                use_container_width=True,
            )

        # ── 4. Convert ───────────────────────────────────────────────────────
        st.subheader("Step 3 — Standardize")
        with st.spinner("Converting..."):
            df_out, report = convert_export_direct(df_raw, date_order)

        st.success("✅ Standardization complete.")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Final Rows", f"{report['rows_out']:,}")
        m3.metric("Account No. padded", f"{report['padded_count']:,}")
        m4.metric("Ambiguous dates", f"{report['ambiguous_count']:,}")
        if report["padded_samples"]:
            first = report["padded_samples"][0]
            st.caption(
                f"Account padding sample: `{first['before']}` → `{first['after']}`"
            )

        # ── 5. Preview ───────────────────────────────────────────────────────
        st.subheader("Preview — Standardized File")
        st.dataframe(df_out.head(20), use_container_width=True)

        # ── 6. Download (keeps uploaded name) ────────────────────────────────
        output = write_output_workbook(df_out)
        st.download_button(
            label="⬇️ Download Standardized File",
            data=output,
            file_name=output_name(uploaded_file.name),
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            ),
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
        st.exception(e)
