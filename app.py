import streamlit as st
import io
import os
from datetime import datetime
from typing import Dict, List, Optional, Tuple
import pandas as pd
import threading

ALLOWED_EXT = {".csv", ".xlsx", ".xls"}

MASTER_FILE = "master_summary.xlsx"
master_lock = threading.Lock()

# -------------------- Helpers --------------------
CANONICAL_US = ["User Alias", "User ID", "Max Loss", "Telegram"]

def _norm(s: str) -> str:
    if s is None: return ""
    s = str(s)
    return "".join(ch for ch in s.strip().lower() if ch.isalnum())

SYNONYMS_US: Dict[str, str] = {
    _norm("User Alias"): "User Alias",
    _norm("User ID"): "User ID",
    _norm("Max Loss"): "Max Loss",
    _norm("Telegram ID(s)"): "Telegram",
}

def _ext_ok(filename: str) -> bool:
    return os.path.splitext(filename or "")[1].lower() in ALLOWED_EXT

DESIRED_ORDER = [
    "SNO","Enabled","UserID","Alias","LoggedIn","SqOff Done","Broker","Qty Multiplier",
    "MTM (All)","ALLOCATION","MAX_LOSS","Available Margin","Total Orders","Total Lots",
    "SERVER","ALGO","REMARK","OPERATOR","EXPIRY"
]

def _reorder_summary_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = list(df.columns)
    desired = [c for c in DESIRED_ORDER if c in cols]
    rest = [c for c in cols if c not in desired]
    return df[desired + rest]

def _read_raw(file_bytes, filename) -> pd.DataFrame:
    name = filename.lower()
    if name.endswith(".csv"):
        return pd.read_csv(
            io.BytesIO(file_bytes),
            header=6,
            dtype=str,
            keep_default_na=False,
            low_memory=False
        )
    return pd.read_excel(io.BytesIO(file_bytes), header=6, dtype=str)

def _select_usersetting_columns(df: pd.DataFrame) -> pd.DataFrame:
    norm_to_orig = {_norm(c): c for c in df.columns}
    resolved = {}
    for canonical in CANONICAL_US:
        candidates = [k for k, v in SYNONYMS_US.items() if v == canonical]
        found = None
        for cand in candidates:
            if cand in norm_to_orig:
                found = norm_to_orig[cand]
                break
        if not found:
            raise ValueError(f"Usersetting missing column: {canonical}")
        resolved[canonical] = found
    cleaned = df[[resolved[c] for c in CANONICAL_US]].copy()
    cleaned.columns = CANONICAL_US
    return cleaned

def _build_lookup(clean_us: pd.DataFrame) -> Dict[str, Tuple[str, str]]:
    return {
        _norm(row["User ID"]): (row["Telegram"], row["Max Loss"])
        for _, row in clean_us.iterrows() if row.get("User ID")
    }

def _insert_allocation_maxloss(df: pd.DataFrame, lookup: Dict[str, Tuple[str, str]], user_id_colname: str) -> pd.DataFrame:
    out = df.copy()
    insert_at = min(9, len(out.columns))
    def fetch(uid):
        tel, mls = lookup.get(_norm(uid), ("", ""))
        return pd.Series({"ALLOCATION": tel, "MAX_LOSS": mls})
    new_cols = out[user_id_colname].apply(fetch)
    out.insert(insert_at, "ALLOCATION", new_cols["ALLOCATION"])
    out.insert(insert_at + 1, "MAX_LOSS", new_cols["MAX_LOSS"])
    return out

def _append_constants(df: pd.DataFrame, consts: Dict[str, str]) -> pd.DataFrame:
    out = df.copy()
    for k in ["SERVER","ALGO","OPERATOR","EXPIRY"]:
        out[k] = consts.get(k, "")
    if "REMARK" not in out.columns:
        out["REMARK"] = ""
    out["REMARK"] = out["REMARK"].astype(str).fillna('') + consts.get("REMARK", "")
    return out

def _read_all_sheets(file_bytes, filename) -> Dict[str, pd.DataFrame]:
    name = filename.lower()
    if name.endswith(".csv"):
        return {"Sheet1": pd.read_csv(io.BytesIO(file_bytes), low_memory=False)}
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    return {sheet: xl.parse(sheet_name=sheet) for sheet in xl.sheet_names}

def _server_from_filename(name: str) -> str:
    base = os.path.splitext(name or "")[0].strip()
    token = base.replace("_"," ").replace("-"," ").split()
    return token[0] if token else ""

def _coerce_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols_to_convert = [
        "ALLOCATION", "MAX_LOSS", "ALGO", "Total Orders", "Total Lots",
        "Available Margin", "MTM (All)", "Qty Multiplier"
    ]
    out_df = df.copy()
    for col in cols_to_convert:
        if col in out_df.columns:
            out_df[col] = pd.to_numeric(out_df[col], errors='coerce')
    return out_df

def _read_saved_mtm(file_bytes, filename) -> pd.DataFrame:
    name = filename.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(file_bytes), dtype=str, keep_default_na=False, low_memory=False)
    else:
        df = pd.read_excel(io.BytesIO(file_bytes), dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def _build_saved_mtm_lookup(df: pd.DataFrame) -> Dict[str, str]:
    norm_cols = { _norm(c): c for c in df.columns }
    uid_col = None
    for cand in ["userid", "user id", "user_id"]:
        if _norm(cand) in norm_cols:
            uid_col = norm_cols[_norm(cand)]
            break
    if not uid_col:
        raise ValueError("Saved MTM file missing 'User ID' column (UserID/User ID).")

    mtm_col = None
    for cand in ["realizedMTM", "realized MTM", "realized_mtm", "xun_realized_mtm"]:
        if _norm(cand) in norm_cols:
            mtm_col = norm_cols[_norm(cand)]
            break
    if not mtm_col:
        raise ValueError("Saved MTM file missing 'realizedMTM' column.")

    lut = {}
    for _, r in df.iterrows():
        key = _norm(r.get(uid_col))
        if key:
            lut[key] = str(r.get(mtm_col, "")).strip()
    return lut

def _apply_saved_mtm(enriched_first: pd.DataFrame, saved_lut: Dict[str, str], uid_col: Optional[str]) -> pd.DataFrame:
    if not uid_col or "MTM (All)" not in enriched_first.columns:
        return enriched_first

    df = enriched_first.copy()
    if "REMARK" not in df.columns:
        df["REMARK"] = ""

    uid_idx = df.columns.get_loc(uid_col)
    mtm_idx = df.columns.get_loc("MTM (All)")
    rem_idx = df.columns.get_loc("REMARK")

    def mutate_row(row):
        uid = _norm(row[uid_idx])
        if uid in saved_lut and saved_lut[uid] != "":
            old_mtm = row[mtm_idx]
            old_mtm_str = str(old_mtm).strip()
            if old_mtm_str != "":
                existing = str(row[rem_idx] or "").strip()
                addition = f"MTM={old_mtm_str}"
                row[rem_idx] = (existing + (" " if existing else "") + addition)
            row[mtm_idx] = saved_lut[uid]
        return row

    df = df.apply(mutate_row, axis=1)
    return df

def apply_remarks(df: pd.DataFrame) -> pd.DataFrame:
    remark_col = "REMARK"
    def remark_logic(row):
        try:
            max_loss_allocation = pd.to_numeric(row['MAX_LOSS'], errors='coerce') / pd.to_numeric(row['ALLOCATION'], errors='coerce') + 0.1
            mtm_allocation = -(pd.to_numeric(row['MTM (All)'], errors='coerce') / pd.to_numeric(row['ALLOCATION'], errors='coerce'))
            if pd.notna(max_loss_allocation) and pd.notna(mtm_allocation) and max_loss_allocation <= mtm_allocation:
                existing = str(row.get(remark_col, '')).strip()
                return (existing + (" " if existing else "") + "Slippage")
            return row.get(remark_col, '')
        except (ValueError, TypeError, ZeroDivisionError):
            return row.get(remark_col, '')

    df[remark_col] = df.apply(remark_logic, axis=1)
    return df

# -------------------- Streamlit App --------------------
st.set_page_config(page_title="Summary Enricher", layout="wide")

if 'stage' not in st.session_state:
    st.session_state.stage = 'upload'
    st.session_state.show_bulk = False

if st.session_state.stage == 'upload':
    st.title("Summary Enricher")

    usersetting = st.file_uploader("Usersetting file (.csv / .xlsx)", type=["csv", "xlsx", "xls"])
    summary = st.file_uploader("Summary file (.xlsx recommended; multi-sheet supported)", type=["csv", "xlsx", "xls"])

    col1, col2 = st.columns(2)
    with col1:
        algo = st.selectbox("ALGO", options=["", "1", "2", "5", "7", "8", "12", "15", "102"])
    with col2:
        operator = st.selectbox("OPERATOR", options=["", "GAURAVK", "CHETANB", "SAHILM", "BANSHIP", "VIKASA", "GULSHANS", "PRADYUMANS", "ASHUTOSHM", "JITESHS"])

    col3, col4 = st.columns(2)
    with col3:
        expiry = st.selectbox("EXPIRY", options=["", "NF 0DTE", "NF 1DTE", "SX 0DTE", "SX 1DTE", "BNF 1DTE", "BNF 0DTE"])
    with col4:
        remark = st.text_input("REMARK (optional)", placeholder="Only write if you want to Fill 1 text in remark in all Users")

    need_saved_mtm = (algo == "8") and ("1DTE" in (expiry or "").upper())
    saved_mtm = None
    if need_saved_mtm:
        saved_mtm = st.file_uploader("Saved MTM file (.csv / .xlsx)", type=["csv", "xlsx", "xls"])

    if st.button("Run"):
        if not usersetting or not summary:
            st.error("Please upload both files.")
        elif need_saved_mtm and not saved_mtm:
            st.error("Saved MTM is required when ALGO=8 and EXPIRY contains 1DTE.")
        else:
            try:
                consts = {
                    "ALGO": algo,
                    "OPERATOR": operator,
                    "EXPIRY": expiry,
                    "REMARK": remark,
                }
                consts["SERVER"] = _server_from_filename(usersetting.name) or _server_from_filename(summary.name)

                us_bytes = usersetting.read()
                raw_us = _read_raw(us_bytes, usersetting.name)
                us_clean = _select_usersetting_columns(raw_us)

                us_buf = io.BytesIO()
                with pd.ExcelWriter(us_buf, engine="openpyxl") as xw:
                    us_clean.to_excel(xw, index=False, sheet_name="Usersetting")
                us_buf.seek(0)
                st.session_state.us_buf = us_buf

                sm_bytes = summary.read()
                sheets = _read_all_sheets(sm_bytes, summary.name)
                names = list(sheets.keys())
                first_name = names[0]
                first_df = sheets[first_name]

                lookup = _build_lookup(us_clean)
                uid_col = "UserID" if "UserID" in first_df.columns else ("User ID" if "User ID" in first_df.columns else None)

                enriched_first = first_df.copy()
                if uid_col:
                    enriched_first = _insert_allocation_maxloss(enriched_first, lookup, uid_col)
                else:
                    insert_at = min(9, len(enriched_first.columns))
                    enriched_first.insert(insert_at, "ALLOCATION", "")
                    enriched_first.insert(insert_at + 1, "MAX_LOSS", "")

                enriched_first = _append_constants(enriched_first, consts)

                if need_saved_mtm:
                    sv_bytes = saved_mtm.read()
                    saved_df = _read_saved_mtm(sv_bytes, saved_mtm.name)
                    saved_lut = _build_saved_mtm_lookup(saved_df)
                    enriched_first = _apply_saved_mtm(enriched_first, saved_lut, uid_col)

                enriched_first = apply_remarks(enriched_first)

                enriched_first = _reorder_summary_columns(enriched_first)

                st.session_state.enriched_first = enriched_first
                st.session_state.sheets = sheets
                st.session_state.first_sheet_name = first_name
                st.session_state.original_summary_filename = summary.name
                st.session_state.consts = consts
                st.session_state.uid_col = uid_col
                st.session_state.stage = 'preview'
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")

elif st.session_state.stage == 'preview':
    st.title("Preview (Sheet 1)")

    st.write(f"Detected SERVER: {st.session_state.consts['SERVER']} • ALGO: {st.session_state.consts['ALGO']} • OPERATOR: {st.session_state.consts['OPERATOR']} • EXPIRY: {st.session_state.consts['EXPIRY']}")

    enriched_first = st.session_state.enriched_first.copy()

    # P&L calculation
    enriched_first['P&L'] = 0.00
    for idx, row in enriched_first.iterrows():
        try:
            mtm = pd.to_numeric(row["MTM (All)"], errors='coerce')
            alloc = pd.to_numeric(row["ALLOCATION"], errors='coerce')
            if pd.notna(mtm) and pd.notna(alloc) and alloc != 0:
                pnl = mtm / alloc
                enriched_first.at[idx, "P&L"] = pnl
        except:
            pass

    # Editable table
    if 'edited_df' not in st.session_state:
        st.session_state.edited_df = enriched_first.copy()

    edited_df = st.data_editor(
        st.session_state.edited_df,
        column_config={
            "REMARK": st.column_config.TextColumn("REMARK"),
            "P&L": st.column_config.NumberColumn("P&L", disabled=True)
        },
        hide_index=False,
        use_container_width=True
    )

    st.session_state.edited_df = edited_df

    if st.button("Bulk Remark"):
        st.session_state.show_bulk = not st.session_state.show_bulk
        st.rerun()

    if st.session_state.show_bulk:
        bulk_remark = st.text_area("Remark to Apply")

        # Prepare bulk table
        bulk_data = []
        for idx, row in edited_df.iterrows():
            user_id = row.get(st.session_state.uid_col)
            if user_id:
                bulk_data.append({
                    "Select": False,
                    "User ID": user_id,
                    "Alias": row.get("Alias", row.get("User Alias", "N/A")),
                    "MTM (All)": row.get("MTM (All)", "N/A"),
                    "ALLOCATION": row.get("ALLOCATION", "N/A"),
                    "MAX_LOSS": row.get("MAX_LOSS", "N/A")
                })

        bulk_df = pd.DataFrame(bulk_data)

        if 'bulk_edited' not in st.session_state:
            st.session_state.bulk_edited = bulk_df.copy()

        bulk_edited = st.data_editor(
            st.session_state.bulk_edited,
            column_config={"Select": st.column_config.CheckboxColumn("Select")},
            hide_index=True,
            use_container_width=True
        )

        st.session_state.bulk_edited = bulk_edited

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            if st.button("Apply"):
                selected_users = bulk_edited[bulk_edited["Select"] == True]["User ID"].tolist()
                for user_id in selected_users:
                    mask = edited_df[st.session_state.uid_col] == user_id
                    edited_df.loc[mask, "REMARK"] = bulk_remark
                st.session_state.edited_df = edited_df
                st.session_state.show_bulk = False
                st.rerun()
        with col2:
            if st.button("Select All"):
                bulk_edited["Select"] = True
                st.session_state.bulk_edited = bulk_edited
                st.rerun()
        with col3:
            if st.button("Clear All"):
                bulk_edited["Select"] = False
                st.session_state.bulk_edited = bulk_edited
                st.rerun()
        with col4:
            if st.button("Cancel"):
                st.session_state.show_bulk = False
                st.rerun()

    if st.button("Submit (Build Final Workbook)"):
        st.session_state.enriched_first = st.session_state.edited_df.drop(columns=["P&L"], errors='ignore')
        st.session_state.stage = 'final'
        st.rerun()

elif st.session_state.stage == 'final':
    st.title("Done ✓")

    enriched_first = st.session_state.enriched_first
    enriched_first = _coerce_numeric_columns(enriched_first)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        enriched_first.to_excel(xw, index=False, sheet_name=st.session_state.first_sheet_name[:31])
        for name, df in st.session_state.sheets.items():
            if name != st.session_state.first_sheet_name:
                df.to_excel(xw, index=False, sheet_name=name[:31])
    out.seek(0)

    original_filename = st.session_state.original_summary_filename
    if not original_filename.lower().endswith('.xlsx'):
        original_filename += '.xlsx'

    st.download_button("Download Enriched Summary", data=out, file_name=original_filename)

    st.download_button("Download Cleaned Usersetting", data=st.session_state.us_buf, file_name="Cleaned_Usersetting.xlsx")

    # Append to master with lock
    with master_lock:
        try:
            if os.path.exists(MASTER_FILE):
                master_df = pd.read_excel(MASTER_FILE)
                master_df = pd.concat([master_df, enriched_first], ignore_index=True)
            else:
                master_df = enriched_first
            master_df.to_excel(MASTER_FILE, index=False)
        except Exception as e:
            st.warning(f"Could not append to master file: {e}")

    if st.button("Start Over"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
