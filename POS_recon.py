# -*- coding: utf-8 -*-
"""
Created on Thu Oct  8 12:19:30 2026

@author: Rajasekaran.Sankaran

POS Reconciliation Tool — Partner POS file vs VCL POS file
=====================================================================

Run:
    pip install streamlit pandas openpyxl
    streamlit run POS_Recon_Testing_V2.0.py

Flow:
    1. Upload the Partner POS file and the VCL POS file (.xlsx / .xls / .csv)
    2. Confirm the column mapping (Loan ID, Status, POS) for each file
    3. Review / adjust the status mapping (so "Live" in one file can equal "Active" in the other)
    4. Click "Run Reconciliation" -> one Excel workbook is saved to ./output and offered for download

Everything in the output comes only from the uploaded files. Nothing is assumed:
column names are confirmed by the user, unparseable values are counted and reported,
and duplicates are listed on their own sheets.
"""
from __future__ import annotations



import io
import re
from collections import Counter
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# Constants
# ----------------------------------------------------------------------------
APP_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT_DIR = APP_DIR / "output"

FIELDS = {
    "loan_id": "Loan ID / LAN",
    "status": "Loan Status",
    "pos": "POS (Principal Outstanding)",
}

# Keyword hints used ONLY to pre-select a column; the user always confirms.
KEYWORDS = {
    "loan_id": ["loan id", "loanid", "loan account number", "loan account no", "loan no",
                "lan", "loan number", "account number", "loan account", "agreement no"],
    "status": ["loan status", "account status", "status"],
    "pos": ["pos", "principal outstanding", "principal os", "outstanding principal",
            "closing pos", "principal o/s", "pos amount"],
}

# Category definitions: key -> (sheet name, display name)
CATEGORIES = [
    ("s1", "Partner POS GT VCL POS", "Partner POS > than VCL POS"),
    ("s2", "Loan Status Not Matching", "Loan Status Not matching"),
    ("s3", "Collection Not Updated", "Collection Not updated. VCL POS > Partner POS"),
    ("s4", "Missing in Partner File", "Loan Missing in Partner File"),
    ("s5", "Missing in VCL File", "Loan Missing in VCL File"),
    ("s6", "POS Less Than 1", "List of Loan where 0 < POS < 1"),
    ("s7", "POS Less Than 0", "List of Loan where POS < 0"),
]
SUMMARY_TITLE = "Partner POS VS VCL POS - Scenarios"

# Difference reconciliation: every loan that makes Partner totals differ from VCL totals
# falls into exactly one bucket. key -> (label, linked scenario sheet or None)
RECON_SHEET = "Difference Reconciliation"
BRIDGE = [
    ("b1", "Loan Missing in VCL File (Partner only)", "Missing in VCL File"),
    ("b2", "Loan Missing in Partner File (VCL only)", "Missing in Partner File"),
    ("b3", "Partner POS > VCL POS", "Partner POS GT VCL POS"),
    ("b4", "VCL POS > Partner POS (Collection not updated)", "Collection Not Updated"),
    ("b5", "POS difference within tolerance", None),
    ("b6", "POS blank in one file", None),
]

ACTIONS = {
    "s1": "Check loans where collections have been updated in VCL but are not reflected in the "
          "partner file. Reconcile and correct the POS.",
    "s2": "Compare loan statuses between Partner and VCL and update the mismatched statuses.",
    "s3": "Update the missing collections in VCL and confirm once completed.",
    "s4": "Validate the missing loans and include them in the next partner data file.",
    "s5": "Investigate why these loans are available in the partner file but not on the VCL "
          "platform, and update accordingly.",
    "s6": "Review these loans and close them if there are no pending dues.",
    "s7": "Investigate negative POS (excess collection, reversal or refund not posted) and "
          "correct the balance in the respective system.",
}

DETAIL_COLUMNS = [
    "Loan ID",
    "Partner Status", "VCL Status", "Status Match (after mapping)",
    "Partner POS", "VCL POS", "POS Difference (Partner - VCL)", "Remarks",
]
# Fixed column letters used by the Summary formulas
COL_PARTNER_POS, COL_VCL_POS, COL_DIFF = "E", "F", "G"

# Styling
FONT = "Arial"
HDR_FILL = PatternFill("solid", start_color="1F3864")
SUB_FILL = PatternFill("solid", start_color="D9E1F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
AMT_FMT = '#,##0.00;(#,##0.00);"-"'
INT_FMT = '#,##0;(#,##0);"-"'


# ----------------------------------------------------------------------------
# Reading & cleaning helpers
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def list_sheets(file_bytes: bytes, name: str) -> list[str]:
    if name.lower().endswith(".csv"):
        return ["(csv)"]
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names


@st.cache_data(show_spinner=False)
def read_table(file_bytes: bytes, name: str, sheet: str, header_row: int) -> pd.DataFrame:
    """Read every column as text so loan IDs / amounts are never silently altered."""
    hdr = header_row - 1
    if name.lower().endswith(".csv"):
        df = pd.read_csv(io.BytesIO(file_bytes), dtype=str, header=hdr,
                         keep_default_na=False, encoding_errors="replace")
    else:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, dtype=str,
                           header=hdr, keep_default_na=False)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, [c for c in df.columns if not c.lower().startswith("unnamed")]]
    # drop rows that are entirely blank
    df = df[~(df.apply(lambda r: "".join(map(str, r)).strip() == "", axis=1))]
    return df.reset_index(drop=True)


def _norm_header(s: str) -> str:
    return re.sub(r"[^a-z0-9/ ]", " ", str(s).lower()).strip()


def guess_column(columns: list[str], field: str) -> int:
    """Return index of the best keyword match, or 0 (= '-- select --')."""
    normed = [re.sub(r"\s+", " ", _norm_header(c)) for c in columns]
    for kw in KEYWORDS[field]:                      # exact match first
        for i, c in enumerate(normed):
            if c == kw:
                return i + 1
    for kw in KEYWORDS[field]:                      # then whole-word containment
        for i, c in enumerate(normed):
            if re.search(rf"\b{re.escape(kw)}\b", c):
                return i + 1
    return 0


def clean_loan_id(s: pd.Series, strip_leading_zeros: bool) -> pd.Series:
    out = (s.astype(str).str.strip().str.lstrip("'").str.upper()
           .str.replace(r"\s+", "", regex=True)
           .str.replace(r"^(\d+)\.0+$", r"\1", regex=True))   # 12345.0 -> 12345
    if strip_leading_zeros:
        out = out.str.replace(r"^0+(?=\w)", "", regex=True)
    return out.replace({"": np.nan, "NAN": np.nan, "NONE": np.nan, "NULL": np.nan})


def clean_amount(s: pd.Series) -> pd.Series:
    txt = s.astype(str).str.strip()
    neg = txt.str.match(r"^\(.*\)$")                       # (1,234.00) -> -1234
    txt = txt.str.replace(r"[₹,\s()]|INR|Rs\.?", "", regex=True, flags=re.IGNORECASE)
    txt = txt.replace({"": np.nan, "-": np.nan, "nan": np.nan, "None": np.nan})
    num = pd.to_numeric(txt, errors="coerce")
    num[neg & num.notna()] = -num[neg & num.notna()].abs()
    return num


def norm_status(s: pd.Series) -> pd.Series:
    return (s.astype(str).str.strip().str.upper().str.replace(r"\s+", " ", regex=True)
            .replace({"": np.nan, "NAN": np.nan, "NONE": np.nan}))


def prepare(df: pd.DataFrame, mapping: dict, extras: list[str], prefix: str,
            strip_zeros: bool) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """Standardise one file. Returns (clean unique-loan df, duplicates df, stats)."""
    stats = {"rows_read": len(df)}
    out = pd.DataFrame({
        "loan_id": clean_loan_id(df[mapping["loan_id"]], strip_zeros),
        "status_raw": df[mapping["status"]].astype(str).str.strip(),
        "status": norm_status(df[mapping["status"]]),
        "pos": clean_amount(df[mapping["pos"]]),
    })
    for c in extras:
        out[f"{prefix} {c}"] = df[c].values

    raw_pos = df[mapping["pos"]].astype(str).str.strip()
    stats["blank_loan_id"] = int(out["loan_id"].isna().sum())
    stats["pos_unparseable"] = int((out["pos"].isna() & (raw_pos != "")).sum())
    stats["pos_blank"] = int((raw_pos == "").sum())

    out = out[out["loan_id"].notna()].copy()
    dup_mask = out.duplicated("loan_id", keep=False)
    dups = out[dup_mask].sort_values("loan_id").copy()
    stats["duplicate_ids"] = int(dups["loan_id"].nunique())
    stats["duplicate_rows"] = int(len(dups))
    out = out.drop_duplicates("loan_id", keep="first")
    stats["unique_loans"] = len(out)
    stats["os_accounts"] = int((out["pos"] > 0).sum())
    stats["total_pos"] = float(out["pos"].sum(skipna=True))
    return out, dups, stats


# ----------------------------------------------------------------------------
# Reconciliation
# ----------------------------------------------------------------------------
def reconcile(p: pd.DataFrame, y: pd.DataFrame, status_map: dict, tol: float) -> dict:
    p = p.copy(); y = y.copy()
    p["status_m"] = p["status"].map(lambda v: status_map.get(v, v) if pd.notna(v) else v)
    y["status_m"] = y["status"].map(lambda v: status_map.get(v, v) if pd.notna(v) else v)

    m = p.merge(y, on="loan_id", how="outer", suffixes=("_p", "_y"), indicator=True)
    both = m["_merge"] == "both"
    m["diff"] = (m["pos_p"] - m["pos_y"]).round(2)
    # POS rounded to paise for the 0 < POS < 1 test, so values that display as 0.00 are not picked up
    m["pos_p_r"] = m["pos_p"].round(2)
    m["pos_y_r"] = m["pos_y"].round(2)

    def same(a, b):
        return (a == b) | (a.isna() & b.isna())

    m["status_match"] = np.where(both, np.where(same(m["status_m_p"], m["status_m_y"]), "Yes", "No"), "")

    masks = {
        "s1": both & (m["diff"] > tol),
        "s2": both & (m["status_match"] == "No"),
        "s3": both & (m["diff"] < -tol),
        "s4": m["_merge"] == "right_only",
        "s5": m["_merge"] == "left_only",
        "s6": m["pos_p_r"].between(0, 1, inclusive="neither")
              | m["pos_y_r"].between(0, 1, inclusive="neither"),
        "s7": (m["pos_p_r"] < 0) | (m["pos_y_r"] < 0),
    }

    def fmt(v):
        return "blank" if pd.isna(v) else f"{v:,.2f}"

    def row_remark(r, key):
        if key == "s1":
            return f"Partner POS higher by {r['diff']:,.2f}"
        if key == "s3":
            return f"VCL POS higher by {abs(r['diff']):,.2f} - collection not reflected in VCL"
        if key == "s2":
            return f"Partner status '{r['status_raw_p']}' vs VCL status '{r['status_raw_y']}'"
        if key == "s4":
            return f"Present only in VCL file (VCL POS {fmt(r['pos_y'])})"
        if key == "s5":
            return f"Present only in Partner file (Partner POS {fmt(r['pos_p'])})"
        if key == "s6":
            parts = []
            if pd.notna(r["pos_p_r"]) and 0 < r["pos_p_r"] < 1:
                parts.append(f"Partner POS {r['pos_p']:,.2f}")
            if pd.notna(r["pos_y_r"]) and 0 < r["pos_y_r"] < 1:
                parts.append(f"VCL POS {r['pos_y']:,.2f}")
            return "POS between 0 and 1 in " + " and ".join(parts)
        if key == "s7":
            parts = []
            if pd.notna(r["pos_p_r"]) and r["pos_p_r"] < 0:
                parts.append(f"Partner POS {r['pos_p']:,.2f}")
            if pd.notna(r["pos_y_r"]) and r["pos_y_r"] < 0:
                parts.append(f"VCL POS {r['pos_y']:,.2f}")
            return "Negative POS in " + " and ".join(parts)
        if key == "b5":
            return f"Difference {r['c_pos']:,.2f} is within tolerance {tol:,.2f}"
        if key == "b6":
            return ("Partner POS blank" if pd.isna(r["pos_p"]) else "VCL POS blank") + \
                f" (Partner {fmt(r['pos_p'])} | VCL {fmt(r['pos_y'])})"
        return ""

    extra_cols = [c for c in m.columns if c.startswith("Partner ") or c.startswith("VCL ")]

    def to_detail(sub: pd.DataFrame, key: str | None) -> pd.DataFrame:
        d = pd.DataFrame({
            "Loan ID": sub["loan_id"],
            "Partner Status": sub["status_raw_p"],
            "VCL Status": sub["status_raw_y"],
            "Status Match (after mapping)": sub["status_match"],
            "Partner POS": sub["pos_p"],
            "VCL POS": sub["pos_y"],
            "POS Difference (Partner - VCL)": sub["diff"],
            "Remarks": sub.apply(lambda r: row_remark(r, key), axis=1) if key and len(sub) else "",
        })
        for c in extra_cols:
            d[c] = sub[c]
        return d.reset_index(drop=True)

    details = {}
    for key, _, _ in CATEGORIES:
        sub = m[masks[key]]
        if key in ("s1",):
            sub = sub.sort_values("diff", ascending=False)
        elif key == "s3":
            sub = sub.sort_values("diff")
        details[key] = to_detail(sub, key)

    # All-loans view with flags
    all_df = to_detail(m, None)
    flags = []
    for i in range(len(m)):
        f = [name for key, _, name in CATEGORIES if masks[key].iloc[i]]
        flags.append("; ".join(f) if f else "Matched")
    all_df["Remarks"] = flags
    all_df.insert(1, "Presence", m["_merge"].map(
        {"both": "Both files", "left_only": "Partner only", "right_only": "VCL only"}).astype(str).values)

    # ---- Difference reconciliation (Partner minus VCL) ----------------------
    # Each loan's contribution to the three Observation differences. Summing these
    # contributions over all loans gives exactly Partner total - VCL total.
    m["c_acc"] = np.select([m["_merge"] == "left_only", m["_merge"] == "right_only"], [1, -1], 0)
    m["c_os"] = (m["pos_p"] > 0).astype(int) - (m["pos_y"] > 0).astype(int)
    m["c_pos"] = m["pos_p"].fillna(0) - m["pos_y"].fillna(0)

    b_cond = {
        "b1": m["_merge"] == "left_only",
        "b2": m["_merge"] == "right_only",
        "b3": masks["s1"],
        "b4": masks["s3"],
        "b6": both & (m["pos_p"].isna() != m["pos_y"].isna()),
        "b5": both & m["pos_p"].notna() & m["pos_y"].notna() & ~masks["s1"] & ~masks["s3"]
              & (m["c_pos"] != 0),
    }
    order = ["b1", "b2", "b3", "b4", "b6", "b5"]           # first match wins -> exclusive
    m["bucket"] = np.select([b_cond[k] for k in order], order, default="")
    bucket_label = {k: lbl for k, lbl, _ in BRIDGE}
    remark_key = {"b1": "s5", "b2": "s4", "b3": "s1", "b4": "s3", "b5": "b5", "b6": "b6"}

    dsub = m[m["bucket"] != ""].copy()
    dsub["_ord"] = dsub["bucket"].map({k: i for i, (k, _, _) in enumerate(BRIDGE)})
    dsub["_abs"] = dsub["c_pos"].abs()
    dsub = dsub.sort_values(["_ord", "_abs"], ascending=[True, False])
    diff_detail = pd.DataFrame({
        "Difference Type": dsub["bucket"].map(bucket_label),
        "Loan ID": dsub["loan_id"],
        "Presence": dsub["_merge"].map({"both": "Both files", "left_only": "Partner only",
                                        "right_only": "VCL only"}).astype(str),
        "Partner Status": dsub["status_raw_p"],
        "VCL Status": dsub["status_raw_y"],
        "Partner POS": dsub["pos_p"],
        "VCL POS": dsub["pos_y"],
        "POS Difference (Partner - VCL)": dsub["diff"],
        "Impact on No. of Accounts": dsub["c_acc"],
        "Impact on Accounts with O/S": dsub["c_os"],
        "Impact on Total POS": dsub["c_pos"],
        "Remarks": [row_remark(r, remark_key[r["bucket"]]) for _, r in dsub.iterrows()],
    })
    for c in extra_cols:
        diff_detail[c] = dsub[c].values
    diff_detail = diff_detail.reset_index(drop=True)

    bridge = []
    for k, lbl, _ in BRIDGE:
        s = m[m["bucket"] == k]
        bridge.append({"Difference Type": lbl, "No. of Loans": len(s),
                       "No. of Accounts": int(s["c_acc"].sum()),
                       "No. of Accounts with O/S": int(s["c_os"].sum()),
                       "Total POS": float(s["c_pos"].sum())})

    return {"merged": m, "masks": masks, "details": details, "all": all_df,
            "matched": int(both.sum()), "diff_detail": diff_detail,
            "bridge": pd.DataFrame(bridge)}


# ----------------------------------------------------------------------------
# Data-driven commentary (every number below is computed from the uploaded data)
# ----------------------------------------------------------------------------
def build_comments(res: dict, tol: float, p_name: str, y_name: str) -> dict:
    """One crisp line per scenario - every number is computed from the uploaded data."""
    m, masks = res["merged"], res["masks"]
    a = lambda v: f"{v:,.2f}"
    c = {}

    sub = m[masks["s1"]]
    c["s1"] = (f"Partner POS higher by {a(sub['diff'].sum())} in total. "
               f"Check collections posted in VCL but not by partner." if len(sub) else "Nil")

    sub = m[masks["s2"]]
    if len(sub):
        (pp, yy), n = Counter(zip(sub["status_raw_p"], sub["status_raw_y"])).most_common(1)[0]
        c["s2"] = f"Most common: Partner '{pp}' vs VCL '{yy}' ({n:,} loans)."
    else:
        c["s2"] = "Nil"

    sub = m[masks["s3"]]
    c["s3"] = (f"VCL POS higher by {a(abs(sub['diff'].sum()))} in total. "
               f"Partner collections not yet updated in VCL." if len(sub) else "Nil")

    sub = m[masks["s4"]]
    c["s4"] = (f"VCL POS on these loans: {a(sub['pos_y'].sum())}." if len(sub) else "Nil")

    sub = m[masks["s5"]]
    c["s5"] = (f"Partner POS on these loans: {a(sub['pos_p'].sum())}." if len(sub) else "Nil")

    sub = m[masks["s6"]]
    if len(sub):
        pc = int(sub["pos_p_r"].between(0, 1, inclusive="neither").sum())
        yc = int(sub["pos_y_r"].between(0, 1, inclusive="neither").sum())
        c["s6"] = f"Partner: {pc:,} | VCL: {yc:,}. Candidates for closure."
    else:
        c["s6"] = "Nil"

    sub = m[masks["s7"]]
    if len(sub):
        pn = sub["pos_p_r"] < 0
        yn = sub["pos_y_r"] < 0
        c["s7"] = (f"Partner: {int(pn.sum()):,} loans ({a(sub.loc[pn, 'pos_p'].sum())}) | "
                   f"VCL: {int(yn.sum()):,} loans ({a(sub.loc[yn, 'pos_y'].sum())}).")
    else:
        c["s7"] = "Nil"
    return c


# ----------------------------------------------------------------------------
# Excel writer
# ----------------------------------------------------------------------------
def _style_header(ws, row, ncols, fill=HDR_FILL, color="FFFFFF"):
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = Font(name=FONT, bold=True, color=color, size=10)
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def write_df(ws, df: pd.DataFrame, start_row: int = 1):
    """Write df with its header on start_row. Returns (first data row, last data row)."""
    cols = list(df.columns)
    for c_i, h in enumerate(cols, start=1):
        ws.cell(row=start_row, column=c_i, value=h)
    _style_header(ws, start_row, len(cols))
    amt_idx = {i + 1 for i, c in enumerate(cols) if c in (
        "Partner POS", "VCL POS", "POS Difference (Partner - VCL)", "Impact on Total POS")}

    for r_i, row in enumerate(df.itertuples(index=False), start=start_row + 1):
        for c_i, v in enumerate(row, start=1):
            if isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            elif v is pd.NaT or (isinstance(v, float) and np.isnan(v)):
                v = None
            elif isinstance(v, np.generic):
                v = v.item()
            cell = ws.cell(row=r_i, column=c_i, value=v)
            cell.font = Font(name=FONT, size=10)
            if c_i in amt_idx:
                cell.number_format = AMT_FMT

    for i, c in enumerate(cols, start=1):
        width = max(12, min(45, len(str(c)) + 2))
        if c == "Remarks":
            width = 60
        ws.column_dimensions[get_column_letter(i)].width = width
    if start_row == 1:
        ws.freeze_panes = "B2"
    last = start_row + len(df)
    if len(df):
        ws.auto_filter.ref = f"A{start_row}:{get_column_letter(len(cols))}{last}"
    return start_row + 1, last


def write_recon_sheet(ws, res):
    """Bridge (formulas) on top + loan-level list below. Every bridge number is a
    COUNTIF/SUMIF over the loan list, and the Check row ties it to the Summary."""
    dd = res["diff_detail"]
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Partner POS VS VCL POS - Difference Reconciliation (Partner minus VCL)"
    ws["A1"].font = Font(name=FONT, bold=True, size=12)
    ws["A2"] = ("Every loan that causes the Summary differences is listed below. Totals are formulas "
                "over that list; the Check row must be 0.")
    ws["A2"].font = Font(name=FONT, size=9, italic=True)

    hdr_row = 4
    heads = ["Difference Type", "No. of Loans", "No. of Accounts", "No. of Accounts with O/S",
             "Total POS", "Detail sheet"]
    for c_i, h in enumerate(heads, start=1):
        ws.cell(row=hdr_row, column=c_i, value=h)
    _style_header(ws, hdr_row, len(heads))

    first_b = hdr_row + 1
    total_r = first_b + len(BRIDGE)
    summ_r, check_r = total_r + 1, total_r + 2
    det_hdr = check_r + 3

    cols = list(dd.columns)
    L = lambda name: get_column_letter(cols.index(name) + 1)
    d0, d1 = det_hdr + 1, max(det_hdr + 1, det_hdr + len(dd))
    rng = lambda name: f"${L(name)}${d0}:${L(name)}${d1}"
    typ = rng("Difference Type")
    sum_cols = {3: "Impact on No. of Accounts", 4: "Impact on Accounts with O/S",
                5: "Impact on Total POS"}

    def put(r, c, v, fmt=None, bold=False, fill=None, link=None):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = Font(name=FONT, size=10, bold=bold, color="0563C1" if link else None,
                         underline="single" if link else None)
        if link:
            cell.hyperlink = link
        if fmt:
            cell.number_format = fmt
        if fill:
            cell.fill = fill
        cell.border = BORDER
        return cell

    for r, (_, lbl, sheet) in enumerate(BRIDGE, start=first_b):
        put(r, 1, lbl)
        put(r, 2, f'=COUNTIF({typ},A{r})', INT_FMT)
        for c, name in sum_cols.items():
            put(r, c, f'=SUMIF({typ},A{r},{rng(name)})', AMT_FMT if c == 5 else INT_FMT)
        put(r, 6, sheet, link=f"#'{sheet}'!A1" if sheet else None)
    put(total_r, 1, "Total difference (from loan list)", bold=True, fill=SUB_FILL)
    for c in range(2, 6):
        col = get_column_letter(c)
        put(total_r, c, f"=SUM({col}{first_b}:{col}{total_r - 1})",
            AMT_FMT if c == 5 else INT_FMT, bold=True, fill=SUB_FILL)
    put(total_r, 6, None, fill=SUB_FILL)
    put(summ_r, 1, "Difference as per Summary", bold=True)
    put(summ_r, 2, None)
    for c, srow in ((3, 3), (4, 4), (5, 5)):           # Summary observation rows 3..5
        put(summ_r, c, f"=Summary!D{srow}", AMT_FMT if c == 5 else INT_FMT, bold=True)
    put(summ_r, 6, "Summary", link="#'Summary'!A1")
    green = PatternFill("solid", start_color="E2EFDA")
    put(check_r, 1, "Check (should be 0)", bold=True, fill=green)
    put(check_r, 2, None, fill=green)
    for c in (3, 4, 5):
        col = get_column_letter(c)
        put(check_r, c, f"=ROUND({col}{total_r}-{col}{summ_r},2)", AMT_FMT if c == 5 else INT_FMT,
            bold=True, fill=green)
    put(check_r, 6, None, fill=green)

    ws.cell(row=det_hdr - 1, column=1, value="Loan-level differences").font = \
        Font(name=FONT, bold=True, size=11)
    write_df(ws, dd, start_row=det_hdr)

    # Impact columns as formulas so they always agree with the POS values shown
    P, Y, PR = L("Partner POS"), L("VCL POS"), L("Presence")
    for r in range(d0, det_hdr + len(dd) + 1):
        ws[f"{L('Impact on No. of Accounts')}{r}"] = \
            f'=IF({PR}{r}="Partner only",1,IF({PR}{r}="VCL only",-1,0))'
        ws[f"{L('Impact on Accounts with O/S')}{r}"] = f"=(N({P}{r})>0)-(N({Y}{r})>0)"
        ws[f"{L('Impact on Total POS')}{r}"] = f"=N({P}{r})-N({Y}{r})"
    ws.column_dimensions["A"].width = 46


def build_workbook(res, comments, p_stats, y_stats) -> Workbook:
    """Output = Summary + Difference Reconciliation + one sheet per scenario."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    grey = PatternFill("solid", start_color="D9D9D9")
    hdr_grey = PatternFill("solid", start_color="BFBFBF")
    thin = Side(style="thin", color="FFFFFF")
    grid = Border(left=thin, right=thin, top=thin, bottom=thin)

    def _cell(row, col, value, bold=False, fmt=None, align="left", fill=grey, link=None):
        c = ws.cell(row=row, column=col, value=value)
        c.font = Font(name=FONT, size=10, bold=bold,
                      color="0563C1" if link else None, underline="single" if link else None)
        if link:
            c.hyperlink = link
        if fmt:
            c.number_format = fmt
        c.fill = fill
        c.border = grid
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        return c

    # ---- Block 1: Observation table (Partner vs VCL totals) ----
    ws.merge_cells("A1:D1")
    ws["A1"] = "Partner POS VS VCL POS - Observation"
    ws["A1"].font = Font(name=FONT, bold=True, size=12)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22
    for col, h in enumerate(["Observation", "Partner POS", "VCL POS",
                             "Difference (Partner - VCL)"], start=1):
        _cell(2, col, h, bold=True, align="left" if col == 1 else "center", fill=hdr_grey)
    ws.row_dimensions[2].height = 30
    obs = [
        ("No. of Accounts", p_stats["unique_loans"], y_stats["unique_loans"], INT_FMT),
        ("No. of Accounts with O/S", p_stats["os_accounts"], y_stats["os_accounts"], INT_FMT),
        ("Total POS", p_stats["total_pos"], y_stats["total_pos"], AMT_FMT),
    ]
    for r, (label, pv, yv, fmt) in enumerate(obs, start=3):
        _cell(r, 1, label, bold=True)
        _cell(r, 2, pv, fmt=fmt, align="right")
        _cell(r, 3, yv, fmt=fmt, align="right")
        _cell(r, 4, f"=B{r}-C{r}", bold=True, fmt=fmt, align="right")
        ws.cell(row=r, column=4).hyperlink = f"#'{RECON_SHEET}'!A1"   # drill into the differences

    # ---- Block 2: Scenario table ----
    top = 3 + len(obs) + 1          # one blank row after the observation table
    ws.merge_cells(start_row=top, start_column=1, end_row=top, end_column=5)
    ws.cell(row=top, column=1, value=SUMMARY_TITLE).font = Font(name=FONT, bold=True, size=12)
    ws.cell(row=top, column=1).alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[top].height = 22
    hdr = top + 1
    _cell(hdr, 1, "Scenario", bold=True, fill=hdr_grey)
    _cell(hdr, 2, "No. of Loans", bold=True, align="center", fill=hdr_grey)
    _cell(hdr, 3, "Comments", bold=True, fill=hdr_grey)
    _cell(hdr, 4, None, bold=True, fill=hdr_grey)
    ws.merge_cells(start_row=hdr, start_column=3, end_row=hdr, end_column=4)
    _cell(hdr, 5, "Action Required", bold=True, fill=hdr_grey)
    ws.row_dimensions[hdr].height = 30

    for r, (key, sheet, name) in enumerate(CATEGORIES, start=hdr + 1):
        _cell(r, 1, name, link=f"#'{sheet}'!A1")
        # live count of the scenario sheet, so Summary and sheet can never disagree
        _cell(r, 2, f"=MAX(COUNTA('{sheet}'!A:A)-1,0)", bold=True, fmt="#,##0", align="center")
        _cell(r, 3, comments[key])
        _cell(r, 4, None)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        _cell(r, 5, ACTIONS[key])
        ws.row_dimensions[r].height = 45

    # Data-quality footnote, only when something was found in the files
    r = hdr + len(CATEGORIES) + 2
    for side, s_ in (("Partner", p_stats), ("VCL", y_stats)):
        issues = []
        if s_["duplicate_ids"]:
            issues.append(f"{s_['duplicate_ids']:,} Loan IDs repeated (first row used)")
        if s_["blank_loan_id"]:
            issues.append(f"{s_['blank_loan_id']:,} rows with blank Loan ID excluded")
        if s_["pos_unparseable"]:
            issues.append(f"{s_['pos_unparseable']:,} POS values not numeric")
        if issues:
            ws.cell(row=r, column=1, value=f"Note - {side} file: " + "; ".join(issues)).font = \
                Font(name=FONT, size=9, italic=True, color="C00000")
            r += 1

    for col, w in zip("ABCDE", (30, 22, 22, 26, 55)):
        ws.column_dimensions[col].width = w
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    write_recon_sheet(wb.create_sheet(RECON_SHEET), res)
    # Scenario detail sheets
    for key, sheet, _ in CATEGORIES:
        write_df(wb.create_sheet(sheet), res["details"][key])
    return wb


# ----------------------------------------------------------------------------
# Streamlit UI
# ----------------------------------------------------------------------------
def file_block(label: str, key: str):
    up = st.file_uploader(label, type=["xlsx", "xls", "xlsm", "csv"], key=f"up_{key}")
    if not up:
        return None, None
    data = up.getvalue()
    sheets = list_sheets(data, up.name)
    c1, c2 = st.columns(2)
    sheet = c1.selectbox("Sheet", sheets, key=f"sh_{key}")
    hdr = c2.number_input("Header row", min_value=1, max_value=50, value=1, key=f"hd_{key}")
    df = read_table(data, up.name, sheet, int(hdr))
    st.caption(f"{len(df):,} rows · {len(df.columns)} columns")
    with st.expander("Preview (first 5 rows)"):
        st.dataframe(df.head(), use_container_width=True)
    return df, up.name


def mapping_block(df: pd.DataFrame, key: str) -> tuple[dict, list[str]]:
    opts = ["-- select --"] + list(df.columns)
    mapping = {}
    for f, label in FIELDS.items():
        idx = guess_column(list(df.columns), f)
        sel = st.selectbox(label, opts, index=idx, key=f"map_{key}_{f}")
        mapping[f] = None if sel == "-- select --" else sel
    used = {v for v in mapping.values() if v}
    extras = st.multiselect("Extra columns to carry into output (optional)",
                            [c for c in df.columns if c not in used], key=f"ex_{key}")
    return mapping, extras


def main():
    st.set_page_config(page_title="POS Reconciliation", page_icon="📊", layout="wide")
    st.title("📊 POS Reconciliation — Partner vs VCL")

    with st.sidebar:
        st.header("Settings")
        tol = st.number_input("POS tolerance (absolute)", min_value=0.0, value=0.0, step=1.0,
                              help="A POS difference is flagged only if it is greater than this value.")
        strip_zeros = st.checkbox("Ignore leading zeros in Loan ID", value=False)
        out_dir = Path(st.text_input("Output folder", str(DEFAULT_OUTPUT_DIR)))

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("1. Partner POS file")
        p_df, p_name = file_block("Upload Partner POS file", "p")
    with c2:
        st.subheader("2. VCL POS file")
        y_df, y_name = file_block("Upload VCL POS file", "y")

    if p_df is None or y_df is None:
        st.info("Upload both files to continue.")
        return

    st.divider()
    st.subheader("3. Confirm column mapping")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Partner file**")
        p_map, p_extra = mapping_block(p_df, "p")
    with c2:
        st.markdown("**VCL file**")
        y_map, y_extra = mapping_block(y_df, "y")

    errors = []
    for side, mp in (("Partner", p_map), ("VCL", y_map)):
        missing = [FIELDS[f] for f, v in mp.items() if not v]
        if missing:
            errors.append(f"{side}: select {', '.join(missing)}")
        chosen = [v for v in mp.values() if v]
        if len(chosen) != len(set(chosen)):
            errors.append(f"{side}: the same column is mapped to more than one field")
    if errors:
        for e in errors:
            st.warning(e)
        return

    # --- Status mapping -------------------------------------------------
    st.divider()
    st.subheader("4. Status mapping")
    st.caption("Statuses are compared after trimming and upper-casing. If the two systems use different "
               "words for the same status (e.g. LIVE vs ACTIVE), set the same value in **Treat as**.")
    p_st = norm_status(p_df[p_map["status"]]).dropna()
    y_st = norm_status(y_df[y_map["status"]]).dropna()
    vals = sorted(set(p_st) | set(y_st))
    pc, yc = Counter(p_st), Counter(y_st)
    sm_df = pd.DataFrame({
        "Status value": vals,
        "Rows in Partner": [pc.get(v, 0) for v in vals],
        "Rows in VCL": [yc.get(v, 0) for v in vals],
        "Treat as": vals,
    })
    sm_df = st.data_editor(sm_df, disabled=["Status value", "Rows in Partner", "Rows in VCL"],
                           hide_index=True, use_container_width=True, key="status_map")
    status_map = {r["Status value"]: str(r["Treat as"]).strip().upper() or r["Status value"]
                  for _, r in sm_df.iterrows()}

    st.divider()
    if not st.button("▶ Run Reconciliation", type="primary"):
        return

    with st.spinner("Reconciling..."):
        p_clean, p_dups, p_stats = prepare(p_df, p_map, p_extra, "Partner", strip_zeros)
        y_clean, y_dups, y_stats = prepare(y_df, y_map, y_extra, "VCL", strip_zeros)
        res = reconcile(p_clean, y_clean, status_map, tol)
        comments = build_comments(res, tol, p_name, y_name)
        meta = {
            "generated": datetime.now().strftime("%d-%b-%Y %H:%M:%S"),
            "p_name": p_name, "y_name": y_name, "tol": tol,
            "p_map": " | ".join(f"{FIELDS[k]} = {v}" for k, v in p_map.items()),
            "y_map": " | ".join(f"{FIELDS[k]} = {v}" for k, v in y_map.items()),
            "strip_zeros": "Yes" if strip_zeros else "No",
        }
        wb = build_workbook(res, comments, p_stats, y_stats)
        out_dir.mkdir(parents=True, exist_ok=True)
        fname = f"POS_Reconciliation_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        out_path = out_dir / fname
        wb.save(out_path)
        buf = io.BytesIO(); wb.save(buf)

    st.success(f"Saved: {out_path}")
    st.download_button("⬇ Download Excel", buf.getvalue(), file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # On-screen summary (same numbers as the workbook)
    st.subheader("Summary")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Partner unique loans", f"{p_stats['unique_loans']:,}")
    k2.metric("VCL unique loans", f"{y_stats['unique_loans']:,}")
    k3.metric("In both files", f"{res['matched']:,}")
    k4.metric("POS difference (total)", f"{p_stats['total_pos'] - y_stats['total_pos']:,.2f}")

    st.markdown("**Partner POS VS VCL POS - Observation**")
    obs_rows = [("No. of Accounts", "unique_loans", "{:,.0f}"),
                ("No. of Accounts with O/S", "os_accounts", "{:,.0f}"),
                ("Total POS", "total_pos", "{:,.2f}")]
    st.dataframe(pd.DataFrame({
        "Observation": [o for o, _, _ in obs_rows],
        "Partner POS": [f.format(p_stats[k]) for _, k, f in obs_rows],
        "VCL POS": [f.format(y_stats[k]) for _, k, f in obs_rows],
        "Difference (Partner - VCL)":
            [f.format(p_stats[k] - y_stats[k]) for _, k, f in obs_rows],
    }), use_container_width=True, hide_index=True)

    st.markdown("**Difference Reconciliation (Partner - VCL)**")
    br = res["bridge"]
    tot = br[["No. of Loans", "No. of Accounts", "No. of Accounts with O/S", "Total POS"]].sum()
    target = {"No. of Accounts": p_stats["unique_loans"] - y_stats["unique_loans"],
              "No. of Accounts with O/S": p_stats["os_accounts"] - y_stats["os_accounts"],
              "Total POS": p_stats["total_pos"] - y_stats["total_pos"]}
    show = pd.concat([br, pd.DataFrame([
        {"Difference Type": "Total difference (from loan list)", **tot.to_dict()},
        {"Difference Type": "Difference as per Summary", "No. of Loans": None, **target}])],
        ignore_index=True)
    st.dataframe(show.style.format({"No. of Loans": "{:,.0f}", "No. of Accounts": "{:,.0f}",
                                    "No. of Accounts with O/S": "{:,.0f}", "Total POS": "{:,.2f}"},
                                   na_rep=""),
                 use_container_width=True, hide_index=True)
    off = [k for k, v in target.items() if round(tot[k] - v, 2) != 0]
    if off:
        st.error(f"Difference reconciliation does not tie for: {', '.join(off)}")
    else:
        st.success("Difference reconciliation ties to the Summary (count, O/S count and POS).")

    st.markdown(f"**{SUMMARY_TITLE}**")
    st.dataframe(pd.DataFrame({"Scenario": [n for _, _, n in CATEGORIES],
                               "No. of Loans": [len(res["details"][k]) for k, _, _ in CATEGORIES],
                               "Comments": [comments[k] for k, _, _ in CATEGORIES],
                               "Action Required": [ACTIONS[k] for k, _, _ in CATEGORIES]}),
                 use_container_width=True, hide_index=True)
    for side, s in (("Partner", p_stats), ("VCL", y_stats)):
        if s["duplicate_ids"] or s["pos_unparseable"] or s["blank_loan_id"]:
            st.warning(f"{side} file: {s['duplicate_ids']:,} duplicate Loan IDs, "
                       f"{s['blank_loan_id']:,} blank Loan IDs, {s['pos_unparseable']:,} unreadable POS "
                       f"— see workbook.")


if __name__ == "__main__":
    main()