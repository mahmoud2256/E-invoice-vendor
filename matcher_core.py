"""
matcher_core.py
----------------
Core parsing + matching logic for the ETA / TINA / Supplier-profile
invoice reconciliation tool. Kept separate from the Streamlit UI so it
can be tested and reused independently (same spirit as Caesar.py).
"""

import re
from io import BytesIO
from xml.etree import ElementTree as ET

import pandas as pd

SS_NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}

# ---------------------------------------------------------------------------
# Placeholder detection for "Voucher No./EInv Internal ID"
# ---------------------------------------------------------------------------
def _is_counting_sequence(digits: str) -> bool:
    """
    True if `digits` is the start of the concatenated counting sequence
    '0123456789101112...' or '123456789101112...' (catches junk like
    '0123456789' and '12345678910').
    """
    if len(digits) < 4:
        return False
    for start in (0, 1):
        s = ""
        n = start
        while len(s) < len(digits):
            s += str(n)
            n += 1
        if s[: len(digits)] == digits:
            return True
    return False


def is_placeholder_voucher(value) -> bool:
    """True if a voucher value should be treated as 'not filled in'."""
    if value is None:
        return True
    s = str(value).strip()
    if s == "" or s.upper() in {"N/A", "NA", "NONE"}:
        return True
    digits_only = re.sub(r"\D", "", s)
    if digits_only == "":
        return True
    if set(digits_only) == {"0"}:  # all-zero: 000, 0000000000, ...
        return True
    if _is_counting_sequence(digits_only):  # 12345678910, 0123456789, ...
        return True
    return False


def _clean_voucher_series(series: pd.Series) -> pd.Series:
    def clean(v):
        if is_placeholder_voucher(v):
            return None
        return str(v).strip()
    return series.apply(clean)


# ---------------------------------------------------------------------------
# Generic reader for the TINA "SpreadsheetML" .xls export
# (used for both the sales report and the supplier profile file)
# ---------------------------------------------------------------------------
def read_tina_spreadsheetml(file_like) -> pd.DataFrame:
    """
    Parse a TINA-exported .xls file that is actually Excel 2003 XML
    (SpreadsheetML), not a binary .xls. Returns the first worksheet as
    a DataFrame using row 1 as the header.
    """
    if hasattr(file_like, "read"):
        data = file_like.read()
    else:
        with open(file_like, "rb") as f:
            data = f.read()

    root = ET.fromstring(data)
    ws = root.find("ss:Worksheet", SS_NS)
    table = ws.find("ss:Table", SS_NS)
    rows = table.findall("ss:Row", SS_NS)

    def cell_text(cell):
        d = cell.find("ss:Data", SS_NS)
        return d.text if d is not None else None

    # Rows can have an explicit ss:Index on cells to skip blank columns;
    # rebuild each row into a dense list matching the header width.
    def row_values(row, width=None):
        cells = row.findall("ss:Cell", SS_NS)
        values = {}
        pos = 0
        for c in cells:
            idx_attr = c.attrib.get("{urn:schemas-microsoft-com:office:spreadsheet}Index")
            if idx_attr is not None:
                pos = int(idx_attr) - 1
            values[pos] = cell_text(c)
            pos += 1
        if width is None:
            width = max(values.keys(), default=-1) + 1
        return [values.get(i) for i in range(width)]

    header = row_values(rows[0])
    width = len(header)
    data_rows = [row_values(r, width=width) for r in rows[1:]]

    df = pd.DataFrame(data_rows, columns=header)
    return df


# ---------------------------------------------------------------------------
# Reader for the ETA portal export (.xlsx)
# ---------------------------------------------------------------------------
def read_eta_portal_file(file_like, sheet_name="جميع الفواتير") -> pd.DataFrame:
    xls = pd.ExcelFile(file_like)
    sheet = sheet_name if sheet_name in xls.sheet_names else xls.sheet_names[0]
    return xls.parse(sheet)


# ---------------------------------------------------------------------------
# Column name constants (adjust here if a source file's headers change)
# ---------------------------------------------------------------------------
TINA_COL_INVOICE_NO = "Invoice no."
TINA_COL_VOUCHER = "Voucher No./EInv Internal ID"
TINA_COL_SUPPLIER_ID = "Supplier ID"
TINA_COL_TOTAL = "TOTAL"
TINA_COL_CURRENCY = "Invoice currency"
TINA_COL_CLIENT_NAME = "Client Name"

PROFILE_COL_SUPPLIER_ID = "Supplier id"
PROFILE_COL_FISCAL_CODE = "Fiscal code"
PROFILE_COL_INCOME_TAX = "Income Tax Y/N"
PROFILE_COL_COMPANY = "Company"

ETA_COL_INTERNAL_NO = "الرقم الداخلى"
ETA_COL_TAX_ID = "الرقم الضريبى للبائع"
ETA_COL_TOTAL = "إجمالى الفاتورة"
ETA_COL_CURRENCY = "عملة الفاتورة"
ETA_COL_SELLER_NAME = "إسم البائع"
ETA_COL_DOC_TYPE = "نوع المستند"


# ---------------------------------------------------------------------------
# Step 1: group TINA sales lines into supplier-level invoice records
# ---------------------------------------------------------------------------
def build_tina_supplier_invoices(tina_df: pd.DataFrame) -> pd.DataFrame:
    """
    Collapse TINA's line-level rows into one row per
    (Invoice no., Supplier ID, Voucher No.) when a voucher is present,
    else per (Invoice no., Supplier ID). Sums TOTAL within each group.
    """
    df = tina_df.copy()
    df[TINA_COL_TOTAL] = pd.to_numeric(df[TINA_COL_TOTAL], errors="coerce")
    df["_voucher_clean"] = _clean_voucher_series(df[TINA_COL_VOUCHER])

    # group key: voucher when present, else a placeholder so rows with no
    # voucher for the same (invoice, supplier) still collapse together
    df["_group_voucher"] = df["_voucher_clean"].fillna("__NO_VOUCHER__")

    grouped = df.groupby(
        [TINA_COL_INVOICE_NO, TINA_COL_SUPPLIER_ID, "_group_voucher"],
        dropna=False,
    ).agg(
        total_amount=(TINA_COL_TOTAL, "sum"),
        currency=(TINA_COL_CURRENCY, "first"),
        line_count=(TINA_COL_TOTAL, "size"),
        client_name=(TINA_COL_CLIENT_NAME, "first") if TINA_COL_CLIENT_NAME in df.columns else (TINA_COL_TOTAL, "size"),
    ).reset_index()

    grouped = grouped.rename(columns={
        TINA_COL_INVOICE_NO: "invoice_no",
        TINA_COL_SUPPLIER_ID: "supplier_id",
        "_group_voucher": "voucher_no",
    })
    grouped["voucher_no"] = grouped["voucher_no"].replace("__NO_VOUCHER__", None)
    return grouped


# ---------------------------------------------------------------------------
# Step 2: enrich with supplier profile (tax id + taxable status)
# ---------------------------------------------------------------------------
def enrich_with_supplier_profile(tina_invoices: pd.DataFrame, profile_df: pd.DataFrame) -> pd.DataFrame:
    prof = profile_df[[PROFILE_COL_SUPPLIER_ID, PROFILE_COL_FISCAL_CODE,
                        PROFILE_COL_INCOME_TAX, PROFILE_COL_COMPANY]].copy()
    prof = prof.rename(columns={
        PROFILE_COL_SUPPLIER_ID: "supplier_id",
        PROFILE_COL_FISCAL_CODE: "fiscal_code",
        PROFILE_COL_INCOME_TAX: "income_tax_flag",
        PROFILE_COL_COMPANY: "supplier_company",
    })
    prof["supplier_id"] = prof["supplier_id"].astype(str).str.strip()

    out = tina_invoices.copy()
    out["supplier_id"] = out["supplier_id"].astype(str).str.strip()
    out = out.merge(prof, on="supplier_id", how="left")

    def taxable_label(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "Unknown"
        s = str(v).strip()
        if s == "1":
            return "Taxable"
        if s == "0":
            return "Non-taxable"
        return "Unknown"

    out["taxable_status"] = out["income_tax_flag"].apply(taxable_label)
    out["has_tax_id"] = out["fiscal_code"].notna() & (out["fiscal_code"].astype(str).str.strip() != "")
    return out


# ---------------------------------------------------------------------------
# Step 3: load + normalize ETA portal invoices
# ---------------------------------------------------------------------------
def normalize_eta_invoices(eta_df: pd.DataFrame) -> pd.DataFrame:
    df = eta_df.copy()
    df[ETA_COL_TOTAL] = pd.to_numeric(df[ETA_COL_TOTAL], errors="coerce")
    df[ETA_COL_TAX_ID] = df[ETA_COL_TAX_ID].astype(str).str.strip()
    df[ETA_COL_INTERNAL_NO] = df[ETA_COL_INTERNAL_NO].astype(str).str.strip()
    df["_matched"] = False
    return df.rename(columns={
        ETA_COL_INTERNAL_NO: "eta_internal_no",
        ETA_COL_TAX_ID: "eta_tax_id",
        ETA_COL_TOTAL: "eta_total",
        ETA_COL_CURRENCY: "eta_currency",
        ETA_COL_SELLER_NAME: "eta_seller_name",
        ETA_COL_DOC_TYPE: "eta_doc_type",
    })


# ---------------------------------------------------------------------------
# Step 4: matching
# ---------------------------------------------------------------------------
def match_invoices(tina_enriched: pd.DataFrame, eta_norm: pd.DataFrame, tolerance: float = 0.01):
    tina = tina_enriched.copy()
    eta = eta_norm.copy()
    tina["_matched"] = False
    tina["match_type"] = None
    tina["matched_eta_index"] = None

    fiscal_str = tina["fiscal_code"].astype(str).str.strip()

    # --- Pass 1: exact match on tax id + voucher number, amount within tolerance ---
    for i, row in tina[tina["voucher_no"].notna() & tina["has_tax_id"]].iterrows():
        candidates = eta[
            (~eta["_matched"])
            & (eta["eta_tax_id"] == fiscal_str.loc[i])
            & (eta["eta_internal_no"] == str(row["voucher_no"]).strip())
        ]
        if candidates.empty:
            continue
        amt_candidates = candidates[(candidates["eta_total"] - row["total_amount"]).abs() <= tolerance]
        chosen = amt_candidates.index[0] if not amt_candidates.empty else candidates.index[0]
        eta.loc[chosen, "_matched"] = True
        tina.loc[i, "_matched"] = True
        tina.loc[i, "matched_eta_index"] = chosen
        tina.loc[i, "match_type"] = (
            "Matched (voucher + amount)" if not amt_candidates.empty
            else "Matched by voucher, amount MISMATCH"
        )

    # --- Pass 2: no voucher recorded -> match on tax id + amount (unique candidate only) ---
    for i, row in tina[(~tina["_matched"]) & tina["has_tax_id"]].iterrows():
        candidates = eta[
            (~eta["_matched"])
            & (eta["eta_tax_id"] == fiscal_str.loc[i])
            & ((eta["eta_total"] - row["total_amount"]).abs() <= tolerance)
        ]
        if len(candidates) == 1:
            chosen = candidates.index[0]
            eta.loc[chosen, "_matched"] = True
            tina.loc[i, "_matched"] = True
            tina.loc[i, "matched_eta_index"] = chosen
            tina.loc[i, "match_type"] = "Matched (amount only, no voucher on file)"
        elif len(candidates) > 1:
            tina.loc[i, "match_type"] = f"Ambiguous - {len(candidates)} possible ETA matches (amount only)"

    matched_mask = tina["_matched"]
    matched = tina[matched_mask].copy()
    # attach ETA-side fields for the matched rows
    for col in ["eta_internal_no", "eta_total", "eta_currency", "eta_seller_name", "eta_doc_type"]:
        matched[col] = matched["matched_eta_index"].map(eta[col])
    matched["amount_diff"] = (matched["eta_total"] - matched["total_amount"]).round(4)

    no_tax_id = tina[~tina["has_tax_id"]].copy()
    unmatched_tina = tina[(~matched_mask) & tina["has_tax_id"]].copy()
    unmatched_eta = eta[~eta["_matched"]].copy()

    return matched, unmatched_tina, unmatched_eta, no_tax_id
