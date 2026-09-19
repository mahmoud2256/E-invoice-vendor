"""
report_builder.py
------------------
Combines matched / unmatched / no-tax-id results into ONE Excel sheet
with a Status column and professional color-coded rows, ready to send
to a finance manager.
"""

from io import BytesIO

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

STATUS_LABELS = {
    "exact": "Matched - Exact",
    "mismatch": "Matched - Amount Mismatch",
    "amount_only": "Matched - No Supplier Invoice No. On File",
    "unmatched_sales": "Unmatched - Not Found in Supplier Invoices",
    "unmatched_supplier": "Unmatched - Not Found in Sales Report",
    "no_tax_id": "No Tax ID On File",
}

# fill colors per status (soft, professional palette — each hue kept visually distinct)
STATUS_FILLS = {
    STATUS_LABELS["exact"]: "C6EFCE",             # green
    STATUS_LABELS["mismatch"]: "FFEB9C",          # amber/yellow
    STATUS_LABELS["amount_only"]: "D9D2E9",       # lavender/purple (was too close to white before)
    STATUS_LABELS["unmatched_sales"]: "F8CBAD",   # orange
    STATUS_LABELS["unmatched_supplier"]: "F8CBAD",
    STATUS_LABELS["no_tax_id"]: "D9D9D9",         # gray
}

MATCH_TYPE_TO_STATUS = {
    "Matched (voucher + amount)": STATUS_LABELS["exact"],
    "Matched by voucher, amount MISMATCH": STATUS_LABELS["mismatch"],
    "Matched (amount only, no voucher on file)": STATUS_LABELS["amount_only"],
}

COLUMN_ORDER = [
    "Invoice No.",
    "Supplier ID",
    "Supplier Name",
    "Supplier Invoice No.",
    "Sales Amount",
    "Supplier Invoice Amount",
    "Amount Difference",
    "Currency",
    "Taxable Status",
    "Status",
]


def build_combined_dataframe(matched, unmatched_sales, unmatched_supplier, no_tax_id) -> pd.DataFrame:
    rows = []

    for _, r in matched.iterrows():
        rows.append({
            "Invoice No.": r.get("invoice_no"),
            "Supplier ID": r.get("supplier_id"),
            "Supplier Name": r.get("supplier_company"),
            "Supplier Invoice No.": r.get("voucher_no") or r.get("eta_internal_no"),
            "Sales Amount": r.get("total_amount"),
            "Supplier Invoice Amount": r.get("eta_total"),
            "Amount Difference": r.get("amount_diff"),
            "Currency": r.get("currency"),
            "Taxable Status": r.get("taxable_status"),
            "Status": MATCH_TYPE_TO_STATUS.get(r.get("match_type"), r.get("match_type")),
        })

    for _, r in unmatched_sales.iterrows():
        rows.append({
            "Invoice No.": r.get("invoice_no"),
            "Supplier ID": r.get("supplier_id"),
            "Supplier Name": r.get("supplier_company"),
            "Supplier Invoice No.": r.get("voucher_no"),
            "Sales Amount": r.get("total_amount"),
            "Supplier Invoice Amount": None,
            "Amount Difference": None,
            "Currency": r.get("currency"),
            "Taxable Status": r.get("taxable_status"),
            "Status": STATUS_LABELS["unmatched_sales"],
        })

    for _, r in unmatched_supplier.iterrows():
        rows.append({
            "Invoice No.": None,
            "Supplier ID": None,
            "Supplier Name": r.get("eta_seller_name"),
            "Supplier Invoice No.": r.get("eta_internal_no"),
            "Sales Amount": None,
            "Supplier Invoice Amount": r.get("eta_total"),
            "Amount Difference": None,
            "Currency": r.get("eta_currency"),
            "Taxable Status": None,
            "Status": STATUS_LABELS["unmatched_supplier"],
        })

    for _, r in no_tax_id.iterrows():
        rows.append({
            "Invoice No.": r.get("invoice_no"),
            "Supplier ID": r.get("supplier_id"),
            "Supplier Name": None,
            "Supplier Invoice No.": r.get("voucher_no"),
            "Sales Amount": r.get("total_amount"),
            "Supplier Invoice Amount": None,
            "Amount Difference": None,
            "Currency": r.get("currency"),
            "Taxable Status": r.get("taxable_status"),
            "Status": STATUS_LABELS["no_tax_id"],
        })

    df = pd.DataFrame(rows, columns=COLUMN_ORDER)
    return df


def write_colored_excel(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Reconciliation", index=False)
        ws = writer.sheets["Reconciliation"]

        # header styling
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for col_idx, _ in enumerate(df.columns, start=1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        status_col_idx = df.columns.get_loc("Status") + 1

        for row_idx in range(2, ws.max_row + 1):
            status_val = ws.cell(row=row_idx, column=status_col_idx).value
            color = STATUS_FILLS.get(status_val)
            if color:
                fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                for col_idx in range(1, len(df.columns) + 1):
                    ws.cell(row=row_idx, column=col_idx).fill = fill

        # auto column width
        for col_idx, col_name in enumerate(df.columns, start=1):
            max_len = max([len(str(col_name))] + [len(str(v)) for v in df.iloc[:, col_idx - 1].fillna("")])
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 40)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    return buf.getvalue()
