"""
app.py - Supplier Invoice Reconciliation
------------------------------------------
Matches supplier e-invoices from the Egyptian Tax Authority (ETA) portal
against your own Sales Report and a Supplier Profile file.

Run locally:
    streamlit run app.py
"""

from io import BytesIO

import pandas as pd
import streamlit as st

from matcher_core import (
    read_tina_spreadsheetml,
    read_eta_portal_file,
    build_tina_supplier_invoices,
    enrich_with_supplier_profile,
    normalize_eta_invoices,
    match_invoices,
)
from report_builder import build_combined_dataframe, write_colored_excel
import auth

st.set_page_config(page_title="Supplier Invoice Reconciliation", layout="wide", page_icon="📋")

# ---------------------------------------------------------------------------
# Global styling
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    :root {
        --navy: #0B2545;
        --navy-light: #13315C;
        --teal: #2EC4B6;
        --teal-dark: #1A9E92;
        --bg-soft: #F4F6F9;
        --green: #1E8E5A;
        --amber: #D98E04;
        --orange: #D9531E;
        --gray: #7A8699;
    }

    .stApp { background-color: var(--bg-soft); }

    /* Hero banner */
    .hero-banner {
        background: linear-gradient(120deg, var(--navy) 0%, var(--navy-light) 55%, var(--teal-dark) 100%);
        padding: 28px 32px;
        border-radius: 14px;
        color: #FFFFFF;
        margin-bottom: 22px;
        box-shadow: 0 6px 18px rgba(11, 37, 69, 0.18);
    }
    .hero-banner h1 { color: #FFFFFF; margin: 0; font-size: 28px; }
    .hero-banner p { color: #D7E3F0; margin: 6px 0 0 0; font-size: 14px; }

    /* Buttons */
    .stButton > button, .stDownloadButton > button {
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.15s ease-in-out;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 10px rgba(0,0,0,0.12);
    }

    /* Metric cards */
    .metric-card {
        background: #FFFFFF;
        border-radius: 12px;
        padding: 16px 18px;
        box-shadow: 0 2px 8px rgba(20, 30, 60, 0.08);
        border-top: 4px solid var(--gray);
        text-align: center;
    }
    .metric-card .metric-value { font-size: 30px; font-weight: 700; color: var(--navy); }
    .metric-card .metric-label { font-size: 13px; color: #5B6472; margin-top: 4px; }
    .metric-card.green { border-top-color: var(--green); }
    .metric-card.orange { border-top-color: var(--orange); }
    .metric-card.amber { border-top-color: var(--amber); }
    .metric-card.gray { border-top-color: var(--gray); }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E4E8EF;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)


def metric_card(label, value, color="gray"):
    st.markdown(
        f"""<div class="metric-card {color}">
                <div class="metric-value">{value}</div>
                <div class="metric-label">{label}</div>
            </div>""",
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Login gate
# ---------------------------------------------------------------------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = None

if not st.session_state.logged_in:
    st.markdown(
        """<div class="hero-banner" style="text-align:center; margin-top:40px;">
                <h1>📋 Supplier Invoice Reconciliation</h1>
                <p>Sign in to continue</p>
            </div>""",
        unsafe_allow_html=True,
    )
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Sign in", use_container_width=True, type="primary")
        if submitted:
            if auth.verify_login(username, password):
                st.session_state.logged_in = True
                st.session_state.username = username
                st.rerun()
            else:
                st.error("Invalid username or password.")
    st.stop()

# ---------------------------------------------------------------------------
# Header / logout
# ---------------------------------------------------------------------------
top_left, top_right = st.columns([5, 1])
with top_left:
    st.markdown(
        """<div class="hero-banner">
                <h1>📋 Supplier Invoice Reconciliation</h1>
                <p>Matches supplier e-invoices from the Egyptian Tax Authority (ETA) portal against your
                Sales Report, using the Supplier Profile file to link them.</p>
            </div>""",
        unsafe_allow_html=True,
    )
with top_right:
    st.write("")
    st.write("")
    st.write(f"👤 **{st.session_state.username}**")
    if st.button("Log out"):
        st.session_state.logged_in = False
        st.session_state.username = None
        st.rerun()

# ---------------------------------------------------------------------------
# Sidebar - inputs
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Required Files")
    sales_file = st.file_uploader(
        "1. Sales Report (.xls)",
        type=["xls"],
        help="Your exported daily/period sales report.",
    )
    eta_file = st.file_uploader(
        "2. Supplier ETA Portal File (.xlsx)",
        type=["xlsx"],
        help="Invoice export from the Egyptian Tax Authority portal.",
    )
    profile_file = st.file_uploader(
        "3. Supplier Profile File (.xls)",
        type=["xls"],
        help="Supplier list with Tax ID and taxable status.",
    )

    st.divider()
    tolerance = st.number_input(
        "Amount matching tolerance",
        min_value=0.0, max_value=100.0, value=0.01, step=0.01,
        help="Amount differences at or below this value are treated as an exact match.",
    )
    run_button = st.button("Run Reconciliation", type="primary", use_container_width=True)

    if auth.is_admin(st.session_state.username):
        st.divider()
        with st.expander("👥 Manage Users"):
            st.caption("Users added here reset if the app is redeployed, unless saved permanently.")
            for u in auth.list_usernames():
                c1, c2 = st.columns([3, 1])
                c1.write(u)
                if u != "Mahmoud" and c2.button("Remove", key=f"rm_{u}"):
                    auth.delete_user(u)
                    st.rerun()
            with st.form("add_user_form", clear_on_submit=True):
                new_user = st.text_input("New username")
                new_pass = st.text_input("New password", type="password")
                new_admin = st.checkbox("Admin access")
                add_submitted = st.form_submit_button("Add user")
            if add_submitted and new_user and new_pass:
                if auth.add_user(new_user, new_pass, new_admin):
                    st.success(f"User '{new_user}' added.")
                    st.rerun()
                else:
                    st.error("That username already exists.")

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _load_and_match(sales_bytes, eta_bytes, profile_bytes, tolerance):
    sales_df = read_tina_spreadsheetml(BytesIO(sales_bytes))
    profile_df = read_tina_spreadsheetml(BytesIO(profile_bytes))
    eta_df = read_eta_portal_file(BytesIO(eta_bytes))

    grouped = build_tina_supplier_invoices(sales_df)
    enriched = enrich_with_supplier_profile(grouped, profile_df)
    eta_norm = normalize_eta_invoices(eta_df)

    matched, unmatched_sales, unmatched_supplier, no_tax_id = match_invoices(enriched, eta_norm, tolerance)
    return matched, unmatched_sales, unmatched_supplier, no_tax_id, len(sales_df), len(eta_df), len(profile_df)


DISPLAY_COLS_MATCHED = [
    "invoice_no", "supplier_id", "supplier_company", "voucher_no",
    "total_amount", "eta_internal_no", "eta_total", "amount_diff",
    "currency", "taxable_status", "match_type",
]
DISPLAY_COLS_UNMATCHED_SALES = [
    "invoice_no", "supplier_id", "supplier_company", "voucher_no",
    "total_amount", "currency", "fiscal_code", "taxable_status",
]
DISPLAY_COLS_UNMATCHED_SUPPLIER = [
    "eta_internal_no", "eta_seller_name", "eta_tax_id", "eta_total",
    "eta_currency", "eta_doc_type",
]
DISPLAY_COLS_NO_TAX_ID = [
    "invoice_no", "supplier_id", "voucher_no", "total_amount", "currency", "taxable_status",
]

RENAME_FOR_DISPLAY = {
    "invoice_no": "Invoice No.",
    "supplier_id": "Supplier ID",
    "supplier_company": "Supplier Name",
    "voucher_no": "Supplier Invoice No.",
    "total_amount": "Sales Amount",
    "eta_internal_no": "Supplier Invoice No. (ETA)",
    "eta_total": "Supplier Invoice Amount",
    "amount_diff": "Amount Difference",
    "currency": "Currency",
    "taxable_status": "Taxable Status",
    "match_type": "Match Type",
    "fiscal_code": "Tax ID",
    "eta_seller_name": "Supplier Name",
    "eta_tax_id": "Tax ID",
    "eta_currency": "Currency",
    "eta_doc_type": "Document Type",
}


def _safe_cols(df, cols):
    return [c for c in cols if c in df.columns]


def _prep_display(df, cols):
    out = df[_safe_cols(df, cols)].rename(columns=RENAME_FOR_DISPLAY)
    return out


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
if run_button:
    if not (sales_file and eta_file and profile_file):
        st.warning("Please upload all three files first.")
        st.stop()

    with st.spinner("Running reconciliation..."):
        matched, unmatched_sales, unmatched_supplier, no_tax_id, n_sales, n_eta, n_profile = _load_and_match(
            sales_file.getvalue(), eta_file.getvalue(), profile_file.getvalue(), tolerance
        )

    st.success(f"Loaded {n_sales} sales lines, {n_eta} ETA invoices, {n_profile} suppliers.")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("Matched", len(matched), "green")
    with c2:
        metric_card("Unmatched (Sales)", len(unmatched_sales), "orange")
    with c3:
        metric_card("Unmatched (Supplier)", len(unmatched_supplier), "orange")
    with c4:
        metric_card("No Tax ID on File", len(no_tax_id), "amber")
    st.write("")

    exact = (matched["match_type"] == "Matched (voucher + amount)").sum()
    mismatch = (matched["match_type"] == "Matched by voucher, amount MISMATCH").sum()
    amount_only = (matched["match_type"] == "Matched (amount only, no voucher on file)").sum()
    st.caption(
        f"Of the matched invoices: {exact} matched exactly (number + amount), "
        f"{mismatch} matched by invoice number but with a different amount (needs review), "
        f"{amount_only} matched by amount only (no supplier invoice number was recorded)."
    )

    tab1, tab2, tab3, tab4 = st.tabs([
        f"✅ Matched ({len(matched)})",
        f"❌ Unmatched - Sales ({len(unmatched_sales)})",
        f"❌ Unmatched - Supplier ({len(unmatched_supplier)})",
        f"⚠️ No Tax ID ({len(no_tax_id)})",
    ])

    with tab1:
        st.dataframe(_prep_display(matched, DISPLAY_COLS_MATCHED), use_container_width=True)
    with tab2:
        st.caption("Sales invoices with a supplier tax ID on file, but no matching ETA invoice was found.")
        st.dataframe(_prep_display(unmatched_sales, DISPLAY_COLS_UNMATCHED_SALES), use_container_width=True)
    with tab3:
        st.caption("Supplier invoices found in ETA with no matching entry in the Sales Report.")
        st.dataframe(_prep_display(unmatched_supplier, DISPLAY_COLS_UNMATCHED_SUPPLIER), use_container_width=True)
    with tab4:
        st.caption("Suppliers found in the Sales Report with no Tax ID on file - these cannot be matched.")
        st.dataframe(_prep_display(no_tax_id, DISPLAY_COLS_NO_TAX_ID), use_container_width=True)

    st.divider()
    combined_df = build_combined_dataframe(matched, unmatched_sales, unmatched_supplier, no_tax_id)
    excel_bytes = write_colored_excel(combined_df)
    st.download_button(
        "⬇️ Download Full Report (Excel)",
        data=excel_bytes,
        file_name="invoice_reconciliation_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
else:
    st.info("Upload the three files from the sidebar, then click 'Run Reconciliation'.")
