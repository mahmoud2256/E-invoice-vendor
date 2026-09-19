"""
app.py - ETA / TINA Invoice Matcher (Streamlit)
------------------------------------------------
يطابق فواتير الموردين الواردة من بوابة الضرائب (ETA) مع فواتير المبيعات
الصادرة من نظام تينا (TINA)، بالاستعانة بملف بروفايل المورد لربط
Supplier ID الداخلي بالرقم الضريبي.

تشغيل محلي:
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

st.set_page_config(page_title="مطابقة فواتير الموردين - ETA", layout="wide")

st.title("🧾 مطابقة فواتير الموردين (ETA) مع فواتير المبيعات (TINA)")
st.caption(
    "بيقارن فواتير الموردين الواردة من بوابة الضرائب المصرية مع فواتير المبيعات "
    "الصادرة من نظام تينا، بالاستعانة بملف بروفايل المورد."
)

# ---------------------------------------------------------------------------
# Sidebar - inputs
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("📂 الملفات المطلوبة")
    tina_file = st.file_uploader(
        "1️⃣ ملف مبيعات تينا (.xls)",
        type=["xls"],
        help="التقرير المصدر من تينا - All_Processed_Transactions...",
    )
    eta_file = st.file_uploader(
        "2️⃣ ملف بورتال الموردين ETA (.xlsx)",
        type=["xlsx"],
        help="الملف المصدر من بوابة الضرائب - eInvoices_...",
    )
    profile_file = st.file_uploader(
        "3️⃣ ملف بروفايل الموردين (.xls)",
        type=["xls"],
        help="Egypt-Suppliers_List...",
    )

    st.divider()
    tolerance = st.number_input(
        "سماحية الفرق في المبلغ",
        min_value=0.0, max_value=100.0, value=0.01, step=0.01,
        help="أي فرق أقل من أو يساوي القيمة دي يعتبر تطابق كامل.",
    )
    run_button = st.button("🚀 ابدأ المطابقة", type="primary", use_container_width=True)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _load_and_match(tina_bytes, eta_bytes, profile_bytes, tolerance):
    tina_df = read_tina_spreadsheetml(BytesIO(tina_bytes))
    profile_df = read_tina_spreadsheetml(BytesIO(profile_bytes))
    eta_df = read_eta_portal_file(BytesIO(eta_bytes))

    grouped = build_tina_supplier_invoices(tina_df)
    enriched = enrich_with_supplier_profile(grouped, profile_df)
    eta_norm = normalize_eta_invoices(eta_df)

    matched, unmatched_tina, unmatched_eta, no_tax_id = match_invoices(enriched, eta_norm, tolerance)
    return matched, unmatched_tina, unmatched_eta, no_tax_id, len(tina_df), len(eta_df), len(profile_df)


def _to_excel_bytes(sheets: dict) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets.items():
            (df if not df.empty else pd.DataFrame({"info": ["لا توجد بيانات"]})).to_excel(
                writer, sheet_name=name[:31], index=False
            )
    return buf.getvalue()


DISPLAY_COLS_MATCHED = [
    "invoice_no", "supplier_id", "supplier_company", "voucher_no",
    "total_amount", "eta_internal_no", "eta_total", "amount_diff",
    "currency", "taxable_status", "match_type",
]
DISPLAY_COLS_UNMATCHED_TINA = [
    "invoice_no", "supplier_id", "supplier_company", "voucher_no",
    "total_amount", "currency", "fiscal_code", "taxable_status", "match_type",
]
DISPLAY_COLS_UNMATCHED_ETA = [
    "eta_internal_no", "eta_seller_name", "eta_tax_id", "eta_total",
    "eta_currency", "eta_doc_type",
]
DISPLAY_COLS_NO_TAX_ID = [
    "invoice_no", "supplier_id", "voucher_no", "total_amount", "currency", "taxable_status",
]


def _safe_cols(df, cols):
    return [c for c in cols if c in df.columns]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
if run_button:
    if not (tina_file and eta_file and profile_file):
        st.warning("لازم ترفع الثلاث ملفات الأول.")
        st.stop()

    with st.spinner("بتتم المطابقة..."):
        matched, unmatched_tina, unmatched_eta, no_tax_id, n_tina, n_eta, n_profile = _load_and_match(
            tina_file.getvalue(), eta_file.getvalue(), profile_file.getvalue(), tolerance
        )

    st.success(
        f"تم تحميل {n_tina} سطر مبيعات، {n_eta} فاتورة ETA، {n_profile} مورد."
    )

    # ---- summary metrics ----
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("✅ متطابقة", len(matched))
    c2.metric("❌ غير متطابقة (تينا)", len(unmatched_tina))
    c3.metric("❌ غير متطابقة (ETA)", len(unmatched_eta))
    c4.metric("⚠️ بدون رقم ضريبي", len(no_tax_id))

    exact = (matched["match_type"] == "Matched (voucher + amount)").sum()
    mismatch = (matched["match_type"] == "Matched by voucher, amount MISMATCH").sum()
    amount_only = (matched["match_type"] == "Matched (amount only, no voucher on file)").sum()
    st.caption(
        f"من ضمن المتطابقة: {exact} تطابق كامل (رقم + مبلغ)، "
        f"{mismatch} تطابق برقم الفاتورة لكن فيه فرق في المبلغ (محتاج مراجعة)، "
        f"{amount_only} تطابق بالمبلغ فقط (مفيش رقم فاتورة مورد مسجل)."
    )

    tab1, tab2, tab3, tab4 = st.tabs([
        f"✅ متطابقة ({len(matched)})",
        f"❌ غير متطابقة - تينا ({len(unmatched_tina)})",
        f"❌ غير متطابقة - ETA ({len(unmatched_eta)})",
        f"⚠️ بدون رقم ضريبي ({len(no_tax_id)})",
    ])

    with tab1:
        st.dataframe(matched[_safe_cols(matched, DISPLAY_COLS_MATCHED)], use_container_width=True)
    with tab2:
        st.caption("فواتير مسجل لها رقم ضريبي للمورد بس ملقتلها نظير في ملف ETA.")
        st.dataframe(unmatched_tina[_safe_cols(unmatched_tina, DISPLAY_COLS_UNMATCHED_TINA)], use_container_width=True)
    with tab3:
        st.caption("فواتير موجودة في ETA ومالهاش نظير في ملف المبيعات.")
        st.dataframe(unmatched_eta[_safe_cols(unmatched_eta, DISPLAY_COLS_UNMATCHED_ETA)], use_container_width=True)
    with tab4:
        st.caption("موردين ظهروا في ملف المبيعات بس مالهمش رقم ضريبي مسجل في ملف البروفايل - مستحيل مطابقتهم.")
        st.dataframe(no_tax_id[_safe_cols(no_tax_id, DISPLAY_COLS_NO_TAX_ID)], use_container_width=True)

    st.divider()
    excel_bytes = _to_excel_bytes({
        "Matched": matched[_safe_cols(matched, DISPLAY_COLS_MATCHED)],
        "Unmatched_TINA": unmatched_tina[_safe_cols(unmatched_tina, DISPLAY_COLS_UNMATCHED_TINA)],
        "Unmatched_ETA": unmatched_eta[_safe_cols(unmatched_eta, DISPLAY_COLS_UNMATCHED_ETA)],
        "No_Tax_ID": no_tax_id[_safe_cols(no_tax_id, DISPLAY_COLS_NO_TAX_ID)],
    })
    st.download_button(
        "⬇️ تحميل التقرير الكامل (Excel)",
        data=excel_bytes,
        file_name="invoice_matching_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
else:
    st.info("ارفع الثلاث ملفات من القائمة على الشمال واضغط 'ابدأ المطابقة'.")
