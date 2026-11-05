import os
import sys
from pathlib import Path

# Ensure project root is in sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import re
import json
import base64
import pandas as pd
import streamlit as st
from datetime import datetime
from typing import List, Dict, Any, Optional

from core.config import (
    EXCEL_TEMPLATE_PATH,
    EXCEL_SHEET_NAME,
    INVOICE_SAMPLES_DIR,
    GEMINI_PRIMARY_MODEL,
    OUTPUT_DIR
)
from core.pipeline import InvoiceProcessingPipeline
from core.outlook_service import MicrosoftGraphOutlookService
from core.validator import InvoiceValidator
from core.models import DynamicInvoiceRecord, ExtractedField, FinancialAudit

st.set_page_config(
    page_title="Gemini Invoice Document Intelligence",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.1rem;
        font-weight: 700;
        color: #1A73E8;
        margin-bottom: 0.1rem;
    }
    .sub-header {
        font-size: 1.0rem;
        color: #5F6368;
        margin-bottom: 1.2rem;
    }
    .status-approved {
        background-color: #E6F4EA;
        color: #137333;
        padding: 4px 10px;
        border-radius: 4px;
        font-weight: 600;
    }
    .status-pending {
        background-color: #FEF7E0;
        color: #B06000;
        padding: 4px 10px;
        border-radius: 4px;
        font-weight: 600;
    }
    .status-rejected {
        background-color: #FCE8E6;
        color: #C5221F;
        padding: 4px 10px;
        border-radius: 4px;
        font-weight: 600;
    }
    .badge-verified {
        color: #137333;
        font-weight: 600;
    }
    .badge-mismatch {
        color: #C5221F;
        font-weight: 600;
    }
    .badge-incomplete {
        color: #B06000;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_resource
def get_pipeline():
    return InvoiceProcessingPipeline()

@st.cache_resource
def get_outlook_service():
    return MicrosoftGraphOutlookService()

pipeline = get_pipeline()
outlook_service = get_outlook_service()

# --- Sidebar Configuration ---
st.sidebar.image("https://www.gstatic.com/lamda/images/gemini_sparkle_v002_d4735304ff6292a690345.svg", width=36)
st.sidebar.title("Automation Controls")

model_choice = st.sidebar.selectbox(
    "Active Gemini Model",
    options=["gemini-3.7-flash", "gemini-2.5-flash", "gemini-2.5-flash-lite"],
    index=0,
    help="Primary document intelligence model with automatic cascade"
)

skip_processed = st.sidebar.checkbox(
    "Skip previously verified files (SHA-256)",
    value=True,
    help="Guarantees idempotency and avoids unnecessary API credits"
)

st.sidebar.markdown("---")
st.sidebar.subheader("Template & Storage")
st.sidebar.caption(f"**Excel Template:** `{Path(EXCEL_TEMPLATE_PATH).name}`")
st.sidebar.caption(f"**Worksheet:** `{EXCEL_SHEET_NAME}`")
st.sidebar.caption(f"**Ledger Status:** Cumulative Ledger in `output/`")

with st.sidebar.expander("Discovered Template Columns", expanded=False):
    for c in pipeline.get_template_columns():
        st.markdown(f"- **{c.header_name}** (`{c.inferred_type}`)")

# --- Main Layout ---
st.markdown('<div class="main-header">🧾 Gemini Invoice Document Intelligence</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Dynamic template schema, Decimal financial cross-footing, review workflow & cumulative ledger</div>', unsafe_allow_html=True)

tab_upload, tab_review, tab_ledger, tab_history, tab_outlook = st.tabs([
    "📥 Ingestion & Processing",
    "🧐 Invoice Review & Approval",
    "📊 Cumulative Excel Ledger",
    "🗄️ SQLite Audit Trail",
    "📬 Outlook Ingestion (MS Graph)"
])

# Helper to safely retrieve field values across possible column names
def get_val(rec, *field_names, default=""):
    """Safely retrieves field value across multiple potential column naming variants."""
    if not rec or not getattr(rec, "fields", None):
        return default
    for name in field_names:
        f = rec.fields.get(name)
        if f is not None and getattr(f, "value", None) is not None and str(f.value).strip() != "":
            return f.value
    target_keys = [re.sub(r'[^a-z0-9]', '', n.lower()) for n in field_names]
    for k, f in rec.fields.items():
        k_canon = re.sub(r'[^a-z0-9]', '', k.lower())
        if k_canon in target_keys and f is not None and getattr(f, "value", None) is not None:
            return f.value
    return default

# Helper to load DB records
def load_db_records(limit=200):
    raw_records = pipeline.db.get_all_records(limit=limit)
    cols = pipeline.get_template_columns()
    records = []
    for r in raw_records:
        rec = None
        # Priority 1: If reviewer modified fields_data in SQLite, preserve those edited values
        if r.get("fields_data"):
            try:
                f_dict = json.loads(r["fields_data"]) if isinstance(r["fields_data"], str) else r["fields_data"]
                fields = {}
                for k, v in f_dict.items():
                    fields[k] = ExtractedField(**v) if isinstance(v, dict) else ExtractedField(value=v)
                
                raw_json = json.loads(r["raw_response"]) if r.get("raw_response") else {}
                _, raw_audit = InvoiceValidator.unwrap_raw_response(raw_json)
                if r.get("taxable_value") is not None:
                    raw_audit["taxable_value"] = r["taxable_value"]
                if r.get("total_amount") is not None:
                    raw_audit["inv_amount"] = r["total_amount"]
                fin_audit = InvoiceValidator.validate_financials(raw_audit=raw_audit, fields=fields)
                
                v_flags = []
                if r.get("validation_flags"):
                    try:
                        v_flags = json.loads(r["validation_flags"]) if isinstance(r["validation_flags"], str) else r["validation_flags"]
                    except Exception:
                        v_flags = []

                rec = DynamicInvoiceRecord(
                    id=r["id"],
                    file_name=r["file_name"],
                    file_path=r.get("relative_path"),
                    relative_path=r.get("relative_path") or r["file_name"],
                    fields=fields,
                    financial_audit=fin_audit,
                    validation_flags=v_flags,
                    raw_response=raw_json
                )
            except Exception:
                rec = None

        if rec is None and r.get("raw_response"):
            try:
                raw_json = json.loads(r["raw_response"]) if isinstance(r["raw_response"], str) else r["raw_response"]
                rec = pipeline.validator.validate_and_reconcile(
                    raw_json=raw_json,
                    columns=cols,
                    file_name=r["file_name"],
                    relative_path=r.get("relative_path"),
                    business_rules=pipeline.business_rules
                )
            except Exception:
                rec = None

        if rec:
            rec.id = r["id"]
            rec.file_hash = r.get("file_hash")
            rec.processing_status = r.get("processing_status", "SUCCESS")
            rec.review_status = r.get("review_status", "PENDING")
            rec.export_status = r.get("export_status", "NOT_EXPORTED")
            rec.is_file_duplicate = bool(r.get("is_file_duplicate", 0))
            rec.is_invoice_duplicate = bool(r.get("is_invoice_duplicate", 0) or r.get("is_duplicate", 0))
            rec.duplicate_reason = r.get("duplicate_reason")
            rec.error_message = r.get("error_message")
            if rec.processing_status == "FAILED":
                rec.financial_audit.status = "INCOMPLETE"
            elif r.get("financial_status"):
                rec.financial_audit.status = r["financial_status"]
            records.append(rec)
    return records

# --- TAB 1: INGESTION & PROCESSING ---
with tab_upload:
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("Upload Invoice Documents")
        uploaded_files = st.file_uploader(
            "Upload one or more invoice files (PDF, PNG, JPG, JPEG)",
            type=["pdf", "png", "jpg", "jpeg"],
            accept_multiple_files=True
        )
        folder_ctx = st.text_input("Batch Context / Department", value="Direct Upload")

        if uploaded_files and st.button(f"Process {len(uploaded_files)} Document(s)", type="primary"):
            upload_dir = Path("uploads")
            upload_dir.mkdir(exist_ok=True)
            prog = st.progress(0)
            status = st.empty()

            upload_results = []
            for i, uf in enumerate(uploaded_files):
                status.text(f"Processing ({i+1}/{len(uploaded_files)}): {uf.name}")
                dest_path = upload_dir / uf.name
                with open(dest_path, "wb") as f:
                    f.write(uf.getbuffer())

                u_rec = pipeline.process_single_invoice(
                    file_path=dest_path,
                    relative_path=uf.name,
                    folder_context=folder_ctx,
                    model_name=model_choice,
                    skip_if_processed=skip_processed,
                    auto_save_excel=False
                )
                upload_results.append(u_rec)
                prog.progress((i + 1) / len(uploaded_files))

            status.empty()
            prog.empty()
            failed_uploads = [r for r in upload_results if r.processing_status == "FAILED"]
            if failed_uploads:
                st.error(f"⚠️ {len(failed_uploads)} of {len(upload_results)} uploaded document(s) failed extraction.")
                with st.expander("View Failure Details", expanded=True):
                    for fr in failed_uploads:
                        st.markdown(f"**{fr.file_name}**: `{fr.error_message}`")
                    if any("429" in str(fr.error_message or "") or "RESOURCE_EXHAUSTED" in str(fr.error_message or "") for fr in failed_uploads):
                        st.warning("⏳ **Google Gemini API Rate Limit / Quota Exceeded (HTTP 429):** Free tier limit (20 requests/day per model) reached. Check your Gemini API billing plan or configure a new key in `.env`.")
            else:
                st.success("Processing complete! Go to the '🧐 Invoice Review & Approval' tab to review extracted values.")

    with col2:
        st.subheader("Recursive Directory Batch Processing")
        st.write(f"Source Folder: `{INVOICE_SAMPLES_DIR}`")
        recursive_scan = st.checkbox("Scan subdirectories recursively", value=True)

        if Path(INVOICE_SAMPLES_DIR).exists():
            discovered = pipeline.scan_directory_recursive(INVOICE_SAMPLES_DIR)
            st.info(f"Discovered **{len(discovered)}** supported invoice documents.")
            batch_count = st.slider("Documents to process in this run", 1, max(1, len(discovered)), min(3, len(discovered)))

            if st.button("Run Batch Extraction Pipeline", type="primary"):
                prog = st.progress(0)
                status = st.empty()

                target_batch = discovered[:batch_count]
                batch_results = []
                for idx, (abs_path, rel_path) in enumerate(target_batch):
                    status.text(f"Processing ({idx+1}/{len(target_batch)}): {rel_path}")
                    b_rec = pipeline.process_single_invoice(
                        file_path=abs_path,
                        relative_path=rel_path,
                        folder_context=Path(rel_path).parent.as_posix(),
                        model_name=model_choice,
                        skip_if_processed=skip_processed,
                        auto_save_excel=False
                    )
                    batch_results.append(b_rec)
                    prog.progress((idx + 1) / len(target_batch))

                status.empty()
                prog.empty()
                failed_batches = [r for r in batch_results if r.processing_status == "FAILED"]
                if failed_batches:
                    st.error(f"⚠️ {len(failed_batches)} of {len(batch_results)} document(s) failed extraction.")
                    with st.expander("View Failure Details", expanded=True):
                        for fr in failed_batches:
                            st.markdown(f"**{fr.file_name}**: `{fr.error_message}`")
                        if any("429" in str(fr.error_message or "") or "RESOURCE_EXHAUSTED" in str(fr.error_message or "") for fr in failed_batches):
                            st.warning("⏳ **Google Gemini API Rate Limit / Quota Exceeded (HTTP 429):** Free tier limit (20 requests/day per model) reached. Check your Gemini API billing plan or configure a new key in `.env`.")
                else:
                    st.success("Batch completed! Records logged to SQLite. Please review in Tab 2.")

# --- TAB 2: INVOICE REVIEW & APPROVAL WORKFLOW ---
with tab_review:
    # 1. Dashboard Overview Metrics
    stats = pipeline.db.get_stats()
    st.subheader("Invoice Review, Verification & Approval Dashboard")
    
    row1_c1, row1_c2, row1_c3, row1_c4, row1_c5 = st.columns(5)
    row1_c1.metric("Total Processed", stats.get("total_processed", 0))
    row1_c2.metric("Pending Review", stats.get("pending_review", 0))
    row1_c3.metric("Approved Invoices", stats.get("approved_count", 0))
    row1_c4.metric("Rejected Invoices", stats.get("rejected_count", 0))
    row1_c5.metric("Failed Processing", stats.get("failed_count", 0))

    row2_c1, row2_c2, row2_c3, row2_c4, row2_c5 = st.columns(5)
    row2_c1.metric("Financially Verified", stats.get("financially_verified", 0))
    row2_c2.metric("Incomplete Tax / Subtotal", stats.get("financial_incomplete", 0))
    row2_c3.metric("Financial Mismatches", stats.get("financial_mismatch", 0))
    row2_c4.metric("Awaiting Excel Export", stats.get("awaiting_export", 0))
    row2_c5.metric("Successfully Exported", stats.get("exported_count", 0))

    st.markdown("---")

    records = load_db_records(limit=300)
    if not records:
        st.info("No invoice records in database yet. Process documents in Tab 1.")
    else:
        # Search & Filter Controls
        st.markdown("### Search & Filter Invoices")
        f_row1_c1, f_row1_c2, f_row1_c3, f_row1_c4 = st.columns([2, 1, 1, 1])
        with f_row1_c1:
            search_query = st.text_input("🔍 Search Vendor or Invoice #", placeholder="e.g. Shaligram, 071, Pinkey...").strip().lower()
        with f_row1_c2:
            status_filter = st.selectbox("Review Status", ["All", "PENDING", "APPROVED", "REJECTED"], index=0)
        with f_row1_c3:
            math_filter = st.selectbox("Financial Status", ["All", "VERIFIED", "INCOMPLETE", "MISMATCH"], index=0)
        with f_row1_c4:
            export_filter = st.selectbox("Export Status", ["All", "NOT_EXPORTED", "EXPORTED"], index=0)

        f_row2_c1, f_row2_c2 = st.columns([2, 2])
        with f_row2_c1:
            sort_by = st.selectbox("Sort By", ["ID", "Vendor Name", "Invoice Date", "Total Amount"], index=0)
        with f_row2_c2:
            sort_order = st.selectbox("Sort Order", ["Descending", "Ascending"], index=0)

        # Apply filtering
        filtered_records = records
        if search_query:
            def matches_search(r):
                v_name = str(get_val(r, "Vendor name ", "Vendor name", default="")).lower()
                inv_num = str(get_val(r, "Inv#", "Inv No", default="")).lower()
                f_name = (r.file_name or "").lower()
                rel_p = (r.relative_path or "").lower()
                return search_query in v_name or search_query in inv_num or search_query in f_name or search_query in rel_p
            filtered_records = [r for r in filtered_records if matches_search(r)]

        if status_filter != "All":
            filtered_records = [r for r in filtered_records if r.review_status == status_filter]
        if math_filter != "All":
            filtered_records = [r for r in filtered_records if r.financial_audit.status == math_filter]
        if export_filter != "All":
            if export_filter == "NOT_EXPORTED":
                filtered_records = [r for r in filtered_records if r.export_status != "EXPORTED"]
            else:
                filtered_records = [r for r in filtered_records if r.export_status == "EXPORTED"]

        # Apply sorting
        reverse_order = (sort_order == "Descending")
        if sort_by == "Vendor Name":
            filtered_records.sort(key=lambda r: str(get_val(r, "Vendor name ", "Vendor name", default="")).lower(), reverse=reverse_order)
        elif sort_by == "Invoice Date":
            filtered_records.sort(key=lambda r: str(get_val(r, "Inv Date", default="")), reverse=reverse_order)
        elif sort_by == "Total Amount":
            filtered_records.sort(key=lambda r: float(r.financial_audit.inv_amount or 0.0), reverse=reverse_order)
        else:
            filtered_records.sort(key=lambda r: r.id or 0, reverse=reverse_order)

        st.caption(f"Displaying **{len(filtered_records)}** of **{len(records)}** total invoice(s)")

        # Compact Table Overview
        table_rows = []
        for r in filtered_records:
            v_val = get_val(r, "Vendor name ", "Vendor name", default="Unknown")
            inv_val = get_val(r, "Inv#", "Inv No", default="N/A")
            dt_val = get_val(r, "Inv Date", default="—")
            tax_val = f"₹ {r.financial_audit.taxable_value:,.2f}" if r.financial_audit.taxable_value is not None else "—"
            tot_val = f"₹ {r.financial_audit.inv_amount:,.2f}" if r.financial_audit.inv_amount is not None else "—"
            
            missing_cols = [k for k, f in r.fields.items() if f.value is None or str(f.value).strip() == ""]
            missing_str = ", ".join(missing_cols[:3]) + (f" (+{len(missing_cols)-3} more)" if len(missing_cols) > 3 else "") if missing_cols else "None"
            
            dup_tag = "No"
            if r.is_file_duplicate and r.is_invoice_duplicate:
                dup_tag = "File & Inv Duplicate"
            elif r.is_file_duplicate:
                dup_tag = "File Duplicate"
            elif r.is_invoice_duplicate:
                dup_tag = "Invoice Duplicate"

            table_rows.append({
                "ID": r.id,
                "Vendor": v_val,
                "Invoice #": inv_val,
                "Invoice Date": dt_val,
                "Taxable Value": tax_val,
                "Invoice Amount": tot_val,
                "Processing": r.processing_status,
                "Financial Status": r.financial_audit.status,
                "Review Status": r.review_status,
                "Export Status": r.export_status,
                "Duplicate": dup_tag,
                "Missing Fields": missing_str
            })

        if table_rows:
            st.dataframe(pd.DataFrame(table_rows), width="stretch")

        # Batch Actions Section
        with st.container(border=True):
            st.markdown("#### ⚡ Batch Review & Export Actions")
            batch_col1, batch_col2 = st.columns([3, 1])
            with batch_col1:
                batch_selection = st.multiselect(
                    "Select Invoices for Batch Actions",
                    options=[r.id for r in filtered_records],
                    format_func=lambda x: f"Record #{x} - {get_val(next((r for r in filtered_records if r.id == x), None), 'Vendor name ', 'Vendor name', default='Unknown')} (Inv# {get_val(next((r for r in filtered_records if r.id == x), None), 'Inv#', 'Inv No', default='N/A')})",
                    help="Select multiple invoices to approve or reject together"
                )
            with batch_col2:
                b_reviewer = st.text_input("Batch Reviewer ID", value="AP Reviewer", key="batch_rev_id")

            if batch_selection:
                selected_objs = [r for r in filtered_records if r.id in batch_selection]
                num_verified = sum(1 for r in selected_objs if r.financial_audit.status == "VERIFIED")
                num_incomplete = sum(1 for r in selected_objs if r.financial_audit.status == "INCOMPLETE")
                num_mismatch = sum(1 for r in selected_objs if r.financial_audit.status == "MISMATCH")
                
                st.info(f"Selected **{len(batch_selection)}** invoices: **{num_verified}** Verified, **{num_incomplete}** Incomplete, **{num_mismatch}** Mismatch.")
                
                ack_batch = True
                if num_mismatch > 0:
                    st.warning("⚠️ Warning: One or more selected invoices have a financial MISMATCH. Explicit acknowledgement is required to approve.")
                    ack_batch = st.checkbox("I acknowledge the financial warnings/mismatches for the selected batch", value=False, key="ack_batch_chk")

                b_act1, b_act2, b_act3 = st.columns([1, 1, 2])
                with b_act1:
                    if st.button("✅ Approve Selected", type="primary", disabled=(num_mismatch > 0 and not ack_batch), key="btn_b_approve"):
                        res = pipeline.db.batch_approve(
                            record_ids=batch_selection,
                            reviewer_id=b_reviewer,
                            acknowledge_warnings=ack_batch
                        )
                        st.success(f"Batch approved {len(res['approved'])} invoices!")
                        st.rerun()
                with b_act2:
                    if st.button("❌ Reject Selected", key="btn_b_reject"):
                        res = pipeline.db.batch_reject(
                            record_ids=batch_selection,
                            reviewer_id=b_reviewer,
                            reason="Batch rejected by reviewer"
                        )
                        st.success(f"Batch rejected {len(res['rejected'])} invoices!")
                        st.rerun()
                with b_act3:
                    if st.button("🔄 Export Approved Invoices to Cumulative Ledger", key="btn_b_export"):
                        try:
                            l_path, count = pipeline.export_approved_to_excel()
                            st.success(f"Appended {count} approved rows to: {Path(l_path).name}")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Export error: {e}")

        # 4. Detailed Invoice Review & Edit Drawer
        st.markdown("---")
        st.markdown("### 🔍 Detailed Invoice Inspection & Review")
        
        detail_id = st.selectbox(
            "Select Invoice to Review & Edit",
            options=[r.id for r in filtered_records],
            format_func=lambda x: f"Record #{x} | {get_val(next((r for r in filtered_records if r.id == x), None), 'Vendor name ', 'Vendor name', default='Unknown')} | Inv# {get_val(next((r for r in filtered_records if r.id == x), None), 'Inv#', 'Inv No', default='N/A')} [{next((r.review_status for r in filtered_records if r.id == x), 'PENDING')}]",
            index=0
        )
        
        rec = next((r for r in filtered_records if r.id == detail_id), None)
        if rec:
            if rec.processing_status == "FAILED" or rec.error_message:
                st.error(f"🚨 **Processing / Extraction Failure for this Document:**\n\n`{rec.error_message}`")
                if "429" in str(rec.error_message or "") or "RESOURCE_EXHAUSTED" in str(rec.error_message or ""):
                    st.warning("⏳ **Google Gemini API Free Tier Quota Exceeded (HTTP 429):** The daily request limit on the free tier (20 requests/day) has been reached. Please check your Gemini API plan, configure a paid key or new project API key in `.env`, or wait for the quota reset period.")

            col_doc, col_details = st.columns([1, 1])

            # LEFT: Original Document Preview
            with col_doc:
                st.markdown(f"#### 📄 Original Document: `{rec.file_name}`")
                st.caption(f"**Relative Path:** `{rec.relative_path or rec.file_name}`")
                st.caption(f"**SHA-256:** `{rec.file_hash[:20]}...`" if rec.file_hash else "")
                
                if rec.is_file_duplicate:
                    st.info(f"ℹ️ {rec.duplicate_reason or 'Identical document SHA-256 previously scanned'}")
                if rec.is_invoice_duplicate:
                    st.warning(f"🚨 {rec.duplicate_reason or 'Matching vendor and invoice number already exists'}")

                # Locate document file safely
                doc_path = None
                p_cand1 = (Path(INVOICE_SAMPLES_DIR) / (rec.relative_path or rec.file_name)).resolve()
                p_cand2 = (PROJECT_ROOT / "uploads" / (rec.relative_path or rec.file_name)).resolve()
                p_cand3 = (Path(INVOICE_SAMPLES_DIR) / Path(rec.file_name).name).resolve()
                p_cand4 = (PROJECT_ROOT / (rec.relative_path or rec.file_name)).resolve()
                for p in [p_cand1, p_cand2, p_cand3, p_cand4]:
                    if p.is_file():
                        doc_path = p
                        break

                if doc_path and doc_path.exists():
                    file_ext = doc_path.suffix.lower()
                    with open(doc_path, "rb") as f_bytes:
                        file_data = f_bytes.read()

                    st.download_button(
                        label=f"📥 Download Original File ({doc_path.name})",
                        data=file_data,
                        file_name=doc_path.name,
                        key=f"dl_detail_{rec.id}"
                    )

                    if file_ext in [".png", ".jpg", ".jpeg"]:
                        st.image(file_data, caption=doc_path.name, width="stretch")
                    elif file_ext == ".pdf":
                        b64_pdf = base64.b64encode(file_data).decode('utf-8')
                        pdf_display = f'<iframe src="data:application/pdf;base64,{b64_pdf}" width="100%" height="550" type="application/pdf"></iframe>'
                        st.html(pdf_display)
                else:
                    st.warning(f"Document file not found at `{rec.relative_path or rec.file_name}`")

            # RIGHT: Extracted Information & Financial Audit
            with col_details:
                st.markdown("#### 📋 Extracted Template Fields & Verification")
                st.markdown(f"**Review Status:** `{rec.review_status}` &nbsp;|&nbsp; **Financial Status:** `{rec.financial_audit.status}`")
                
                field_rows = []
                for k, f in rec.fields.items():
                    val_display = str(f.value) if (f.value is not None and str(f.value).strip() != "") else "[Missing - Null]"
                    field_rows.append({
                        "Excel Column": k,
                        "Value": val_display,
                        "Confidence": f"{f.confidence:.0%}",
                        "Source": f.source,
                        "Uncertain?": "⚠️ Yes" if f.is_uncertain else "✅ Clean"
                    })
                st.table(pd.DataFrame(field_rows))

                # Financial Validation breakdown
                fa = rec.financial_audit
                st.markdown("#### 💰 Financial Reconciliation")
                fin_c1, fin_c2, fin_c3 = st.columns(3)
                fin_c1.metric("Taxable Value", f"₹ {fa.taxable_value:,.2f}" if fa.taxable_value is not None else "—")
                fin_c2.metric("Total Tax (GST)", f"₹ {fa.tax_amount:,.2f}" if fa.tax_amount is not None else "—")
                fin_c3.metric("Calculated Total", f"₹ {fa.calculated_invoice_total or fa.calculated_total:,.2f}" if (fa.calculated_invoice_total or fa.calculated_total) is not None else "—")
                
                fin_c4, fin_c5, fin_c6 = st.columns(3)
                fin_c4.metric("Extracted Total", f"₹ {fa.inv_amount or fa.extracted_invoice_total:,.2f}" if (fa.inv_amount or fa.extracted_invoice_total) is not None else "—")
                fin_c5.metric("Reconciliation Diff", f"₹ {fa.difference or fa.variance:,.2f}")
                fin_c6.metric("Math Valid?", "✅ Valid" if fa.is_math_valid else "❌ Invalid")

                if fa.cgst or fa.sgst or fa.igst or fa.cess:
                    st.caption(f"Tax Breakdown — CGST: {fa.cgst or 0} | SGST: {fa.sgst or 0} | IGST: {fa.igst or 0} | Cess: {fa.cess or 0}")

                if rec.processing_status == "FAILED":
                    st.error("❌ Document extraction failed. No financial fields were extracted.")
                elif fa.status == "VERIFIED":
                    st.success("✅ Financially Verified: Taxable Value + Taxes = Extracted Total within tolerance.")
                elif fa.status == "INCOMPLETE":
                    st.warning(f"⚠️ Financial Incomplete: {fa.notes or fa.explanation}")
                else:
                    st.error(f"❌ Financial Mismatch: {fa.notes or fa.explanation}")

                if getattr(fa, "is_ambiguous_tax", False):
                    st.warning("⚠️ Ambiguous Tax Structure: Both interstate (IGST) and intrastate (CGST/SGST) components reported.")

                # Corrections Form
                with st.expander("✏️ Correct Extracted Values", expanded=False):
                    with st.form(key=f"detail_edit_form_{rec.id}"):
                        edit_inputs = {}
                        e_cols = st.columns(2)
                        for idx, (col_name, f_data) in enumerate(rec.fields.items()):
                            col_slot = e_cols[idx % 2]
                            val_str = str(f_data.value) if f_data.value is not None else ""
                            edit_inputs[col_name] = col_slot.text_input(col_name, value=val_str)

                        edit_notes = st.text_input("Correction Note (optional)", placeholder="e.g. Fixed OCR typo in vendor name")
                        rev_name = st.text_input("Reviewer Name", value="AP Reviewer", key=f"rev_name_{rec.id}")

                        if st.form_submit_button("💾 Save Corrections"):
                            clean_updates = {}
                            for k, v in edit_inputs.items():
                                clean_updates[k] = None if (not v or v.strip() == "" or v.strip().lower() in ("none", "null")) else v.strip()
                            ok, err, _ = pipeline.db.update_invoice_fields(
                                record_id=rec.id,
                                fields_dict=clean_updates,
                                reviewer_id=rev_name,
                                notes=edit_notes
                            )
                            if ok:
                                st.success("Corrections saved and financial validation recalculated! Status remains PENDING.")
                                st.rerun()
                            else:
                                st.error(f"Save error: {err}")

                # Decision Controls
                st.markdown("#### 🎯 Approval Decision")
                dec_c1, dec_c2 = st.columns([2, 1])
                with dec_c1:
                    decision_reviewer = st.text_input("Reviewer Identity", value="AP Reviewer", key=f"dec_rev_{rec.id}")
                    decision_notes = st.text_input("Approval / Rejection Notes (optional)", key=f"dec_notes_{rec.id}")

                requires_ack = (fa.status == "MISMATCH" or getattr(fa, "is_ambiguous_tax", False))
                ack_single = True
                if requires_ack:
                    st.warning("⚠️ Warning: Approval requires explicit acknowledgement because financial status is MISMATCH or tax structure is ambiguous.")
                    ack_single = st.checkbox("I acknowledge the financial warnings/mismatches and authorize approval", key=f"ack_single_{rec.id}")

                btn_c1, btn_c2, btn_c3 = st.columns(3)
                with btn_c1:
                    if st.button("✅ Approve Invoice", type="primary", disabled=(requires_ack and not ack_single), key=f"btn_app_{rec.id}"):
                        ok, err = pipeline.db.approve_invoice(
                            record_id=rec.id,
                            reviewer_id=decision_reviewer,
                            acknowledge_warnings=ack_single,
                            notes=decision_notes
                        )
                        if ok:
                            st.success(f"Invoice #{rec.id} approved successfully!")
                            st.rerun()
                        else:
                            st.error(f"Approval failed: {err}")

                with btn_c2:
                    if st.button("❌ Reject Invoice", key=f"btn_rej_{rec.id}"):
                        ok, err = pipeline.db.reject_invoice(
                            record_id=rec.id,
                            reviewer_id=decision_reviewer,
                            reason=decision_notes or "Rejected by AP Reviewer"
                        )
                        if ok:
                            st.success(f"Invoice #{rec.id} marked as REJECTED.")
                            st.rerun()
                        else:
                            st.error(f"Rejection failed: {err}")

                with btn_c3:
                    if st.button("🔄 Return to Pending", key=f"btn_pend_{rec.id}"):
                        ok, err = pipeline.db.return_to_pending(
                            record_id=rec.id,
                            reviewer_id=decision_reviewer,
                            reason=decision_notes or "Returned to review queue"
                        )
                        if ok:
                            st.success(f"Invoice #{rec.id} returned to PENDING.")
                            st.rerun()
                        else:
                            st.error(f"Failed: {err}")

                # Audit History Timeline Expander
                audit_logs = pipeline.db.get_audit_history(rec.id)
                with st.expander(f"📜 Audit Trail & History ({len(audit_logs)} events)", expanded=False):
                    if audit_logs:
                        for ev in reversed(audit_logs):
                            st.markdown(f"**{ev.get('timestamp', '—')}** | Action: `{ev.get('action')}` by `{ev.get('reviewer_id')}`")
                            st.caption(f"Details: {ev.get('details')}")
                            if ev.get("changes"):
                                st.json(ev["changes"])
                            st.markdown("---")
                    else:
                        st.caption("No reviewer audit events recorded yet.")

# --- TAB 3: CUMULATIVE EXCEL LEDGER ---
with tab_ledger:
    st.subheader("Cumulative Excel Ledger Synchronization (`AP_BLR_Cumulative_Ledger.xlsx`)")
    st.markdown("""
    - **Pristine Template Preservation:** The master template (`Book3.xlsx`) is preserved unmodified.
    - **Cumulative Ledger:** Approved records are synchronized into `output/AP_BLR_Cumulative_Ledger.xlsx`.
    - **Duplicate Prevention:** Existing invoice numbers and vendors are detected to prevent duplicate rows.
    - **Review Gating:** Only invoices with **APPROVED** status are synchronized into the ledger.
    """)

    approved_list = pipeline.db.get_approved_unexported_records()
    all_approved = pipeline.db.get_approved_records()
    ledger_stats = pipeline.db.get_stats()
    c_m1, c_m2, c_m3 = st.columns(3)
    c_m1.metric("Approved Awaiting Export", ledger_stats.get("awaiting_export", len(approved_list)))
    c_m2.metric("Successfully Exported to Ledger", ledger_stats.get("exported_count", 0))
    c_m3.metric("Total Approved Invoices", len(all_approved))

    col_sync, col_dl = st.columns(2)
    with col_sync:
        if st.button("🔄 Synchronize Approved Invoices to Ledger", type="primary", disabled=(len(approved_list) == 0)):
            try:
                ledger_path, count = pipeline.export_approved_to_excel()
                st.success(f"Successfully appended **{count}** approved rows into: `{Path(ledger_path).name}`")
                st.rerun()
            except Exception as e:
                st.error(f"Synchronization failed: {e}")

    with col_dl:
        ledger_file = Path(OUTPUT_DIR) / "AP_BLR_Cumulative_Ledger.xlsx"
        if ledger_file.exists():
            with open(ledger_file, "rb") as f:
                st.download_button(
                    label="📥 Download Cumulative Ledger (.xlsx)",
                    data=f.read(),
                    file_name=ledger_file.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="secondary"
                )
        else:
            st.info("No cumulative ledger created yet. Approve records and click synchronize above.")

# --- TAB 4: SQLITE AUDIT TRAIL ---
with tab_history:
    st.subheader("SQLite Processing Audit Trail")
    stats = pipeline.db.get_stats()

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total Invoices Logged", stats["total_processed"])
    m2.metric("Approved Records", stats["approved_count"])
    m3.metric("Duplicates Prevented", stats["duplicates_found"])
    m4.metric("Math Verified", stats["math_verified"])
    m5.metric("Total Spend", f"₹ {stats['total_spend']:,.2f}")

    raw_table = pipeline.db.get_all_records(limit=200)
    if raw_table:
        df_hist = pd.DataFrame(raw_table)
        columns_to_show = [
            "id", "file_name", "relative_path", "vendor_name", "invoice_number",
            "invoice_date", "total_amount", "financial_status", "processing_status",
            "review_status", "export_status", "reviewed_at", "reviewer_id", "approval_notes",
            "is_math_valid", "is_duplicate", "is_file_duplicate", "created_at"
        ]
        available_cols = [c for c in columns_to_show if c in df_hist.columns]
        st.dataframe(df_hist[available_cols], width="stretch")
    else:
        st.info("No audit logs yet.")

# --- TAB 5: OUTLOOK INGESTION (MS GRAPH) ---
with tab_outlook:
    st.subheader("Microsoft Graph API Outlook Mailbox Monitoring")
    st.caption("Module ready for enterprise mailbox integration.")
    st.json(outlook_service.get_auth_status())
