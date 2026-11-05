import os
import sys
import shutil
import json
from pathlib import Path
from typing import List, Optional, Dict, Any

# Ensure project root is in sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks, Query, Body
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, Field

from core.config import (
    EXCEL_TEMPLATE_PATH,
    EXCEL_SHEET_NAME,
    SQLITE_DB_PATH,
    INVOICE_SAMPLES_DIR,
    OUTPUT_DIR,
    GEMINI_PRIMARY_MODEL
)
from core.pipeline import InvoiceProcessingPipeline
from core.models import DynamicInvoiceRecord
from core.outlook_service import MicrosoftGraphOutlookService

app = FastAPI(
    title="Gemini Invoice Document Intelligence API",
    description="Enterprise invoice automation engine with dynamic template schema, decimal audit, and review workflow",
    version="2.1.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

pipeline = InvoiceProcessingPipeline()
app.state.pipeline = pipeline
outlook_service = MicrosoftGraphOutlookService()

@app.get("/api/health")
def health_check():
    stats = pipeline.db.get_stats()
    return {
        "status": "healthy",
        "primary_model": GEMINI_PRIMARY_MODEL,
        "template_path": EXCEL_TEMPLATE_PATH,
        "sheet_name": EXCEL_SHEET_NAME,
        "stats": stats
    }

@app.get("/api/template/columns")
def get_columns():
    cols = pipeline.get_template_columns()
    return {
        "template": EXCEL_TEMPLATE_PATH,
        "sheet": EXCEL_SHEET_NAME,
        "columns_count": len(cols),
        "columns": [c.model_dump() for c in cols]
    }

@app.post("/api/invoices/process-file")
async def process_file(
    file: UploadFile = File(...),
    folder_context: Optional[str] = Form("Direct Upload"),
    model_name: Optional[str] = Form(None),
    skip_if_processed: bool = Form(True)
):
    upload_dir = Path("uploads")
    upload_dir.mkdir(parents=True, exist_ok=True)
    temp_path = upload_dir / file.filename

    try:
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        record = pipeline.process_single_invoice(
            file_path=temp_path,
            relative_path=file.filename,
            folder_context=folder_context,
            model_name=model_name,
            skip_if_processed=skip_if_processed,
            auto_save_excel=False
        )
        return {"success": True, "record": record.model_dump()}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/invoices/process-batch")
def process_batch(
    directory_path: Optional[str] = Form(None),
    model_name: Optional[str] = Form(None),
    skip_if_processed: bool = Form(True)
):
    target_dir = directory_path or INVOICE_SAMPLES_DIR
    try:
        records = pipeline.process_batch_recursive(
            directory_path=target_dir,
            model_name=model_name,
            skip_if_processed=skip_if_processed
        )
        return {
            "success": True,
            "processed_count": len(records),
            "records": [r.model_dump() for r in records]
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# --- Pydantic Request Models for Review & Approval ---

class UpdateFieldsRequest(BaseModel):
    fields: Dict[str, Any] = Field(..., description="Map of Excel column name to new value")
    reviewer_id: Optional[str] = Field("AP Reviewer", description="Identity of reviewer making corrections")
    notes: Optional[str] = Field(None, description="Optional note explaining correction")

class ApproveRequest(BaseModel):
    reviewer_id: Optional[str] = Field("AP Reviewer", description="Reviewer making approval")
    acknowledge_warnings: bool = Field(False, description="Explicit acknowledgement if invoice has MISMATCH or ambiguous tax")
    notes: Optional[str] = Field(None, description="Approval notes")

class RejectRequest(BaseModel):
    reviewer_id: Optional[str] = Field("AP Reviewer", description="Reviewer rejecting invoice")
    reason: Optional[str] = Field(None, description="Reason for rejection")

class ReturnToPendingRequest(BaseModel):
    reviewer_id: Optional[str] = Field("AP Reviewer", description="Reviewer returning invoice to pending")
    reason: Optional[str] = Field(None, description="Reason for returning to pending")

class BatchApproveRequest(BaseModel):
    record_ids: List[int] = Field(..., description="List of invoice IDs to approve")
    reviewer_id: Optional[str] = Field("AP Reviewer", description="Reviewer identity")
    acknowledge_warnings: bool = Field(False, description="Acknowledge warnings/mismatches for batch")
    notes: Optional[str] = Field(None, description="Batch approval notes")

class BatchRejectRequest(BaseModel):
    record_ids: List[int] = Field(..., description="List of invoice IDs to reject")
    reviewer_id: Optional[str] = Field("AP Reviewer", description="Reviewer identity")
    reason: Optional[str] = Field(None, description="Reason for batch rejection")


@app.get("/api/dashboard/stats")
def get_dashboard_stats():
    """Returns comprehensive dashboard KPIs and financial reconciliation counts."""
    return pipeline.db.get_stats()

@app.get("/api/invoices")
def list_invoices(
    search: Optional[str] = None,
    vendor: Optional[str] = None,
    invoice_number: Optional[str] = None,
    processing_status: Optional[str] = None,
    review_status: Optional[str] = None,
    financial_status: Optional[str] = None,
    export_status: Optional[str] = None,
    sort_by: str = Query("id"),
    sort_order: str = Query("desc"),
    limit: int = 100,
    offset: int = 0
):
    """Lists, filters, searches, and sorts invoice records."""
    records, total = pipeline.db.search_and_filter_invoices(
        search=search,
        vendor=vendor,
        invoice_number=invoice_number,
        processing_status=processing_status,
        review_status=review_status,
        financial_status=financial_status,
        export_status=export_status,
        sort_by=sort_by,
        sort_order=sort_order,
        limit=limit,
        offset=offset
    )
    return {"total": total, "limit": limit, "offset": offset, "records": records}

@app.get("/api/invoices/{record_id}")
def get_invoice_detail(record_id: int):
    """Retrieves full details, parsed fields, audit history, and corrections for an invoice."""
    rec = pipeline.db.get_record_by_id(record_id)
    if not rec:
        raise HTTPException(status_code=404, detail=f"Invoice #{record_id} not found")
    
    parsed_fields = {}
    if rec.get("fields_data"):
        try:
            parsed_fields = json.loads(rec["fields_data"])
        except Exception:
            parsed_fields = {}
    
    audit_history = pipeline.db.get_audit_history(record_id)
    
    corrections = {}
    if rec.get("corrections_json"):
        try:
            corrections = json.loads(rec["corrections_json"])
        except Exception:
            corrections = {}
            
    return {
        "record": rec,
        "fields": parsed_fields,
        "audit_history": audit_history,
        "corrections": corrections
    }

@app.put("/api/invoices/{record_id}/fields")
def update_invoice_fields(record_id: int, payload: UpdateFieldsRequest):
    """Updates extracted fields, reruns financial cross-footing, and records audit trail."""
    ok, err, updated_row = pipeline.db.update_invoice_fields(
        record_id=record_id,
        fields_dict=payload.fields,
        reviewer_id=payload.reviewer_id or "AP Reviewer",
        notes=payload.notes
    )
    if not ok:
        status_code = 404 if "not found" in (err or "").lower() else 400
        raise HTTPException(status_code=status_code, detail=err)
    return {"success": True, "record_id": record_id, "updated_record": updated_row}

@app.post("/api/invoices/{record_id}/approve")
def approve_invoice(record_id: int, payload: ApproveRequest):
    """Approves an invoice. Enforces explicit acknowledgement if there is a mismatch or ambiguous tax."""
    ok, err = pipeline.db.approve_invoice(
        record_id=record_id,
        reviewer_id=payload.reviewer_id or "AP Reviewer",
        acknowledge_warnings=payload.acknowledge_warnings,
        notes=payload.notes
    )
    if not ok:
        status_code = 404 if "not found" in (err or "").lower() else 400
        raise HTTPException(status_code=status_code, detail=err)
    return {"success": True, "record_id": record_id, "review_status": "APPROVED"}

@app.post("/api/invoices/{record_id}/reject")
def reject_invoice(record_id: int, payload: RejectRequest):
    """Rejects an invoice with a reason. Excluded from cumulative Excel export."""
    ok, err = pipeline.db.reject_invoice(
        record_id=record_id,
        reviewer_id=payload.reviewer_id or "AP Reviewer",
        reason=payload.reason
    )
    if not ok:
        raise HTTPException(status_code=404, detail=err)
    return {"success": True, "record_id": record_id, "review_status": "REJECTED"}

@app.post("/api/invoices/{record_id}/return-to-pending")
def return_invoice_to_pending(record_id: int, payload: ReturnToPendingRequest):
    """Returns an invoice to PENDING review status."""
    ok, err = pipeline.db.return_to_pending(
        record_id=record_id,
        reviewer_id=payload.reviewer_id or "AP Reviewer",
        reason=payload.reason
    )
    if not ok:
        raise HTTPException(status_code=404, detail=err)
    return {"success": True, "record_id": record_id, "review_status": "PENDING"}

@app.post("/api/invoices/batch/approve")
def batch_approve(payload: BatchApproveRequest):
    """Batch approves selected invoices with warning checks."""
    result = pipeline.db.batch_approve(
        record_ids=payload.record_ids,
        reviewer_id=payload.reviewer_id or "AP Reviewer",
        acknowledge_warnings=payload.acknowledge_warnings,
        notes=payload.notes
    )
    return {"success": True, **result}

@app.post("/api/invoices/batch/reject")
def batch_reject(payload: BatchRejectRequest):
    """Batch rejects selected invoices."""
    result = pipeline.db.batch_reject(
        record_ids=payload.record_ids,
        reviewer_id=payload.reviewer_id or "AP Reviewer",
        reason=payload.reason
    )
    return {"success": True, **result}

@app.post("/api/export/approved")
def export_approved_alias():
    """Alias for Phase 4 cumulative ledger synchronization."""
    try:
        path, count = pipeline.export_approved_to_excel()
        return {"success": True, "ledger_path": path, "rows_exported": count}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/invoices/{record_id}/document")
def get_invoice_document(record_id: int):
    """
    Secure document preview endpoint.
    Strictly verifies path containment inside configured invoice directories.
    Prevents directory traversal attacks.
    """
    rec = pipeline.db.get_record_by_id(record_id)
    if not rec:
        raise HTTPException(status_code=404, detail=f"Record #{record_id} not found")
    
    rel_path = rec.get("relative_path") or rec.get("file_name")
    if not rel_path:
        raise HTTPException(status_code=404, detail="No document reference stored for this record")
    
    # Check for obvious traversal attack sequences
    if ".." in rel_path or rel_path.startswith("/") or rel_path.startswith("\\"):
        # Normalize relative path securely
        rel_path = Path(rel_path).name

    allowed_roots = [
        Path(INVOICE_SAMPLES_DIR).resolve(),
        (PROJECT_ROOT / "uploads").resolve(),
        PROJECT_ROOT.resolve()
    ]
    
    candidate_file: Optional[Path] = None
    # 1. Check under INVOICE_SAMPLES_DIR
    p1 = (Path(INVOICE_SAMPLES_DIR) / rel_path).resolve()
    if p1.is_file():
        candidate_file = p1
    else:
        # 2. Check under uploads
        p2 = (PROJECT_ROOT / "uploads" / rel_path).resolve()
        if p2.is_file():
            candidate_file = p2
        else:
            # 3. Check direct filename under INVOICE_SAMPLES_DIR
            p3 = (Path(INVOICE_SAMPLES_DIR) / Path(rel_path).name).resolve()
            if p3.is_file():
                candidate_file = p3
            else:
                p4 = (PROJECT_ROOT / rel_path).resolve()
                if p4.is_file():
                    candidate_file = p4

    if not candidate_file or not candidate_file.exists():
        raise HTTPException(status_code=404, detail="Invoice file not found on disk")

    # STRICT ACCESS CHECK: Ensure candidate_file is strictly within an allowed root
    is_safe = False
    for root in allowed_roots:
        try:
            candidate_file.relative_to(root)
            is_safe = True
            break
        except ValueError:
            continue

    if not is_safe:
        raise HTTPException(status_code=403, detail="Access denied: file is outside allowed invoice directories")

    ext = candidate_file.suffix.lower()
    media_types = {
        ".pdf": "application/pdf",
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg"
    }
    m_type = media_types.get(ext, "application/octet-stream")
    return FileResponse(str(candidate_file), media_type=m_type, filename=candidate_file.name)

@app.get("/api/records")
def list_records(
    limit: int = 100,
    status: Optional[str] = None,
    review_status: Optional[str] = None
):
    records = pipeline.db.get_all_records(limit=limit, status=status, review_status=review_status)
    stats = pipeline.db.get_stats()
    return {
        "stats": stats,
        "records": records
    }

@app.post("/api/invoices/{record_id}/review")
def update_review(
    record_id: int,
    review_status: str = Body(..., embed=True),
    edited_fields: Optional[Dict[str, Any]] = Body(None, embed=True)
):
    """Updates invoice review status (APPROVED, REJECTED, PENDING) and optional corrections."""
    success = pipeline.db.update_review_status(
        record_id=record_id,
        review_status=review_status.upper(),
        edited_fields=edited_fields
    )
    if not success:
        raise HTTPException(status_code=404, detail=f"Record #{record_id} not found.")
    return {"success": True, "record_id": record_id, "review_status": review_status.upper()}

@app.post("/api/invoices/export-approved")
def export_approved():
    """Exports only approved records to the cumulative Excel ledger."""
    try:
        path, count = pipeline.export_approved_to_excel()
        return {"success": True, "ledger_path": path, "rows_exported": count}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/download-excel")
def download_excel(file_path: Optional[str] = None):
    if file_path and os.path.exists(file_path):
        target = file_path
    else:
        out_dir = Path(OUTPUT_DIR)
        cumulative = out_dir / "AP_BLR_Cumulative_Ledger.xlsx"
        if cumulative.exists():
            target = str(cumulative)
        else:
            files = list(out_dir.glob("Book3_*.xlsx"))
            if not files:
                target = EXCEL_TEMPLATE_PATH
            else:
                files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
                target = str(files[0])

    return FileResponse(
        target,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=os.path.basename(target)
    )

@app.get("/api/outlook/status")
def outlook_status():
    return outlook_service.get_auth_status()
