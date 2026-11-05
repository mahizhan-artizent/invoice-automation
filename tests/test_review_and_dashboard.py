import os
import json
import pytest
from pathlib import Path
from decimal import Decimal
from fastapi.testclient import TestClient

from core.database import Database
from core.models import DynamicInvoiceRecord, ExtractedField, FinancialAudit
from core.pipeline import InvoiceProcessingPipeline
from core.template_engine import TemplateEngine
from core.config import EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME, INVOICE_SAMPLES_DIR
from backend.app import app

def create_mock_record(
    db: Database,
    file_name: str = "test_inv.pdf",
    vendor: str = "Acme Corp",
    inv_no: str = "INV-001",
    inv_date: str = "2026-06-01",
    taxable: float = 1000.0,
    total: float = 1180.0,
    review_status: str = "PENDING",
    financial_status: str = "VERIFIED",
    cgst: float = 90.0,
    sgst: float = 90.0,
    igst: float = 0.0,
    is_ambiguous_tax: bool = False
) -> int:
    rec = DynamicInvoiceRecord(
        file_name=file_name,
        relative_path=file_name,
        review_status=review_status,
        fields={
            "Entity": ExtractedField(value="Artizent"),
            "Vendor name ": ExtractedField(value=vendor),
            "Inv#": ExtractedField(value=inv_no),
            "Inv Date": ExtractedField(value=inv_date),
            "Taxable Value": ExtractedField(value=taxable),
            "Inv Amount": ExtractedField(value=total),
            "NOE": ExtractedField(value="Consulting")
        },
        financial_audit=FinancialAudit(
            status=financial_status,
            validation_status=financial_status,
            taxable_value=Decimal(str(taxable)),
            inv_amount=Decimal(str(total)),
            cgst=Decimal(str(cgst)) if cgst else None,
            sgst=Decimal(str(sgst)) if sgst else None,
            igst=Decimal(str(igst)) if igst else None,
            tax_amount=Decimal(str(cgst + sgst + igst)),
            is_ambiguous_tax=is_ambiguous_tax,
            is_math_valid=(financial_status == "VERIFIED")
        ),
        raw_response={
            "financial_audit": {
                "taxable_value": taxable,
                "inv_amount": total,
                "cgst": cgst,
                "sgst": sgst,
                "igst": igst,
                "tax_amount": cgst + sgst + igst,
                "is_ambiguous_tax": is_ambiguous_tax
            }
        }
    )
    return db.insert_invoice_record(rec)


# 1. Invoice listing and filtering
def test_invoice_listing_and_filtering(tmp_path):
    db_file = tmp_path / "test_filter.db"
    db = Database(str(db_file))

    id1 = create_mock_record(db, vendor="Vendor A", inv_no="INV-101", taxable=100.0, total=118.0, review_status="PENDING", financial_status="VERIFIED")
    id2 = create_mock_record(db, vendor="Vendor B", inv_no="INV-102", taxable=200.0, total=236.0, review_status="APPROVED", financial_status="INCOMPLETE")
    id3 = create_mock_record(db, vendor="Alpha Tech", inv_no="INV-201", taxable=300.0, total=354.0, review_status="REJECTED", financial_status="MISMATCH")

    # Search by vendor name
    records, total = db.search_and_filter_invoices(search="Alpha")
    assert total == 1
    assert records[0]["id"] == id3

    # Search by invoice number
    records, total = db.search_and_filter_invoices(invoice_number="102")
    assert total == 1
    assert records[0]["id"] == id2

    # Filter by review status
    records, total = db.search_and_filter_invoices(review_status="PENDING")
    assert total == 1
    assert records[0]["id"] == id1

    # Filter by financial status
    records, total = db.search_and_filter_invoices(financial_status="MISMATCH")
    assert total == 1
    assert records[0]["id"] == id3

    # Sorting by total_amount asc / desc
    records_desc, _ = db.search_and_filter_invoices(sort_by="total_amount", sort_order="desc")
    assert records_desc[0]["id"] == id3
    assert records_desc[-1]["id"] == id1

    records_asc, _ = db.search_and_filter_invoices(sort_by="total_amount", sort_order="asc")
    assert records_asc[0]["id"] == id1
    assert records_asc[-1]["id"] == id3


# 2. Invoice detail retrieval
def test_invoice_detail_retrieval(tmp_path):
    db_file = tmp_path / "test_detail.db"
    db = Database(str(db_file))

    rec_id = create_mock_record(db, vendor="Shaligram Infotech", inv_no="GST/0626/071", taxable=364800.0, total=430464.0)
    
    rec = db.get_record_by_id(rec_id)
    assert rec is not None
    assert rec["vendor_name"] == "Shaligram Infotech"
    assert rec["invoice_number"] == "GST/0626/071"
    assert rec["taxable_value"] == 364800.0
    assert rec["total_amount"] == 430464.0
    assert rec["review_status"] == "PENDING"

    # Verify fields_data is serialized valid JSON
    fields = json.loads(rec["fields_data"])
    assert "Vendor name " in fields or "Vendor name" in fields


# 3. Field editing and correction persistence
def test_field_editing_and_correction_persistence(tmp_path):
    db_file = tmp_path / "test_edit.db"
    db = Database(str(db_file))

    rec_id = create_mock_record(db, vendor="Original Vendor", inv_no="INV-ORIG", taxable=100.0, total=118.0)

    # Edit fields
    new_fields = {
        "Vendor name ": "Corrected Vendor LLP",
        "Inv#": "INV-CORRECTED",
        "Taxable Value": 200.0,
        "Inv Amount": 236.0
    }
    ok, err, updated_row = db.update_invoice_fields(
        record_id=rec_id,
        fields_dict=new_fields,
        reviewer_id="Jane Doe",
        notes="Corrected vendor name and values"
    )
    assert ok is True
    assert updated_row["vendor_name"] == "Corrected Vendor LLP"
    assert updated_row["invoice_number"] == "INV-CORRECTED"
    assert updated_row["taxable_value"] == 200.0
    assert updated_row["total_amount"] == 236.0
    # Rule: Editing must keep invoice in PENDING status (no auto-approval!)
    assert updated_row["review_status"] == "PENDING"

    # Verify corrections_json persisted the old and new values
    corrections = json.loads(updated_row["corrections_json"])
    assert "Vendor name " in corrections
    assert corrections["Vendor name "]["old"] == "Original Vendor"
    assert corrections["Vendor name "]["new"] == "Corrected Vendor LLP"


# 4. Financial validation after editing
def test_financial_validation_after_editing(tmp_path):
    db_file = tmp_path / "test_reval.db"
    db = Database(str(db_file))

    # Insert invoice with mismatch: taxable 100, cgst 9, sgst 9 (total should be 118), but inv_amount is 500!
    rec_id = create_mock_record(
        db,
        taxable=100.0,
        total=500.0,
        cgst=9.0,
        sgst=9.0,
        financial_status="MISMATCH"
    )
    r_initial = db.get_record_by_id(rec_id)
    assert r_initial["financial_status"] == "MISMATCH"

    # Edit invoice amount to exactly balance the math: 118.00
    ok, err, updated = db.update_invoice_fields(
        record_id=rec_id,
        fields_dict={"Inv Amount": 118.0, "Taxable Value": 100.0},
        reviewer_id="Reviewer 1"
    )
    assert ok is True
    # Financial status should have automatically recalculated to VERIFIED!
    assert updated["financial_status"] == "VERIFIED"
    assert updated["is_math_valid"] == 1


# 5. Individual approval, rejection, and returning to pending
def test_individual_approval_rejection_return_to_pending(tmp_path):
    db_file = tmp_path / "test_actions.db"
    db = Database(str(db_file))

    rec_id = create_mock_record(db, review_status="PENDING", financial_status="VERIFIED")

    # Approve
    ok, err = db.approve_invoice(rec_id, reviewer_id="Lead Reviewer", notes="Looks good")
    assert ok is True
    r_app = db.get_record_by_id(rec_id)
    assert r_app["review_status"] == "APPROVED"
    assert r_app["reviewer_id"] == "Lead Reviewer"
    assert r_app["reviewed_at"] is not None
    assert r_app["approval_notes"] == "Looks good"

    # Reject
    ok, err = db.reject_invoice(rec_id, reviewer_id="Auditor", reason="Invalid GSTIN")
    assert ok is True
    r_rej = db.get_record_by_id(rec_id)
    assert r_rej["review_status"] == "REJECTED"
    assert r_rej["reviewer_id"] == "Auditor"
    assert r_rej["approval_notes"] == "Invalid GSTIN"

    # Return to Pending
    ok, err = db.return_to_pending(rec_id, reviewer_id="Supervisor", reason="Need re-check")
    assert ok is True
    r_pend = db.get_record_by_id(rec_id)
    assert r_pend["review_status"] == "PENDING"
    assert r_pend["approval_notes"] == "Need re-check"


# 6. Batch approval and batch rejection
def test_batch_approval_and_rejection(tmp_path):
    db_file = tmp_path / "test_batch.db"
    db = Database(str(db_file))

    id1 = create_mock_record(db, inv_no="B-1", financial_status="VERIFIED")
    id2 = create_mock_record(db, inv_no="B-2", financial_status="VERIFIED")
    id3 = create_mock_record(db, inv_no="B-3", financial_status="VERIFIED")

    # Batch approve
    res = db.batch_approve([id1, id2], reviewer_id="Batch Reviewer")
    assert len(res["approved"]) == 2
    assert id1 in res["approved"]
    assert id2 in res["approved"]
    assert db.get_record_by_id(id1)["review_status"] == "APPROVED"
    assert db.get_record_by_id(id2)["review_status"] == "APPROVED"

    # Batch reject
    res_rej = db.batch_reject([id3], reviewer_id="Batch Reviewer", reason="Duplicate batch")
    assert len(res_rej["rejected"]) == 1
    assert db.get_record_by_id(id3)["review_status"] == "REJECTED"


# 7. Approval acknowledgement for financial warnings and mismatches
def test_approval_acknowledgement_for_financial_warnings(tmp_path):
    db_file = tmp_path / "test_ack.db"
    db = Database(str(db_file))

    # Mismatch record
    id_mismatch = create_mock_record(db, financial_status="MISMATCH")
    
    # Attempt approval without acknowledgement -> must fail
    ok, err = db.approve_invoice(id_mismatch, acknowledge_warnings=False)
    assert ok is False
    assert "acknowledgement is required" in err
    assert db.get_record_by_id(id_mismatch)["review_status"] == "PENDING"

    # Attempt approval with explicit acknowledgement -> succeeds
    ok, err = db.approve_invoice(id_mismatch, acknowledge_warnings=True, notes="Override authorized by CFO")
    assert ok is True
    assert db.get_record_by_id(id_mismatch)["review_status"] == "APPROVED"

    # Ambiguous tax structure
    id_ambig = create_mock_record(db, financial_status="VERIFIED", is_ambiguous_tax=True)
    ok, err = db.approve_invoice(id_ambig, acknowledge_warnings=False)
    assert ok is False
    assert "ambiguous tax" in err.lower()

    ok, err = db.approve_invoice(id_ambig, acknowledge_warnings=True)
    assert ok is True


# 8. Audit history timeline
def test_audit_history(tmp_path):
    db_file = tmp_path / "test_audit.db"
    db = Database(str(db_file))

    rec_id = create_mock_record(db, vendor="Beta Corp", taxable=100.0, total=118.0)

    # 1. Update field
    db.update_invoice_fields(rec_id, {"Vendor name ": "Beta Corp Ltd"}, reviewer_id="UserA")
    # 2. Reject
    db.reject_invoice(rec_id, reviewer_id="UserB", reason="Price discrepancy")
    # 3. Return to pending
    db.return_to_pending(rec_id, reviewer_id="UserC", reason="Discount clarified")
    # 4. Approve (with acknowledgement if warnings exist)
    ok, err = db.approve_invoice(rec_id, reviewer_id="UserD", acknowledge_warnings=True, notes="Final signoff")
    assert ok is True

    history = db.get_audit_history(rec_id)
    assert len(history) == 4
    actions = [ev["action"] for ev in history]
    assert actions == ["FIELD_UPDATE", "REJECT", "RETURN_TO_PENDING", "APPROVE"]
    assert history[0]["reviewer_id"] == "UserA"
    assert history[3]["reviewer_id"] == "UserD"


# 9. Pending and Rejected invoices excluded from cumulative Excel export
def test_pending_and_rejected_invoice_exclusion_from_export(tmp_path):
    db_file = tmp_path / "test_export_exclusion.db"
    db = Database(str(db_file))
    target_ledger = tmp_path / "Cumulative_Ledger.xlsx"

    id_pend = create_mock_record(db, inv_no="PEND-1", review_status="PENDING")
    id_rej = create_mock_record(db, inv_no="REJ-1", review_status="REJECTED")

    pipeline = InvoiceProcessingPipeline(db_path=str(db_file))
    saved_path, count = pipeline.export_approved_to_excel(target_path=str(target_ledger))

    # Nothing should be exported
    assert count == 0
    assert db.get_record_by_id(id_pend)["export_status"] == "NOT_EXPORTED"
    assert db.get_record_by_id(id_rej)["export_status"] == "NOT_EXPORTED"


# 10. Approved invoice export integration
def test_approved_invoice_export(tmp_path):
    db_file = tmp_path / "test_export_approved.db"
    db = Database(str(db_file))
    target_ledger = tmp_path / "Cumulative_Ledger.xlsx"

    id_app = create_mock_record(db, vendor="Exported Vendor", inv_no="EXP-001", review_status="PENDING")
    # Approve it
    db.approve_invoice(id_app, reviewer_id="Export Reviewer")

    pipeline = InvoiceProcessingPipeline(db_path=str(db_file))
    saved_path, count = pipeline.export_approved_to_excel(target_path=str(target_ledger))

    assert count == 1
    rec_after = db.get_record_by_id(id_app)
    assert rec_after["export_status"] == "EXPORTED"
    assert rec_after["exported_at"] is not None

    # Repeated export should not re-export
    saved_path2, count2 = pipeline.export_approved_to_excel(target_path=str(target_ledger))
    assert count2 == 0


# 11. Document preview access restrictions (prevent path traversal)
def test_document_preview_access_restrictions(tmp_path):
    from unittest.mock import patch
    db_file = tmp_path / "test_preview_api.db"
    temp_db = Database(str(db_file))

    with patch("backend.app.pipeline.db", temp_db):
        client = TestClient(app)

        # 1. Non-existent record returns 404
        resp = client.get("/api/invoices/999999/document")
        assert resp.status_code == 404

        # 2. Existing sample file preview succeeds
        samples = list(Path(INVOICE_SAMPLES_DIR).glob("*.pdf"))
        if samples:
            sample_name = samples[0].name
            rec_id = temp_db.insert_invoice_record(
                DynamicInvoiceRecord(
                    file_name=sample_name,
                    relative_path=sample_name,
                    fields={"Vendor name ": ExtractedField(value="Test")}
                )
            )
            resp_good = client.get(f"/api/invoices/{rec_id}/document")
            assert resp_good.status_code == 200
            assert resp_good.headers["content-type"] == "application/pdf"

        # 3. Path traversal attack attempt is blocked
        traversal_id = temp_db.insert_invoice_record(
            DynamicInvoiceRecord(
                file_name="traversal.pdf",
                relative_path="../../etc/passwd",
                fields={"Vendor name ": ExtractedField(value="Hacker")}
            )
        )
        resp_bad = client.get(f"/api/invoices/{traversal_id}/document")
        # Must be 404 (file not found) or 403 (access denied), never 200
        assert resp_bad.status_code in (403, 404)


# 12. Dashboard compatibility and KPIs
def test_existing_dashboard_compatibility(tmp_path):
    from unittest.mock import patch
    db_file = tmp_path / "test_kpis.db"
    db = Database(str(db_file))

    create_mock_record(db, review_status="PENDING", financial_status="VERIFIED")
    create_mock_record(db, review_status="APPROVED", financial_status="VERIFIED")
    create_mock_record(db, review_status="REJECTED", financial_status="MISMATCH")

    stats = db.get_stats()
    # Check all required overview keys
    assert stats["total_processed"] == 3
    assert stats["pending_review"] == 1
    assert stats["approved_count"] == 1
    assert stats["rejected_count"] == 1
    assert stats["financially_verified"] == 2
    assert stats["financial_mismatch"] == 1
    assert stats["awaiting_export"] == 1
    assert stats["exported_count"] == 0

    # Test FastAPI stats endpoint using mocked db
    with patch("backend.app.pipeline.db", db):
        client = TestClient(app)
        resp = client.get("/api/dashboard/stats")
        assert resp.status_code == 200
        json_stats = resp.json()
        assert json_stats["total_processed"] == 3
        assert json_stats["pending_review"] == 1
        assert json_stats["approved_count"] == 1
        assert json_stats["rejected_count"] == 1
