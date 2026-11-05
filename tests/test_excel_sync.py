import os
import shutil
import hashlib
import openpyxl
import pytest
from pathlib import Path
from datetime import datetime, date
from decimal import Decimal
from unittest.mock import patch

from core.template_engine import TemplateEngine
from core.models import DynamicInvoiceRecord, ExtractedField, FinancialAudit, ExcelColumnDefinition
from core.database import Database
from core.pipeline import InvoiceProcessingPipeline
from core.config import EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME

def get_file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()

# 1. Original Template Preservation
def test_original_template_preservation(tmp_path):
    orig_path = Path(EXCEL_TEMPLATE_PATH).resolve()
    orig_sha = get_file_sha256(orig_path)
    
    target_ledger = tmp_path / "ledger.xlsx"
    engine = TemplateEngine(EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME)
    
    rec = DynamicInvoiceRecord(
        file_name="test.pdf",
        review_status="APPROVED",
        fields={
            "Vendor name ": ExtractedField(value="Vendor One"),
            "Inv#": ExtractedField(value="INV-101"),
            "Taxable Value": ExtractedField(value=10000.0),
            "Inv Amount": ExtractedField(value=11800.0)
        }
    )
    engine.write_records_to_excel([rec], target_path=str(target_ledger), only_approved=True)
    
    # Verify original Book3.xlsx was completely untouched
    assert get_file_sha256(orig_path) == orig_sha
    
    # Verify attempting to overwrite Book3.xlsx directly raises a safety ValueError
    with pytest.raises(ValueError, match="Safety violation: Target path"):
        engine.write_records_to_excel([rec], target_path=str(orig_path), only_approved=True)

# 2. Existing Row Preservation
def test_existing_row_preservation(tmp_path):
    target_ledger = tmp_path / "ledger_existing.xlsx"
    engine = TemplateEngine(EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME)
    
    # Initial load of template to check baseline rows
    base_wb = openpyxl.load_workbook(EXCEL_TEMPLATE_PATH)
    base_ws = base_wb[EXCEL_SHEET_NAME]
    base_max_row = base_ws.max_row
    base_rows_content = [
        (r, [base_ws.cell(r, c).value for c in range(1, 13)])
        for r in range(engine.header_row + 1, base_max_row + 1)
        if any(base_ws.cell(r, c).value is not None for c in range(1, 13))
    ]
    
    rec = DynamicInvoiceRecord(
        file_name="test_preserve.pdf",
        review_status="APPROVED",
        fields={
            "Vendor name ": ExtractedField(value="Unique Vendor XYZ"),
            "Inv#": ExtractedField(value="XYZ-001"),
            "Taxable Value": ExtractedField(value=25000.0),
            "Inv Amount": ExtractedField(value=29500.0)
        }
    )
    engine.write_records_to_excel([rec], target_path=str(target_ledger), only_approved=True)
    
    # Check that previous populated rows are completely intact
    new_wb = openpyxl.load_workbook(target_ledger)
    new_ws = new_wb[EXCEL_SHEET_NAME]
    for row_idx, expected_row in base_rows_content:
        actual_row = [new_ws.cell(row_idx, c).value for c in range(1, 13)]
        assert actual_row == expected_row, f"Row {row_idx} was modified!"

# 3. Existing Formula Preservation
def test_existing_formula_preservation(tmp_path):
    test_template = tmp_path / "template_with_formula.xlsx"
    shutil.copy2(EXCEL_TEMPLATE_PATH, test_template)
    
    # Insert a dummy formula in an existing cell
    wb = openpyxl.load_workbook(test_template)
    ws = wb[EXCEL_SHEET_NAME]
    ws.cell(5, 12, "=SUM(J5, 100)")
    wb.save(test_template)
    
    engine = TemplateEngine(str(test_template), EXCEL_SHEET_NAME)
    target_ledger = tmp_path / "ledger_formula.xlsx"
    
    rec = DynamicInvoiceRecord(
        file_name="test_form.pdf",
        review_status="APPROVED",
        fields={"Vendor name ": ExtractedField(value="V1"), "Inv#": ExtractedField(value="I1")}
    )
    engine.write_records_to_excel([rec], target_path=str(target_ledger), only_approved=True)
    
    check_wb = openpyxl.load_workbook(target_ledger, data_only=False)
    assert check_wb[EXCEL_SHEET_NAME].cell(5, 12).value == "=SUM(J5, 100)"

# 4. Formatting and Number-Format Preservation
def test_formatting_and_number_format_preservation(tmp_path):
    target_ledger = tmp_path / "ledger_formatting.xlsx"
    engine = TemplateEngine(EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME)
    
    rec = DynamicInvoiceRecord(
        file_name="test_fmt.pdf",
        review_status="APPROVED",
        fields={
            "Vendor name ": ExtractedField(value="Style Test Vendor"),
            "Inv#": ExtractedField(value="FMT-01"),
            "Taxable Value": ExtractedField(value=50000.0),
            "Inv Amount": ExtractedField(value=59000.0)
        }
    )
    engine.write_records_to_excel([rec], target_path=str(target_ledger), only_approved=True)
    
    wb = openpyxl.load_workbook(target_ledger)
    ws = wb[EXCEL_SHEET_NAME]
    new_row = ws.max_row
    
    proto_row = engine.header_row + 1
    # Check currency number format copied to Col 10 (Taxable Value) and Col 11 (Inv Amount)
    assert ws.cell(new_row, 10).number_format == ws.cell(proto_row, 10).number_format
    assert ws.cell(new_row, 11).number_format == ws.cell(proto_row, 11).number_format
    # Check font preserved
    assert ws.cell(new_row, 4).font.name == ws.cell(proto_row, 4).font.name
    # Check border preserved
    assert ws.cell(new_row, 1).border is not None

# 5. Sequential S.No Generation
def test_sequential_sno_generation(tmp_path):
    target_ledger = tmp_path / "ledger_sno.xlsx"
    engine = TemplateEngine(EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME)
    
    # Existing max S.No in Book3.xlsx is 18
    rec1 = DynamicInvoiceRecord(
        file_name="rec1.pdf",
        review_status="APPROVED",
        fields={"Vendor name ": ExtractedField(value="V A"), "Inv#": ExtractedField(value="A-1")}
    )
    rec2 = DynamicInvoiceRecord(
        file_name="rec2.pdf",
        review_status="APPROVED",
        fields={"Vendor name ": ExtractedField(value="V B"), "Inv#": ExtractedField(value="B-1")}
    )
    engine.write_records_to_excel([rec1, rec2], target_path=str(target_ledger), only_approved=True)
    
    wb = openpyxl.load_workbook(target_ledger)
    ws = wb[EXCEL_SHEET_NAME]
    
    # Sequential S.No generation: verifies consecutive integers
    sno1 = ws.cell(ws.max_row - 1, 1).value
    sno2 = ws.cell(ws.max_row, 1).value
    assert sno1 is not None and sno2 is not None
    assert sno2 == sno1 + 1

# 6. Cumulative Exports Across Multiple Batches
def test_cumulative_exports_across_multiple_batches(tmp_path):
    target_ledger = tmp_path / "ledger_batches.xlsx"
    engine = TemplateEngine(EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME)
    
    # Batch 1
    rec1 = DynamicInvoiceRecord(
        file_name="batch1.pdf",
        review_status="APPROVED",
        fields={"Vendor name ": ExtractedField(value="Batch1 Vendor"), "Inv#": ExtractedField(value="B1")}
    )
    engine.write_records_to_excel([rec1], target_path=str(target_ledger), only_approved=True)
    
    wb1 = openpyxl.load_workbook(target_ledger)
    count1 = wb1[EXCEL_SHEET_NAME].max_row
    
    # Batch 2
    rec2 = DynamicInvoiceRecord(
        file_name="batch2.pdf",
        review_status="APPROVED",
        fields={"Vendor name ": ExtractedField(value="Batch2 Vendor"), "Inv#": ExtractedField(value="B2")}
    )
    engine.write_records_to_excel([rec2], target_path=str(target_ledger), only_approved=True)
    
    wb2 = openpyxl.load_workbook(target_ledger)
    count2 = wb2[EXCEL_SHEET_NAME].max_row
    assert count2 == count1 + 1

# 7. Duplicate Export Prevention
def test_duplicate_export_prevention(tmp_path):
    target_ledger = tmp_path / "ledger_dupes.xlsx"
    engine = TemplateEngine(EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME)
    
    rec = DynamicInvoiceRecord(
        file_name="inv.pdf",
        review_status="APPROVED",
        fields={"Vendor name ": ExtractedField(value="Duplicate Vendor"), "Inv#": ExtractedField(value="DUP-01")}
    )
    # First write
    _, count1, _ = engine.write_records_to_excel([rec], target_path=str(target_ledger), only_approved=True)
    assert count1 == 1
    
    # Second write with same invoice
    _, count2, _ = engine.write_records_to_excel([rec], target_path=str(target_ledger), only_approved=True)
    assert count2 == 0, "Duplicate invoice must not be re-appended!"

# 8. Pending Invoice Exclusion
def test_pending_invoice_exclusion(tmp_path):
    target_ledger = tmp_path / "ledger_pending.xlsx"
    engine = TemplateEngine(EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME)
    
    rec_pending = DynamicInvoiceRecord(
        file_name="pending.pdf",
        review_status="PENDING",
        fields={"Vendor name ": ExtractedField(value="Pending Vendor"), "Inv#": ExtractedField(value="P-01")}
    )
    _, count, _ = engine.write_records_to_excel([rec_pending], target_path=str(target_ledger), only_approved=True)
    assert count == 0

# 9. Rejected Invoice Exclusion
def test_rejected_invoice_exclusion(tmp_path):
    target_ledger = tmp_path / "ledger_rejected.xlsx"
    engine = TemplateEngine(EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME)
    
    rec_rejected = DynamicInvoiceRecord(
        file_name="rejected.pdf",
        review_status="REJECTED",
        fields={"Vendor name ": ExtractedField(value="Rejected Vendor"), "Inv#": ExtractedField(value="R-01")}
    )
    _, count, _ = engine.write_records_to_excel([rec_rejected], target_path=str(target_ledger), only_approved=True)
    assert count == 0

# 10. Approved Invoice Export
def test_approved_invoice_export(tmp_path):
    target_ledger = tmp_path / "ledger_approved.xlsx"
    engine = TemplateEngine(EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME)
    
    rec_approved = DynamicInvoiceRecord(
        file_name="approved.pdf",
        review_status="APPROVED",
        fields={"Vendor name ": ExtractedField(value="Approved Vendor"), "Inv#": ExtractedField(value="APP-01")}
    )
    _, count, mappings = engine.write_records_to_excel([rec_approved], target_path=str(target_ledger), only_approved=True)
    assert count == 1
    assert len(mappings) == 1

# 11. Missing Values Exported as Blank
def test_missing_values_exported_as_blank(tmp_path):
    target_ledger = tmp_path / "ledger_blank.xlsx"
    engine = TemplateEngine(EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME)
    
    rec = DynamicInvoiceRecord(
        file_name="partial.pdf",
        review_status="APPROVED",
        fields={
            "Vendor name ": ExtractedField(value="Partial Vendor"),
            "Inv#": ExtractedField(value="PART-01"),
            "Inv received date ": ExtractedField(value=None),  # Blank
            "Sender": ExtractedField(value=None)               # Blank
        }
    )
    engine.write_records_to_excel([rec], target_path=str(target_ledger), only_approved=True)
    
    wb = openpyxl.load_workbook(target_ledger)
    ws = wb[EXCEL_SHEET_NAME]
    last_row = ws.max_row
    
    # Inv received date is Col 3
    assert ws.cell(last_row, 3).value is None
    # Sender is Col 12
    assert ws.cell(last_row, 12).value is None

# 12. Monetary Values Exported as Numeric Cells
def test_monetary_values_exported_as_numeric_cells(tmp_path):
    target_ledger = tmp_path / "ledger_numeric.xlsx"
    engine = TemplateEngine(EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME)
    
    rec = DynamicInvoiceRecord(
        file_name="numeric.pdf",
        review_status="APPROVED",
        fields={
            "Vendor name ": ExtractedField(value="Num Vendor"),
            "Inv#": ExtractedField(value="NUM-01"),
            "Taxable Value": ExtractedField(value="325,000.50"),
            "Inv Amount": ExtractedField(value=383500.0)
        }
    )
    engine.write_records_to_excel([rec], target_path=str(target_ledger), only_approved=True)
    
    wb = openpyxl.load_workbook(target_ledger)
    ws = wb[EXCEL_SHEET_NAME]
    last_row = ws.max_row
    
    tax_val = ws.cell(last_row, 10).value
    inv_amt = ws.cell(last_row, 11).value
    
    assert isinstance(tax_val, (int, float))
    assert tax_val == 325000.50
    assert isinstance(inv_amt, (int, float))
    assert inv_amt == 383500.0

# 13. Workbook Write Failure & Atomic Recovery
def test_workbook_write_failure_atomic_recovery(tmp_path):
    target_ledger = tmp_path / "ledger_fail.xlsx"
    engine = TemplateEngine(EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME)
    
    # Initial valid write
    rec1 = DynamicInvoiceRecord(
        file_name="init.pdf",
        review_status="APPROVED",
        fields={"Vendor name ": ExtractedField(value="Init Vendor"), "Inv#": ExtractedField(value="INIT-01")}
    )
    engine.write_records_to_excel([rec1], target_path=str(target_ledger), only_approved=True)
    mtime_before = os.path.getmtime(target_ledger)
    
    rec2 = DynamicInvoiceRecord(
        file_name="fail.pdf",
        review_status="APPROVED",
        fields={"Vendor name ": ExtractedField(value="Fail Vendor"), "Inv#": ExtractedField(value="FAIL-01")}
    )
    
    # Simulate write error (e.g. PermissionError or disk failure during wb.save)
    with patch("openpyxl.Workbook.save", side_effect=PermissionError("File locked by Excel")):
        with pytest.raises(RuntimeError, match="Excel synchronization failure"):
            engine.write_records_to_excel([rec2], target_path=str(target_ledger), only_approved=True)
            
    # Verify temporary file was cleaned up and target ledger is untouched
    assert not (tmp_path / "ledger_fail.tmp.xlsx").exists()
    assert os.path.getmtime(target_ledger) == mtime_before

# 14. Retry After Export Failure
def test_retry_after_export_failure(tmp_path):
    test_db_path = tmp_path / "test_retry.db"
    test_ledger = tmp_path / "test_retry_ledger.xlsx"
    db = Database(str(test_db_path))
    
    # Insert an approved invoice in DB
    rec = DynamicInvoiceRecord(
        file_name="retry_inv.pdf",
        review_status="APPROVED",
        fields={
            "Entity": ExtractedField(value="Nous Info"),
            "Vendor name ": ExtractedField(value="Retry Vendor"),
            "Inv#": ExtractedField(value="RET-01"),
            "Taxable Value": ExtractedField(value=10000.0),
            "Inv Amount": ExtractedField(value=11800.0)
        }
    )
    rec_id = db.insert_invoice_record(rec)
    db.update_review_status(rec_id, "APPROVED")
    
    pipeline = InvoiceProcessingPipeline(
        template_path=EXCEL_TEMPLATE_PATH,
        sheet_name=EXCEL_SHEET_NAME,
        db_path=str(test_db_path)
    )
    
    # 1. First attempt fails due to simulated write lock
    with patch("openpyxl.Workbook.save", side_effect=PermissionError("Locked")):
        with pytest.raises(RuntimeError):
            pipeline.export_approved_to_excel(target_path=str(test_ledger))
            
    # Check that database recorded failure and status is FAILED
    row_fail = db.get_record_by_id(rec_id)
    assert row_fail["export_status"] == "FAILED"
    assert "Locked" in row_fail["export_error"]
    
    # 2. Retry attempt succeeds
    saved_path, count = pipeline.export_approved_to_excel(target_path=str(test_ledger))
    assert count == 1
    
    row_success = db.get_record_by_id(rec_id)
    assert row_success["export_status"] == "EXPORTED"
    assert row_success["exported_row"] is not None

# 15. Database Migration Compatibility
def test_database_migration_compatibility(tmp_path):
    test_db = tmp_path / "legacy.db"
    
    # Create legacy table missing export columns
    import sqlite3
    conn = sqlite3.connect(test_db)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE invoice_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT,
            vendor_name TEXT,
            invoice_number TEXT,
            total_amount REAL,
            processing_status TEXT DEFAULT 'SUCCESS',
            review_status TEXT DEFAULT 'PENDING'
        )
    """)
    cursor.execute("INSERT INTO invoice_records (file_name, vendor_name, total_amount) VALUES ('old.pdf', 'Old Vendor', 5000.0)")
    conn.commit()
    conn.close()
    
    # Instantiate Database and verify safe migration
    db = Database(str(test_db))
    with db.get_connection() as c:
        cur = c.cursor()
        cur.execute("PRAGMA table_info(invoice_records)")
        cols = {row["name"] for row in cur.fetchall()}
        assert "export_status" in cols
        assert "exported_at" in cols
        assert "exported_row" in cols
        assert "exported_sheet" in cols
        assert "export_error" in cols
        
        # Verify old record is preserved
        cur.execute("SELECT * FROM invoice_records WHERE file_name = 'old.pdf'")
        row = cur.fetchone()
        assert row["vendor_name"] == "Old Vendor"
        assert row["export_status"] == "NOT_EXPORTED"

# 16. Restart and Resynchronization Behavior
def test_restart_and_resynchronization_behavior(tmp_path):
    test_db_path = tmp_path / "restart.db"
    test_ledger = tmp_path / "restart_ledger.xlsx"
    db = Database(str(test_db_path))
    
    # Insert and approve an invoice
    rec = DynamicInvoiceRecord(
        file_name="restart_test.pdf",
        review_status="APPROVED",
        fields={
            "Vendor name ": ExtractedField(value="Restart Vendor"),
            "Inv#": ExtractedField(value="RST-01"),
            "Taxable Value": ExtractedField(value=20000.0),
            "Inv Amount": ExtractedField(value=23600.0)
        }
    )
    rec_id = db.insert_invoice_record(rec)
    db.update_review_status(rec_id, "APPROVED")
    
    # Session 1: Export to Excel
    p1 = InvoiceProcessingPipeline(template_path=EXCEL_TEMPLATE_PATH, sheet_name=EXCEL_SHEET_NAME, db_path=str(test_db_path))
    path1, count1 = p1.export_approved_to_excel(target_path=str(test_ledger))
    assert count1 == 1
    
    # Session 2: Complete application restart (new pipeline instance)
    p2 = InvoiceProcessingPipeline(template_path=EXCEL_TEMPLATE_PATH, sheet_name=EXCEL_SHEET_NAME, db_path=str(test_db_path))
    path2, count2 = p2.export_approved_to_excel(target_path=str(test_ledger))
    
    # Already exported invoice must NOT create another row
    assert count2 == 0
    wb = openpyxl.load_workbook(test_ledger)
    ws = wb[EXCEL_SHEET_NAME]
    
    # Count occurrences of RST-01
    rst_count = sum(1 for r in range(1, ws.max_row + 1) if ws.cell(r, 5).value == "RST-01")
    assert rst_count == 1
