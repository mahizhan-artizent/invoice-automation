import tempfile
from pathlib import Path
from core.pipeline import calculate_sha256, InvoiceProcessingPipeline
from core.database import Database
from core.models import DynamicInvoiceRecord, ExtractedField, FinancialAudit

def test_sha256_calculation():
    with tempfile.NamedTemporaryFile(delete=False) as f:
        f.write(b"Test Invoice Content")
        f_path = Path(f.name)
        
    try:
        hash1 = calculate_sha256(f_path)
        hash2 = calculate_sha256(f_path)
        assert hash1 == hash2
        assert len(hash1) == 64
    finally:
        f_path.unlink()

def test_file_level_duplicate_vs_invoice_duplicate(tmp_path):
    db_file = tmp_path / "test_dup.db"
    db = Database(str(db_file))
    
    # 1. Insert original invoice record
    rec1 = DynamicInvoiceRecord(
        file_name="inv1.pdf",
        relative_path="inv1.pdf",
        file_hash="hash_abc_123",
        processing_status="SUCCESS",
        review_status="PENDING",
        fields={
            "Vendor name ": ExtractedField(value="Tech Vendor"),
            "Inv#": ExtractedField(value="INV-999"),
            "Inv Amount": ExtractedField(value=5000.0)
        },
        financial_audit=FinancialAudit(inv_amount=5000.0)
    )
    rec1_id = db.insert_invoice_record(rec1)
    
    # Check document lookup by hash (file-level check)
    record_by_hash = db.get_record_by_hash("hash_abc_123")
    assert record_by_hash is not None
    assert record_by_hash["id"] == rec1_id
    
    # Check file with different hash: not a file duplicate
    assert db.get_record_by_hash("hash_xyz_789") is None
    
    # Check invoice-level duplicate (same vendor + inv#):
    is_inv_dup, reason = db.check_invoice_duplicate("Tech Vendor", "INV-999", 5000.0)
    assert is_inv_dup is True
    assert f"Record #{rec1_id}" in reason
    
    # Different vendor: not an invoice duplicate
    is_inv_dup2, _ = db.check_invoice_duplicate("Other Vendor", "INV-999", 5000.0)
    assert is_inv_dup2 is False

