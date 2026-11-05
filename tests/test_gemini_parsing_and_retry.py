import json
from unittest.mock import MagicMock
from core.pipeline import InvoiceProcessingPipeline
from core.gemini_service import GeminiDocumentService
from core.validator import InvoiceValidator
from core.database import Database
from core.template_engine import TemplateEngine

def test_gemini_parsing_nested_extracted_data():
    engine = TemplateEngine()
    cols = engine.get_column_definitions()
    
    # Simulate Gemini nested response
    raw_response = {
        "extracted_data": [
            {
                "Entity": {"value": "Acme Corp", "confidence": 0.98},
                "Vendor name ": {"value": "Global Tech", "confidence": 0.99},
                "Inv#": {"value": "GT-1001", "confidence": 0.99},
                "Inv Date": {"value": "2026-05-20", "confidence": 0.99},
                "TaxableValue": {"value": 10000.0, "confidence": 0.99},  # No space
                "Inv Amount": {"value": 11800.0, "confidence": 0.99}
            }
        ],
        "financial_audit": {
            "taxable_value": 10000.0,
            "tax_amount": 1800.0,
            "inv_amount": 11800.0
        }
    }
    
    record = InvoiceValidator.validate_and_reconcile(raw_response, cols, "sample.pdf")
    
    # Verify values parsed cleanly
    assert record.fields["Entity"].value == "Acme Corp"
    assert record.fields["Vendor name "].value == "Global Tech"
    assert record.fields["Inv#"].value == "GT-1001"
    assert record.fields["Taxable Value"].value == 10000.0
    assert record.fields["Inv Amount"].value == 11800.0
    assert record.financial_audit.status == "VERIFIED"

def test_failure_recording_and_retry(tmp_path):
    db_file = tmp_path / "test_retry.db"
    dummy_invoice = tmp_path / "fail_invoice.pdf"
    dummy_invoice.write_bytes(b"%PDF-1.4 dummy invoice")
    
    # Mock Gemini service to simulate API failure
    mock_gemini = MagicMock(spec=GeminiDocumentService)
    mock_gemini.extract_invoice_data.side_effect = RuntimeError("Simulated Gemini 503 Overload")
    
    pipeline = InvoiceProcessingPipeline(db_path=str(db_file))
    pipeline.gemini_service = mock_gemini
    
    # 1. First run fails
    rec1 = pipeline.process_single_invoice(dummy_invoice, skip_if_processed=False)
    assert rec1.processing_status == "FAILED"
    assert "Simulated Gemini 503 Overload" in rec1.error_message
    
    # Check DB recorded failure
    db_rec1 = pipeline.db.get_record_by_hash(rec1.file_hash)
    assert db_rec1 is not None
    assert db_rec1["processing_status"] == "FAILED"
    assert "Simulated Gemini 503 Overload" in db_rec1["error_message"]
    
    # 2. Retry with recovered Gemini mock response
    mock_gemini.extract_invoice_data.side_effect = None
    mock_gemini.extract_invoice_data.return_value = (
        {
            "fields": {
                "Vendor name ": {"value": "Recovered Vendor"},
                "Inv#": {"value": "REC-001"},
                "Inv Date": {"value": "2026-06-01"},
                "Taxable Value": {"value": 2000.0},
                "Inv Amount": {"value": 2360.0}
            },
            "financial_audit": {
                "tax_amount": 360.0
            }
        },
        "gemini-3.7-flash",
        1.2
    )
    
    # Second run should allow retry because previous attempt was FAILED
    rec2 = pipeline.process_single_invoice(dummy_invoice, skip_if_processed=True)
    assert rec2.processing_status == "SUCCESS"
    assert rec2.fields["Vendor name "].value == "Recovered Vendor"
    
    # Verify DB now contains the successful retry
    db_rec2 = pipeline.db.get_record_by_hash(rec2.file_hash)
    assert db_rec2["processing_status"] == "SUCCESS"
    assert db_rec2["vendor_name"] == "Recovered Vendor"

