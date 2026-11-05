from datetime import date
from core.validator import InvoiceValidator
from core.template_engine import TemplateEngine
from core.models import ExcelColumnDefinition

def test_date_normalization_formats():
    assert InvoiceValidator._normalize_date("2026-06-01") == "2026-06-01"
    assert InvoiceValidator._normalize_date("01-06-2026") == "2026-06-01"
    assert InvoiceValidator._normalize_date("01/06/2026") == "2026-06-01"
    assert InvoiceValidator._normalize_date("1 June 2026") == "2026-06-01"
    assert InvoiceValidator._normalize_date("Jun 01, 2026") == "2026-06-01"
    assert InvoiceValidator._normalize_date("2026-06-01T00:00:00") == "2026-06-01"

def test_no_fallback_to_today_when_date_missing():
    # Crucial test: When invoice date is missing from document, it must return None
    raw_fields = {
        "Vendor": "ABC Corp",
        "Amount": 1000
    }
    found_date = InvoiceValidator._find_date_value(raw_fields)
    assert found_date is None, "Should NEVER fallback to date.today()!"

def test_uncertain_accounting_month_when_derived():
    engine = TemplateEngine()
    cols = engine.get_column_definitions()
    
    # Simulate LLM response where Inv Date is extracted but accounting month is missing
    raw_json = {
        "fields": {
            "Vendor name ": {"value": "Acme Inc"},
            "Inv#": {"value": "INV-100"},
            "Inv Date": {"value": "2026-05-15"},
            "Taxable Value": {"value": 1000.0},
            "Inv Amount": {"value": 1180.0}
        },
        "financial_audit": {"tax_amount": 180.0}
    }
    
    record = InvoiceValidator.validate_and_reconcile(raw_json, cols, "test.pdf")
    
    # Check that Inv accMonth1 was derived AND flagged as uncertain for manual review
    acc_month1 = record.fields.get(" Inv accMonth1")
    assert acc_month1 is not None
    assert acc_month1.value == "26_05"
    assert acc_month1.is_uncertain is True
    assert "manual review" in acc_month1.flag_reason.lower()

