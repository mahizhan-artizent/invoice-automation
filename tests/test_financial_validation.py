from decimal import Decimal
from datetime import date
from core.validator import InvoiceValidator
from core.models import ExtractedField, FinancialAudit, ExcelColumnDefinition
from core.business_rules import BusinessRulesConfig, AccountingPeriodConfig

# 1. CGST + SGST Reconciliation
def test_cgst_and_sgst_reconciliation():
    raw_audit = {
        "taxable_value": 200000.0,
        "cgst": 18000.0,
        "sgst": 18000.0,
        "inv_amount": 236000.0
    }
    fields = {
        "Taxable Value": ExtractedField(value=200000.0),
        "Inv Amount": ExtractedField(value=236000.0)
    }
    report = InvoiceValidator.validate_financials(raw_audit, fields)
    assert report.validation_status == "VERIFIED"
    assert report.status == "VERIFIED"
    assert report.is_math_valid is True
    assert report.cgst == Decimal("18000.00")
    assert report.sgst == Decimal("18000.00")
    assert report.tax_amount == Decimal("36000.00")
    assert report.calculated_invoice_total == Decimal("236000.00")
    assert report.difference == Decimal("0.00")
    assert len(report.missing_components) == 0

# 2. IGST Reconciliation
def test_igst_reconciliation():
    raw_audit = {
        "taxable_value": 150000.0,
        "igst": 27000.0,
        "inv_amount": 177000.0
    }
    fields = {
        "Taxable Value": ExtractedField(value=150000.0),
        "Inv Amount": ExtractedField(value=177000.0)
    }
    report = InvoiceValidator.validate_financials(raw_audit, fields)
    assert report.validation_status == "VERIFIED"
    assert report.is_math_valid is True
    assert report.igst == Decimal("27000.00")
    assert report.cgst is None
    assert report.sgst is None
    assert report.calculated_invoice_total == Decimal("177000.00")
    assert report.difference == Decimal("0.00")

# 3. Missing CGST or SGST (Never infer missing component!)
def test_missing_cgst_or_sgst():
    # Case 3a: CGST present, SGST missing
    raw_audit_missing_sgst = {
        "taxable_value": 200000.0,
        "cgst": 18000.0,
        "sgst": None,  # Intentionally missing
        "inv_amount": 236000.0
    }
    report_a = InvoiceValidator.validate_financials(raw_audit_missing_sgst)
    assert report_a.validation_status == "INCOMPLETE"
    assert report_a.is_math_valid is False
    assert "sgst" in report_a.missing_components
    assert report_a.requires_manual_review is True
    assert "matching SGST is missing" in report_a.explanation

    # Case 3b: SGST present, CGST missing
    raw_audit_missing_cgst = {
        "taxable_value": 200000.0,
        "cgst": None,  # Intentionally missing
        "sgst": 18000.0,
        "inv_amount": 236000.0
    }
    report_b = InvoiceValidator.validate_financials(raw_audit_missing_cgst)
    assert report_b.validation_status == "INCOMPLETE"
    assert report_b.is_math_valid is False
    assert "cgst" in report_b.missing_components
    assert report_b.requires_manual_review is True
    assert "matching CGST is missing" in report_b.explanation

# 4. Missing Tax Information (Never infer missing tax!)
def test_missing_tax_information():
    raw_audit = {
        "taxable_value": 100000.0,
        "inv_amount": 118000.0
        # No tax components provided at all
    }
    fields = {
        "Taxable Value": ExtractedField(value=100000.0),
        "Inv Amount": ExtractedField(value=118000.0)
    }
    report = InvoiceValidator.validate_financials(raw_audit, fields)
    assert report.validation_status == "INCOMPLETE"
    assert report.is_math_valid is False
    assert "tax_amount" in report.missing_components
    assert report.requires_manual_review is True
    assert "Tax breakdown (GST) missing" in report.explanation

# 5. Discounts Before Tax (Pre-tax discount; do NOT double-count!)
def test_discounts_before_tax():
    # Gross: 120,000; Discount: 20,000; Taxable value: 100,000; IGST (18%): 18,000; Total: 118,000
    # In pre-tax discount, taxable_value already reflects the discount.
    # Subtracting discount again would wrongly result in 98,000.
    raw_audit = {
        "taxable_value": 100000.0,
        "igst": 18000.0,
        "discount_amount": 20000.0,
        "discount_timing": "pre_tax",
        "inv_amount": 118000.0
    }
    report = InvoiceValidator.validate_financials(raw_audit)
    assert report.validation_status == "VERIFIED"
    assert report.is_math_valid is True
    assert report.discount_applied_before_tax is True
    assert report.calculated_invoice_total == Decimal("118000.00")
    assert report.difference == Decimal("0.00")

# 6. Discounts After Tax (Post-tax discount deducted from subtotal + tax)
def test_discounts_after_tax():
    # Taxable: 100,000; IGST (18%): 18,000; Post-tax discount: 5,000; Total: 113,000
    # 100,000 + 18,000 - 5,000 = 113,000
    raw_audit = {
        "taxable_value": 100000.0,
        "igst": 18000.0,
        "discount_amount": 5000.0,
        "discount_timing": "post_tax",
        "inv_amount": 113000.0
    }
    report = InvoiceValidator.validate_financials(raw_audit)
    assert report.validation_status == "VERIFIED"
    assert report.is_math_valid is True
    assert report.discount_applied_before_tax is False
    assert report.calculated_invoice_total == Decimal("113000.00")
    assert report.difference == Decimal("0.00")

# 7. Additional Charges & Shipping / Freight
def test_additional_charges_and_freight():
    raw_audit = {
        "taxable_value": 50000.0,
        "igst": 9000.0,
        "additional_charges": 500.0,
        "shipping_charges": 1200.0,
        "inv_amount": 60700.0  # 50,000 + 9,000 + 500 + 1,200 = 60,700
    }
    report = InvoiceValidator.validate_financials(raw_audit)
    assert report.validation_status == "VERIFIED"
    assert report.is_math_valid is True
    assert report.additional_charges == Decimal("500.00")
    assert report.shipping_charges == Decimal("1200.00")
    assert report.calculated_invoice_total == Decimal("60700.00")
    assert report.difference == Decimal("0.00")

# 8. Cess Reconciliation
def test_cess_reconciliation():
    raw_audit = {
        "taxable_value": 500000.0,
        "igst": 140000.0,      # 28% GST
        "cess": 60000.0,       # 12% Compensation Cess
        "inv_amount": 700000.0 # 500,000 + 140,000 + 60,000 = 700,000
    }
    report = InvoiceValidator.validate_financials(raw_audit)
    assert report.validation_status == "VERIFIED"
    assert report.is_math_valid is True
    assert report.cess == Decimal("60000.00")
    assert report.tax_amount == Decimal("200000.00")
    assert report.calculated_invoice_total == Decimal("700000.00")
    assert report.difference == Decimal("0.00")

# 9. Rounding Differences (Within vs Outside Tolerance)
def test_rounding_differences():
    # Case 9a: Within tolerance (0.35 variance <= 1.00 tolerance)
    raw_audit_small_round = {
        "taxable_value": 100000.0,
        "igst": 18000.35,
        "rounding_adjustment": -0.35,
        "inv_amount": 118000.00
    }
    report_a = InvoiceValidator.validate_financials(raw_audit_small_round, rounding_tolerance=Decimal("1.00"))
    assert report_a.validation_status == "VERIFIED"
    assert report_a.difference == Decimal("0.00")

    # Case 9b: Outside tolerance (2.50 variance > 1.00 tolerance)
    raw_audit_exceed_round = {
        "taxable_value": 100000.0,
        "igst": 18000.0,
        "inv_amount": 118002.50  # 2.50 difference without explanation
    }
    report_b = InvoiceValidator.validate_financials(raw_audit_exceed_round, rounding_tolerance=Decimal("1.00"))
    assert report_b.validation_status == "MISMATCH"
    assert report_b.is_math_valid is False
    assert report_b.difference == Decimal("2.50")
    assert "difference=2.50 > tolerance 1.00" in report_b.explanation

# 10. Mismatched Invoice Totals
def test_mismatched_invoice_totals():
    raw_audit = {
        "taxable_value": 100000.0,
        "igst": 18000.0,
        "inv_amount": 150000.0  # Calculated: 118,000; Expected: 150,000
    }
    report = InvoiceValidator.validate_financials(raw_audit)
    assert report.validation_status == "MISMATCH"
    assert report.is_math_valid is False
    assert report.difference == Decimal("32000.00")
    assert "Mismatch: Calculated total 118000.00 differs from Invoice total 150000.00" in report.explanation

# 11. Missing Invoice Date & Accounting Month Handling (Never guess!)
def test_missing_invoice_date_and_accounting_month():
    raw_json = {
        "fields": {
            "Entity": "Nous Infosolutions",
            "Inv Date": None,         # Intentionally missing date
            " Inv accMonth1": None,   # Accounting month not provided
            " Inv acc Month2": None
        },
        "financial_audit": {
            "taxable_value": 50000.0,
            "igst": 9000.0,
            "inv_amount": 59000.0
        }
    }
    columns = [
        ExcelColumnDefinition(col_idx=1, header_name="Entity", clean_key="entity", inferred_type="string"),
        ExcelColumnDefinition(col_idx=2, header_name="Inv Date", clean_key="invdate", inferred_type="date"),
        ExcelColumnDefinition(col_idx=3, header_name=" Inv accMonth1", clean_key="invaccmonth1", inferred_type="string"),
        ExcelColumnDefinition(col_idx=4, header_name=" Inv acc Month2", clean_key="invaccmonth2", inferred_type="date")
    ]
    cfg = BusinessRulesConfig(
        accounting_config=AccountingPeriodConfig(allow_guessing_without_date=False)
    )
    record = InvoiceValidator.validate_and_reconcile(
        raw_json=raw_json,
        columns=columns,
        file_name="invoice_no_date.pdf",
        business_rules=cfg
    )
    
    # 1. Invoice Date must remain None (NEVER today's date!)
    assert record.fields["Inv Date"].value is None
    assert record.fields["Inv Date"].value != date.today().isoformat()

    # 2. Accounting Month must NOT be guessed
    assert record.fields[" Inv accMonth1"].value is None
    assert record.fields[" Inv accMonth1"].is_uncertain is True
    assert "cannot be guessed" in record.fields[" Inv accMonth1"].flag_reason

    assert record.fields[" Inv acc Month2"].value is None
    assert record.fields[" Inv acc Month2"].is_uncertain is True
    assert "cannot be guessed" in record.fields[" Inv acc Month2"].flag_reason

# 12. Decimal Precision (Prevents float drift like 0.1 + 0.2 != 0.3)
def test_decimal_precision():
    assert InvoiceValidator.parse_decimal("1,234.56") == Decimal("1234.56")
    assert InvoiceValidator.parse_decimal("₹ 48,387.00") == Decimal("48387.00")
    assert InvoiceValidator.parse_decimal("  -50.25  ") == Decimal("-50.25")
    assert InvoiceValidator.parse_decimal(None) is None
    assert InvoiceValidator.parse_decimal("invalid") is None

    # Multi-term decimal addition without binary float inaccuracies
    d1 = InvoiceValidator.parse_decimal("0.10")
    d2 = InvoiceValidator.parse_decimal("0.20")
    assert (d1 + d2) == Decimal("0.30")

    # High precision penny quantizing
    raw_audit = {
        "taxable_value": "100000.333",
        "igst": "18000.067",
        "inv_amount": "118000.40"
    }
    report = InvoiceValidator.validate_financials(raw_audit)
    assert report.validation_status == "VERIFIED"
    assert report.taxable_value == Decimal("100000.33")
    assert report.igst == Decimal("18000.07")
    assert report.calculated_invoice_total == Decimal("118000.40")

# 13. Ambiguous Tax Structures (Both IGST and CGST/SGST present)
def test_ambiguous_tax_structures():
    raw_audit = {
        "taxable_value": 100000.0,
        "cgst": 9000.0,
        "sgst": 9000.0,
        "igst": 18000.0,      # Both inter and intra-state taxes present!
        "inv_amount": 136000.0
    }
    report = InvoiceValidator.validate_financials(raw_audit)
    assert report.is_ambiguous_tax is True
    assert report.requires_manual_review is True
    assert any("Ambiguous tax structure" in r for r in report.fields_requiring_manual_review)
    assert "Ambiguous tax structure" in report.explanation

# 14. Structured Validation Report Dictionary Export
def test_structured_validation_report_to_dict():
    raw_audit = {
        "taxable_value": 50000.0,
        "igst": 9000.0,
        "discount_amount": 1000.0,
        "discount_timing": "post_tax",
        "additional_charges": 200.0,
        "shipping_charges": 300.0,
        "rounding_adjustment": -0.20,
        "inv_amount": 58499.80
    }
    report = InvoiceValidator.validate_financials(raw_audit)
    rep_dict = report.to_report_dict()
    assert rep_dict["taxable_value"] == 50000.0
    assert rep_dict["igst"] == 9000.0
    assert rep_dict["cgst"] is None  # Unavailable values are null
    assert rep_dict["sgst"] is None
    assert rep_dict["cess"] is None
    assert rep_dict["discount"] == 1000.0
    assert rep_dict["additional_charges"] == 200.0
    assert rep_dict["shipping_charges"] == 300.0
    assert rep_dict["rounding_adjustment"] == -0.20
    assert rep_dict["extracted_invoice_total"] == 58499.80
    assert rep_dict["calculated_invoice_total"] == 58499.80
    assert rep_dict["difference"] == 0.0
    assert rep_dict["validation_status"] == "VERIFIED"
    assert rep_dict["missing_components"] == []
    assert rep_dict["discount_applied_before_tax"] is False
