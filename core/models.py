from datetime import datetime, date
from decimal import Decimal
from typing import Dict, Any, Optional, List, Union
from pydantic import BaseModel, Field, field_serializer

class ExtractedField(BaseModel):
    """Encapsulates an extracted column value, preserving its source and certainty status."""
    value: Optional[Union[str, float, int, Decimal]] = Field(
        default=None,
        description="The extracted and normalized value matching the column requirement. Missing values are null."
    )
    raw_text: Optional[str] = Field(
        default=None,
        description="Original unformatted text extracted directly from document or context"
    )
    source: str = Field(
        default="unknown",
        description="Source of data: 'invoice_header', 'bill_to', 'table_line_items', 'filename_context', 'directory_context', or 'derived'"
    )
    confidence: float = Field(
        default=1.0,
        description="Confidence score between 0.0 and 1.0"
    )
    is_uncertain: bool = Field(
        default=False,
        description="True if value is ambiguous, blurry, inferred, or conflicting"
    )
    flag_reason: Optional[str] = Field(
        default=None,
        description="Detailed reason if value is missing, uncertain, or synthesized"
    )

    @field_serializer("value")
    def serialize_value(self, v):
        if isinstance(v, Decimal):
            return float(v)
        return v

class FinancialAudit(BaseModel):
    """
    Financial integrity validation results using Decimal arithmetic.
    Distinguishes verified, incomplete, and mismatched calculations.
    Supports comprehensive GST/tax components, pre/post-tax discounts, cess, and freight.
    """
    # Validation status: VERIFIED, INCOMPLETE, MISMATCH
    status: str = Field(
        default="VERIFIED",
        description="'VERIFIED' (math balanced), 'INCOMPLETE' (missing tax or subtotal), 'MISMATCH' (math discrepancies)"
    )
    validation_status: str = Field(
        default="VERIFIED",
        description="Structured validation status: 'VERIFIED', 'INCOMPLETE', 'MISMATCH'"
    )

    # Core components
    taxable_value: Optional[Decimal] = None
    cgst: Optional[Decimal] = None
    sgst: Optional[Decimal] = None
    igst: Optional[Decimal] = None
    cess: Optional[Decimal] = None
    tax_amount: Optional[Decimal] = None

    # Discounts and timings
    discount: Optional[Decimal] = None
    discount_amount: Optional[Decimal] = None  # Backward-compatible alias
    discount_applied_before_tax: Optional[bool] = None

    # Additional charges, freight, and rounding
    additional_charges: Optional[Decimal] = None
    extra_charges: Optional[Decimal] = None     # Backward-compatible alias
    shipping_charges: Optional[Decimal] = None
    rounding_adjustment: Optional[Decimal] = None

    # Extracted vs Calculated Totals
    inv_amount: Optional[Decimal] = None       # Backward-compatible alias
    extracted_invoice_total: Optional[Decimal] = None
    calculated_total: Optional[Decimal] = None # Backward-compatible alias
    calculated_invoice_total: Optional[Decimal] = None

    # Variance / Reconciliation Difference
    variance: Decimal = Decimal("0.00")         # Backward-compatible alias
    difference: Optional[Decimal] = Decimal("0.00")

    # Math and tax metadata
    is_math_valid: bool = True
    tax_rate_percent: Optional[Decimal] = None

    # Explanations and review flags
    notes: Optional[str] = None                # Backward-compatible alias
    explanation: Optional[str] = None
    missing_components: List[str] = Field(default_factory=list)
    fields_requiring_manual_review: List[str] = Field(default_factory=list)
    requires_manual_review: bool = False
    is_ambiguous_tax: bool = False

    @field_serializer(
        "taxable_value", "cgst", "sgst", "igst", "cess", "tax_amount",
        "discount", "discount_amount", "additional_charges", "extra_charges",
        "shipping_charges", "rounding_adjustment", "inv_amount",
        "extracted_invoice_total", "calculated_total", "calculated_invoice_total",
        "variance", "difference", "tax_rate_percent"
    )
    def serialize_decimal(self, v):
        if isinstance(v, Decimal):
            return float(v)
        return v

    def to_report_dict(self) -> Dict[str, Any]:
        """Returns structured validation report with exact nulls for unavailable values."""
        return {
            "taxable_value": float(self.taxable_value) if self.taxable_value is not None else None,
            "cgst": float(self.cgst) if self.cgst is not None else None,
            "sgst": float(self.sgst) if self.sgst is not None else None,
            "igst": float(self.igst) if self.igst is not None else None,
            "cess": float(self.cess) if self.cess is not None else None,
            "discount": float(self.discount) if self.discount is not None else None,
            "additional_charges": float(self.additional_charges) if self.additional_charges is not None else None,
            "shipping_charges": float(self.shipping_charges) if self.shipping_charges is not None else None,
            "rounding_adjustment": float(self.rounding_adjustment) if self.rounding_adjustment is not None else None,
            "extracted_invoice_total": float(self.extracted_invoice_total) if self.extracted_invoice_total is not None else None,
            "calculated_invoice_total": float(self.calculated_invoice_total) if self.calculated_invoice_total is not None else None,
            "difference": float(self.difference) if self.difference is not None else None,
            "validation_status": self.validation_status,
            "explanation": self.explanation,
            "missing_components": self.missing_components,
            "fields_requiring_manual_review": self.fields_requiring_manual_review,
            "requires_manual_review": self.requires_manual_review,
            "is_ambiguous_tax": self.is_ambiguous_tax,
            "discount_applied_before_tax": self.discount_applied_before_tax
        }

class DynamicInvoiceRecord(BaseModel):
    """Complete record for an invoice extracted dynamically according to template columns."""
    id: Optional[int] = None
    file_name: str
    file_path: Optional[str] = None
    relative_path: Optional[str] = None
    folder_context: Optional[str] = None
    file_hash: Optional[str] = None  # SHA-256
    
    processing_status: str = Field(default="SUCCESS", description="'SUCCESS', 'FAILED', 'SKIPPED'")
    review_status: str = Field(default="PENDING", description="'PENDING', 'APPROVED', 'REJECTED'")
    export_status: str = Field(default="NOT_EXPORTED", description="'NOT_EXPORTED', 'EXPORTED', 'FAILED'")
    exported_at: Optional[str] = None
    exported_sheet: Optional[str] = None
    exported_row: Optional[int] = None
    export_error: Optional[str] = None
    
    is_file_duplicate: bool = False
    is_invoice_duplicate: bool = False
    is_duplicate: bool = False  # Backward compatibility flag
    duplicate_reason: Optional[str] = None
    error_message: Optional[str] = None
    
    fields: Dict[str, ExtractedField] = Field(
        default_factory=dict,
        description="Mapping of exact Excel column header names to their extracted field metadata"
    )
    financial_audit: FinancialAudit = Field(default_factory=FinancialAudit)
    validation_flags: List[str] = Field(default_factory=list)
    raw_response: Optional[Dict[str, Any]] = None
    extracted_at: str = Field(default_factory=lambda: datetime.now().isoformat())

class ExcelColumnDefinition(BaseModel):
    """Discovered column metadata from the Excel template."""
    col_idx: int
    header_name: str
    clean_key: str
    inferred_type: str
    number_format: Optional[str] = None
    sample_values: List[str] = Field(default_factory=list)
    description: Optional[str] = None
