import re
from datetime import datetime, date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Dict, Any, List, Optional, Tuple, Union

from core.models import ExtractedField, FinancialAudit, DynamicInvoiceRecord, ExcelColumnDefinition
from core.business_rules import BusinessRulesConfig

def canonical_key(s: Any) -> str:
    """Normalize any string to lowercase alphanumeric only for robust dictionary matching."""
    return re.sub(r'[^a-z0-9]', '', str(s).lower())

class InvoiceValidator:
    """
    Performs Pydantic-based structural validation, Decimal-precision financial integrity verification,
    and business rule enforcement on extracted invoice data.
    """

    SEMANTIC_ALIASES = {
        "taxablevalue": ["taxablevalue", "taxableamount", "subtotal", "pretaxamount", "taxable", "netamount"],
        "invamount": ["invamount", "invoiceamount", "totalamount", "grandtotal", "total", "amount", "balancedue"],
        "inv": ["inv", "invno", "invnumber", "invoiceno", "invoicenumber", "invoicenum", "billno"],
        "invdate": ["invdate", "invoicedate", "dateofissue", "billdate"],
        "invreceiveddate": ["invreceiveddate", "receiveddate", "receiptdate"],
        "vendorname": ["vendorname", "vendor", "supplier", "suppliername", "contractor", "serviceprovider"],
        "entity": ["entity", "billedto", "customer", "customername", "companyname", "billto"],
        "noe": ["noe", "natureofexpense", "description", "itemdescription", "particulars", "service", "services"],
        "sender": ["sender", "consultant", "signatory", "authorizedsignatory", "submittedby", "contactperson"],
        "invaccmonth1": ["invaccmonth1", "accmonth1", "accountingmonth1", "accmonth"],
        "invaccmonth2": ["invaccmonth2", "accmonth2", "accountingmonth2"]
    }

    def __init__(self, business_rules: Optional[BusinessRulesConfig] = None):
        self.business_rules = business_rules or BusinessRulesConfig()

    @staticmethod
    def parse_decimal(val: Any) -> Optional[Decimal]:
        """Converts arbitrary input to 2-decimal-place Decimal using ROUND_HALF_UP."""
        if val is None:
            return None
        if isinstance(val, Decimal):
            return val.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        cleaned = re.sub(r'[^\d.-]', '', str(val).strip())
        if not cleaned or cleaned == "-" or cleaned == ".":
            return None
        try:
            return Decimal(cleaned).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        except (InvalidOperation, ValueError):
            return None

    @classmethod
    def unwrap_raw_response(cls, raw_json: Any) -> Tuple[Dict[str, Any], Dict[str, Any]]:
        """
        Safely extracts the dictionary of extracted fields and the financial audit dict
        regardless of whether Gemini nested them inside a list, 'extracted_data', 'fields', etc.
        """
        if isinstance(raw_json, list):
            raw_json = raw_json[0] if raw_json else {}

        if not isinstance(raw_json, dict):
            return {}, {}

        financial_audit_raw = raw_json.get("financial_audit", {})
        if not isinstance(financial_audit_raw, dict):
            financial_audit_raw = {}

        # Look for nested containers
        for container_key in ("extracted_data", "fields", "data", "invoices", "invoice_data", "columns"):
            if container_key in raw_json:
                inner = raw_json[container_key]
                if isinstance(inner, list) and len(inner) > 0 and isinstance(inner[0], dict):
                    return inner[0], financial_audit_raw
                elif isinstance(inner, dict):
                    return inner, financial_audit_raw

        return raw_json, financial_audit_raw

    @classmethod
    def _find_field_value(
        cls,
        raw_fields: Dict[str, Any],
        col_header: str
    ) -> Optional[Any]:
        """Finds field entry matching column header using exact, stripped, and canonical aliases."""
        if col_header in raw_fields:
            return raw_fields[col_header]

        target_canon = canonical_key(col_header)
        canon_map = {canonical_key(k): v for k, v in raw_fields.items()}
        if target_canon in canon_map:
            return canon_map[target_canon]

        aliases = cls.SEMANTIC_ALIASES.get(target_canon, [])
        for alias in aliases:
            if alias in canon_map:
                return canon_map[alias]

        return None

    @classmethod
    def validate_and_reconcile(
        cls,
        raw_json: Any,
        columns: List[ExcelColumnDefinition],
        file_name: str,
        file_path: Optional[str] = None,
        relative_path: Optional[str] = None,
        folder_context: Optional[str] = None,
        business_rules: Optional[BusinessRulesConfig] = None
    ) -> DynamicInvoiceRecord:
        """
        Transforms raw Gemini response into a strictly validated DynamicInvoiceRecord.
        Performs Decimal-precision financial checks and flags missing or uncertain data.
        """
        rules_cfg = business_rules or BusinessRulesConfig()
        extracted_fields: Dict[str, ExtractedField] = {}
        validation_flags: List[str] = []

        raw_fields, raw_audit = cls.unwrap_raw_response(raw_json)

        # Pre-extract common date for derivations if available
        inv_date = cls._find_date_value(raw_fields)

        # Map each discovered Excel column
        for col in columns:
            col_name = col.header_name
            target_canon = canonical_key(col_name)
            rule = rules_cfg.get_rule_for_key(target_canon)

            field_entry = cls._find_field_value(raw_fields, col_name)

            # Fallback to financial_audit if Taxable Value or Inv Amount wasn't in field dict
            if field_entry is None:
                if target_canon in ("taxablevalue", "taxableamount") and "taxable_value" in raw_audit:
                    field_entry = {
                        "value": raw_audit.get("taxable_value"),
                        "source": "financial_summary",
                        "confidence": 0.95,
                        "is_uncertain": False,
                        "raw_text": str(raw_audit.get("taxable_value"))
                    }
                elif target_canon in ("invamount", "totalamount") and "inv_amount" in raw_audit:
                    field_entry = {
                        "value": raw_audit.get("inv_amount"),
                        "source": "financial_summary",
                        "confidence": 0.95,
                        "is_uncertain": False,
                        "raw_text": str(raw_audit.get("inv_amount"))
                    }

            if isinstance(field_entry, dict):
                raw_val = field_entry.get("value")
                source = field_entry.get("source", "invoice_document")
                confidence = float(field_entry.get("confidence", 0.95))
                is_uncertain = bool(field_entry.get("is_uncertain", False))
                flag_reason = field_entry.get("flag_reason")
                raw_text = field_entry.get("raw_text", str(raw_val) if raw_val is not None else None)
            else:
                raw_val = field_entry
                source = "invoice_document" if field_entry is not None else "missing"
                confidence = 0.90 if field_entry is not None else 0.0
                is_uncertain = (field_entry is None)
                flag_reason = "Field missing from document" if field_entry is None else None
                raw_text = str(raw_val) if raw_val is not None else None

            # Normalization based on inferred column type
            norm_val = raw_val
            if col.inferred_type in ("number", "integer"):
                dec = cls.parse_decimal(raw_val)
                norm_val = int(dec) if (col.inferred_type == "integer" and dec is not None) else (float(dec) if dec is not None else None)
            elif col.inferred_type == "date" and norm_val:
                norm_val = cls._normalize_date(str(norm_val))

            # Business rule derivations for accounting month
            acc_cfg = rules_cfg.accounting_config
            if target_canon == "invaccmonth1":
                if not norm_val and inv_date:
                    norm_val = inv_date.strftime(acc_cfg.code_format)
                    source = "derived"
                    confidence = 0.85
                    is_uncertain = acc_cfg.require_review_for_derived
                    flag_reason = "Derived from Invoice Date; marked for manual review." if is_uncertain else None
                elif not norm_val and not inv_date:
                    # Strict requirement: Never guess accounting month when invoice date is missing
                    norm_val = None
                    source = "missing"
                    confidence = 0.0
                    is_uncertain = True
                    flag_reason = "Missing invoice date; accounting month cannot be guessed and is left blank for manual review."
                elif norm_val and rule and rule.requires_review_if_derived:
                    is_uncertain = True
                    flag_reason = flag_reason or "Accounting month requires verification."
            elif target_canon == "invaccmonth2":
                if not norm_val and inv_date:
                    norm_val = inv_date.strftime(acc_cfg.date_format)
                    source = "derived"
                    confidence = 0.85
                    is_uncertain = acc_cfg.require_review_for_derived
                    flag_reason = "Derived from Invoice Date; marked for manual review." if is_uncertain else None
                elif not norm_val and not inv_date:
                    # Strict requirement: Never guess accounting month date when invoice date is missing
                    norm_val = None
                    source = "missing"
                    confidence = 0.0
                    is_uncertain = True
                    flag_reason = "Missing invoice date; accounting month date cannot be guessed and is left blank for manual review."
                elif norm_val and rule and rule.requires_review_if_derived:
                    is_uncertain = True
                    flag_reason = flag_reason or "Accounting month requires verification."

            # Sequencer handling (S.No)
            if target_canon in ("sno", "slno", "serialno") or (rule and rule.derivation_type == "sequencer"):
                norm_val = None  # Always null in extracted record; auto-generated at Excel write
                source = "derived"
                confidence = 1.0
                is_uncertain = False
                flag_reason = "Auto-generated by ledger sequencer"

            # Missing and low confidence checks
            if norm_val is None and target_canon not in ("sno", "slno", "serialno"):
                is_uncertain = True
                flag_reason = flag_reason or f"Missing required value for '{col_name.strip()}'"
                validation_flags.append(f"Missing '{col_name.strip()}'")

            if confidence < 0.6 and norm_val is not None:
                is_uncertain = True
                validation_flags.append(f"Low confidence ({confidence:.2f}) on '{col_name.strip()}'")

            extracted_fields[col_name] = ExtractedField(
                value=norm_val,
                raw_text=raw_text,
                source=source,
                confidence=confidence,
                is_uncertain=is_uncertain,
                flag_reason=flag_reason
            )

        # Decimal Financial Cross-Footing Validation
        audit = cls.validate_financials(
            raw_audit=raw_audit,
            fields=extracted_fields,
            rounding_tolerance=rules_cfg.rounding_tolerance
        )
        if audit.status != "VERIFIED":
            validation_flags.append(f"Financial verification: {audit.explanation or audit.notes}")
        if audit.requires_manual_review:
            for review_item in audit.fields_requiring_manual_review:
                validation_flags.append(f"Financial Review Required: {review_item}")

        return DynamicInvoiceRecord(
            file_name=file_name,
            file_path=file_path,
            relative_path=relative_path or file_name,
            folder_context=folder_context,
            fields=extracted_fields,
            financial_audit=audit,
            validation_flags=validation_flags,
            raw_response=raw_json
        )

    @classmethod
    def _audit_finances(
        cls,
        raw_audit: Dict[str, Any],
        fields: Optional[Dict[str, ExtractedField]] = None,
        rounding_tolerance: Optional[Decimal] = None
    ) -> FinancialAudit:
        """Backward-compatible delegate to validate_financials."""
        return cls.validate_financials(raw_audit, fields, rounding_tolerance)

    @classmethod
    def validate_financials(
        cls,
        raw_audit: Dict[str, Any],
        fields: Optional[Dict[str, ExtractedField]] = None,
        rounding_tolerance: Optional[Decimal] = None
    ) -> FinancialAudit:
        """
        Robust, Decimal-precision financial reconciliation engine for invoices.
        Supports:
        - Taxable value, CGST, SGST, IGST, Cess
        - Discounts (detecting pre-tax vs post-tax to avoid double-counting)
        - Additional charges, freight/shipping charges, rounding adjustments
        - Final invoice amount
        
        Statuses:
        - 'VERIFIED': All required components independently available and reconcile within tolerance.
        - 'INCOMPLETE': One or more required components missing (tax breakdown, subtotal, etc.).
        - 'MISMATCH': Independently available components do not reconcile within tolerance.
        
        Flags ambiguous tax structures (both IGST and CGST/SGST present) for manual review.
        """
        tolerance = rounding_tolerance or Decimal("1.00")
        fields = fields or {}
        missing_components: List[str] = []
        review_fields: List[str] = []
        requires_manual_review = False
        is_ambiguous_tax = False

        # 1. Parse all monetary components using exact Decimal arithmetic
        taxable_val = None
        inv_amt = None

        # Check extracted fields first for Taxable Value and Invoice Amount
        for k, f in fields.items():
            k_canon = canonical_key(k)
            if k_canon in ("taxablevalue", "taxableamount") and f.value is not None:
                taxable_val = cls.parse_decimal(f.value)
            elif k_canon in ("invamount", "totalamount") and f.value is not None:
                inv_amt = cls.parse_decimal(f.value)

        # Fallback to raw_audit if not found in fields
        if taxable_val is None:
            taxable_val = cls.parse_decimal(raw_audit.get("taxable_value") or raw_audit.get("subtotal") or raw_audit.get("taxable_amount"))
        if inv_amt is None:
            inv_amt = cls.parse_decimal(raw_audit.get("inv_amount") or raw_audit.get("total_amount") or raw_audit.get("grand_total"))

        # Taxes
        cgst = cls.parse_decimal(raw_audit.get("cgst"))
        sgst = cls.parse_decimal(raw_audit.get("sgst"))
        igst = cls.parse_decimal(raw_audit.get("igst"))
        cess = cls.parse_decimal(raw_audit.get("cess"))
        tax_amt = cls.parse_decimal(raw_audit.get("tax_amount") or raw_audit.get("total_tax"))

        # Discounts
        discount = cls.parse_decimal(raw_audit.get("discount") or raw_audit.get("discount_amount")) or Decimal("0.00")
        discount_timing = raw_audit.get("discount_timing") or raw_audit.get("discount_type")  # 'pre_tax', 'post_tax', or None

        # Additional charges and freight
        additional_charges = cls.parse_decimal(raw_audit.get("additional_charges") or raw_audit.get("extra_charges")) or Decimal("0.00")
        shipping_charges = cls.parse_decimal(raw_audit.get("shipping_charges") or raw_audit.get("freight_charges") or raw_audit.get("freight")) or Decimal("0.00")
        rounding_adj = cls.parse_decimal(raw_audit.get("rounding_adjustment") or raw_audit.get("round_off")) or Decimal("0.00")

        # 2. Check for missing primary anchors (Taxable Value or Grand Total)
        if taxable_val is None:
            missing_components.append("taxable_value")
        if inv_amt is None:
            missing_components.append("extracted_invoice_total")

        if missing_components:
            notes = f"Incomplete: Missing required component(s): {', '.join(missing_components)}. Financial verification cannot proceed."
            return FinancialAudit(
                status="INCOMPLETE",
                validation_status="INCOMPLETE",
                taxable_value=taxable_val,
                cgst=cgst,
                sgst=sgst,
                igst=igst,
                cess=cess,
                tax_amount=tax_amt,
                discount=discount if discount > Decimal("0.00") else None,
                discount_amount=discount if discount > Decimal("0.00") else None,
                additional_charges=additional_charges if additional_charges > Decimal("0.00") else None,
                extra_charges=additional_charges if additional_charges > Decimal("0.00") else None,
                shipping_charges=shipping_charges if shipping_charges > Decimal("0.00") else None,
                rounding_adjustment=rounding_adj if rounding_adj != Decimal("0.00") else None,
                inv_amount=inv_amt,
                extracted_invoice_total=inv_amt,
                calculated_total=None,
                calculated_invoice_total=None,
                is_math_valid=False,
                variance=Decimal("0.00"),
                difference=None,
                notes=notes,
                explanation=notes,
                missing_components=missing_components,
                fields_requiring_manual_review=review_fields,
                requires_manual_review=True
            )

        # 3. Check Tax Components & Missing Tax Rules
        has_cgst = (cgst is not None)
        has_sgst = (sgst is not None)
        has_igst = (igst is not None)
        has_cess = (cess is not None)
        has_tax_amt = (tax_amt is not None)

        # Rule 3A: Intra-State GST requires both CGST and SGST. If one is missing, NEVER infer!
        if has_cgst and not has_sgst:
            missing_components.append("sgst")
            notes = f"Incomplete: CGST is present (₹{cgst}) but matching SGST is missing. Tax cannot be inferred."
            return FinancialAudit(
                status="INCOMPLETE",
                validation_status="INCOMPLETE",
                taxable_value=taxable_val,
                cgst=cgst,
                sgst=None,
                igst=igst,
                cess=cess,
                tax_amount=tax_amt,
                discount=discount if discount > Decimal("0.00") else None,
                discount_amount=discount if discount > Decimal("0.00") else None,
                additional_charges=additional_charges if additional_charges > Decimal("0.00") else None,
                extra_charges=additional_charges if additional_charges > Decimal("0.00") else None,
                shipping_charges=shipping_charges if shipping_charges > Decimal("0.00") else None,
                rounding_adjustment=rounding_adj if rounding_adj != Decimal("0.00") else None,
                inv_amount=inv_amt,
                extracted_invoice_total=inv_amt,
                calculated_total=None,
                calculated_invoice_total=None,
                is_math_valid=False,
                variance=Decimal("0.00"),
                difference=None,
                notes=notes,
                explanation=notes,
                missing_components=missing_components,
                fields_requiring_manual_review=["sgst"],
                requires_manual_review=True
            )

        if has_sgst and not has_cgst:
            missing_components.append("cgst")
            notes = f"Incomplete: SGST is present (₹{sgst}) but matching CGST is missing. Tax cannot be inferred."
            return FinancialAudit(
                status="INCOMPLETE",
                validation_status="INCOMPLETE",
                taxable_value=taxable_val,
                cgst=None,
                sgst=sgst,
                igst=igst,
                cess=cess,
                tax_amount=tax_amt,
                discount=discount if discount > Decimal("0.00") else None,
                discount_amount=discount if discount > Decimal("0.00") else None,
                additional_charges=additional_charges if additional_charges > Decimal("0.00") else None,
                extra_charges=additional_charges if additional_charges > Decimal("0.00") else None,
                shipping_charges=shipping_charges if shipping_charges > Decimal("0.00") else None,
                rounding_adjustment=rounding_adj if rounding_adj != Decimal("0.00") else None,
                inv_amount=inv_amt,
                extracted_invoice_total=inv_amt,
                calculated_total=None,
                calculated_invoice_total=None,
                is_math_valid=False,
                variance=Decimal("0.00"),
                difference=None,
                notes=notes,
                explanation=notes,
                missing_components=missing_components,
                fields_requiring_manual_review=["cgst"],
                requires_manual_review=True
            )

        # Rule 3B: Missing tax information entirely
        individual_taxes = [t for t in (cgst, sgst, igst, cess) if t is not None]
        if not individual_taxes and tax_amt is None:
            # Check if invoice is explicitly marked exempt or 0% tax
            is_explicit_zero_tax = (
                raw_audit.get("is_exempt") is True or
                raw_audit.get("tax_exempt") is True or
                "exempt" in str(raw_audit.get("notes", "")).lower() or
                "0%" in str(raw_audit.get("notes", "")).lower() or
                "rcm" in str(raw_audit.get("notes", "")).lower()
            )
            if is_explicit_zero_tax and taxable_val == inv_amt:
                total_tax = Decimal("0.00")
            else:
                missing_components.append("tax_amount")
                notes = "Incomplete: Tax breakdown (GST) missing from invoice. Tax cannot be inferred."
                return FinancialAudit(
                    status="INCOMPLETE",
                    validation_status="INCOMPLETE",
                    taxable_value=taxable_val,
                    cgst=None,
                    sgst=None,
                    igst=None,
                    cess=None,
                    tax_amount=None,
                    discount=discount if discount > Decimal("0.00") else None,
                    discount_amount=discount if discount > Decimal("0.00") else None,
                    additional_charges=additional_charges if additional_charges > Decimal("0.00") else None,
                    extra_charges=additional_charges if additional_charges > Decimal("0.00") else None,
                    shipping_charges=shipping_charges if shipping_charges > Decimal("0.00") else None,
                    rounding_adjustment=rounding_adj if rounding_adj != Decimal("0.00") else None,
                    inv_amount=inv_amt,
                    extracted_invoice_total=inv_amt,
                    calculated_total=None,
                    calculated_invoice_total=None,
                    is_math_valid=False,
                    variance=Decimal("0.00"),
                    difference=None,
                    notes=notes,
                    explanation=notes,
                    missing_components=missing_components,
                    fields_requiring_manual_review=["tax_amount"],
                    requires_manual_review=True
                )
        else:
            # Rule 3C: Ambiguous tax structures
            # Check if both IGST > 0 and (CGST > 0 or SGST > 0) are present on the same invoice
            has_positive_igst = (igst is not None and igst > Decimal("0.00"))
            has_positive_cgst_sgst = (
                (cgst is not None and cgst > Decimal("0.00")) or
                (sgst is not None and sgst > Decimal("0.00"))
            )
            if has_positive_igst and has_positive_cgst_sgst:
                is_ambiguous_tax = True
                requires_manual_review = True
                review_fields.append("Ambiguous tax structure: Both inter-state (IGST) and intra-state (CGST/SGST) taxes present on single invoice")

            # Sum available individual taxes
            combined_individual = sum(individual_taxes, Decimal("0.00"))
            if tax_amt is None or tax_amt == Decimal("0.00"):
                total_tax = combined_individual
            else:
                total_tax = tax_amt

        # 4. Handle Discounts (Pre-tax vs Post-tax without double-counting)
        total_charges = additional_charges + shipping_charges
        
        # Hypothesis A: Post-tax discount (applied after taxes to subtotal + taxes)
        calc_post_tax = (taxable_val + total_tax + total_charges - discount + rounding_adj).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        var_post_tax = abs(calc_post_tax - inv_amt).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

        # Hypothesis B: Pre-tax discount (already deducted to reach taxable_val; do NOT subtract again)
        calc_pre_tax = (taxable_val + total_tax + total_charges + rounding_adj).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        var_pre_tax = abs(calc_pre_tax - inv_amt).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

        applied_before_tax: Optional[bool] = None
        if str(discount_timing).lower() in ("pre_tax", "pretax", "before_tax", "before"):
            calculated_total = calc_pre_tax
            variance = var_pre_tax
            applied_before_tax = True
        elif str(discount_timing).lower() in ("post_tax", "posttax", "after_tax", "after"):
            calculated_total = calc_post_tax
            variance = var_post_tax
            applied_before_tax = False
        elif discount > Decimal("0.00"):
            # Auto-detect discount timing based on mathematical reconciliation
            if var_post_tax <= tolerance and var_pre_tax > tolerance:
                calculated_total = calc_post_tax
                variance = var_post_tax
                applied_before_tax = False
            elif var_pre_tax <= tolerance and var_post_tax > tolerance:
                calculated_total = calc_pre_tax
                variance = var_pre_tax
                applied_before_tax = True
            elif var_post_tax <= tolerance and var_pre_tax <= tolerance:
                # Both within tolerance (discount is tiny)
                calculated_total = calc_post_tax
                variance = var_post_tax
                applied_before_tax = False
            else:
                # Neither reconciles: report post-tax with explanation
                calculated_total = calc_post_tax
                variance = var_post_tax
                applied_before_tax = False
        else:
            calculated_total = calc_pre_tax
            variance = var_pre_tax
            applied_before_tax = None

        # 5. Primary Status Determination
        is_valid = (variance <= tolerance)
        status = "VERIFIED" if is_valid else "MISMATCH"

        # 6. Formulate Explanation
        notes = f"Taxable={taxable_val}, Tax={total_tax}, Total={inv_amt}"
        if discount > Decimal("0.00"):
            timing_label = "Pre-tax (factored into taxable value)" if applied_before_tax else "Post-tax"
            notes += f", Discount={discount} ({timing_label})"
        if additional_charges > Decimal("0.00"):
            notes += f", ExtraCharges={additional_charges}"
        if shipping_charges > Decimal("0.00"):
            notes += f", Freight={shipping_charges}"
        if cess and cess > Decimal("0.00"):
            notes += f", Cess={cess}"
        if rounding_adj != Decimal("0.00"):
            notes += f", Rounding={rounding_adj}"

        if is_valid:
            notes += " | Financial math verified."
        else:
            notes += f" | Mismatch: Calculated total {calculated_total} differs from Invoice total {inv_amt} (difference={variance} > tolerance {tolerance})."
            requires_manual_review = True
            review_fields.append("invoice_total_mismatch")

        if is_ambiguous_tax:
            notes += " | Warning: Ambiguous tax structure detected."

        # Compute effective tax rate
        tax_rate = None
        if taxable_val > Decimal("0.00") and total_tax > Decimal("0.00"):
            tax_rate = ((total_tax / taxable_val) * Decimal("100")).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)

        return FinancialAudit(
            status=status,
            validation_status=status,
            taxable_value=taxable_val,
            cgst=cgst,
            sgst=sgst,
            igst=igst,
            cess=cess,
            tax_amount=total_tax,
            discount=discount if discount > Decimal("0.00") else None,
            discount_amount=discount if discount > Decimal("0.00") else None,
            discount_applied_before_tax=applied_before_tax,
            additional_charges=additional_charges if additional_charges > Decimal("0.00") else None,
            extra_charges=additional_charges if additional_charges > Decimal("0.00") else None,
            shipping_charges=shipping_charges if shipping_charges > Decimal("0.00") else None,
            rounding_adjustment=rounding_adj if rounding_adj != Decimal("0.00") else None,
            inv_amount=inv_amt,
            extracted_invoice_total=inv_amt,
            calculated_total=calculated_total,
            calculated_invoice_total=calculated_total,
            is_math_valid=is_valid,
            variance=variance,
            difference=variance,
            tax_rate_percent=tax_rate,
            notes=notes,
            explanation=notes,
            missing_components=missing_components,
            fields_requiring_manual_review=review_fields,
            requires_manual_review=requires_manual_review,
            is_ambiguous_tax=is_ambiguous_tax
        )

    @staticmethod
    def _normalize_date(date_str: str) -> Optional[str]:
        cleaned = date_str.strip()
        for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d", "%d %b %Y", "%d %B %Y", "%b %d, %Y", "%B %d, %Y"):
            try:
                dt = datetime.strptime(cleaned, fmt)
                return dt.strftime("%Y-%m-%d")
            except ValueError:
                pass
        return cleaned

    @staticmethod
    def _find_date_value(raw_fields: Dict[str, Any]) -> Optional[date]:
        """
        Searches extracted fields for a valid invoice date.
        NEVER returns date.today() as a fallback! Returns None if not found.
        """
        for k, v in raw_fields.items():
            if "date" in str(k).lower():
                val = v.get("value") if isinstance(v, dict) else v
                if val:
                    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d", "%d %b %Y", "%d %B %Y"):
                        try:
                            return datetime.strptime(str(val).strip(), fmt).date()
                        except ValueError:
                            pass
        return None
