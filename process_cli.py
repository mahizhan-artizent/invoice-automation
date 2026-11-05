#!/usr/bin/env python3
import os
import sys
from pathlib import Path

# Ensure project root is in sys.path
PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import json
from core.pipeline import InvoiceProcessingPipeline

def main():
    if len(sys.argv) < 2:
        print("Usage: python3 process_cli.py <path_to_invoice_pdf_or_image>")
        print("Example: python3 process_cli.py 'samples inv/NIIN-009131-ARNA TECHNOSOFT LLP-Nailesh R Agrawal.pdf'")
        sys.exit(1)

    file_path = sys.argv[1]
    print(f"\n=======================================================")
    print(f"🚀 Processing: {file_path}")
    print(f"=======================================================\n")

    pipeline = InvoiceProcessingPipeline()
    record = pipeline.process_single_invoice(file_path, auto_save_excel=True)

    print("📋 DYNAMICALLY EXTRACTED EXCEL COLUMNS:")
    print("-" * 55)
    for col_name, field in record.fields.items():
        status_icon = "⚠️" if field.is_uncertain else "✅"
        print(f"{status_icon} {col_name.strip():<20}: {str(field.value):<35} (Source: {field.source}, Conf: {field.confidence:.0%})")

    print("\n⚖️ FINANCIAL CROSS-FOOTING CHECK:")
    print("-" * 55)
    fa = record.financial_audit
    print(f"• Taxable Value  : ₹ {fa.taxable_value:,.2f}" if fa.taxable_value else "• Taxable Value  : N/A")
    print(f"• Tax (GST)      : ₹ {fa.tax_amount:,.2f}" if fa.tax_amount else "• Tax (GST)      : N/A")
    print(f"• Invoice Total  : ₹ {fa.inv_amount:,.2f}" if fa.inv_amount else "• Invoice Total  : N/A")
    print(f"• Math Valid?    : {'✅ Yes (Balanced)' if fa.is_math_valid else '❌ No (' + fa.status + ')'}")
    print(f"• Status         : {fa.status}")
    print(f"• Notes          : {fa.notes}")

    print("\n🔍 INTEGRITY & DUPLICATE CHECKS:")
    print("-" * 55)
    print(f"• Document Hash  : {record.file_hash[:16]}..." if record.file_hash else "• Document Hash  : N/A")
    print(f"• File Duplicate : {'🚨 YES (Identical SHA-256)' if record.is_file_duplicate else '✅ Unique File'}")
    print(f"• Inv Duplicate  : {'🚨 YES (Same Vendor/Inv#)' if record.is_invoice_duplicate else '✅ Unique Invoice'}")
    if record.duplicate_reason:
        print(f"  Reason         : {record.duplicate_reason}")
    if record.validation_flags:
        print(f"• Alerts         : {record.validation_flags}")
    else:
        print(f"• Alerts         : None (Clean)")

    print("\n💾 OUTPUTS SAVED:")
    print("-" * 55)
    print("• SQLite Database : Updated in invoice_automation.db")
    print("• Excel Workbook  : Updated in output/ directory")
    print(f"=======================================================\n")

if __name__ == "__main__":
    main()

