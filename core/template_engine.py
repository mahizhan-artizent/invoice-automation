import os
import re
import json
import shutil
import logging
from copy import copy
from datetime import datetime, date
from decimal import Decimal
from pathlib import Path
from typing import List, Dict, Any, Tuple, Optional, Set
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from core.models import ExcelColumnDefinition, ExtractedField, DynamicInvoiceRecord
from core.config import EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME, OUTPUT_DIR
from core.business_rules import BusinessRulesConfig, FieldBusinessRule

logger = logging.getLogger(__name__)

class TemplateEngine:
    def __init__(
        self,
        template_path: Optional[str] = None,
        sheet_name: Optional[str] = None,
        business_rules: Optional[BusinessRulesConfig] = None
    ):
        self.template_path = Path(template_path or EXCEL_TEMPLATE_PATH)
        self.sheet_name = sheet_name or EXCEL_SHEET_NAME
        self.business_rules = business_rules or BusinessRulesConfig()
        self.columns: List[ExcelColumnDefinition] = []
        self.header_row: int = 2
        self._inspect_template()

    def _inspect_template(self) -> None:
        """Dynamically inspects the Excel template to discover all required columns and formats."""
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template not found at: {self.template_path}")

        wb = openpyxl.load_workbook(self.template_path, data_only=True)
        if self.sheet_name in wb.sheetnames:
            ws = wb[self.sheet_name]
        else:
            ws = wb.active

        # Find header row (row with most non-empty string headers in first 5 rows)
        best_row = 1
        max_headers = 0
        for r in range(1, min(ws.max_row + 1, 6)):
            valid_headers = [
                ws.cell(r, c).value for c in range(1, ws.max_column + 1)
                if ws.cell(r, c).value is not None and str(ws.cell(r, c).value).strip() != ""
            ]
            if len(valid_headers) > max_headers:
                max_headers = len(valid_headers)
                best_row = r

        self.header_row = best_row
        self.columns = []

        # Inspect each column
        for col_idx in range(1, ws.max_column + 1):
            raw_header = ws.cell(self.header_row, col_idx).value
            if raw_header is None or str(raw_header).strip() == "":
                continue

            header_str = str(raw_header)
            clean_key = re.sub(r'[^a-zA-Z0-9_]', '_', header_str.strip().lower())
            clean_key = re.sub(r'_+', '_', clean_key).strip('_')

            # Inspect subsequent rows to infer type, format, and sample values
            samples = []
            num_format = ws.cell(self.header_row + 1, col_idx).number_format
            inferred_type = "string"

            for r in range(self.header_row + 1, min(ws.max_row + 1, self.header_row + 10)):
                val = ws.cell(r, col_idx).value
                if val is not None:
                    samples.append(str(val))
                    if isinstance(val, (int, float)) and inferred_type == "string":
                        inferred_type = "number"
                    elif isinstance(val, (datetime, date)):
                        inferred_type = "date"

            # Check header name heuristics
            lower_h = header_str.lower()
            if "date" in lower_h or "month2" in lower_h:
                inferred_type = "date"
            elif any(k in lower_h for k in ["amount", "value", "price", "rate", "tax"]):
                inferred_type = "number"
            elif "s.no" in lower_h or "sno" in lower_h or "sl.no" in lower_h:
                inferred_type = "integer"

            self.columns.append(ExcelColumnDefinition(
                col_idx=col_idx,
                header_name=header_str,
                clean_key=clean_key,
                inferred_type=inferred_type,
                number_format=num_format,
                sample_values=samples[:3],
                description=f"Excel column '{header_str}' (type: {inferred_type})"
            ))

    def get_column_definitions(self) -> List[ExcelColumnDefinition]:
        return self.columns

    def build_extraction_prompt_spec(self, folder_name: str, filename: str) -> str:
        """
        Constructs genuinely dynamic extraction instructions and a dynamic JSON schema
        based on whatever columns are discovered in the target template.
        Uses configurable business rules without hardcoding columns.
        """
        cols_desc = []
        schema_fields_example = {}

        for col in self.columns:
            samples_str = f" (e.g. {', '.join(col.sample_values)})" if col.sample_values else ""
            rule = self.business_rules.get_rule_for_key(re.sub(r'[^a-z0-9]', '', col.header_name.lower()))
            
            guidance = ""
            if rule and rule.prompt_guidance:
                guidance = f" - Rule: {rule.prompt_guidance}"
            else:
                guidance = f" - Extract '{col.header_name}' directly from document. Return null if missing."

            cols_desc.append(f"- **{col.header_name}** [Type: {col.inferred_type}]{samples_str}{guidance}")

            # Build schema entry for this column
            type_placeholder = "YYYY-MM-DD" if col.inferred_type == "date" else (0.00 if col.inferred_type == "number" else "...")
            schema_fields_example[col.header_name] = {
                "value": f"<{col.inferred_type} value, or null if missing>",
                "raw_text": "<verbatim text from document or null>",
                "source": "<'invoice_header', 'bill_to', 'table_line_items', 'filename_context', or 'derived'>",
                "confidence": 0.95,
                "is_uncertain": False,
                "flag_reason": None
            }

        cols_formatted = "\n".join(cols_desc)
        schema_fields_json = json.dumps(schema_fields_example, indent=4)

        prompt = f"""You are an expert financial document intelligence system.
Your task is to extract information from the attached invoice document strictly matching the target Excel columns.

### Document and Directory Context
- Relative Folder Context: {folder_name}
- Filename: {filename}
(Note: Directory and filenames may contain reference ticket numbers, contractor, or vendor details.)

### Dynamic Target Excel Columns to Extract:
{cols_formatted}

### General Extraction Guidelines:
1. Missing Values: If any field cannot be found on the document, return its 'value' as null. NEVER invent or hallucinate data.
2. Field-Level Provenance: For every column, specify the exact 'source' where the data was located.
3. Uncertainty Flagging: If a value is inferred, ambiguous, blurry, or derived from context, set 'is_uncertain': true and explain in 'flag_reason'.
4. Financial Detail: Break down subtotal (taxable value), CGST, SGST, IGST, cess, discounts, extra/shipping charges, rounding adjustment, and total invoice amount in the financial_audit section. Specify whether discounts were applied pre_tax or post_tax.

### Output JSON Structure:
You MUST respond with valid JSON following this exact top-level schema:
```json
{{
  "fields": {schema_fields_json},
  "financial_audit": {{
    "taxable_value": 0.00,
    "cgst": 0.00,
    "sgst": 0.00,
    "igst": 0.00,
    "cess": 0.00,
    "tax_amount": 0.00,
    "discount_amount": 0.00,
    "discount_timing": "pre_tax or post_tax or null",
    "extra_charges": 0.00,
    "shipping_charges": 0.00,
    "rounding_adjustment": 0.00,
    "inv_amount": 0.00,
    "notes": "..."
  }}
}}
```
Respond ONLY with the JSON object. Do not include conversational markdown around it."""
        return prompt.strip()

    def get_max_serial_number(self, ws) -> int:
        """Finds the current maximum S.No in the sheet."""
        max_sno = 0
        for r in range(self.header_row + 1, ws.max_row + 1):
            val = ws.cell(r, 1).value
            if isinstance(val, int) and val > max_sno:
                max_sno = val
            elif isinstance(val, str) and val.isdigit() and int(val) > max_sno:
                max_sno = int(val)
        return max_sno

    def _get_existing_invoice_signatures(self, ws) -> Set[Tuple[str, str]]:
        """
        Scans existing rows in the worksheet to identify already recorded invoices.
        Returns a set of normalized (vendor_name, invoice_number) tuples.
        """
        existing = set()
        vendor_col = None
        inv_no_col = None

        for col in self.columns:
            canon = re.sub(r'[^a-z0-9]', '', col.header_name.lower())
            if "vendor" in canon:
                vendor_col = col.col_idx
            elif "inv" in canon and ("#" in col.header_name or "no" in canon or "num" in canon):
                inv_no_col = col.col_idx

        if vendor_col and inv_no_col:
            for r in range(self.header_row + 1, ws.max_row + 1):
                v_val = ws.cell(r, vendor_col).value
                i_val = ws.cell(r, inv_no_col).value
                if v_val and i_val:
                    norm_sig = (str(v_val).strip().lower(), str(i_val).strip().lower())
                    existing.add(norm_sig)

        return existing

    def write_records_to_excel(
        self,
        records: List[DynamicInvoiceRecord],
        target_path: Optional[str] = None,
        only_approved: bool = True
    ) -> Tuple[str, int, List[Tuple[int, int]]]:
        """
        Safely synchronizes records into a cumulative ledger workbook:
        1. Preserves original template Book3.xlsx untouched (strictly rejects direct template overwrite).
        2. Preserves existing baseline rows, formatting, formulas, column widths, and sheet properties.
        3. Prevents duplicate rows across multiple runs using in-sheet and stable business signatures.
        4. Exports ONLY approved records (if only_approved is True). PENDING and REJECTED are excluded.
        5. Performs atomic temporary writes (.tmp), integrity validation, and backup rotation (.bak).
        6. Ensures numeric values are true numbers and dates are true date objects. Missing values are left blank.

        Returns:
            (saved_file_path, num_rows_appended, list_of_(record_id, assigned_row_idx))
        """
        out_dir = Path(OUTPUT_DIR)
        out_dir.mkdir(parents=True, exist_ok=True)
        
        # Determine target ledger path (defaults to AP_BLR_Cumulative_Ledger.xlsx in output/)
        ledger_path = Path(target_path).resolve() if target_path else (out_dir / "AP_BLR_Cumulative_Ledger.xlsx").resolve()
        template_resolved = Path(self.template_path).resolve()

        # Strict safety check: Never allow writing directly to original template
        if ledger_path == template_resolved:
            raise ValueError(
                f"Safety violation: Target path '{ledger_path}' cannot be the original template Book3.xlsx. "
                "Cumulative ledger must be written to a separate output file."
            )

        # Initialize cumulative ledger from template if not existing yet
        if not ledger_path.exists():
            shutil.copy2(template_resolved, ledger_path)
            logger.info(f"Initialized cumulative ledger from template: {ledger_path.name}")

        # Filter eligible records based on approval status
        records_to_export = [
            r for r in records
            if (not only_approved) or (r.review_status == "APPROVED")
        ]

        if not records_to_export:
            logger.info("No records to export matching review criteria (only approved records may enter ledger).")
            return str(ledger_path), 0, []

        wb = None
        tmp_path = ledger_path.with_name(f"{ledger_path.stem}.tmp.xlsx")
        bak_path = ledger_path.with_name(f"{ledger_path.stem}.bak.xlsx")

        try:
            wb = openpyxl.load_workbook(ledger_path, data_only=False)
            ws = wb[self.sheet_name] if self.sheet_name in wb.sheetnames else wb.active

            # Get existing invoice signatures to prevent duplicates
            existing_signatures = self._get_existing_invoice_signatures(ws)

            # Find next empty row (scan backwards to handle empty trailing rows)
            current_row = ws.max_row + 1
            while current_row > self.header_row:
                prev_row_vals = [ws.cell(current_row - 1, c).value for c in range(1, len(self.columns) + 1)]
                if any(v is not None for v in prev_row_vals):
                    break
                current_row -= 1

            next_sno = self.get_max_serial_number(ws) + 1
            proto_row = self.header_row + 1  # Prototype row (Row 3) for cell styles
            appended_count = 0
            exported_mappings: List[Tuple[int, int]] = []

            for record in records_to_export:
                # Check for duplicate invoice in ledger
                vendor_val = None
                inv_no_val = None
                for col in self.columns:
                    canon = re.sub(r'[^a-z0-9]', '', col.header_name.lower())
                    field_data = record.fields.get(col.header_name)
                    if field_data and field_data.value is not None:
                        if "vendor" in canon:
                            vendor_val = str(field_data.value).strip().lower()
                        elif "inv" in canon and ("#" in col.header_name or "no" in canon or "num" in canon):
                            inv_no_val = str(field_data.value).strip().lower()

                if vendor_val and inv_no_val:
                    sig = (vendor_val, inv_no_val)
                    if sig in existing_signatures:
                        logger.info(f"Skipping duplicate invoice already in ledger: {sig}")
                        continue
                    existing_signatures.add(sig)

                assigned_row = current_row

                # 1. Write S.No (numeric integer)
                sno_cell = ws.cell(current_row, 1, int(next_sno))
                if proto_row <= ws.max_row:
                    proto_sno = ws.cell(proto_row, 1)
                    sno_cell.font = copy(proto_sno.font)
                    sno_cell.alignment = copy(proto_sno.alignment)
                    sno_cell.border = copy(proto_sno.border)

                # 2. Write remaining dynamic columns
                for col in self.columns:
                    if col.col_idx == 1:
                        continue  # S.No already written

                    field_data = record.fields.get(col.header_name)
                    val = field_data.value if field_data else None
                    cell = ws.cell(current_row, col.col_idx)

                    # Copy style & number format from prototype row
                    if proto_row <= ws.max_row:
                        proto_cell = ws.cell(proto_row, col.col_idx)
                        cell.font = copy(proto_cell.font)
                        cell.alignment = copy(proto_cell.alignment)
                        cell.border = copy(proto_cell.border)
                        if proto_cell.number_format:
                            cell.number_format = proto_cell.number_format

                    # Enforce strict field typing
                    if val is not None:
                        canon_col = re.sub(r'[^a-z0-9]', '', col.header_name.lower())
                        if col.inferred_type == "date" or "date" in canon_col or "month2" in canon_col:
                            parsed_date = None
                            if isinstance(val, (datetime, date)):
                                parsed_date = val
                            elif isinstance(val, str):
                                for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%Y", "%d %b %Y", "%d %B %Y", "%Y-%m-%dT%H:%M:%S"):
                                    try:
                                        parsed_date = datetime.strptime(val.strip(), fmt).date()
                                        break
                                    except ValueError:
                                        pass
                            cell.value = parsed_date if parsed_date else str(val)
                        elif col.inferred_type in ("number", "integer") or any(k in canon_col for k in ["amount", "value", "price", "rate"]):
                            cleaned_val = re.sub(r'[^\d.-]', '', str(val).strip())
                            try:
                                num_val = float(cleaned_val) if col.inferred_type != "integer" else int(float(cleaned_val))
                                cell.value = num_val
                            except (ValueError, TypeError):
                                cell.value = str(val)
                        else:
                            cell.value = str(val)
                    else:
                        # Missing values MUST remain blank in Excel
                        cell.value = None

                exported_mappings.append((record.id if record.id is not None else 0, assigned_row))
                next_sno += 1
                current_row += 1
                appended_count += 1

            # Atomic Write: 1. Save to .tmp file
            wb.save(tmp_path)
            wb.close()

            # Atomic Write: 2. Validate that the written workbook can be reopened and is healthy
            check_wb = openpyxl.load_workbook(tmp_path, data_only=False)
            check_ws = check_wb[self.sheet_name] if self.sheet_name in check_wb.sheetnames else check_wb.active
            if appended_count > 0 and check_ws.max_row < current_row - 1:
                check_wb.close()
                raise IOError(f"Validation failed: Temporary workbook max_row {check_ws.max_row} is less than expected {current_row - 1}")
            check_wb.close()

            # Atomic Write: 3. Create backup of previous valid output
            if ledger_path.exists():
                shutil.copy2(ledger_path, bak_path)

            # Atomic Write: 4. Atomic replace
            os.replace(tmp_path, ledger_path)
            logger.info(f"Successfully synchronized {appended_count} approved records into cumulative ledger: {ledger_path.name}")
            return str(ledger_path), appended_count, exported_mappings

        except Exception as e:
            logger.error(f"Failed to synchronize records into Excel ledger at {ledger_path}: {e}")
            if tmp_path.exists():
                try:
                    tmp_path.unlink()
                except Exception:
                    pass
            raise RuntimeError(f"Excel synchronization failure: {e}")
