import os
import json
import hashlib
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional, Union, Tuple, Generator

from core.models import DynamicInvoiceRecord, ExtractedField, ExcelColumnDefinition
from core.template_engine import TemplateEngine
from core.gemini_service import GeminiDocumentService
from core.validator import InvoiceValidator
from core.database import Database
from core.business_rules import BusinessRulesConfig

logger = logging.getLogger(__name__)

SUPPORTED_EXTENSIONS = {".pdf", ".png", ".jpg", ".jpeg"}

def calculate_sha256(file_path: Union[str, Path]) -> str:
    """Calculates SHA-256 document checksum."""
    h = hashlib.sha256()
    with open(file_path, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()

class InvoiceProcessingPipeline:
    """
    Stabilized end-to-end invoice automation pipeline:
    1. Read invoice document with SHA-256 calculation and idempotency check.
    2. Dynamic extraction via Gemini (official google-genai SDK).
    3. Resilient JSON unwrapping and Pydantic validation.
    4. Decimal financial audit and separate invoice/file duplicate detection.
    5. Safe database audit logging with failure recording and retry support.
    6. Review-aware cumulative Excel ledger synchronization.
    """

    def __init__(
        self,
        template_path: Optional[str] = None,
        sheet_name: Optional[str] = None,
        db_path: Optional[str] = None,
        business_rules: Optional[BusinessRulesConfig] = None
    ):
        self.business_rules = business_rules or BusinessRulesConfig()
        self.template_engine = TemplateEngine(template_path, sheet_name, self.business_rules)
        self.gemini_service = GeminiDocumentService()
        self.validator = InvoiceValidator(self.business_rules)
        self.db = Database(db_path)

    def get_template_columns(self) -> List[ExcelColumnDefinition]:
        return self.template_engine.get_column_definitions()

    @staticmethod
    def scan_directory_recursive(base_directory: Union[str, Path]) -> List[Tuple[Path, str]]:
        """
        Recursively discovers all supported invoice documents using pathlib.
        Preserves complete relative directory paths.
        
        Returns: List of (absolute_path, relative_posix_path)
        """
        base_dir = Path(base_directory)
        if not base_dir.exists() or not base_dir.is_dir():
            raise NotADirectoryError(f"Directory not found: {base_dir}")

        found_files: List[Tuple[Path, str]] = []
        for path in sorted(base_dir.rglob("*")):
            if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS:
                # Do not scan hidden directories like .git or .agents
                if any(part.startswith(".") for part in path.parts):
                    continue
                rel_path = path.relative_to(base_dir).as_posix()
                found_files.append((path, rel_path))

        return found_files

    def process_single_invoice(
        self,
        file_path: Union[str, Path],
        relative_path: Optional[str] = None,
        folder_context: Optional[str] = None,
        model_name: Optional[str] = None,
        skip_if_processed: bool = True,
        auto_save_excel: bool = False
    ) -> DynamicInvoiceRecord:
        """
        Processes a single invoice document through the complete pipeline.
        Calculates SHA-256 and handles idempotency, failures, and retries.
        """
        path = Path(file_path).resolve()
        if not path.exists():
            raise FileNotFoundError(f"Invoice file does not exist: {path}")

        rel_path = relative_path or path.name
        folder_name = folder_context or path.parent.name
        file_name = path.name

        # Calculate document hash
        file_hash = calculate_sha256(path)

        # File-level duplicate check & Idempotency
        if skip_if_processed:
            existing_record = self.db.get_record_by_hash(file_hash)
            if existing_record and existing_record.get("processing_status") == "SUCCESS":
                logger.info(f"Skipping already processed file (SHA-256 match): {rel_path}")
                # Load previously stored fields
                try:
                    raw_fields = json.loads(existing_record["fields_data"]) if existing_record.get("fields_data") else {}
                    raw_json = json.loads(existing_record["raw_response"]) if existing_record.get("raw_response") else {}
                except Exception as e:
                    logger.error(f"Error parsing json from existing record: {e}")
                    raw_fields, raw_json = {}, {}

                if not raw_json and raw_fields:
                    raw_json = {"fields": raw_fields}

                columns = self.template_engine.get_column_definitions()
                cached_rec = self.validator.validate_and_reconcile(
                    raw_json=raw_json,
                    columns=columns,
                    file_name=file_name,
                    file_path=str(path),
                    relative_path=rel_path,
                    folder_context=folder_name,
                    business_rules=self.business_rules
                )
                cached_rec.id = existing_record["id"]
                cached_rec.file_hash = file_hash
                cached_rec.processing_status = "SKIPPED"
                cached_rec.review_status = existing_record.get("review_status", "PENDING")
                cached_rec.is_file_duplicate = True
                cached_rec.duplicate_reason = f"File with identical SHA-256 already processed as Record #{existing_record['id']}"
                return cached_rec

        # Step 1: Dynamic prompt construction
        prompt = self.template_engine.build_extraction_prompt_spec(
            folder_name=folder_name,
            filename=file_name
        )

        try:
            # Step 2: Extraction via Gemini
            raw_json, model_used, exec_time = self.gemini_service.extract_invoice_data(
                document_path=path,
                prompt=prompt,
                folder_context=folder_name,
                model_name=model_name
            )

            # Step 3: Validate and reconcile
            columns = self.template_engine.get_column_definitions()
            record = self.validator.validate_and_reconcile(
                raw_json=raw_json,
                columns=columns,
                file_name=file_name,
                file_path=str(path),
                relative_path=rel_path,
                folder_context=folder_name,
                business_rules=self.business_rules
            )
            record.file_hash = file_hash
            record.processing_status = "SUCCESS"

            # Step 4: Separate Invoice-Level Duplicate Check
            vendor_val = None
            inv_no_val = None
            total_val = float(record.financial_audit.inv_amount) if record.financial_audit.inv_amount is not None else None

            for k, f in record.fields.items():
                k_lower = k.lower()
                if "vendor" in k_lower and f.value:
                    vendor_val = str(f.value)
                elif ("inv#" in k_lower or "invoice number" in k_lower or "inv no" in k_lower) and f.value:
                    inv_no_val = str(f.value)

            is_inv_dup, dup_reason = self.db.check_invoice_duplicate(vendor_val, inv_no_val, total_val)
            record.is_invoice_duplicate = is_inv_dup
            record.is_duplicate = is_inv_dup
            record.duplicate_reason = dup_reason
            if is_inv_dup:
                record.validation_flags.append(f"INVOICE DUPLICATE: {dup_reason}")

            # Step 5: Save audit trail to database
            rec_id = self.db.insert_invoice_record(
                record=record,
                model_used=model_used,
                execution_time=exec_time
            )
            record.id = rec_id

            # Step 6: Map to Excel if auto-save requested
            if auto_save_excel:
                # Note: only approved or explicit save
                self.template_engine.write_records_to_excel([record], only_approved=False)

            return record

        except Exception as e:
            logger.error(f"Failed processing invoice {rel_path}: {e}")
            # Record failure in SQLite
            fail_id = self.db.record_failure(
                file_name=file_name,
                relative_path=rel_path,
                file_hash=file_hash,
                error_message=str(e),
                model_used=model_name or "gemini"
            )
            columns = self.template_engine.get_column_definitions()
            fail_rec = DynamicInvoiceRecord(
                id=fail_id,
                file_name=file_name,
                file_path=str(path),
                relative_path=rel_path,
                folder_context=folder_name,
                file_hash=file_hash,
                processing_status="FAILED",
                review_status="REJECTED",
                error_message=str(e),
                validation_flags=[f"Processing failure: {e}"]
            )
            return fail_rec

    def process_batch_recursive(
        self,
        directory_path: Union[str, Path],
        model_name: Optional[str] = None,
        skip_if_processed: bool = True
    ) -> List[DynamicInvoiceRecord]:
        """
        Recursively scans directory and processes all discovered documents.
        Supports retrying previously failed records and skips verified hashes.
        """
        found_files = self.scan_directory_recursive(directory_path)
        logger.info(f"Discovered {len(found_files)} invoice files in {directory_path} (recursive).")

        records: List[DynamicInvoiceRecord] = []
        for abs_path, rel_path in found_files:
            folder_ctx = Path(rel_path).parent.as_posix()
            if folder_ctx == ".":
                folder_ctx = Path(directory_path).name

            rec = self.process_single_invoice(
                file_path=abs_path,
                relative_path=rel_path,
                folder_context=folder_ctx,
                model_name=model_name,
                skip_if_processed=skip_if_processed,
                auto_save_excel=False
            )
            records.append(rec)

        return records

    def export_approved_to_excel(
        self,
        target_path: Optional[str] = None,
        include_already_exported: bool = False
    ) -> Tuple[str, int]:
        """
        Exports approved records from the SQLite database to the cumulative Excel ledger.
        
        Guarantees:
        - Only review_status == 'APPROVED' records are eligible.
        - PENDING and REJECTED records are strictly excluded.
        - Records already marked export_status == 'EXPORTED' are skipped (preventing duplicates).
        - Safe atomic writing to temporary file with integrity check and .bak backup.
        - Records SQLite export status and row location upon successful write.
        - On write failure, preserves DB approval and marks export_status == 'FAILED' for retry.
        
        Returns:
            (ledger_path, num_rows_appended)
        """
        if include_already_exported:
            approved_rows = self.db.get_approved_records()
        else:
            approved_rows = self.db.get_approved_unexported_records()

        if not approved_rows:
            logger.info("No approved records awaiting Excel synchronization.")
            out_dir = Path("output")
            out_dir.mkdir(parents=True, exist_ok=True)
            default_ledger = Path(target_path) if target_path else out_dir / "AP_BLR_Cumulative_Ledger.xlsx"
            return str(default_ledger), 0

        columns = self.template_engine.get_column_definitions()
        records: List[DynamicInvoiceRecord] = []

        for row in approved_rows:
            try:
                # Reconstruct record from fields_data (preserving any reviewer edits)
                if row.get("fields_data"):
                    f_dict = json.loads(row["fields_data"])
                    fields = {}
                    for k, v in f_dict.items():
                        fields[k] = ExtractedField(**v) if isinstance(v, dict) else ExtractedField(value=v)
                    
                    rec = DynamicInvoiceRecord(
                        id=row["id"],
                        file_name=row["file_name"],
                        relative_path=row.get("relative_path"),
                        fields=fields,
                        review_status="APPROVED",
                        export_status=row.get("export_status", "NOT_EXPORTED")
                    )
                else:
                    raw_json = json.loads(row["raw_response"]) if row.get("raw_response") else {}
                    rec = self.validator.validate_and_reconcile(
                        raw_json=raw_json,
                        columns=columns,
                        file_name=row["file_name"],
                        relative_path=row.get("relative_path"),
                        business_rules=self.business_rules
                    )
                    rec.id = row["id"]
                    rec.review_status = "APPROVED"

                records.append(rec)
            except Exception as e:
                logger.warning(f"Error preparing approved record #{row.get('id')} for export: {e}")

        if not records:
            out_dir = Path("output")
            default_ledger = Path(target_path) if target_path else out_dir / "AP_BLR_Cumulative_Ledger.xlsx"
            return str(default_ledger), 0

        try:
            saved_path, num_appended, exported_mappings = self.template_engine.write_records_to_excel(
                records=records,
                target_path=target_path,
                only_approved=True
            )

            # Update database status for successfully exported records
            for rec_id, row_idx in exported_mappings:
                if rec_id > 0:
                    self.db.mark_as_exported(
                        record_id=rec_id,
                        sheet_name=self.template_engine.sheet_name,
                        row_idx=row_idx
                    )

            return saved_path, num_appended

        except Exception as e:
            logger.error(f"Excel synchronization failed: {e}")
            # Mark export failure in database for retry without losing approvals
            for r in records:
                if r.id:
                    self.db.mark_export_failed(r.id, str(e))
            raise
