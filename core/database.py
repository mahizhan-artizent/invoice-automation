import sqlite3
import json
import logging
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple

from core.config import SQLITE_DB_PATH
from core.models import DynamicInvoiceRecord, ExtractedField

logger = logging.getLogger(__name__)

class Database:
    def __init__(self, db_path: Optional[str] = None):
        self.db_path = db_path or SQLITE_DB_PATH
        self.init_db()

    def get_connection(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def init_db(self):
        """Initializes database and safely performs additive migrations for existing tables."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            
            # Base table creation
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS invoice_records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    file_name TEXT NOT NULL,
                    vendor_name TEXT,
                    invoice_number TEXT,
                    invoice_date TEXT,
                    taxable_value REAL,
                    total_amount REAL,
                    model_used TEXT,
                    execution_time REAL,
                    is_duplicate INTEGER DEFAULT 0,
                    duplicate_reason TEXT,
                    is_math_valid INTEGER DEFAULT 1,
                    validation_flags TEXT,
                    fields_data TEXT,
                    raw_response TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            # Safe additive migration for Phase 2 / Phase 4 columns
            cursor.execute("PRAGMA table_info(invoice_records)")
            existing_cols = {row["name"] for row in cursor.fetchall()}

            migrations = [
                ("file_hash", "TEXT"),
                ("relative_path", "TEXT"),
                ("processing_status", "TEXT DEFAULT 'SUCCESS'"),
                ("review_status", "TEXT DEFAULT 'PENDING'"),
                ("error_message", "TEXT"),
                ("is_file_duplicate", "INTEGER DEFAULT 0"),
                ("is_invoice_duplicate", "INTEGER DEFAULT 0"),
                ("financial_status", "TEXT DEFAULT 'VERIFIED'"),
                ("export_status", "TEXT DEFAULT 'NOT_EXPORTED'"),
                ("exported_at", "TIMESTAMP"),
                ("exported_sheet", "TEXT"),
                ("exported_row", "INTEGER"),
                ("export_error", "TEXT"),
                ("reviewed_at", "TIMESTAMP"),
                ("reviewer_id", "TEXT DEFAULT 'AP Reviewer'"),
                ("approval_notes", "TEXT"),
                ("corrections_json", "TEXT"),
                ("audit_history_json", "TEXT")
            ]

            for col_name, col_type in migrations:
                if col_name not in existing_cols:
                    logger.info(f"Applying safe migration: adding column '{col_name}' to invoice_records")
                    cursor.execute(f"ALTER TABLE invoice_records ADD COLUMN {col_name} {col_type}")

            # Indexes
            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_inv_vendor 
                ON invoice_records(vendor_name, invoice_number)
            """)
            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_inv_file_hash
                ON invoice_records(file_hash)
            """)
            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_inv_status
                ON invoice_records(processing_status, review_status)
            """)
            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_inv_export
                ON invoice_records(export_status, review_status)
            """)

            conn.commit()

    def get_record_by_hash(self, file_hash: str) -> Optional[Dict[str, Any]]:
        """Finds previously processed record by SHA-256 document hash."""
        if not file_hash:
            return None
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT * FROM invoice_records 
                WHERE file_hash = ?
                ORDER BY id DESC LIMIT 1
            """, (file_hash,))
            row = cursor.fetchone()
            return dict(row) if row else None

    def get_record_by_id(self, record_id: int) -> Optional[Dict[str, Any]]:
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM invoice_records WHERE id = ?", (record_id,))
            row = cursor.fetchone()
            return dict(row) if row else None

    def check_invoice_duplicate(
        self,
        vendor_name: Optional[str],
        invoice_number: Optional[str],
        total_amount: Optional[float]
    ) -> Tuple[bool, Optional[str]]:
        """
        Invoice-level duplicate check based on business identity:
        Matches normalized (vendor_name, invoice_number) in successful records.
        """
        if not vendor_name or not invoice_number:
            return False, None

        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, file_name, relative_path, invoice_date, total_amount 
                FROM invoice_records 
                WHERE LOWER(TRIM(vendor_name)) = LOWER(TRIM(?))
                  AND LOWER(TRIM(invoice_number)) = LOWER(TRIM(?))
                  AND processing_status = 'SUCCESS'
                ORDER BY id ASC LIMIT 1
            """, (vendor_name, invoice_number))
            row = cursor.fetchone()
            if row:
                doc_ref = row["relative_path"] or row["file_name"]
                return True, f"Invoice duplicate of Record #{row['id']} ({doc_ref}, Inv# {invoice_number}, Total: {row['total_amount']})"

        return False, None

    def insert_invoice_record(
        self,
        record: DynamicInvoiceRecord,
        model_used: str = "gemini",
        execution_time: float = 0.0
    ) -> int:
        vendor = None
        inv_no = None
        inv_date = None
        taxable = None
        total = None

        for k, f in record.fields.items():
            k_lower = k.lower()
            if "vendor" in k_lower and f.value:
                vendor = str(f.value)
            elif ("inv#" in k_lower or "invoice number" in k_lower or "inv no" in k_lower) and f.value:
                inv_no = str(f.value)
            elif "inv date" in k_lower and f.value:
                inv_date = str(f.value)
            elif "taxable" in k_lower and f.value is not None:
                taxable = float(f.value)
            elif ("inv amount" in k_lower or "total" in k_lower) and f.value is not None:
                total = float(f.value)

        # Fallback to financial_audit if total is not in fields
        if total is None and record.financial_audit.inv_amount is not None:
            total = float(record.financial_audit.inv_amount)
        if taxable is None and record.financial_audit.taxable_value is not None:
            taxable = float(record.financial_audit.taxable_value)

        fields_dict = {k: v.model_dump() for k, v in record.fields.items()}

        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO invoice_records (
                    file_name, vendor_name, invoice_number, invoice_date,
                    taxable_value, total_amount, model_used, execution_time,
                    is_duplicate, duplicate_reason, is_math_valid,
                    validation_flags, fields_data, raw_response,
                    file_hash, relative_path, processing_status, review_status,
                    is_file_duplicate, error_message, financial_status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                record.file_name,
                vendor,
                inv_no,
                inv_date,
                taxable,
                total,
                model_used,
                execution_time,
                1 if (record.is_invoice_duplicate or record.is_duplicate) else 0,
                record.duplicate_reason,
                1 if record.financial_audit.is_math_valid else 0,
                json.dumps(record.validation_flags),
                json.dumps(fields_dict),
                json.dumps(record.raw_response or {}),
                record.file_hash,
                record.relative_path or record.file_name,
                record.processing_status,
                record.review_status,
                1 if record.is_file_duplicate else 0,
                record.error_message,
                record.financial_audit.status
            ))
            conn.commit()
            return cursor.lastrowid

    def record_failure(
        self,
        file_name: str,
        relative_path: str,
        file_hash: Optional[str],
        error_message: str,
        model_used: str = "gemini",
        folder_context: Optional[str] = None
    ) -> int:
        """Records an invoice processing failure in SQLite to allow auditing and retries."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO invoice_records (
                    file_name, relative_path, file_hash, processing_status,
                    review_status, error_message, model_used, execution_time,
                    is_math_valid, validation_flags, fields_data, raw_response,
                    financial_status
                ) VALUES (?, ?, ?, 'FAILED', 'REJECTED', ?, ?, 0.0, 0, ?, '{}', '{}', 'INCOMPLETE')
            """, (
                file_name,
                relative_path,
                file_hash,
                error_message,
                model_used,
                json.dumps([f"Processing failure: {error_message}"])
            ))
            conn.commit()
            return cursor.lastrowid

    def _append_audit_event(
        self,
        cursor: sqlite3.Cursor,
        record_id: int,
        action: str,
        reviewer_id: str,
        details: str,
        previous_state: Optional[Dict[str, Any]] = None,
        new_state: Optional[Dict[str, Any]] = None,
        changes: Optional[Dict[str, Any]] = None
    ) -> List[Dict[str, Any]]:
        cursor.execute("SELECT audit_history_json FROM invoice_records WHERE id = ?", (record_id,))
        row = cursor.fetchone()
        history = []
        if row and row["audit_history_json"]:
            try:
                history = json.loads(row["audit_history_json"])
            except Exception:
                history = []

        event = {
            "timestamp": datetime.now().isoformat(),
            "action": action,
            "reviewer_id": reviewer_id or "AP Reviewer",
            "details": details,
            "changes": changes or {},
            "previous_state": previous_state or {},
            "new_state": new_state or {}
        }
        history.append(event)
        cursor.execute("UPDATE invoice_records SET audit_history_json = ? WHERE id = ?", (json.dumps(history), record_id))
        return history

    def get_audit_history(self, record_id: int) -> List[Dict[str, Any]]:
        """Returns the parsed audit trail for an invoice record."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT audit_history_json FROM invoice_records WHERE id = ?", (record_id,))
            row = cursor.fetchone()
            if row and row["audit_history_json"]:
                try:
                    return json.loads(row["audit_history_json"])
                except Exception:
                    return []
            return []

    def update_invoice_fields(
        self,
        record_id: int,
        fields_dict: Dict[str, Any],
        reviewer_id: str = "AP Reviewer",
        notes: Optional[str] = None
    ) -> Tuple[bool, Optional[str], Optional[Dict[str, Any]]]:
        """
        Updates extracted template fields with reviewer corrections.
        Preserves nulls, recalculates financial validation, logs audit history and corrections,
        and ensures the record remains in PENDING review status (never silently auto-approved).
        """
        from core.validator import InvoiceValidator

        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM invoice_records WHERE id = ?", (record_id,))
            existing = cursor.fetchone()
            if not existing:
                return False, f"Record #{record_id} not found.", None

            existing_dict = dict(existing)
            current_fields_data = {}
            if existing_dict.get("fields_data"):
                try:
                    current_fields_data = json.loads(existing_dict["fields_data"])
                except Exception:
                    current_fields_data = {}

            # Parse existing corrections
            existing_corrections = {}
            if existing_dict.get("corrections_json"):
                try:
                    existing_corrections = json.loads(existing_dict["corrections_json"])
                except Exception:
                    existing_corrections = {}

            changes = {}
            updated_fields_obj: Dict[str, ExtractedField] = {}
            for col_name, new_val in fields_dict.items():
                old_val = None
                old_field = current_fields_data.get(col_name)
                if isinstance(old_field, dict):
                    old_val = old_field.get("value")
                elif old_field is not None:
                    old_val = old_field

                clean_val = None if (new_val is None or str(new_val).strip() == "" or str(new_val).strip().lower() in ("none", "null")) else new_val
                
                if str(old_val) != str(clean_val):
                    changes[col_name] = {"old": old_val, "new": clean_val}

                updated_fields_obj[col_name] = ExtractedField(
                    value=clean_val,
                    raw_text=str(clean_val) if clean_val is not None else None,
                    source="reviewer_correction" if str(old_val) != str(clean_val) else (old_field.get("source", "unknown") if isinstance(old_field, dict) else "unknown"),
                    confidence=1.0 if str(old_val) != str(clean_val) else (old_field.get("confidence", 1.0) if isinstance(old_field, dict) else 1.0),
                    is_uncertain=False if str(old_val) != str(clean_val) else (old_field.get("is_uncertain", False) if isinstance(old_field, dict) else False),
                    flag_reason=f"Corrected by {reviewer_id}" if str(old_val) != str(clean_val) else (old_field.get("flag_reason") if isinstance(old_field, dict) else None)
                )

            # Preserve other fields from current_fields_data not in fields_dict
            for k, v in current_fields_data.items():
                if k not in updated_fields_obj:
                    if isinstance(v, dict):
                        updated_fields_obj[k] = ExtractedField(**v)
                    else:
                        updated_fields_obj[k] = ExtractedField(value=v)

            # Map to top-level columns
            vendor = existing_dict.get("vendor_name")
            inv_no = existing_dict.get("invoice_number")
            inv_date = existing_dict.get("invoice_date")
            taxable = existing_dict.get("taxable_value")
            total = existing_dict.get("total_amount")

            for k, f in updated_fields_obj.items():
                k_lower = k.lower()
                if "vendor" in k_lower:
                    vendor = str(f.value) if f.value is not None else None
                elif "inv#" in k_lower or "invoice number" in k_lower or "inv no" in k_lower:
                    inv_no = str(f.value) if f.value is not None else None
                elif "inv date" in k_lower:
                    inv_date = str(f.value) if f.value is not None else None
                elif "taxable" in k_lower:
                    if f.value is not None:
                        try:
                            taxable = float(f.value)
                        except (ValueError, TypeError):
                            taxable = None
                    else:
                        taxable = None
                elif "inv amount" in k_lower or "total" in k_lower:
                    if f.value is not None:
                        try:
                            total = float(f.value)
                        except (ValueError, TypeError):
                            total = None
                    else:
                        total = None

            # Rerun financial validation using InvoiceValidator
            raw_audit = {}
            if existing_dict.get("raw_response"):
                try:
                    raw_res = json.loads(existing_dict["raw_response"])
                    _, raw_audit = InvoiceValidator.unwrap_raw_response(raw_res)
                except Exception:
                    raw_audit = {}

            if taxable is not None:
                raw_audit["taxable_value"] = taxable
            if total is not None:
                raw_audit["inv_amount"] = total

            new_audit = InvoiceValidator.validate_financials(raw_audit=raw_audit, fields=updated_fields_obj)
            financial_status = new_audit.status
            is_math_valid = 1 if new_audit.is_math_valid else 0

            # Merge corrections
            if changes:
                for k, v in changes.items():
                    existing_corrections[k] = v

            serialized_fields = {k: v.model_dump() for k, v in updated_fields_obj.items()}
            # Rule: do not silently approve after editing. Retain PENDING.
            review_status = "PENDING"

            cursor.execute("""
                UPDATE invoice_records
                SET fields_data = ?,
                    vendor_name = ?,
                    invoice_number = ?,
                    invoice_date = ?,
                    taxable_value = ?,
                    total_amount = ?,
                    financial_status = ?,
                    is_math_valid = ?,
                    review_status = ?,
                    corrections_json = ?
                WHERE id = ?
            """, (
                json.dumps(serialized_fields),
                vendor,
                inv_no,
                inv_date,
                taxable,
                total,
                financial_status,
                is_math_valid,
                review_status,
                json.dumps(existing_corrections),
                record_id
            ))

            details = f"Updated {len(changes)} field(s) by {reviewer_id}. Financial status: {financial_status}"
            if notes:
                details += f" (Note: {notes})"

            self._append_audit_event(
                cursor=cursor,
                record_id=record_id,
                action="FIELD_UPDATE",
                reviewer_id=reviewer_id,
                details=details,
                changes=changes,
                new_state={"financial_status": financial_status, "review_status": review_status}
            )

            conn.commit()

            cursor.execute("SELECT * FROM invoice_records WHERE id = ?", (record_id,))
            updated_row = dict(cursor.fetchone())
            return True, None, updated_row

    def approve_invoice(
        self,
        record_id: int,
        reviewer_id: str = "AP Reviewer",
        acknowledge_warnings: bool = False,
        notes: Optional[str] = None
    ) -> Tuple[bool, Optional[str]]:
        """
        Approves an invoice for subsequent cumulative Excel export.
        Requires explicit acknowledge_warnings=True if there is a financial MISMATCH or ambiguous tax.
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM invoice_records WHERE id = ?", (record_id,))
            row = cursor.fetchone()
            if not row:
                return False, f"Record #{record_id} not found."

            record = dict(row)
            fin_status = record.get("financial_status") or "VERIFIED"

            is_ambiguous = False
            raw_res = record.get("raw_response")
            if raw_res:
                try:
                    parsed_res = json.loads(raw_res) if isinstance(raw_res, str) else raw_res
                    fin_audit = parsed_res.get("financial_audit", {})
                    if fin_audit.get("is_ambiguous_tax"):
                        is_ambiguous = True
                except Exception:
                    pass

            v_flags = record.get("validation_flags") or ""
            if "ambiguous" in v_flags.lower():
                is_ambiguous = True

            # Review rules
            if fin_status == "MISMATCH" and not acknowledge_warnings:
                return False, f"Invoice #{record_id} has a financial MISMATCH. Explicit reviewer acknowledgement is required to approve."

            if is_ambiguous and not acknowledge_warnings:
                return False, f"Invoice #{record_id} has an ambiguous tax structure. Explicit reviewer acknowledgement is required to approve."

            cursor.execute("""
                UPDATE invoice_records
                SET review_status = 'APPROVED',
                    reviewed_at = CURRENT_TIMESTAMP,
                    reviewer_id = ?,
                    approval_notes = ?
                WHERE id = ?
            """, (reviewer_id, notes, record_id))

            details = f"Invoice approved by {reviewer_id}"
            if acknowledge_warnings:
                details += " with explicit financial warnings acknowledgement"
            if notes:
                details += f" (Note: {notes})"

            self._append_audit_event(
                cursor=cursor,
                record_id=record_id,
                action="APPROVE",
                reviewer_id=reviewer_id,
                details=details,
                new_state={"review_status": "APPROVED", "reviewed_at": datetime.now().isoformat()}
            )

            conn.commit()
            return True, None

    def reject_invoice(
        self,
        record_id: int,
        reviewer_id: str = "AP Reviewer",
        reason: Optional[str] = None
    ) -> Tuple[bool, Optional[str]]:
        """Rejects an invoice. Rejected invoices are excluded from Excel export."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM invoice_records WHERE id = ?", (record_id,))
            if not cursor.fetchone():
                return False, f"Record #{record_id} not found."

            cursor.execute("""
                UPDATE invoice_records
                SET review_status = 'REJECTED',
                    reviewed_at = CURRENT_TIMESTAMP,
                    reviewer_id = ?,
                    approval_notes = ?
                WHERE id = ?
            """, (reviewer_id, reason, record_id))

            details = f"Invoice rejected by {reviewer_id}"
            if reason:
                details += f" (Reason: {reason})"

            self._append_audit_event(
                cursor=cursor,
                record_id=record_id,
                action="REJECT",
                reviewer_id=reviewer_id,
                details=details,
                new_state={"review_status": "REJECTED", "reason": reason}
            )

            conn.commit()
            return True, None

    def return_to_pending(
        self,
        record_id: int,
        reviewer_id: str = "AP Reviewer",
        reason: Optional[str] = None
    ) -> Tuple[bool, Optional[str]]:
        """Returns an invoice to PENDING review status."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM invoice_records WHERE id = ?", (record_id,))
            if not cursor.fetchone():
                return False, f"Record #{record_id} not found."

            cursor.execute("""
                UPDATE invoice_records
                SET review_status = 'PENDING',
                    reviewed_at = NULL,
                    reviewer_id = ?,
                    approval_notes = ?
                WHERE id = ?
            """, (reviewer_id, reason, record_id))

            details = f"Returned to PENDING by {reviewer_id}"
            if reason:
                details += f" (Reason: {reason})"

            self._append_audit_event(
                cursor=cursor,
                record_id=record_id,
                action="RETURN_TO_PENDING",
                reviewer_id=reviewer_id,
                details=details,
                new_state={"review_status": "PENDING"}
            )

            conn.commit()
            return True, None

    def batch_approve(
        self,
        record_ids: List[int],
        reviewer_id: str = "AP Reviewer",
        acknowledge_warnings: bool = False,
        notes: Optional[str] = None
    ) -> Dict[str, Any]:
        """Batch approves multiple invoices with warning checks."""
        approved = []
        failed = []
        for rid in record_ids:
            ok, err = self.approve_invoice(
                record_id=rid,
                reviewer_id=reviewer_id,
                acknowledge_warnings=acknowledge_warnings,
                notes=notes
            )
            if ok:
                approved.append(rid)
            else:
                failed.append({"id": rid, "reason": err})
        return {"approved": approved, "failed": failed}

    def batch_reject(
        self,
        record_ids: List[int],
        reviewer_id: str = "AP Reviewer",
        reason: Optional[str] = None
    ) -> Dict[str, Any]:
        """Batch rejects multiple invoices."""
        rejected = []
        failed = []
        for rid in record_ids:
            ok, err = self.reject_invoice(
                record_id=rid,
                reviewer_id=reviewer_id,
                reason=reason
            )
            if ok:
                rejected.append(rid)
            else:
                failed.append({"id": rid, "reason": err})
        return {"rejected": rejected, "failed": failed}

    def search_and_filter_invoices(
        self,
        search: Optional[str] = None,
        vendor: Optional[str] = None,
        invoice_number: Optional[str] = None,
        processing_status: Optional[str] = None,
        review_status: Optional[str] = None,
        financial_status: Optional[str] = None,
        export_status: Optional[str] = None,
        sort_by: str = "id",
        sort_order: str = "desc",
        limit: int = 100,
        offset: int = 0
    ) -> Tuple[List[Dict[str, Any]], int]:
        """Search and filter invoices using parameterized queries and dynamic sorting."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            conditions = ["1=1"]
            params: List[Any] = []

            if search:
                s = f"%{search.strip()}%"
                conditions.append("(vendor_name LIKE ? OR invoice_number LIKE ? OR file_name LIKE ? OR relative_path LIKE ?)")
                params.extend([s, s, s, s])

            if vendor:
                conditions.append("vendor_name LIKE ?")
                params.append(f"%{vendor.strip()}%")

            if invoice_number:
                conditions.append("invoice_number LIKE ?")
                params.append(f"%{invoice_number.strip()}%")

            if processing_status and processing_status.upper() != "ALL":
                conditions.append("processing_status = ?")
                params.append(processing_status.upper())

            if review_status and review_status.upper() != "ALL":
                conditions.append("review_status = ?")
                params.append(review_status.upper())

            if financial_status and financial_status.upper() != "ALL":
                conditions.append("financial_status = ?")
                params.append(financial_status.upper())

            if export_status and export_status.upper() != "ALL":
                if export_status.upper() == "NOT_EXPORTED":
                    conditions.append("(export_status IS NULL OR export_status = 'NOT_EXPORTED')")
                else:
                    conditions.append("export_status = ?")
                    params.append(export_status.upper())

            where_clause = " AND ".join(conditions)

            cursor.execute(f"SELECT COUNT(*) FROM invoice_records WHERE {where_clause}", tuple(params))
            total_count = cursor.fetchone()[0]

            valid_sort_cols = {
                "id": "id",
                "invoice_date": "invoice_date",
                "created_at": "created_at",
                "vendor_name": "vendor_name",
                "total_amount": "total_amount",
                "processing_date": "created_at"
            }
            order_col = valid_sort_cols.get(sort_by.lower(), "id")
            order_dir = "ASC" if sort_order.lower() == "asc" else "DESC"

            query = f"""
                SELECT * FROM invoice_records 
                WHERE {where_clause}
                ORDER BY {order_col} {order_dir}
                LIMIT ? OFFSET ?
            """
            select_params = list(params) + [limit, offset]
            cursor.execute(query, tuple(select_params))
            rows = [dict(r) for r in cursor.fetchall()]

            return rows, total_count

    def update_review_status(
        self,
        record_id: int,
        review_status: str,
        edited_fields: Optional[Dict[str, Any]] = None
    ) -> bool:
        """Backward-compatible delegate to update_invoice_fields / approve / reject."""
        rev_upper = review_status.upper()
        if edited_fields is not None:
            ok, _, _ = self.update_invoice_fields(record_id, edited_fields)
            if not ok:
                return False
        
        if rev_upper == "APPROVED":
            ok, _ = self.approve_invoice(record_id, acknowledge_warnings=True)
            return ok
        elif rev_upper == "REJECTED":
            ok, _ = self.reject_invoice(record_id)
            return ok
        elif rev_upper == "PENDING":
            ok, _ = self.return_to_pending(record_id)
            return ok
        else:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("UPDATE invoice_records SET review_status = ? WHERE id = ?", (rev_upper, record_id))
                conn.commit()
                return cursor.rowcount > 0

    def get_all_records(
        self,
        limit: int = 200,
        status: Optional[str] = None,
        review_status: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        with self.get_connection() as conn:
            cursor = conn.cursor()
            query = "SELECT * FROM invoice_records WHERE 1=1"
            params = []
            if status:
                query += " AND processing_status = ?"
                params.append(status)
            if review_status:
                query += " AND review_status = ?"
                params.append(review_status)

            query += " ORDER BY id DESC LIMIT ?"
            params.append(limit)

            cursor.execute(query, tuple(params))
            rows = cursor.fetchall()
            return [dict(r) for r in rows]

    def get_approved_records(self) -> List[Dict[str, Any]]:
        """Returns all approved records ready for Excel ledger synchronization."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT * FROM invoice_records 
                WHERE review_status = 'APPROVED'
                  AND processing_status = 'SUCCESS'
                ORDER BY id ASC
            """)
            return [dict(r) for r in cursor.fetchall()]

    def get_approved_unexported_records(self) -> List[Dict[str, Any]]:
        """Returns all approved records that have not yet been exported to the cumulative ledger."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT * FROM invoice_records 
                WHERE review_status = 'APPROVED'
                  AND processing_status = 'SUCCESS'
                  AND (export_status IS NULL OR export_status != 'EXPORTED')
                ORDER BY id ASC
            """)
            return [dict(r) for r in cursor.fetchall()]

    def mark_as_exported(self, record_id: int, sheet_name: str, row_idx: int) -> bool:
        """Marks a record as successfully synchronized to the cumulative Excel workbook."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE invoice_records
                SET export_status = 'EXPORTED',
                    exported_at = CURRENT_TIMESTAMP,
                    exported_sheet = ?,
                    exported_row = ?,
                    export_error = NULL
                WHERE id = ?
            """, (sheet_name, row_idx, record_id))
            conn.commit()
            return cursor.rowcount > 0

    def mark_export_failed(self, record_id: int, error_message: str) -> bool:
        """Records an export failure against an approved record, preserving approval status for retry."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE invoice_records
                SET export_status = 'FAILED',
                    export_error = ?
                WHERE id = ?
            """, (error_message, record_id))
            conn.commit()
            return cursor.rowcount > 0

    def get_stats(self) -> Dict[str, Any]:
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM invoice_records")
            total = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM invoice_records WHERE is_duplicate = 1 OR is_invoice_duplicate = 1 OR is_file_duplicate = 1")
            duplicates = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM invoice_records WHERE is_math_valid = 1")
            math_valid = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM invoice_records WHERE review_status = 'PENDING'")
            pending = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM invoice_records WHERE review_status = 'APPROVED'")
            approved = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM invoice_records WHERE review_status = 'REJECTED'")
            rejected = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM invoice_records WHERE financial_status = 'VERIFIED' AND processing_status = 'SUCCESS'")
            fin_verified = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM invoice_records WHERE (financial_status = 'INCOMPLETE' OR processing_status = 'FAILED')")
            fin_incomplete = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM invoice_records WHERE financial_status = 'MISMATCH' AND processing_status = 'SUCCESS'")
            fin_mismatch = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM invoice_records WHERE review_status = 'APPROVED' AND processing_status = 'SUCCESS' AND (export_status IS NULL OR export_status != 'EXPORTED')")
            awaiting_export = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM invoice_records WHERE export_status = 'EXPORTED'")
            exported = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM invoice_records WHERE processing_status = 'FAILED'")
            failed = cursor.fetchone()[0]
            cursor.execute("SELECT SUM(total_amount) FROM invoice_records WHERE processing_status = 'SUCCESS'")
            sum_total = cursor.fetchone()[0] or 0.0

            return {
                "total_processed": total,
                "pending_review": pending,
                "approved_count": approved,
                "rejected_count": rejected,
                "financially_verified": fin_verified,
                "financial_incomplete": fin_incomplete,
                "financial_mismatch": fin_mismatch,
                "awaiting_export": awaiting_export,
                "exported_count": exported,
                "pending_export_count": awaiting_export,
                "failed_count": failed,
                "duplicates_found": duplicates,
                "math_verified": math_valid,
                "total_spend": round(sum_total, 2)
            }
