import os
import sys
import json
import time
from pathlib import Path

# Add project root to path
PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.pipeline import InvoiceProcessingPipeline, calculate_sha256
from core.database import Database

def main():
    samples_dir = Path("/Users/mahizhan/Desktop/Madan/samples inv")
    pipeline = InvoiceProcessingPipeline()
    db = pipeline.db

    # Step 1: Inventory
    print("=================================================================")
    print("STEP 1: INVENTORY INSPECTION")
    print("=================================================================")
    discovered = pipeline.scan_directory_recursive(samples_dir)
    print(f"Total supported invoices discovered: {len(discovered)}")
    for abs_p, rel_p in discovered:
        sz = abs_p.stat().st_size
        sha = calculate_sha256(abs_p)
        print(f" - {rel_p:<50} ({sz:,} bytes) [SHA-256: {sha[:12]}...]")

    # Step 2: Process Invoices
    print("\n=================================================================")
    print("STEP 2: PROCESSING INVOICES (Idempotent & Safe)")
    print("=================================================================\n")

    results = []
    skipped_count = 0
    success_count = 0
    failed_count = 0

    for idx, (abs_path, rel_path) in enumerate(discovered, 1):
        sha = calculate_sha256(abs_path)
        existing = db.get_record_by_hash(sha)

        is_cached = (existing is not None and existing.get("processing_status") == "SUCCESS")
        action_label = "RE-CHECKING (CACHED)" if is_cached else "EXTRACTING"
        print(f"[{idx:2d}/{len(discovered)}] {action_label}: {rel_path}...")

        start_t = time.time()
        record = pipeline.process_single_invoice(
            file_path=abs_path,
            relative_path=rel_path,
            folder_context=Path(rel_path).parent.as_posix(),
            skip_if_processed=True,
            auto_save_excel=False
        )
        elapsed = time.time() - start_t

        if record.processing_status == "SKIPPED":
            skipped_count += 1
            print(f"       -> SKIPPED (Idempotent match with Record #{record.id}) [{elapsed:.2f}s]")
        elif record.processing_status == "SUCCESS":
            success_count += 1
            vendor = record.fields.get("Vendor name ", record.fields.get("Vendor name")).value if (record.fields.get("Vendor name ") or record.fields.get("Vendor name")) else "N/A"
            amt = record.financial_audit.inv_amount
            print(f"       -> SUCCESS | Vendor: {vendor} | Amount: {amt} [{elapsed:.2f}s]")
        else:
            failed_count += 1
            print(f"       -> FAILED | Error: {record.error_message} [{elapsed:.2f}s]")

        results.append(record)
        # Gentle pace between API calls to prevent 429
        if not is_cached:
            time.sleep(1.0)

    # Save summary report to JSON
    report_data = {
        "summary": {
            "total_discovered": len(discovered),
            "successfully_processed": success_count,
            "previously_processed_skipped": skipped_count,
            "failed": failed_count,
            "pending_manual_review": len(results),
            "file_duplicates": sum(1 for r in results if r.is_file_duplicate),
            "invoice_duplicates": sum(1 for r in results if r.is_invoice_duplicate)
        },
        "records": [r.model_dump() for r in results]
    }

    report_path = Path("output/validation_run_report.json")
    report_path.parent.mkdir(exist_ok=True)
    with open(report_path, "w") as f:
        json.dump(report_data, f, indent=2, default=str)

    print(f"\nReport written to: {report_path}")

if __name__ == "__main__":
    main()

