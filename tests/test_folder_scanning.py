from pathlib import Path
from core.pipeline import InvoiceProcessingPipeline

def test_recursive_folder_scanning_and_relative_paths(tmp_path):
    # Setup nested directory structure
    dept_a = tmp_path / "Finance" / "DeptA"
    dept_b = tmp_path / "Operations"
    dept_a.mkdir(parents=True)
    dept_b.mkdir(parents=True)
    
    # Create supported files
    (dept_a / "inv1.pdf").write_bytes(b"%PDF-1.4 test")
    (dept_a / "receipt1.png").write_bytes(b"png data")
    (dept_b / "bill1.jpg").write_bytes(b"jpg data")
    (dept_b / "scan1.jpeg").write_bytes(b"jpeg data")
    
    # Create unsupported / hidden files
    (dept_a / "notes.txt").write_text("not an invoice")
    (dept_b / "data.csv").write_text("csv data")
    
    hidden_dir = tmp_path / ".hidden"
    hidden_dir.mkdir()
    (hidden_dir / "hidden.pdf").write_bytes(b"%PDF hidden")

    # Perform recursive scan
    results = InvoiceProcessingPipeline.scan_directory_recursive(tmp_path)
    
    # Should only discover the 4 supported files
    assert len(results) == 4
    
    discovered_rel_paths = [rel for _, rel in results]
    assert "Finance/DeptA/inv1.pdf" in discovered_rel_paths
    assert "Finance/DeptA/receipt1.png" in discovered_rel_paths
    assert "Operations/bill1.jpg" in discovered_rel_paths
    assert "Operations/scan1.jpeg" in discovered_rel_paths
    
    # Verify non-supported files are excluded
    assert not any("notes.txt" in r for r in discovered_rel_paths)
    assert not any("data.csv" in r for r in discovered_rel_paths)
    assert not any(".hidden" in r for r in discovered_rel_paths)

