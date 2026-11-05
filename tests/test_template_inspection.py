import pytest
from pathlib import Path
from core.template_engine import TemplateEngine
from core.config import EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME

def test_template_inspection_headers():
    engine = TemplateEngine(EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME)
    columns = engine.get_column_definitions()
    
    assert len(columns) == 12, f"Expected 12 columns, found {len(columns)}"
    
    col_names = [c.header_name.strip() for c in columns]
    assert "S.No" in col_names
    assert "Entity" in col_names
    assert "Inv received date" in col_names
    assert "Vendor name" in col_names
    assert "Inv#" in col_names
    assert "Inv Date" in col_names
    assert "Inv accMonth1" in col_names
    assert "Inv acc Month2" in col_names
    assert "NOE" in col_names
    assert "Taxable Value" in col_names
    assert "Inv Amount" in col_names
    assert "Sender" in col_names

def test_template_column_types_and_formats():
    engine = TemplateEngine(EXCEL_TEMPLATE_PATH, EXCEL_SHEET_NAME)
    col_map = {c.header_name.strip(): c for c in engine.get_column_definitions()}
    
    assert col_map["S.No"].inferred_type == "integer"
    assert col_map["Inv Date"].inferred_type == "date"
    assert col_map["Taxable Value"].inferred_type == "number"
    assert col_map["Inv Amount"].inferred_type == "number"
    assert col_map["Vendor name"].inferred_type == "string"

def test_template_file_not_found():
    with pytest.raises(FileNotFoundError):
        TemplateEngine("non_existent_file.xlsx", "Sheet1")

