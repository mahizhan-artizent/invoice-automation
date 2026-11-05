import json
from core.template_engine import TemplateEngine
from core.business_rules import BusinessRulesConfig
from core.models import ExcelColumnDefinition

def test_dynamic_prompt_generation_includes_discovered_columns():
    engine = TemplateEngine()
    prompt = engine.build_extraction_prompt_spec(folder_name="DeptA/May", filename="invoice_001.pdf")
    
    # Verify prompt contains folder context and filename
    assert "DeptA/May" in prompt
    assert "invoice_001.pdf" in prompt
    
    # Verify prompt dynamically includes all columns in template
    for col in engine.get_column_definitions():
        assert col.header_name in prompt

def test_dynamic_schema_json_structure():
    engine = TemplateEngine()
    prompt = engine.build_extraction_prompt_spec(folder_name="Invoices", filename="test.pdf")
    
    # Verify schema template was formatted into the prompt
    assert '"fields":' in prompt
    assert '"financial_audit":' in prompt
    assert 'Missing Values: If any field cannot be found on the document, return its \'value\' as null' in prompt

def test_custom_business_rule_override():
    custom_rules = BusinessRulesConfig()
    rule = custom_rules.get_rule_for_key("invaccmonth1")
    assert rule is not None
    assert rule.derivation_type == "accounting_code"
    assert rule.requires_review_if_derived is True

