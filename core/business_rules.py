from dataclasses import dataclass, field
from decimal import Decimal
from typing import Dict, Any, Optional, Callable
import re

@dataclass
class FieldBusinessRule:
    """Configurable business rule for a specific column or semantic field."""
    canonical_name: str
    description: str
    derivation_type: Optional[str] = None  # 'sequencer', 'accounting_code', 'accounting_date', 'constant'
    prompt_guidance: Optional[str] = None
    requires_review_if_derived: bool = False
    default_value: Optional[Any] = None

@dataclass
class AccountingPeriodConfig:
    """Configurable accounting month derivation parameters."""
    code_format: str = "%y_%m"          # e.g. "26_06"
    date_format: str = "%Y-%m-01"       # e.g. "2026-06-01"
    require_review_for_derived: bool = True
    allow_guessing_without_date: bool = False  # Strict rule: NEVER guess without invoice date

class BusinessRulesConfig:
    """
    Central repository for company-specific business rules.
    Decouples company rules from generic template inspection and extraction.
    """
    def __init__(
        self,
        custom_rules: Optional[Dict[str, FieldBusinessRule]] = None,
        rounding_tolerance: Decimal = Decimal("1.00"),
        accounting_config: Optional[AccountingPeriodConfig] = None
    ):
        self.rules: Dict[str, FieldBusinessRule] = custom_rules or self._get_default_rules()
        self.rounding_tolerance: Decimal = rounding_tolerance
        self.accounting_config: AccountingPeriodConfig = accounting_config or AccountingPeriodConfig()

    @staticmethod
    def _get_default_rules() -> Dict[str, FieldBusinessRule]:
        return {
            "sno": FieldBusinessRule(
                canonical_name="sno",
                description="Serial Number",
                derivation_type="sequencer",
                prompt_guidance="Leave value as null. This is auto-generated sequentially by the ledger engine.",
                requires_review_if_derived=False
            ),
            "entity": FieldBusinessRule(
                canonical_name="entity",
                description="Billed-to Company Entity",
                derivation_type="entity_lookup",
                prompt_guidance="Identify the recipient organization / company that was billed (e.g. customer name in 'Bill To').",
                requires_review_if_derived=False
            ),
            "invreceiveddate": FieldBusinessRule(
                canonical_name="invreceiveddate",
                description="Invoice Received Date",
                derivation_type="received_date",
                prompt_guidance="Date the invoice was received from email/metadata (YYYY-MM-DD). If not explicitly stated, leave null.",
                requires_review_if_derived=False
            ),
            "invaccmonth1": FieldBusinessRule(
                canonical_name="invaccmonth1",
                description="Accounting Month Code",
                derivation_type="accounting_code",
                prompt_guidance="Accounting month code in YY_MM format based on invoice or service period (e.g. '26_06' for June 2026). If ambiguous, mark uncertain.",
                requires_review_if_derived=True
            ),
            "invaccmonth2": FieldBusinessRule(
                canonical_name="invaccmonth2",
                description="Accounting Month Date",
                derivation_type="accounting_date",
                prompt_guidance="First day of the accounting month in YYYY-MM-01 format (e.g. '2026-06-01'). If ambiguous, mark uncertain.",
                requires_review_if_derived=True
            ),
            "noe": FieldBusinessRule(
                canonical_name="noe",
                description="Nature of Expense",
                derivation_type="line_items_summary",
                prompt_guidance="Concise description of the services rendered, consulting scope, hours, or items purchased.",
                requires_review_if_derived=False
            ),
            "sender": FieldBusinessRule(
                canonical_name="sender",
                description="Invoice Sender / Signatory",
                derivation_type="sender_context",
                prompt_guidance="Consultant name, contractor, or authorized signatory who submitted/signed the invoice (often in filename or footer).",
                requires_review_if_derived=False
            )
        }

    def get_rule_for_key(self, canon_key: str) -> Optional[FieldBusinessRule]:
        """Looks up rule for a normalized canonical key."""
        return self.rules.get(canon_key)

