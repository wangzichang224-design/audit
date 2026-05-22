from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator


class AuditModel(BaseModel):
    model_config = ConfigDict(extra="forbid", populate_by_name=True)


class SourceRef(AuditModel):
    document_id: str
    document_hash: str
    file_path: str
    sheet_name: str | None = None
    row_number: int | None = None
    page_number: int | None = None
    locator: str | None = None

    def label(self) -> str:
        bits = [self.document_id]
        if self.sheet_name:
            bits.append(f"sheet={self.sheet_name}")
        if self.row_number is not None:
            bits.append(f"row={self.row_number}")
        if self.page_number is not None:
            bits.append(f"page={self.page_number}")
        if self.locator:
            bits.append(self.locator)
        return "; ".join(bits)


DocumentType = Literal[
    "trial_balance",
    "general_ledger",
    "bank_journal",
    "bank_confirmation",
    "bank_statement",
    "customers",
    "suppliers",
    "inventory",
    "fixed_assets",
    "sales",
    "purchase",
    "unknown",
]


class SourceDocument(AuditModel):
    document_id: str
    path: str
    document_type: DocumentType
    sha256: str


class TrialBalanceRow(AuditModel):
    account_code: str = ""
    account_name: str
    ending_debit: float = 0.0
    ending_credit: float = 0.0
    prior_year: float = 0.0
    source: SourceRef

    @property
    def ending_balance(self) -> float:
        return self.ending_debit - self.ending_credit


class BankJournalRow(AuditModel):
    txn_date: date
    bank_name: str = ""
    bank_account: str = ""
    currency: str = "CNY"
    description: str = ""
    debit: float = 0.0
    credit: float = 0.0
    balance: float = 0.0
    counterparty: str = ""
    txn_id: str = ""
    source: SourceRef


class BankStatementTransaction(AuditModel):
    txn_date: date
    bank_name: str = ""
    bank_account: str = ""
    currency: str = "CNY"
    description: str = ""
    debit: float = 0.0
    credit: float = 0.0
    balance: float = 0.0
    voucher_no: str = ""
    statement_status: str = ""
    source: SourceRef


class BankConfirmationRow(AuditModel):
    bank_name: str = ""
    bank_account: str = ""
    currency: str = "CNY"
    confirmed_balance: float
    confirmation_date: date | None = None
    restricted_amount: float = 0.0
    restriction_nature: str = ""
    source: SourceRef


class CaseMetadata(AuditModel):
    case_id: str
    client_name: str = "UNKNOWN_CLIENT"
    period_end: date | None = None
    currency: str = "CNY"
    gaap: str = "企业会计准则"
    te: float = 0.0
    sad: float = 0.0


class StandardizedAuditPackage(AuditModel):
    schema_version: str = "auditpaper.standardized_package.v1"
    meta: CaseMetadata
    source_documents: list[SourceDocument] = Field(default_factory=list)
    trial_balance: list[TrialBalanceRow] = Field(default_factory=list)
    bank_journal: list[BankJournalRow] = Field(default_factory=list)
    bank_statement_transactions: list[BankStatementTransaction] = Field(default_factory=list)
    bank_confirmations: list[BankConfirmationRow] = Field(default_factory=list)


class GeneralLedgerRow(AuditModel):
    posting_date: date | None = None
    voucher_id: str = ""
    line_no: str = ""
    company_code: str = ""
    account_code: str = ""
    account_name: str = ""
    description: str = ""
    debit: float = 0.0
    credit: float = 0.0
    currency: str = "CNY"
    counterparty: str = ""
    customer_id: str = ""
    supplier_id: str = ""
    material_id: str = ""
    asset_id: str = ""
    source: SourceRef

    @property
    def amount(self) -> float:
        return self.debit - self.credit


MasterRecordType = Literal["customer", "supplier", "inventory", "fixed_asset", "sales", "purchase"]


class MasterDataRecord(AuditModel):
    record_type: MasterRecordType
    record_id: str = ""
    name: str = ""
    amount: float = 0.0
    currency: str = "CNY"
    date_value: date | None = None
    attributes: dict[str, Any] = Field(default_factory=dict)
    source: SourceRef


class ErpFieldMapping(AuditModel):
    canonical_field: str
    source_header: str = ""
    confidence: float = 0.0
    required: bool = False
    notes: str = ""


class ErpTableMapping(AuditModel):
    source_type: DocumentType
    document_id: str
    path: str
    sheet_name: str = ""
    row_count: int = 0
    header_row: int = 1
    matched_fields: list[ErpFieldMapping] = Field(default_factory=list)
    missing_required: list[str] = Field(default_factory=list)
    confidence: float = 0.0
    sample_rows: list[dict[str, Any]] = Field(default_factory=list)
    debit_total: float = 0.0
    credit_total: float = 0.0
    balance_total: float = 0.0


class ErpMappingManifest(AuditModel):
    schema_version: str = "auditpaper.erp_mapping_manifest.v1"
    provider: Literal["sap", "yonyou", "auto"] = "auto"
    root_path: str
    confirmed: bool = False
    generated_at: str
    tables: list[ErpTableMapping] = Field(default_factory=list)
    blocking_issues: list[str] = Field(default_factory=list)

    @property
    def can_import(self) -> bool:
        required_types = {"trial_balance", "general_ledger"}
        present = {table.source_type for table in self.tables if not table.missing_required}
        return self.confirmed and not self.blocking_issues and required_types.issubset(present)


class StandardizedErpAuditPackage(AuditModel):
    schema_version: str = "auditpaper.erp_standardized_package.v1"
    meta: CaseMetadata
    provider: Literal["sap", "yonyou", "auto"] = "auto"
    source_documents: list[SourceDocument] = Field(default_factory=list)
    mapping_confirmed: bool = False
    trial_balance: list[TrialBalanceRow] = Field(default_factory=list)
    general_ledger: list[GeneralLedgerRow] = Field(default_factory=list)
    bank_journal: list[BankJournalRow] = Field(default_factory=list)
    customers: list[MasterDataRecord] = Field(default_factory=list)
    suppliers: list[MasterDataRecord] = Field(default_factory=list)
    inventory: list[MasterDataRecord] = Field(default_factory=list)
    fixed_assets: list[MasterDataRecord] = Field(default_factory=list)
    sales: list[MasterDataRecord] = Field(default_factory=list)
    purchase: list[MasterDataRecord] = Field(default_factory=list)


class AuditFinding(AuditModel):
    finding_type: Literal["余额差异", "截止风险", "波动异常", "资料缺失", "受限资金", "提示"]
    severity: Literal["低", "中", "高"]
    description: str
    amount: float | None = None
    sources: list[SourceRef] = Field(default_factory=list)


class WriteCellCommand(AuditModel):
    sheet_name: str
    cell: str
    value: Any
    purpose: str
    source: SourceRef

    @field_validator("cell")
    @classmethod
    def uppercase_cell(cls, value: str) -> str:
        return value.upper()


class WritePlan(AuditModel):
    schema_version: str = "auditpaper.write_plan.v1"
    template_profile: str
    allowed_cells: dict[str, list[str]]
    commands: list[WriteCellCommand] = Field(default_factory=list)
    findings: list[AuditFinding] = Field(default_factory=list)


class ProvenanceEntry(AuditModel):
    sheet_name: str
    cell: str
    value_repr: str
    purpose: str
    source: SourceRef


class HarnessResult(AuditModel):
    output_path: str
    commands_applied: int
    provenance: list[ProvenanceEntry] = Field(default_factory=list)


class CashWorkflowResult(AuditModel):
    case_dir: str
    output_path: str | None = None
    findings_path: str
    write_plan_path: str
    provenance_path: str | None = None
    findings: list[AuditFinding]


def ensure_path(value: str | Path) -> Path:
    return value if isinstance(value, Path) else Path(value)
