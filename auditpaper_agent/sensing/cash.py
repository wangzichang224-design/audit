from __future__ import annotations

import json
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd

from auditpaper_agent.contracts import (
    BankConfirmationRow,
    BankJournalRow,
    BankStatementTransaction,
    CaseMetadata,
    SourceDocument,
    SourceRef,
    StandardizedAuditPackage,
    TrialBalanceRow,
)
from auditpaper_agent.sensing.ocr import get_ocr_provider
from auditpaper_agent.trace import AuditTracer, ensure_tracer
from auditpaper_agent.utils import parse_float, sha256_file, write_json


def ingest_cash_case(
    case_dir: str | Path,
    trial_balance_path: str | Path,
    journal_path: str | Path,
    bank_statement_path: str | Path | None = None,
    ocr_provider: str = "pdf-text",
    client_name: str | None = None,
    period_end: str | date | None = None,
    tracer: AuditTracer | None = None,
) -> StandardizedAuditPackage:
    """Parse external cash materials into `case_package.json`.

    This function consumes materials only. It does not create synthetic data and
    does not generate a blank template.
    """
    tracer = ensure_tracer(tracer)
    case_dir = Path(case_dir)
    case_dir.mkdir(parents=True, exist_ok=True)
    tb_path = Path(trial_balance_path)
    journal = Path(journal_path)
    bank_path = Path(bank_statement_path) if bank_statement_path else None

    profile = _load_nearby_profile(tb_path)
    inferred_client = _infer_client_name(tb_path.parent)
    meta = CaseMetadata(
        case_id=case_dir.name,
        client_name=client_name or profile.get("company_name") or profile.get("client_name") or inferred_client or "UNKNOWN_CLIENT",
        period_end=_coerce_date(period_end or profile.get("period_end") or _infer_period_end(tb_path, journal)),
        currency=profile.get("currency", "CNY"),
        gaap=profile.get("gaap", "企业会计准则"),
        te=parse_float(profile.get("te"), 0.0),
        sad=parse_float(profile.get("sad"), 0.0),
    )

    docs = [
        _source_document(tb_path, "trial_balance"),
        _source_document(journal, "bank_journal"),
    ]
    if bank_path:
        docs.append(_source_document(bank_path, "bank_confirmation" if bank_path.is_dir() else "bank_statement"))

    tracer.emit("Sense", "starting cash material ingestion", f"case_dir={case_dir}")
    tracer.emit("Sense", "reading trial balance", str(tb_path))
    trial_balance = _parse_trial_balance(tb_path)
    tracer.emit("Sense", "trial balance parsed", f"rows={len(trial_balance)}")
    tracer.emit("Sense", "reading bank journal", str(journal))
    bank_journal = _parse_bank_journal(journal)
    tracer.emit("Sense", "bank journal parsed", f"rows={len(bank_journal)}")
    confirmations = _parse_confirmations(bank_path, ocr_provider, tracer=tracer) if bank_path else []
    statement_transactions = _parse_bank_statement_transactions(bank_path, tracer=tracer) if bank_path else []

    package = StandardizedAuditPackage(
        meta=meta,
        source_documents=docs,
        trial_balance=trial_balance,
        bank_journal=bank_journal,
        bank_statement_transactions=statement_transactions,
        bank_confirmations=confirmations,
    )
    write_json(case_dir / "case_package.json", package)
    tracer.emit("Output", "wrote standardized audit package", str(case_dir / "case_package.json"))
    return package


def load_case_package(case_dir: str | Path) -> StandardizedAuditPackage:
    data = json.loads((Path(case_dir) / "case_package.json").read_text(encoding="utf-8"))
    return StandardizedAuditPackage.model_validate(data)


def _read_table(path: Path) -> pd.DataFrame:
    return _read_table_with_sheet(path)[0]


def _read_table_with_sheet(path: Path) -> tuple[pd.DataFrame, str | None]:
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        xls = pd.ExcelFile(path)
        sheet_name = str(xls.sheet_names[0])
        return pd.read_excel(xls, sheet_name=sheet_name), sheet_name
    if suffix == ".csv":
        return pd.read_csv(path, encoding="utf-8-sig"), None
    raise ValueError(f"Unsupported table file: {path}")


def _source_document(path: Path, document_type: str) -> SourceDocument:
    if path.is_dir():
        digest = _dir_hash(path)
    else:
        digest = sha256_file(path)
    return SourceDocument(
        document_id=path.stem if path.is_file() else path.name,
        path=str(path),
        document_type=document_type,
        sha256=digest,
    )


def _source_ref(path: Path, document_type: str, sheet: str | None = None, row: int | None = None, page: int | None = None) -> SourceRef:
    digest = _dir_hash(path) if path.is_dir() else sha256_file(path)
    return SourceRef(
        document_id=path.stem if path.is_file() else path.name,
        document_hash=digest,
        file_path=str(path),
        sheet_name=sheet,
        row_number=row,
        page_number=page,
        locator=document_type,
    )


def _parse_trial_balance(path: Path) -> list[TrialBalanceRow]:
    df, sheet = _read_table_with_sheet(path)
    sheet = sheet or "试算平衡表"
    rows: list[TrialBalanceRow] = []
    for idx, raw in df.iterrows():
        account_name = _first(raw, "科目名称", "account_name", "账户名称", "科目")
        if not str(account_name).strip() or str(account_name).lower() == "nan":
            continue
        ending_debit = parse_float(_first(raw, "期末借方余额", "ending_debit", "debit"))
        ending_credit = parse_float(_first(raw, "期末贷方余额", "ending_credit", "credit"))
        if abs(ending_debit) < 0.000001 and abs(ending_credit) < 0.000001:
            ending_balance = parse_float(_first(raw, "期末余额", "ending_balance", "balance"))
            direction = str(_first(raw, "余额方向", "direction", default="借") or "借")
            if "贷" in direction:
                ending_credit = abs(ending_balance)
            else:
                ending_debit = ending_balance
        rows.append(
            TrialBalanceRow(
                account_code=str(_first(raw, "科目编码", "account_code", default="") or ""),
                account_name=str(account_name).strip(),
                ending_debit=ending_debit,
                ending_credit=ending_credit,
                prior_year=parse_float(_first(raw, "上年末审定数", "期初余额", "prior_year", "prior_year_balance")),
                source=_source_ref(path, "trial_balance", sheet=sheet, row=int(idx) + 2),
            )
        )
    return rows


def _parse_bank_journal(path: Path) -> list[BankJournalRow]:
    df, sheet = _read_table_with_sheet(path)
    sheet = sheet or "银行存款日记账"
    rows: list[BankJournalRow] = []
    for idx, raw in df.iterrows():
        if not _is_cash_account_row(raw):
            continue
        txn_date = _coerce_date(_first(raw, "日期", "凭证日期", "记账日期", "业务日期", "交易日期", "date", "txn_date"))
        if txn_date is None:
            continue
        rows.append(
            BankJournalRow(
                txn_date=txn_date,
                bank_name=str(_first(raw, "银行名称", "bank_name", default="") or ""),
                bank_account=_normalize_account(str(_first(raw, "账号", "银行账号", "bank_account", default="") or "")),
                currency=str(_first(raw, "币种", "currency", default="CNY") or "CNY"),
                description=str(_first(raw, "摘要", "description", default="") or ""),
                debit=parse_float(_first(raw, "借方金额", "debit")),
                credit=parse_float(_first(raw, "贷方金额", "credit")),
                balance=parse_float(_first(raw, "余额", "balance")),
                counterparty=str(_first(raw, "对方单位", "对方科目", "来源单号", "counterparty", default="") or ""),
                txn_id=str(_first(raw, "交易编号", "流水号", "凭证号", "txn_id", default="") or ""),
                source=_source_ref(path, "bank_journal", sheet=sheet, row=int(idx) + 2),
            )
        )
    return rows


def _parse_confirmations(path: Path | None, provider_name: str, tracer: AuditTracer | None = None) -> list[BankConfirmationRow]:
    tracer = ensure_tracer(tracer)
    if path is None:
        return []
    if path.is_file() and path.suffix.lower() in {".xlsx", ".xls", ".csv"}:
        tracer.emit("OCR", "reading bank statement/confirmation table", str(path))
        return _parse_confirmation_table(path)

    if path.is_dir():
        table_files = sorted([p for p in path.iterdir() if p.suffix.lower() in {".xlsx", ".xls", ".csv"}])
        table_rows: list[BankConfirmationRow] = []
        for file_path in table_files:
            tracer.emit("OCR", "reading bank statement table", file_path.name)
            table_rows.extend(_parse_confirmation_table(file_path))
        if table_rows:
            tracer.emit("OCR", "statement table extraction complete", f"rows={len(table_rows)}")
            return table_rows

    provider = get_ocr_provider(provider_name, tracer=tracer)
    tracer.emit("OCR", "selected OCR provider", provider.name)
    files = [path] if path.is_file() else sorted([p for p in path.iterdir() if p.suffix.lower() in {".pdf", ".png", ".jpg", ".jpeg"}])
    rows: list[BankConfirmationRow] = []
    for page_idx, file_path in enumerate(files, start=1):
        tracer.emit("OCR", "extracting bank confirmation", file_path.name)
        data = provider.extract_confirmation(file_path)
        if data is None:
            tracer.emit("OCR", "no structured confirmation extracted", file_path.name)
            continue
        rows.append(
            BankConfirmationRow(
                bank_name=str(data.get("bank_name", "")),
                bank_account=_normalize_account(str(data.get("bank_account", ""))),
                currency=str(data.get("currency", "CNY") or "CNY"),
                confirmed_balance=parse_float(data.get("confirmed_balance"), 0.0),
                confirmation_date=_coerce_date(data.get("confirmation_date")),
                restricted_amount=parse_float(data.get("restricted_amount"), 0.0),
                restriction_nature=str(data.get("restriction_nature", "")),
                source=_source_ref(file_path, "bank_confirmation", page=page_idx),
            )
        )
        tracer.emit("OCR", "confirmation parsed", f"{rows[-1].bank_name} {rows[-1].bank_account} balance={rows[-1].confirmed_balance}")
    tracer.emit("OCR", "confirmation extraction complete", f"rows={len(rows)}")
    return rows


def _parse_confirmation_table(path: Path) -> list[BankConfirmationRow]:
    df, sheet = _read_table_with_sheet(path)
    rows: list[BankConfirmationRow] = []
    for idx, raw in df.iterrows():
        bal = parse_float(_first(raw, "confirmed_balance", "回函余额", "确认余额", "银行余额"))
        if abs(bal) < 0.01:
            continue
        rows.append(
            BankConfirmationRow(
                bank_name=str(_first(raw, "bank_name", "银行名称", default="") or ""),
                bank_account=_normalize_account(str(_first(raw, "bank_account", "银行账号", "账号", default="") or "")),
                currency=str(_first(raw, "currency", "币种", default="CNY") or "CNY"),
                confirmed_balance=bal,
                confirmation_date=_coerce_date(_first(raw, "confirmation_date", "回函日期", "函证日期", default=None)),
                restricted_amount=parse_float(_first(raw, "restricted_amount", "受限金额", default=0.0)),
                restriction_nature=str(_first(raw, "restriction_nature", "受限原因", default="") or ""),
                source=_source_ref(path, "bank_confirmation", row=int(idx) + 2),
            )
        )
    if rows:
        return rows

    if "余额" in df.columns and not df.empty:
        valid = df[df["余额"].notna()]
        if valid.empty:
            return []
        last_idx = valid.index[-1]
        raw = valid.loc[last_idx]
        bal = parse_float(raw["余额"])
        if abs(bal) < 0.01:
            return []
        rows.append(
            BankConfirmationRow(
                bank_name=str(sheet or _bank_name_from_filename(path)),
                bank_account=_normalize_account(str(_first(raw, "账号", "银行账号", "bank_account", default="") or "")),
                currency=str(_first(raw, "币种", "currency", default="CNY") or "CNY"),
                confirmed_balance=bal,
                confirmation_date=_coerce_date(_first(raw, "交易日期", "日期", "confirmation_date", default=None)),
                restricted_amount=0.0,
                restriction_nature="",
                source=_source_ref(path, "bank_confirmation", sheet=sheet, row=int(last_idx) + 2),
            )
        )
    return rows


def _parse_bank_statement_transactions(path: Path | None, tracer: AuditTracer | None = None) -> list[BankStatementTransaction]:
    tracer = ensure_tracer(tracer)
    if path is None:
        return []
    files: list[Path]
    if path.is_dir():
        files = sorted([p for p in path.iterdir() if p.suffix.lower() in {".xlsx", ".xls", ".csv"}])
    elif path.suffix.lower() in {".xlsx", ".xls", ".csv"}:
        files = [path]
    else:
        return []

    rows: list[BankStatementTransaction] = []
    for file_path in files:
        df, sheet = _read_table_with_sheet(file_path)
        bank_name = str(sheet or _bank_name_from_filename(file_path))
        for idx, raw in df.iterrows():
            txn_date = _coerce_date(_first(raw, "交易日期", "日期", "date", "txn_date"))
            if txn_date is None:
                continue
            rows.append(
                BankStatementTransaction(
                    txn_date=txn_date,
                    bank_name=bank_name,
                    bank_account=_normalize_account(str(_first(raw, "账号", "银行账号", "bank_account", default="") or "")),
                    currency=str(_first(raw, "币种", "currency", default="CNY") or "CNY"),
                    description=str(_first(raw, "摘要", "description", default="") or ""),
                    debit=parse_float(_first(raw, "借方发生额", "借方金额", "debit")),
                    credit=parse_float(_first(raw, "贷方发生额", "贷方金额", "credit")),
                    balance=parse_float(_first(raw, "余额", "balance")),
                    voucher_no=str(_first(raw, "凭证号", "交易编号", "流水号", "voucher_no", default="") or ""),
                    statement_status=str(_first(raw, "对账状态", "status", default="") or ""),
                    source=_source_ref(file_path, "bank_statement", sheet=sheet, row=int(idx) + 2),
                )
            )
    if rows:
        tracer.emit("Sense", "bank statement transactions parsed", f"rows={len(rows)}")
    return rows


def _first(row: Any, *names: str, default: Any = None) -> Any:
    for name in names:
        if name in row and pd.notna(row[name]):
            return row[name]
    return default


def _is_cash_account_row(row: Any) -> bool:
    account_code = str(_first(row, "科目编码", "account_code", default="") or "")
    account_name = str(_first(row, "科目名称", "账户名称", "account_name", default="") or "")
    if not account_code and not account_name:
        return True
    return account_code.startswith(("1001", "1002", "1009")) or account_name.startswith(("库存现金", "银行存款", "其他货币资金"))


def _coerce_date(value: Any) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.date()


def _infer_period_end(*paths: Path) -> date | None:
    for path in paths:
        match = re.search(r"(20\d{2})", path.name)
        if match:
            return date(int(match.group(1)), 12, 31)
    journal_path = paths[-1]
    try:
        df = _read_table(journal_path)
        date_values = []
        for name in ("日期", "凭证日期", "记账日期", "业务日期", "交易日期", "date", "txn_date"):
            if name in df:
                date_values = list(df.get(name, []))
                break
        dates = [_coerce_date(v) for v in date_values]
        dates = [d for d in dates if d is not None]
        return max(dates) if dates else None
    except Exception:
        return None


def _infer_client_name(materials_dir: Path) -> str | None:
    for path in sorted(materials_dir.glob("企业信用报告_*")):
        stem = path.stem
        if "_" in stem:
            name = stem.split("_", 1)[1].strip()
            if name:
                return name
    return None


def _bank_name_from_filename(path: Path) -> str:
    parts = path.stem.split("_")
    return parts[0] if parts else path.stem


def _load_nearby_profile(path: Path) -> dict[str, Any]:
    for candidate in [path.parent / "client_profile.json", path.parent / "case_metadata.json"]:
        if candidate.exists():
            return json.loads(candidate.read_text(encoding="utf-8"))
    return {}


def _normalize_account(value: str) -> str:
    return "".join(ch for ch in value if ch.isdigit())


def _dir_hash(path: Path) -> str:
    import hashlib

    h = hashlib.sha256()
    for item in sorted(p for p in path.rglob("*") if p.is_file()):
        h.update(str(item.relative_to(path)).encode("utf-8"))
        h.update(sha256_file(item).encode("ascii"))
    return h.hexdigest()
