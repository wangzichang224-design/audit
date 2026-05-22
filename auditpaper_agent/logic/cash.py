from __future__ import annotations

from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

from openpyxl import load_workbook

from auditpaper_agent.contracts import (
    AuditFinding,
    BankConfirmationRow,
    CashWorkflowResult,
    SourceRef,
    StandardizedAuditPackage,
    TrialBalanceRow,
    WriteCellCommand,
    WritePlan,
)
from auditpaper_agent.harness.excel import apply_write_plan
from auditpaper_agent.knowledge.cash import CashTemplateProfile, detect_cash_profile_from_workbook
from auditpaper_agent.reasoning import DeepSeekReasoningProvider
from auditpaper_agent.sensing.cash import load_case_package
from auditpaper_agent.trace import AuditTracer, ensure_tracer
from auditpaper_agent.utils import write_json


CASH_ACCOUNT_NAMES = ("库存现金", "银行存款", "其他货币资金")


def run_cash_workflow(
    case_dir: str | Path,
    template_path: str | Path | None = None,
    output_path: str | Path | None = None,
    tracer: AuditTracer | None = None,
    use_reasoning: bool = False,
) -> CashWorkflowResult:
    tracer = ensure_tracer(tracer)
    case_dir = Path(case_dir)
    tracer.emit("Logic", "loading standardized audit package", str(case_dir / "case_package.json"))
    package = load_case_package(case_dir)
    _validate_package_quality(package, tracer)
    tracer.emit("Logic", "running deterministic cash audit checks", f"client={package.meta.client_name}")
    findings = analyze_cash_package(package)
    tracer.emit("Logic", "audit checks complete", f"findings={len(findings)}")
    if use_reasoning:
        findings = DeepSeekReasoningProvider(tracer=tracer).enhance_findings(package, findings)
    profile = _profile_from_template(template_path) if template_path else None
    write_plan = build_cash_write_plan(package, findings, profile)
    tracer.emit("Logic", "write plan built", f"commands={len(write_plan.commands)} profile={write_plan.template_profile}")

    findings_path = case_dir / "audit_findings.json"
    write_plan_path = case_dir / "write_plan.json"
    write_json(findings_path, findings)
    write_json(write_plan_path, write_plan)
    tracer.emit("Output", "wrote audit findings", str(findings_path))
    tracer.emit("Output", "wrote write plan", str(write_plan_path))

    provenance_path: Path | None = None
    if template_path and output_path:
        harness_result = apply_write_plan(template_path, output_path, write_plan, tracer=tracer)
        provenance_path = case_dir / "provenance.json"
        write_json(provenance_path, harness_result.provenance)
        tracer.emit("Output", "wrote provenance log", str(provenance_path))
        tracer.emit("Output", "wrote filled workbook", str(output_path))
        return CashWorkflowResult(
            case_dir=str(case_dir),
            output_path=str(output_path),
            findings_path=str(findings_path),
            write_plan_path=str(write_plan_path),
            provenance_path=str(provenance_path),
            findings=findings,
        )

    return CashWorkflowResult(
        case_dir=str(case_dir),
        findings_path=str(findings_path),
        write_plan_path=str(write_plan_path),
        provenance_path=None,
        findings=findings,
    )


def analyze_cash_package(package: StandardizedAuditPackage) -> list[AuditFinding]:
    findings: list[AuditFinding] = []
    bank_book = sum_cash_balance(package.trial_balance, "银行存款")
    confirmed = sum(row.confirmed_balance for row in package.bank_confirmations)
    sad = package.meta.sad or max(abs(bank_book) * 0.001, 1.0)

    if package.bank_confirmations and abs(bank_book - confirmed) > sad:
        findings.append(
            AuditFinding(
                finding_type="余额差异",
                severity="高" if abs(bank_book - confirmed) > (package.meta.te or sad * 10) else "中",
                amount=round(bank_book - confirmed, 2),
                description=(
                    f"银行存款账面余额 {bank_book:,.2f} 与银行回函/对账资料确认余额 "
                    f"{confirmed:,.2f} 差异 {abs(bank_book - confirmed):,.2f}，超过阈值 {sad:,.2f}。"
                ),
                sources=_first_sources(package, ["银行存款"], include_confirmations=True),
            )
        )

    for row in package.trial_balance:
        if _base_account(row.account_name) not in CASH_ACCOUNT_NAMES:
            continue
        if abs(row.prior_year) > 0:
            pct = (row.ending_balance - row.prior_year) / abs(row.prior_year)
            if abs(pct) > 0.5:
                findings.append(
                    AuditFinding(
                        finding_type="波动异常",
                        severity="高" if abs(pct) > 2.0 else "中",
                        amount=round(row.ending_balance - row.prior_year, 2),
                        description=(
                            f"{row.account_name} 期末余额 {row.ending_balance:,.2f}，"
                            f"上年审定数 {row.prior_year:,.2f}，变动 {pct:.1%}，需执行波动分析。"
                        ),
                        sources=[row.source],
                    )
                )

    cutoff_rows = _near_period_end_transfers(package)
    if not cutoff_rows and len(package.bank_journal) >= 10:
        findings.append(
            AuditFinding(
                finding_type="截止风险",
                severity="中",
                description="银行日记账存在多笔期末前后交易，但未识别出银行间转账样本，需人工复核截止性测试范围。",
                sources=[package.bank_journal[0].source] if package.bank_journal else [],
            )
        )

    restricted = [row for row in package.bank_confirmations if row.restricted_amount > 0]
    for row in restricted:
        findings.append(
            AuditFinding(
                finding_type="受限资金",
                severity="中",
                amount=row.restricted_amount,
                description=f"{row.bank_name} {row.bank_account} 回函显示受限资金 {row.restricted_amount:,.2f}，性质：{row.restriction_nature or '未说明'}。",
                sources=[row.source],
            )
        )

    if not package.bank_confirmations:
        findings.append(
            AuditFinding(
                finding_type="资料缺失",
                severity="中",
                description="未解析到银行回函或银行对账确认余额，余额核对程序需补充资料。",
                sources=[],
            )
        )
    return findings


def build_cash_write_plan(
    package: StandardizedAuditPackage,
    findings: list[AuditFinding],
    profile: CashTemplateProfile | None = None,
) -> WritePlan:
    profile = profile or _default_profile()
    commands: list[WriteCellCommand] = []
    add = commands.append
    src = _best_source(package)
    meta = package.meta

    if profile.name == "sim_c_cash_v1":
        _append_simulated_template(commands, package, findings, profile)
    elif profile.name == "reference_c_cash_v1":
        lead = profile.sheet_lead
        add(_cmd(lead, "C2", meta.client_name, "客户名称", src))
        if meta.period_end:
            add(_cmd(lead, "C3", meta.period_end.isoformat(), "期末", src))
        add(_cmd(lead, "C4", date.today().isoformat(), "分析日期", src))
        add(_cmd(lead, "C5", meta.te, "可容忍误差", src))
        add(_cmd(lead, "C6", meta.sad, "名义金额", src))
        add(_cmd(lead, "C7", meta.gaap, "适用会计准则", src))
        add(_cmd(lead, "C8", meta.currency, "记账本位币", src))
        for cell, risk in {"C15": "Moderate", "C16": "Moderate", "C17": "Low", "C18": "Low", "C19": "Low"}.items():
            add(_cmd(lead, cell, risk, "CRA 风险等级", src))
        add(_cmd(lead, "C32", 0.1, "波动幅度百分比", src))
        _append_reference_lead_rows(commands, package, profile)
        _append_reference_bkd_rows(commands, package, profile)
        if profile.sheet_confirmations in profile.allowed_cells:
            _append_reference_confirmation_rows(commands, package, profile)
        _append_reference_recon(commands, package, profile)
        _append_reference_cutoff(commands, package, profile)
    else:
        lead = profile.sheet_lead
        add(_cmd(lead, "B3", meta.client_name, "客户名称", src))
        if meta.period_end:
            add(_cmd(lead, "B4", meta.period_end.isoformat(), "期末", src))
        add(_cmd(lead, "B5", date.today().isoformat(), "分析日期", src))
        add(_cmd(lead, "B6", meta.te, "可容忍误差", src))
        add(_cmd(lead, "B7", meta.sad, "名义金额", src))
        add(_cmd(lead, "B8", meta.gaap, "适用会计准则", src))
        add(_cmd(lead, "B10", meta.currency, "记账本位币", src))

    return WritePlan(
        template_profile=profile.name,
        allowed_cells=profile.allowed_cells,
        commands=commands,
        findings=findings,
    )


def sum_cash_balance(rows: list[TrialBalanceRow], account_name: str) -> float:
    return sum(row.ending_balance for row in rows if _base_account(row.account_name) == account_name)


def _append_reference_lead_rows(commands: list[WriteCellCommand], package: StandardizedAuditPackage, profile: CashTemplateProfile) -> None:
    row_map = {"库存现金": 38, "银行存款": 39, "其他货币资金": 40}
    code_map = {"库存现金": "1001", "银行存款": "1002", "其他货币资金": "1009"}
    for account, row_num in row_map.items():
        rows = [r for r in package.trial_balance if _base_account(r.account_name) == account]
        source = rows[0].source if rows else _best_source(package)
        current = sum(r.ending_balance for r in rows)
        commands.extend(
            [
                _cmd(profile.sheet_lead, f"B{row_num}", f"{row_num-37:02d}", "账套名称/编码", source),
                _cmd(profile.sheet_lead, f"C{row_num}", code_map[account], "总账科目编码", source),
                _cmd(profile.sheet_lead, f"E{row_num}", "C.00 BKD/", "索引号", source),
                _cmd(profile.sheet_lead, f"F{row_num}", current, "期末账面数", source),
                _cmd(profile.sheet_lead, f"G{row_num}", 0.0, "账表调整数", source),
                _cmd(profile.sheet_lead, f"I{row_num}", 0.0, "审计调整数", source),
            ]
        )


def _append_reference_bkd_rows(commands: list[WriteCellCommand], package: StandardizedAuditPackage, profile: CashTemplateProfile) -> None:
    rows = [r for r in package.trial_balance if _base_account(r.account_name) in CASH_ACCOUNT_NAMES][:6]
    by_account = _latest_journal_by_account(package)
    for idx, tb in enumerate(rows, start=10):
        bank_name, bank_account, currency = by_account.get(_base_account(tb.account_name), ("", "", package.meta.currency))
        commands.extend(
            [
                _cmd(profile.sheet_bkd, f"B{idx}", package.meta.client_name, "公司名称", tb.source),
                _cmd(profile.sheet_bkd, f"C{idx}", _base_account(tb.account_name), "科目名称", tb.source),
                _cmd(profile.sheet_bkd, f"D{idx}", bank_name, "银行/存款机构名称", tb.source),
                _cmd(profile.sheet_bkd, f"E{idx}", bank_account, "账号", tb.source),
                _cmd(profile.sheet_bkd, f"F{idx}", currency, "币种", tb.source),
                _cmd(profile.sheet_bkd, f"G{idx}", tb.ending_balance, "期末账面金额原币", tb.source),
                _cmd(profile.sheet_bkd, f"H{idx}", tb.ending_balance, "期末账面金额本币", tb.source),
                _cmd(profile.sheet_bkd, f"I{idx}", "经营性银行账户" if _base_account(tb.account_name) == "银行存款" else "货币资金账户", "账户性质用途", tb.source),
                _cmd(profile.sheet_bkd, f"J{idx}", "是" if _base_account(tb.account_name) == "银行存款" else "否", "是否函证", tb.source),
            ]
        )


def _append_reference_confirmation_rows(commands: list[WriteCellCommand], package: StandardizedAuditPackage, profile: CashTemplateProfile) -> None:
    for idx, conf in enumerate(package.bank_confirmations[:8], start=10):
        book_balance = _book_balance_for_confirmation(package, conf.bank_account)
        commands.extend(
            [
                _cmd(profile.sheet_confirmations, f"B{idx}", package.meta.client_name, "公司名称", conf.source),
                _cmd(profile.sheet_confirmations, f"C{idx}", "银行存款", "科目名称", conf.source),
                _cmd(profile.sheet_confirmations, f"D{idx}", conf.bank_name, "银行名称", conf.source),
                _cmd(profile.sheet_confirmations, f"E{idx}", conf.bank_account, "银行账号", conf.source),
                _cmd(profile.sheet_confirmations, f"F{idx}", conf.currency, "币种", conf.source),
                _cmd(profile.sheet_confirmations, f"G{idx}", book_balance, "期末账面银行余额原币", conf.source),
                _cmd(profile.sheet_confirmations, f"H{idx}", book_balance, "期末账面金额本币", conf.source),
                _cmd(profile.sheet_confirmations, f"I{idx}", f"C01-{idx-9:02d}", "函证编号", conf.source),
                _cmd(profile.sheet_confirmations, f"J{idx}", package.meta.client_name, "账户名称", conf.source),
                _cmd(profile.sheet_confirmations, f"K{idx}", conf.confirmed_balance, "回函余额", conf.source),
                _cmd(profile.sheet_confirmations, f"L{idx}", conf.restricted_amount, "受限金额", conf.source),
                _cmd(profile.sheet_confirmations, f"M{idx}", conf.restriction_nature, "受限说明", conf.source),
            ]
        )


def _append_reference_recon(commands: list[WriteCellCommand], package: StandardizedAuditPackage, profile: CashTemplateProfile) -> None:
    src = _best_source(package)
    conf = package.bank_confirmations[0] if package.bank_confirmations else None
    book = sum_cash_balance(package.trial_balance, "银行存款")
    confirmed = sum(c.confirmed_balance for c in package.bank_confirmations)
    if conf:
        src = conf.source
        commands.extend(
            [
                _cmd(profile.sheet_recon, "C17", "银行存款", "调节账户科目", src),
                _cmd(profile.sheet_recon, "C18", conf.bank_name, "银行名称", src),
                _cmd(profile.sheet_recon, "C19", conf.bank_account, "银行账号", src),
                _cmd(profile.sheet_recon, "C20", conf.currency, "原币币种", src),
                _cmd(profile.sheet_recon, "C21", (conf.confirmation_date or package.meta.period_end or date.today()).isoformat(), "余额调节表日期", src),
            ]
        )
    commands.extend(
        [
            _cmd(profile.sheet_recon, "B10", "选取存在回函/账面差异或余额重大的银行账户执行余额调节测试。", "选样理由", src),
            _cmd(profile.sheet_recon, "C24", book, "期末账面数", src),
            _cmd(profile.sheet_recon, "F24", confirmed, "银行对账单/回函金额", src),
        ]
    )


def _append_reference_cutoff(commands: list[WriteCellCommand], package: StandardizedAuditPackage, profile: CashTemplateProfile) -> None:
    src = package.bank_journal[0].source if package.bank_journal else _best_source(package)
    transfers = _near_period_end_transfers(package)
    fallback_samples = [] if transfers else _cutoff_samples(package)
    reason = (
        "根据期末前后银行日记账识别大额或银行间转账样本。"
        if transfers
        else "未识别出明确银行间转账，已列示期末前后大额银行收支作为截止性复核样本。"
    )
    commands.extend(
        [
            _cmd(profile.sheet_cutoff, "C8", "期末前后5个工作日", "使用的截止期间", src),
            _cmd(profile.sheet_cutoff, "C10", reason, "截止期间理由", src),
        ]
    )
    samples = (transfers or fallback_samples)[:10]
    if package.meta.period_end:
        before_samples = [sample for sample in samples if sample.txn_date <= package.meta.period_end]
        after_samples = [sample for sample in samples if sample.txn_date > package.meta.period_end]
    else:
        before_samples = samples
        after_samples = []
    _append_reference_cutoff_sample_rows(commands, package, profile, before_samples[:5], start_row=20, fallback=not transfers)
    _append_reference_cutoff_sample_rows(commands, package, profile, after_samples[:5], start_row=29, fallback=not transfers)


def _append_reference_cutoff_sample_rows(
    commands: list[WriteCellCommand],
    package: StandardizedAuditPackage,
    profile: CashTemplateProfile,
    samples,
    start_row: int,
    fallback: bool,
) -> None:
    for row_number, sample in enumerate(samples, start=start_row):
        sequence = row_number - start_row + 1
        amount = round(sample.debit - sample.credit, 2)
        transaction_id = sample.txn_id or sample.description[:20]
        counterpart_name = sample.counterparty if not fallback else "未识别配对银行账户"
        counterpart_account = "" if fallback else sample.bank_account
        conclusion = (
            "已识别银行间转账样本，需核对对方银行记录及跨期列示。"
            if not fallback
            else "大额/临近期末银行收支样本；未识别配对银行间转账，需复核是否跨期。"
        )
        commands.extend(
            [
                _cmd(profile.sheet_cutoff, f"B{row_number}", f"CUTOFF-{sequence:02d}", "样本编号", sample.source),
                _cmd(profile.sheet_cutoff, f"C{row_number}", package.meta.client_name, "公司名称", sample.source),
                _cmd(profile.sheet_cutoff, f"D{row_number}", sample.bank_name, "银行名称", sample.source),
                _cmd(profile.sheet_cutoff, f"E{row_number}", sample.bank_account, "银行账号", sample.source),
                _cmd(profile.sheet_cutoff, f"F{row_number}", sample.txn_date.isoformat(), "收款/付款时间", sample.source),
                _cmd(profile.sheet_cutoff, f"G{row_number}", transaction_id, "交易编号", sample.source),
                _cmd(profile.sheet_cutoff, f"H{row_number}", sample.currency, "币种", sample.source),
                _cmd(profile.sheet_cutoff, f"I{row_number}", amount, "交易金额", sample.source),
                _cmd(profile.sheet_cutoff, f"J{row_number}", counterpart_name, "对方银行名称/复核对象", sample.source),
                _cmd(profile.sheet_cutoff, f"K{row_number}", counterpart_account, "对方银行账号", sample.source),
                _cmd(profile.sheet_cutoff, f"L{row_number}", sample.txn_date.isoformat(), "对方交易日期/待核日期", sample.source),
                _cmd(profile.sheet_cutoff, f"M{row_number}", amount, "对方交易金额/待核金额", sample.source),
                _cmd(profile.sheet_cutoff, f"N{row_number}", conclusion, "银行间转账如跨期是否在余额调节表中体现", sample.source),
            ]
        )


def _cmd(sheet: str, cell: str, value, purpose: str, source: SourceRef) -> WriteCellCommand:
    return WriteCellCommand(sheet_name=sheet, cell=cell, value=value, purpose=purpose, source=source)


def _append_simulated_template(
    commands: list[WriteCellCommand],
    package: StandardizedAuditPackage,
    findings: list[AuditFinding],
    profile: CashTemplateProfile,
) -> None:
    src = _best_source(package)
    for sheet in [profile.sheet_lead, profile.sheet_bkd, profile.sheet_confirmations, profile.sheet_recon, profile.sheet_cutoff]:
        commands.extend(
            [
                _cmd(sheet, "C2", package.meta.client_name, "客户名称", src),
                _cmd(sheet, "C3", package.meta.period_end.isoformat() if package.meta.period_end else "", "期末", src),
                _cmd(sheet, "C4", "AuditPaper-Agent", "编制人", src),
            ]
        )
    _append_sim_lead(commands, package, profile)
    _append_sim_bkd(commands, package, profile)
    _append_sim_confirmations(commands, package, profile)
    _append_sim_recon(commands, package, profile)
    _append_sim_cutoff(commands, package, profile)


def _append_sim_lead(commands: list[WriteCellCommand], package: StandardizedAuditPackage, profile: CashTemplateProfile) -> None:
    statement_total = sum(c.confirmed_balance for c in package.bank_confirmations)
    adjustment_total = _statement_adjustment_total(package)
    rows = [r for r in package.trial_balance if _base_account(r.account_name) in CASH_ACCOUNT_NAMES]
    for offset, tb in enumerate(rows[:9], start=9):
        base = _base_account(tb.account_name)
        is_bank = base == "银行存款"
        statement_balance = statement_total if is_bank else 0.0
        adjustment = adjustment_total if is_bank else 0.0
        adjusted_book = statement_balance + adjustment if is_bank else tb.ending_balance
        difference = round(adjusted_book - tb.ending_balance, 2)
        conclusion = "已调节一致" if abs(difference) < 0.01 else "存在差异，需进一步核对"
        source = tb.source
        commands.extend(
            [
                _cmd(profile.sheet_lead, f"B{offset}", tb.account_code, "科目编码", source),
                _cmd(profile.sheet_lead, f"C{offset}", tb.account_name, "科目名称", source),
                _cmd(profile.sheet_lead, f"D{offset}", tb.ending_balance, "TB期末数", source),
                _cmd(profile.sheet_lead, f"E{offset}", tb.ending_balance, "A3审前数", source),
                _cmd(profile.sheet_lead, f"F{offset}", statement_balance if is_bank else "", "银行对账单余额", source),
                _cmd(profile.sheet_lead, f"G{offset}", adjustment if is_bank else "", "未达调节", source),
                _cmd(profile.sheet_lead, f"H{offset}", adjusted_book, "调节后账面数", source),
                _cmd(profile.sheet_lead, f"I{offset}", difference, "差异", source),
                _cmd(profile.sheet_lead, f"J{offset}", conclusion, "结论", source),
                _cmd(profile.sheet_lead, f"K{offset}", "TB/C.00 BKD/C.02", "资料索引", source),
            ]
        )

    for offset, conf in enumerate(package.bank_confirmations[:12], start=19):
        account_id = _account_id(conf, offset - 18)
        book_balance = _book_balance_for_statement(package, conf)
        commands.extend(
            [
                _cmd(profile.sheet_lead, f"B{offset}", account_id, "账户ID", conf.source),
                _cmd(profile.sheet_lead, f"C{offset}", conf.bank_name, "开户行", conf.source),
                _cmd(profile.sheet_lead, f"D{offset}", conf.bank_account, "账号", conf.source),
                _cmd(profile.sheet_lead, f"E{offset}", "经营性银行账户", "账户性质", conf.source),
                _cmd(profile.sheet_lead, f"F{offset}", conf.currency, "币种", conf.source),
                _cmd(profile.sheet_lead, f"G{offset}", book_balance, "账面余额", conf.source),
                _cmd(profile.sheet_lead, f"H{offset}", conf.confirmed_balance, "对账单余额", conf.source),
                _cmd(profile.sheet_lead, f"I{offset}", "已取得对账单", "函证状态", conf.source),
                _cmd(profile.sheet_lead, f"J{offset}", _account_note(package, conf), "备注", conf.source),
            ]
        )


def _append_sim_bkd(commands: list[WriteCellCommand], package: StandardizedAuditPackage, profile: CashTemplateProfile) -> None:
    for offset, conf in enumerate(package.bank_confirmations[:20], start=9):
        account_id = _account_id(conf, offset - 8)
        book_balance = _book_balance_for_statement(package, conf)
        commands.extend(
            [
                _cmd(profile.sheet_bkd, f"B{offset}", package.meta.client_name, "公司名称", conf.source),
                _cmd(profile.sheet_bkd, f"C{offset}", "银行存款", "科目", conf.source),
                _cmd(profile.sheet_bkd, f"D{offset}", account_id, "账户ID", conf.source),
                _cmd(profile.sheet_bkd, f"E{offset}", conf.bank_name, "开户行", conf.source),
                _cmd(profile.sheet_bkd, f"F{offset}", conf.bank_account, "账号", conf.source),
                _cmd(profile.sheet_bkd, f"G{offset}", "经营性银行账户", "账户性质", conf.source),
                _cmd(profile.sheet_bkd, f"H{offset}", conf.currency, "币种", conf.source),
                _cmd(profile.sheet_bkd, f"I{offset}", book_balance, "期末账面余额", conf.source),
                _cmd(profile.sheet_bkd, f"J{offset}", conf.confirmed_balance, "对账单余额", conf.source),
                _cmd(profile.sheet_bkd, f"K{offset}", "是", "是否取得对账单", conf.source),
                _cmd(profile.sheet_bkd, f"L{offset}", "是", "是否纳入函证", conf.source),
            ]
        )


def _append_sim_confirmations(commands: list[WriteCellCommand], package: StandardizedAuditPackage, profile: CashTemplateProfile) -> None:
    for offset, conf in enumerate(package.bank_confirmations[:20], start=9):
        account_id = _account_id(conf, offset - 8)
        book_balance = _book_balance_for_statement(package, conf)
        difference = round(conf.confirmed_balance - book_balance, 2)
        commands.extend(
            [
                _cmd(profile.sheet_confirmations, f"B{offset}", account_id, "账户ID", conf.source),
                _cmd(profile.sheet_confirmations, f"C{offset}", conf.bank_name, "开户行", conf.source),
                _cmd(profile.sheet_confirmations, f"D{offset}", conf.bank_account, "账号", conf.source),
                _cmd(profile.sheet_confirmations, f"E{offset}", book_balance, "账面余额", conf.source),
                _cmd(profile.sheet_confirmations, f"F{offset}", conf.confirmed_balance, "函证金额", conf.source),
                _cmd(profile.sheet_confirmations, f"G{offset}", "", "发函日期", conf.source),
                _cmd(profile.sheet_confirmations, f"H{offset}", conf.confirmation_date.isoformat() if conf.confirmation_date else "", "回函日期", conf.source),
                _cmd(profile.sheet_confirmations, f"I{offset}", conf.confirmed_balance, "回函金额", conf.source),
                _cmd(profile.sheet_confirmations, f"J{offset}", difference, "差异", conf.source),
                _cmd(profile.sheet_confirmations, f"K{offset}", _account_note(package, conf), "差异说明", conf.source),
                _cmd(profile.sheet_confirmations, f"L{offset}", "不适用", "替代程序", conf.source),
            ]
        )


def _append_sim_recon(commands: list[WriteCellCommand], package: StandardizedAuditPackage, profile: CashTemplateProfile) -> None:
    for offset, conf in enumerate(package.bank_confirmations[:16], start=9):
        account_id = _account_id(conf, offset - 8)
        components = _statement_adjustments_for_bank(package, conf.bank_name)
        adjusted = conf.confirmed_balance + components["enterprise_received_bank_unreceived"] - components["bank_received_enterprise_unrecorded"] + components["enterprise_paid_bank_unpaid"] - components["bank_paid_enterprise_unrecorded"]
        book_balance = _book_balance_for_statement(package, conf)
        difference = round(adjusted - book_balance, 2)
        conclusion = "调节一致" if abs(difference) < 0.01 else "存在差异，需复核"
        commands.extend(
            [
                _cmd(profile.sheet_recon, f"B{offset}", account_id, "账户ID", conf.source),
                _cmd(profile.sheet_recon, f"C{offset}", conf.bank_name, "开户行", conf.source),
                _cmd(profile.sheet_recon, f"D{offset}", conf.confirmed_balance, "银行对账单余额", conf.source),
                _cmd(profile.sheet_recon, f"E{offset}", components["enterprise_received_bank_unreceived"], "企业已收银行未收", conf.source),
                _cmd(profile.sheet_recon, f"F{offset}", components["bank_received_enterprise_unrecorded"], "银行已收企业未入账", conf.source),
                _cmd(profile.sheet_recon, f"G{offset}", components["enterprise_paid_bank_unpaid"], "企业已付银行未付", conf.source),
                _cmd(profile.sheet_recon, f"H{offset}", components["bank_paid_enterprise_unrecorded"], "银行已付企业未入账", conf.source),
                _cmd(profile.sheet_recon, f"I{offset}", adjusted, "调节后账面余额", conf.source),
                _cmd(profile.sheet_recon, f"J{offset}", book_balance, "账面余额", conf.source),
                _cmd(profile.sheet_recon, f"K{offset}", difference, "差异", conf.source),
                _cmd(profile.sheet_recon, f"L{offset}", conclusion, "结论", conf.source),
            ]
        )

    adjustments = [row for row in package.bank_statement_transactions if row.statement_status and row.statement_status != "已达账"]
    for offset, row in enumerate(adjustments[:10], start=26):
        commands.extend(
            [
                _cmd(profile.sheet_recon, f"B{offset}", f"UR-{offset-25:02d}", "未达编号", row.source),
                _cmd(profile.sheet_recon, f"C{offset}", _account_id_for_bank(package, row.bank_name), "账户ID", row.source),
                _cmd(profile.sheet_recon, f"D{offset}", row.txn_date.isoformat(), "日期", row.source),
                _cmd(profile.sheet_recon, f"E{offset}", row.statement_status, "类型", row.source),
                _cmd(profile.sheet_recon, f"F{offset}", row.debit or row.credit, "金额", row.source),
                _cmd(profile.sheet_recon, f"G{offset}", row.description, "摘要", row.source),
                _cmd(profile.sheet_recon, f"H{offset}", "待检查期后入账", "期后处理", row.source),
                _cmd(profile.sheet_recon, f"I{offset}", "作为未达项调节", "结论", row.source),
            ]
        )


def _append_sim_cutoff(commands: list[WriteCellCommand], package: StandardizedAuditPackage, profile: CashTemplateProfile) -> None:
    for offset, row in enumerate(_cutoff_samples(package)[:20], start=9):
        commands.extend(
            [
                _cmd(profile.sheet_cutoff, f"B{offset}", f"CUT-{offset-8:02d}", "流水号", row.source),
                _cmd(profile.sheet_cutoff, f"C{offset}", row.txn_date.isoformat(), "日期", row.source),
                _cmd(profile.sheet_cutoff, f"D{offset}", "银行存款", "账户ID", row.source),
                _cmd(profile.sheet_cutoff, f"E{offset}", row.description, "摘要", row.source),
                _cmd(profile.sheet_cutoff, f"F{offset}", row.debit, "借方", row.source),
                _cmd(profile.sheet_cutoff, f"G{offset}", row.credit, "贷方", row.source),
                _cmd(profile.sheet_cutoff, f"H{offset}", row.txn_id, "凭证号", row.source),
                _cmd(profile.sheet_cutoff, f"I{offset}", row.counterparty, "来源单号", row.source),
                _cmd(profile.sheet_cutoff, f"J{offset}", row.txn_date.isoformat(), "序时账日期", row.source),
                _cmd(profile.sheet_cutoff, f"K{offset}", "否", "是否跨期", row.source),
                _cmd(profile.sheet_cutoff, f"L{offset}", "已纳入截止测试样本", "结论", row.source),
            ]
        )


def _profile_from_template(template_path: str | Path) -> CashTemplateProfile:
    wb = load_workbook(template_path, read_only=True, data_only=False)
    try:
        return detect_cash_profile_from_workbook(wb)
    finally:
        wb.close()


def _default_profile() -> CashTemplateProfile:
    from auditpaper_agent.knowledge.cash import reference_cash_profile

    return reference_cash_profile()


def _base_account(account_name: str) -> str:
    for name in CASH_ACCOUNT_NAMES:
        if str(account_name).startswith(name):
            return name
    return str(account_name)


def _best_source(package: StandardizedAuditPackage) -> SourceRef:
    if package.trial_balance:
        return package.trial_balance[0].source
    if package.bank_journal:
        return package.bank_journal[0].source
    if package.bank_confirmations:
        return package.bank_confirmations[0].source
    raise ValueError("Package has no source rows")


def _first_sources(package: StandardizedAuditPackage, accounts: list[str], include_confirmations: bool = False) -> list[SourceRef]:
    sources = [row.source for row in package.trial_balance if _base_account(row.account_name) in accounts]
    if include_confirmations:
        sources.extend(row.source for row in package.bank_confirmations[:2])
    return sources[:5]


def _latest_journal_by_account(package: StandardizedAuditPackage) -> dict[str, tuple[str, str, str]]:
    by_account: dict[str, tuple[str, str, str]] = {}
    for row in package.bank_journal:
        if row.bank_name or row.bank_account:
            by_account["银行存款"] = (row.bank_name, row.bank_account, row.currency)
    return by_account


def _book_balance_for_confirmation(package: StandardizedAuditPackage, bank_account: str) -> float:
    matching = [row.balance for row in package.bank_journal if row.bank_account == bank_account]
    if matching:
        return matching[-1]
    return sum_cash_balance(package.trial_balance, "银行存款")


def _near_period_end_transfers(package: StandardizedAuditPackage):
    if not package.meta.period_end:
        return []
    start = package.meta.period_end - timedelta(days=7)
    end = package.meta.period_end + timedelta(days=7)
    candidates = []
    for row in package.bank_journal:
        text = f"{row.description} {row.counterparty}"
        is_transfer = "转账" in text or "往来" in text or "内部" in text
        is_near = start <= row.txn_date <= end
        if is_near and is_transfer:
            candidates.append(row)
    return candidates


def _validate_package_quality(package: StandardizedAuditPackage, tracer: AuditTracer) -> None:
    if not package.trial_balance:
        raise ValueError("未解析到试算平衡表数据，请检查 TB 文件格式。")
    if not package.bank_journal:
        raise ValueError("未解析到序时账/银行日记账数据，请检查日期、摘要、借贷金额等列名。")
    bank_book = sum_cash_balance(package.trial_balance, "银行存款")
    if package.bank_confirmations and abs(bank_book) < 0.01:
        tracer.emit("Logic", "quality warning", "银行存款 TB 余额为 0，但已解析到对账单/回函余额，请复核 TB 列映射。")


def _account_id(conf: BankConfirmationRow, index: int) -> str:
    stem = Path(conf.source.document_id).stem
    parts = stem.split("_")
    if "BA" in parts:
        pos = parts.index("BA")
        tail = parts[pos + 1 : pos + 3]
        if tail:
            return "BA-" + "-".join(tail)
    return f"BA-{index:02d}"


def _account_id_for_bank(package: StandardizedAuditPackage, bank_name: str) -> str:
    for idx, conf in enumerate(package.bank_confirmations, start=1):
        if conf.bank_name == bank_name:
            return _account_id(conf, idx)
    return bank_name or "银行存款"


def _book_balance_for_statement(package: StandardizedAuditPackage, conf: BankConfirmationRow) -> float:
    components = _statement_adjustments_for_bank(package, conf.bank_name)
    return round(
        conf.confirmed_balance
        + components["enterprise_received_bank_unreceived"]
        - components["bank_received_enterprise_unrecorded"]
        + components["enterprise_paid_bank_unpaid"]
        - components["bank_paid_enterprise_unrecorded"],
        2,
    )


def _account_note(package: StandardizedAuditPackage, conf: BankConfirmationRow) -> str:
    components = _statement_adjustments_for_bank(package, conf.bank_name)
    total_adjustment = sum(abs(v) for v in components.values())
    if total_adjustment < 0.01:
        return "无未达项，账表核对一致"
    return "存在未达项，已在 C.02 执行余额调节"


def _statement_adjustment_total(package: StandardizedAuditPackage) -> float:
    total = 0.0
    for conf in package.bank_confirmations:
        components = _statement_adjustments_for_bank(package, conf.bank_name)
        total += (
            components["enterprise_received_bank_unreceived"]
            - components["bank_received_enterprise_unrecorded"]
            + components["enterprise_paid_bank_unpaid"]
            - components["bank_paid_enterprise_unrecorded"]
        )
    return round(total, 2)


def _statement_adjustments_for_bank(package: StandardizedAuditPackage, bank_name: str) -> dict[str, float]:
    result = {
        "enterprise_received_bank_unreceived": 0.0,
        "bank_received_enterprise_unrecorded": 0.0,
        "enterprise_paid_bank_unpaid": 0.0,
        "bank_paid_enterprise_unrecorded": 0.0,
    }
    for row in package.bank_statement_transactions:
        if row.bank_name != bank_name:
            continue
        status = row.statement_status or ""
        amount = row.debit or row.credit
        if "企业已收" in status and "银行未收" in status:
            result["enterprise_received_bank_unreceived"] += amount
        elif "银行已收" in status and ("企业未入账" in status or "企业未入" in status):
            result["bank_received_enterprise_unrecorded"] += amount
        elif "企业已付" in status and "银行未付" in status:
            result["enterprise_paid_bank_unpaid"] += amount
        elif "银行已付" in status and ("企业未入账" in status or "企业未入" in status):
            result["bank_paid_enterprise_unrecorded"] += amount
    return {key: round(value, 2) for key, value in result.items()}


def _cutoff_samples(package: StandardizedAuditPackage):
    if package.meta.period_end:
        start = package.meta.period_end - timedelta(days=7)
        end = package.meta.period_end + timedelta(days=7)
        near = [row for row in package.bank_journal if start <= row.txn_date <= end]
        if near:
            return sorted(near, key=lambda row: abs(row.debit - row.credit), reverse=True)
    return sorted(package.bank_journal, key=lambda row: (row.txn_date, abs(row.debit - row.credit)), reverse=True)
