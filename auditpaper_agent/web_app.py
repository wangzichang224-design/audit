from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path
from typing import Any

if __package__ in {None, ""}:  # Streamlit may execute this file outside package context.
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

try:
    import streamlit as st
except ModuleNotFoundError as exc:  # pragma: no cover - runtime environment guard
    raise RuntimeError("Streamlit is required. Install dependencies, then run: streamlit run auditpaper_agent/web_app.py") from exc

from auditpaper_agent.agent import ProviderHealth, check_provider_health
from auditpaper_agent.contracts import ErpMappingManifest
from auditpaper_agent.diagnostics import ClientPackageDiagnostics, diagnose_client_package
from auditpaper_agent.service import AutoCashRunResult, CashMaterialsSummary, inspect_cash_materials, run_auto_cash_case
from auditpaper_agent.sensing.erp import diagnose_erp_export, import_erp_export
from auditpaper_agent.suite import (
    WorkpaperProjectSummary,
    WorkpaperSuiteResult,
    inspect_workpaper_project,
    run_auto_workpaper_suite,
    run_erp_workpaper_suite,
)
from auditpaper_agent.utils import safe_filename


def main() -> None:
    st.set_page_config(page_title="AuditPaper-Agent", page_icon="AP", layout="wide")
    _init_state()

    st.title("AuditPaper-Agent")
    st.caption("本地路径对话框版：粘贴客户资料包或完整拟真项目路径，先识别资料，再确认生成底稿。")

    with st.sidebar:
        st.subheader("运行设置")
        st.session_state.ocr_provider = st.selectbox(
            "OCR provider",
            options=["pdf-text", "qwen-ocr", "textin", "stub"],
            index=0,
            help="默认 pdf-text 不调用付费 API；qwen-ocr / textin 需要 .env 配置。",
        )
        st.session_state.use_reasoning = st.checkbox(
            "使用 DeepSeek 优化审计发现文字",
            value=False,
            help="只优化发现描述，不改变金额、来源或 Excel 写入单元格。",
        )
        st.divider()
        if st.button("测试 API / Provider", use_container_width=True):
            with st.spinner("正在测试 provider..."):
                st.session_state.provider_health = check_provider_health(live=True)
        if st.session_state.get("provider_health"):
            _render_provider_health(st.session_state.provider_health)
        st.divider()
        st.write("完整项目会生成全套底稿 ZIP；单 C 资料包会生成货币资金底稿。")

    _render_history()

    submitted = st.chat_input("粘贴客户资料包或完整拟真项目文件夹路径，例如 D:\\...\\audit_sim_autoparts_2025_v2")
    if submitted:
        _handle_path(submitted)
        st.rerun()

    project: WorkpaperProjectSummary | None = st.session_state.project_summary
    discovery: CashMaterialsSummary | None = st.session_state.discovery
    diagnostics: ClientPackageDiagnostics | None = st.session_state.diagnostics
    erp_manifest: ErpMappingManifest | None = st.session_state.erp_manifest
    result: AutoCashRunResult | WorkpaperSuiteResult | None = st.session_state.result

    if diagnostics:
        _render_diagnostics(diagnostics)

    if project and project.is_ready and (diagnostics is None or diagnostics.can_generate):
        _render_project(project)
        if result is None and st.button("确认生成全套企业级底稿", type="primary", use_container_width=True):
            _run_suite(project.project_dir or project.input_path)
            st.rerun()
    elif discovery:
        _render_discovery(discovery)
        if discovery.is_ready and result is None and (diagnostics is None or diagnostics.can_generate):
            if st.button("确认生成 C/货币资金底稿", type="primary", use_container_width=True):
                _run_cash_case(discovery.materials_dir)
                st.rerun()
    elif erp_manifest:
        _render_erp_mapping(erp_manifest)
        if result is None and not erp_manifest.blocking_issues:
            if st.button("确认字段映射并生成主循环底稿", type="primary", use_container_width=True):
                _run_erp_case(Path(erp_manifest.root_path))
                st.rerun()

    if result:
        if isinstance(result, WorkpaperSuiteResult):
            _render_suite_result(result)
        else:
            _render_cash_result(result)


def _init_state() -> None:
    defaults: dict[str, Any] = {
        "messages": [
            {
                "role": "assistant",
                "content": (
                    "把完整拟真项目路径或 C 底稿资料包路径粘贴给我。"
                    "我会先识别项目结构，再让你确认生成。"
                ),
            }
        ],
        "project_summary": None,
        "discovery": None,
        "diagnostics": None,
        "erp_manifest": None,
        "result": None,
        "ocr_provider": "pdf-text",
        "use_reasoning": False,
        "provider_health": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def _handle_path(raw_path: str) -> None:
    st.session_state.messages.append({"role": "user", "content": raw_path})
    st.session_state.result = None
    st.session_state.project_summary = None
    st.session_state.discovery = None
    st.session_state.erp_manifest = None
    st.session_state.diagnostics = diagnose_client_package(raw_path)

    project = inspect_workpaper_project(raw_path)
    if project.is_ready:
        st.session_state.project_summary = project
        st.session_state.messages.append(
            {
                "role": "assistant",
                "content": "已识别为完整拟真审计项目。请确认下方资料覆盖率和底稿清单，确认后生成全套企业级底稿。",
            }
        )
        return

    discovery = inspect_cash_materials(raw_path)
    st.session_state.discovery = discovery
    if discovery.is_ready:
        st.session_state.messages.append(
            {
                "role": "assistant",
                "content": "已识别为单 C/货币资金资料包。请确认下方识别结果，确认后生成货币资金底稿。",
            }
        )
        return

    erp_manifest = diagnose_erp_export(raw_path)
    if erp_manifest.tables:
        st.session_state.discovery = None
        st.session_state.erp_manifest = erp_manifest
        st.session_state.messages.append(
            {
                "role": "assistant",
                "content": "已识别为 SAP/用友等 ERP 导出资料。请复核字段映射和阻塞项，确认后生成主循环 clean-room 底稿。",
            }
        )
    else:
        missing = "、".join(project.missing_required or discovery.missing_required)
        st.session_state.messages.append(
            {
                "role": "assistant",
                "content": f"这个路径暂时不能运行，缺少：{missing}",
            }
        )


def _run_suite(project_dir: Path) -> None:
    with st.spinner("正在生成全套企业级拟真底稿..."):
        result = run_auto_workpaper_suite(project_dir)
    st.session_state.result = result
    if result.success:
        st.session_state.messages.append({"role": "assistant", "content": f"全套底稿生成完成：{result.output_dir}"})
    else:
        st.session_state.messages.append({"role": "assistant", "content": "全套底稿未完整生成：" + "；".join(result.errors)})


def _run_cash_case(materials_dir: Path) -> None:
    with st.spinner("正在解析资料、生成写入计划并填充底稿..."):
        result = run_auto_cash_case(
            materials_dir,
            ocr_provider=st.session_state.ocr_provider,
            use_reasoning=st.session_state.use_reasoning,
        )
    st.session_state.result = result
    if result.success:
        st.session_state.messages.append({"role": "assistant", "content": f"C 底稿生成完成：{result.output_path}"})
    else:
        st.session_state.messages.append({"role": "assistant", "content": result.error or "运行未完成。"})


def _run_erp_case(export_dir: Path) -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    case_dir = Path("run_out") / f"{safe_filename(export_dir.name)}_erp_case_{stamp}"
    with st.spinner("正在确认字段映射、标准化 ERP 数据并生成主循环底稿..."):
        import_result = import_erp_export(
            export_dir=export_dir,
            case_dir=case_dir,
            provider="auto",
            confirm_mapping=True,
        )
        if not import_result.success:
            st.session_state.messages.append({"role": "assistant", "content": "ERP 导入未完成：" + "；".join(import_result.errors)})
            return
        result = run_erp_workpaper_suite(case_dir)
    st.session_state.result = result
    if result.success:
        st.session_state.messages.append({"role": "assistant", "content": f"ERP 主循环底稿生成完成：{result.output_dir}"})
    else:
        st.session_state.messages.append({"role": "assistant", "content": "ERP 主循环底稿未完整生成：" + "；".join(result.errors)})


def _render_history() -> None:
    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            st.write(message["content"])


def _render_project(project: WorkpaperProjectSummary) -> None:
    st.subheader("完整项目识别")
    cols = st.columns(4)
    cols[0].metric("客户", project.client_name or "UNKNOWN")
    cols[1].metric("期间", project.period_end or "-")
    cols[2].metric("底稿数量", len(project.expected_workbooks))
    cols[3].metric("资料覆盖", f"{sum(project.source_coverage.values())}/{len(project.source_coverage)}")

    st.write(f"项目目录：`{project.project_dir}`")
    st.write(f"C 资料包：`{project.cash_materials_dir}`")

    coverage_rows = [{"资料": key, "状态": "已识别" if value else "缺失"} for key, value in project.source_coverage.items()]
    st.dataframe(coverage_rows, hide_index=True, use_container_width=True)

    st.write("将生成的底稿：")
    st.dataframe([{"底稿": name} for name in project.expected_workbooks], hide_index=True, use_container_width=True)


def _render_discovery(discovery: CashMaterialsSummary) -> None:
    st.subheader("C 资料识别")
    rows = [
        ("资料文件夹", discovery.materials_dir),
        ("试算平衡表 / TB", discovery.trial_balance),
        ("银行日记账 / 序时账", discovery.journal),
        ("银行回函 / 对账单", discovery.bank_statement),
        ("底稿模板", discovery.template),
    ]
    st.dataframe(
        [{"资料": label, "识别结果": str(path) if path else "未识别"} for label, path in rows],
        hide_index=True,
        use_container_width=True,
    )
    if discovery.missing_required:
        st.warning("缺少必要资料：" + "、".join(discovery.missing_required))
    else:
        st.success("必要资料已识别完整。")


def _render_erp_mapping(manifest: ErpMappingManifest) -> None:
    st.subheader("ERP 字段映射诊断")
    cols = st.columns(4)
    cols[0].metric("Provider", manifest.provider)
    cols[1].metric("识别表", len(manifest.tables))
    cols[2].metric("阻塞项", len(manifest.blocking_issues))
    cols[3].metric("确认状态", "已确认" if manifest.confirmed else "待确认")

    if manifest.blocking_issues:
        st.warning("；".join(manifest.blocking_issues))
    else:
        st.success("必要 ERP 表和关键字段已识别；生成前仍需人工确认字段映射。")

    rows = []
    for table in manifest.tables:
        matched = [field for field in table.matched_fields if field.source_header]
        rows.append(
            {
                "资料类型": table.source_type,
                "文件": table.path,
                "sheet": table.sheet_name,
                "行数": table.row_count,
                "置信度": f"{table.confidence:.0%}",
                "匹配字段": ", ".join(f"{field.canonical_field}={field.source_header}" for field in matched[:8]),
                "缺失关键字段": ", ".join(table.missing_required),
            }
        )
    st.dataframe(rows, hide_index=True, use_container_width=True)

    with st.expander("样例行", expanded=False):
        for table in manifest.tables:
            st.write(f"{table.source_type}: `{table.path}`")
            st.dataframe(table.sample_rows, hide_index=True, use_container_width=True)


def _render_diagnostics(diagnostics: ClientPackageDiagnostics) -> None:
    st.subheader("公开 Beta 诊断报告")
    cols = st.columns(4)
    cols[0].metric("模式", diagnostics.mode)
    cols[1].metric("可生成", "是" if diagnostics.can_generate else "否")
    cols[2].metric("置信度", f"{diagnostics.confidence:.0%}")
    cols[3].metric("默认 API", diagnostics.provider_summary.get("default_ocr", "pdf-text/offline"))

    if diagnostics.missing_required:
        st.warning("缺失/阻塞项：" + "、".join(diagnostics.missing_required))
    else:
        st.success("诊断通过：当前资料包可进入生成确认。")

    check_rows = [
        {
            "检查项": check.name,
            "结果": "通过" if check.passed else "未通过",
            "置信度": f"{check.confidence:.0%}",
            "说明": check.detail,
        }
        for check in diagnostics.checks
    ]
    st.dataframe(check_rows, hide_index=True, use_container_width=True)

    st.write("实际使用的 C 资料包：")
    st.dataframe(
        [{"资料": key, "路径": value} for key, value in diagnostics.selected_cash_materials.items()],
        hide_index=True,
        use_container_width=True,
    )

    if diagnostics.cash_candidates:
        with st.expander("候选 C 资料包评分", expanded=False):
            st.dataframe(diagnostics.cash_candidates, hide_index=True, use_container_width=True)

    if diagnostics.agent_used or diagnostics.agent_reason:
        st.info(f"Agent 识别：{diagnostics.agent_reason or '已调用'}")

    with st.expander("识别到的文件和表头", expanded=False):
        file_rows = [
            {
                "类别": file.category,
                "文件": str(file.path),
                "可读": "是" if file.readable else "否",
                "sheet": file.sheet_name,
                "匹配字段": ", ".join(f"{key}={value}" for key, value in file.matched_columns.items()),
                "错误": file.error,
            }
            for file in diagnostics.files[:80]
        ]
        st.dataframe(file_rows, hide_index=True, use_container_width=True)

    with st.expander("API / Provider 状态", expanded=False):
        st.json(diagnostics.provider_summary)


def _render_provider_health(health: ProviderHealth) -> None:
    rows = [
        {
            "provider": health.reasoning.provider,
            "model": health.reasoning.model,
            "configured": health.reasoning.configured,
            "ok": health.reasoning.ok,
            "purpose": health.reasoning.purpose,
            "error": health.reasoning.error,
        },
        {
            "provider": health.vision.provider,
            "model": health.vision.model,
            "configured": health.vision.configured,
            "ok": health.vision.ok,
            "purpose": health.vision.purpose,
            "error": health.vision.error,
        },
    ]
    st.dataframe(rows, hide_index=True, use_container_width=True)


def _render_suite_result(result: WorkpaperSuiteResult) -> None:
    if not result.success:
        st.error("；".join(result.errors) or "全套底稿未完整生成。")

    st.subheader("全套底稿生成结果")
    cols = st.columns(4)
    cols[0].metric("客户", result.client_name or "UNKNOWN")
    cols[1].metric("期间", result.period_end or "-")
    cols[2].metric("底稿", len(result.workbooks))
    cols[3].metric("复核提示", result.findings_count)

    st.write(f"输出目录：`{result.output_dir}`")
    download_cols = st.columns(4)
    if result.zip_path:
        _download_button(download_cols[0], "整套 ZIP", result.zip_path)
    if result.manifest_path:
        _download_button(download_cols[1], "suite_manifest", result.manifest_path)
    for idx, (code, path) in enumerate(result.workbooks.items(), start=2):
        _download_button(download_cols[idx % len(download_cols)], code, path)


def _render_cash_result(result: AutoCashRunResult) -> None:
    if not result.success:
        st.error(result.error or "运行未完成。")
        return

    st.subheader("C 底稿生成结果")
    metrics = st.columns(4)
    metrics[0].metric("客户", result.client_name or "UNKNOWN")
    metrics[1].metric("期间", result.period_end or "-")
    metrics[2].metric("审计发现", result.findings_count)
    metrics[3].metric("写入单元格", result.write_commands_count)

    st.write(f"输出目录：`{result.case_dir}`")
    st.write(f"证据链记录：`{result.provenance_count}` 条")

    download_cols = st.columns(5)
    for idx, (name, path) in enumerate(result.artifacts.items()):
        _download_button(download_cols[idx % len(download_cols)], name, path)

    with st.expander("运行轨迹", expanded=False):
        for event in result.trace_events:
            st.write(f"[{event.stage}] {event.message} {event.detail}".strip())


def _download_button(container, label: str, path: Path) -> None:
    if not path.exists():
        container.caption(f"{label} 未生成")
        return
    container.download_button(
        label=f"下载 {label}",
        data=path.read_bytes(),
        file_name=path.name,
        mime=_mime_for(path),
        use_container_width=True,
    )


def _mime_for(path: Path) -> str:
    if path.suffix.lower() == ".xlsx":
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    if path.suffix.lower() == ".json":
        return "application/json"
    if path.suffix.lower() == ".zip":
        return "application/zip"
    return "application/octet-stream"


if __name__ == "__main__":
    main()
