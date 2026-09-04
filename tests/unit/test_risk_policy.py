from __future__ import annotations

from sow_merge_tool.branch_submit import BranchSubmitBatch, BranchSubmitEngine, SvnChangeItem
from sow_merge_tool.risk_policy import (
    assess_path_risks,
    assess_repository_log_risks,
    risk_badges,
)


def test_common_table_and_file_operation_risks_are_advisory():
    risks = assess_path_risks("language/ErrorCode.xlsx", operation="modify")
    assert "公共路径" in risk_badges(risks)
    assert all(risk.code != "common-delete" for risk in risks)

    delete_risks = assess_path_risks("common/ErrorCode.xlsx", operation="delete")
    assert {"公共路径", "公共删除风险"}.issubset(risk_badges(delete_risks))

    badges = risk_badges(
        assess_path_risks("tmp/full-replacement-rollback.xlsx", operation="modify")
    )
    assert {"回滚风险", "全量替换风险", "临时配置"}.issubset(badges)
    assert "公共路径" in risk_badges(assess_path_risks("language/ErrorCode.xlsx"))


def test_structured_log_risks_flag_cross_area_migration_without_blocking_normal_merge():
    risks = assess_repository_log_risks([
        {
            "message": "rollback full-replacement temp validation",
            "changed_paths": [
                {"path": "/branches/develop/language/Old.xlsx", "action": "D"},
                {
                    "path": "/common/ErrorCode.xlsx",
                    "action": "A",
                    "copyfrom_path": "/branches/develop/language/Old.xlsx",
                },
            ],
        }
    ])
    badges = risk_badges(risks)
    assert {
        "分支到公共迁移风险",
        "日志含回滚提示",
        "日志含全量替换提示",
        "日志含临时配置提示",
        "公共路径复制来源风险",
    }.issubset(badges)

    assert assess_repository_log_risks([
        {"message": "normal merge", "changed_paths": [{"path": "/branches/release/Data.xlsx", "action": "M"}]}
    ]) == ()
    unicode_badges = risk_badges(assess_repository_log_risks([
        {
            "message": "release合并master 全量替换，后续需要回滚，临时配置",
            "changed_paths": [],
        }
    ]))
    assert {"日志含回滚提示", "日志含全量替换提示", "日志含临时配置提示"}.issubset(unicode_badges)


def test_source_snapshot_carries_risk_badges_without_changing_operation(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "local"))
    source = tmp_path / "develop" / "language" / "ErrorCode.xlsx"
    source.parent.mkdir(parents=True)
    source.write_bytes(b"fixture")
    batch = BranchSubmitBatch(
        batch_id="risk-snapshot",
        wc_root=str(tmp_path),
        source_branch="develop",
        target_branches=["release"],
        files=[],
        message="risk",
    )
    engine = BranchSubmitEngine(str(tmp_path))
    item = SvnChangeItem(
        path=str(source),
        relative_path="language/ErrorCode.xlsx",
        extension=".xlsx",
        node_kind="file",
        node_status="unversioned",
        text_status="none",
        prop_status="normal",
        versioned=False,
        selectable=True,
        checked=True,
    )
    plan = engine._source_snapshot(batch, item)
    assert plan.operation == "add"
    assert plan.risk_badges == ["公共路径"]
    assert plan.risk_reasons
