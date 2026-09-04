"""Conservative path-risk hints for multi-branch configuration review.

These checks are intentionally advisory.  They help a reviewer notice shared
tables, rollback/full-replacement files, and temporary configuration, but they
never turn a path name into a delete, move, or overwrite instruction.
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass


@dataclass(frozen=True)
class PathRisk:
    code: str
    label: str
    reason: str


_COMMON_SEGMENTS = frozenset({"common", "shared", "public", "common_table", "public_table"})
_COMMON_FILE_HINTS = frozenset({"errorcode", "error_code", "publictable", "sharedtable"})
_ROLLBACK_TOKENS = frozenset({"rollback", "roll_back", "revert", "restore", "回滚", "还原"})
_REPLACE_TOKENS = frozenset({"replace", "replacement", "full_replace", "fullreplacement", "全量替换"})
_TEMP_TOKENS = frozenset({"temp", "tmp", "temporary", "test", "临时", "测试"})
_TOKEN_RE = re.compile(r"[A-Za-z0-9]+|[\u4e00-\u9fff]+")


def _tokens(value: str) -> set[str]:
    normalized = str(value or "").replace("\\", "/").strip("/").lower()
    pieces = set(_TOKEN_RE.findall(normalized))
    for part in normalized.split("/"):
        pieces.update(
            piece
            for piece in part.rsplit(".", 1)[0].replace("-", "_").split("_")
            if piece
        )
    for special in _ROLLBACK_TOKENS | _REPLACE_TOKENS | _TEMP_TOKENS:
        if special in normalized:
            pieces.add(special)
    pieces.update({normalized})
    return pieces


def _is_common_path(value: str) -> bool:
    normalized = str(value or "").replace("\\", "/").strip("/").lower()
    parts = [part for part in normalized.split("/") if part]
    stem = parts[-1].rsplit(".", 1)[0] if parts else ""
    return any(part in _COMMON_SEGMENTS for part in parts) or stem in _COMMON_FILE_HINTS


def assess_path_risks(
    relative_path: str,
    *,
    operation: str = "modify",
    source_branch: str = "",
    target_branch: str = "",
) -> tuple[PathRisk, ...]:
    """Return deterministic advisory risks for one selected path.

    ``source_branch`` and ``target_branch`` are context for the displayed
    explanation only.  The returned risks do not authorize any filesystem or
    repository operation.
    """
    normalized = str(relative_path or "").replace("\\", "/").strip("/")
    tokens = _tokens(normalized)
    risks: list[PathRisk] = []
    if _is_common_path(normalized):
        scope = "公共/共享配置路径"
        if source_branch and target_branch:
            scope = f"{source_branch} → {target_branch} 的公共/共享配置路径"
        risks.append(PathRisk("common-path", "公共路径", f"检测到{scope}，默认不迁移路径、不删除目标独立内容"))
    if tokens & _ROLLBACK_TOKENS:
        risks.append(PathRisk("rollback", "回滚风险", "路径名称包含回滚语义，仅提示人工核对，不自动反向应用"))
    if tokens & _REPLACE_TOKENS:
        risks.append(PathRisk("full-replacement", "全量替换风险", "路径名称包含全量替换语义，仅提示人工核对，不整文件覆盖目标"))
    if tokens & _TEMP_TOKENS or normalized.startswith("~$"):
        risks.append(PathRisk("temporary-config", "临时配置", "路径名称可能是临时配置，仅提示并保留现有 SVN/文件安全门禁"))
    if operation == "delete" and _is_common_path(normalized):
        risks.append(PathRisk("common-delete", "公共删除风险", "公共路径的删除只保留风险提示，仍需 SVN 明确 deleted 状态和现有安全规则"))
    return tuple(risks)


def risk_badges(risks: tuple[PathRisk, ...] | list[PathRisk]) -> list[str]:
    """Return UI-safe labels without exposing the risk object internals."""
    return list(dict.fromkeys(risk.label for risk in risks))


def assess_repository_log_risks(entries: Iterable[object]) -> tuple[PathRisk, ...]:
    """Turn structured log evidence into advisory badges only.

    Changed-path actions and copyfrom metadata are used to highlight a
    branch-local to common/shared migration.  Commit messages are inspected
    only for explicit rollback/full-replacement/temporary wording; a normal
    ``merge`` message produces no risk and never blocks a batch.
    """
    risks: list[PathRisk] = []
    branch_delete = False
    common_add = False
    for entry in entries:
        if hasattr(entry, "message"):
            message = str(getattr(entry, "message", "") or "")
            changed_paths = getattr(entry, "changed_paths", ()) or ()
        elif isinstance(entry, dict):
            message = str(entry.get("message", "") or "")
            changed_paths = entry.get("changed_paths", ()) or ()
        else:
            continue
        message_tokens = _tokens(message)
        if message_tokens & _ROLLBACK_TOKENS:
            risks.append(PathRisk("rollback-log", "日志含回滚提示", "提交日志含回滚/还原语义，建议人工核对来源和目标路径"))
        if message_tokens & _REPLACE_TOKENS:
            risks.append(PathRisk("replacement-log", "日志含全量替换提示", "提交日志含全量替换语义，建议人工核对，工具不整文件覆盖目标"))
        if message_tokens & _TEMP_TOKENS:
            risks.append(PathRisk("temporary-log", "日志含临时配置提示", "提交日志含临时/测试语义，建议确认是否属于正式配置"))
        for raw_path in changed_paths:
            if hasattr(raw_path, "path"):
                path = str(getattr(raw_path, "path", "") or "")
                action = str(getattr(raw_path, "action", "") or "").upper()
                copyfrom = str(getattr(raw_path, "copyfrom_path", "") or "")
            elif isinstance(raw_path, dict):
                path = str(raw_path.get("path", "") or "")
                action = str(raw_path.get("action", "") or "").upper()
                copyfrom = str(raw_path.get("copyfrom_path", "") or "")
            else:
                continue
            path_is_common = _is_common_path(path)
            copy_is_common = _is_common_path(copyfrom)
            branch_delete = branch_delete or (action in {"D", "R"} and not path_is_common)
            common_add = common_add or (action in {"A", "R"} and path_is_common)
            if copyfrom and path_is_common != copy_is_common:
                risks.append(PathRisk("common-copy", "公共路径复制来源风险", "changed path 的 copyfrom 跨越分支与公共/共享区域，建议人工核对路径历史"))
    if branch_delete and common_add:
        risks.append(PathRisk("common-migration", "分支到公共迁移风险", "日志同时出现分支删除与公共新增；默认不删除目标、不自动迁移路径"))
    return tuple(dict.fromkeys(risks))


__all__ = [
    "PathRisk",
    "assess_path_risks",
    "assess_repository_log_risks",
    "risk_badges",
]
