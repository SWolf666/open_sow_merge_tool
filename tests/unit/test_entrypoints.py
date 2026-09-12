from __future__ import annotations

import sow_merge_tool as public
from sow_merge_tool.launch_center import StartCenterResult
from sow_merge_tool.path_selection import compare_selection_from_paths


def test_public_entrypoint_and_version() -> None:
    assert public.APP_NAME == "sow_merge_tool"
    assert public.APP_VERSION.endswith("update96")
    assert callable(public.run_entrypoint)


def test_compare_prefill_is_normalized_without_opening_or_writing(tmp_path) -> None:
    first = tmp_path / "中文 路径" / "A.xlsx"
    second = tmp_path / "目标" / "A.xlsx"
    first.parent.mkdir()
    second.parent.mkdir()
    first.write_bytes(b"a"); second.write_bytes(b"b")
    paths = compare_selection_from_paths([first, first, second])
    assert paths == (str(first), str(second))
    assert first.read_bytes() == b"a" and second.read_bytes() == b"b"


def test_install_scripts_have_distinct_compare_and_branch_context_keys() -> None:
    from pathlib import Path

    root = Path(__file__).resolve().parents[2]
    install = (root / "安装.bat").read_text(encoding="utf-8")
    uninstall = (root / "卸载.bat").read_text(encoding="utf-8")
    for suffix in (".xlsx", ".xlsm"):
        assert f"SystemFileAssociations\\{suffix}\\shell\\SowExcelCompareMerge" in install
        assert f"SystemFileAssociations\\{suffix}\\shell\\SowExcelCompareMerge" in uninstall
    assert "Excel 文件比较/合并" in install
    assert " --compare " in install
    assert " --branch-submit " in install
    assert "Position -PropertyType String -Value Top" in install


def test_compare_switch_prefills_branch_without_automatic_scan(monkeypatch, tmp_path) -> None:
    from sow_merge_tool import launch_center as center

    captured = []
    monkeypatch.setattr(
        center,
        "launch_start_center",
        lambda paths, **_kwargs: StartCenterResult("branch", tuple(paths)),
    )
    from sow_merge_tool import branch_submit

    monkeypatch.setattr(branch_submit, "launch_ui", lambda paths=None: captured.append(tuple(paths or ())))
    old_argv = __import__("sys").argv
    __import__("sys").argv = ["sow_merge_tool", "--compare", str(tmp_path / "中文.xlsx")]
    try:
        public.main()
    finally:
        __import__("sys").argv = old_argv
    assert captured == [(str(tmp_path / "中文.xlsx"),)]


def test_no_argument_launch_uses_start_center(monkeypatch) -> None:
    import sys

    from sow_merge_tool import branch_submit
    from sow_merge_tool import launch_center as center

    captured = []
    monkeypatch.setattr(center, "launch_start_center", lambda paths: StartCenterResult("branch", tuple(paths)))
    monkeypatch.setattr(branch_submit, "launch_ui", lambda paths=None: captured.append(tuple(paths or ())))
    old_argv = sys.argv
    sys.argv = ["sow_merge_tool"]
    try:
        public.main()
    finally:
        sys.argv = old_argv
    assert captured == [()]
