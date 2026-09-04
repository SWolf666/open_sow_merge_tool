from __future__ import annotations

from openpyxl import Workbook

from sow_merge_tool.launch_center import (
    COMPARISON_STATES,
    ComparisonLaunch,
    ComparisonListModel,
    ComparisonSessionManager,
)
from sow_merge_tool.legacy_core import RolePresentation
from sow_merge_tool.path_selection import PairStatus, RelativeMapping


class _FakeProcess:
    def __init__(self, returncode=None):
        self.returncode = returncode
        self.terminated = False

    def poll(self):
        return self.returncode

    def terminate(self):
        self.terminated = True
        self.returncode = -15


def test_comparison_command_is_direct_and_role_explicit(monkeypatch, tmp_path):
    monkeypatch.setattr("sow_merge_tool.launch_center.sys.frozen", False, raising=False)
    left = tmp_path / "Source 中文.xlsx"
    right = tmp_path / "Target 中文.xlsx"
    launch = ComparisonLaunch.for_paths(str(left), str(right))
    assert "--compare" not in launch.command
    assert launch.command[-3:] == ("--source-target", str(left), str(right))


def test_session_manager_rejects_duplicate_and_reports_close():
    processes = []

    def popen(command, **_kwargs):
        processes.append(command)
        process = _FakeProcess()
        process.command = command
        fake_processes.append(process)
        return process

    fake_processes = []
    states = []
    manager = ComparisonSessionManager(popen_factory=popen)
    first = manager.start("left.xlsx", "right.xlsx", callback=lambda session: states.append(session.state))
    assert first is not None
    assert states[:2] == ["启动中", "比较中"]
    assert manager.start("other-left.xlsx", "other-right.xlsx") is None
    assert len(processes) == 1
    fake_processes[0].returncode = 0
    manager.poll_now()
    assert first.state == "已关闭"
    assert manager.active_session is None
    assert states[-1] == "已关闭"


def test_session_manager_start_failure_is_explicit_and_close_keeps_child():
    def failing(_command, **_kwargs):
        raise OSError("spawn failed")

    manager = ComparisonSessionManager(popen_factory=failing)
    failed = manager.start("left.xlsx", "right.xlsx")
    assert failed is not None and failed.state == "启动失败"
    assert failed.error

    process = _FakeProcess()
    manager = ComparisonSessionManager(popen_factory=lambda _command, **_kwargs: process)
    running = manager.start("left.xlsx", "right.xlsx")
    manager.close()
    assert running is not None and running.state == "比较中"
    assert not process.terminated
    assert set(COMPARISON_STATES) == {"未打开", "启动中", "比较中", "已关闭", "启动失败"}


def test_source_target_presentation_keeps_source_read_only_target_writable():
    app = type("App", (), {"merge_mode": False, "has_base": False, "role_mode": "source-target"})()
    presentation = RolePresentation.for_app(app)
    assert (presentation.left, presentation.right) == ("Source", "Target")
    assert presentation.action_label("A2B") == "应用 Source → Target 行"


def test_headless_list_model_covers_open_double_click_and_parent_close(tmp_path):
    left = tmp_path / "Source.xlsx"
    right = tmp_path / "Target.xlsx"
    for path, value in ((left, "left"), (right, "right")):
        workbook = Workbook()
        workbook.active["A1"] = value
        workbook.save(path)
        workbook.close()
    process = _FakeProcess()
    manager = ComparisonSessionManager(popen_factory=lambda _command, **_kwargs: process)
    mapping = RelativeMapping("Data.xlsx", str(left), str(right), PairStatus.MATCHED)
    model = ComparisonListModel((mapping,), manager=manager)
    session = model.double_click("mapping-0")
    assert session.state == "比较中"
    assert model.row_state("mapping-0") == "比较中"
    assert "--compare" not in session.launch.command
    try:
        model.double_click("mapping-0")
    except RuntimeError as exc:
        assert "已有" in str(exc)
    else:
        raise AssertionError("headless list must reject a second active child")
    model.close()
    assert not process.terminated
    process.returncode = 0
    model.manager.poll_now()
    assert model.row_state("mapping-0") == "已关闭"
