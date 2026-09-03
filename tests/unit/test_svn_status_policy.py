from __future__ import annotations

import os
import tkinter as tk
from tkinter import ttk

import pytest

from sow_merge_tool import branch_submit as bs
from sow_merge_tool import svn_status_provider as sp

ALL_NATIVE_STATUSES = {
    "none",
    "unversioned",
    "normal",
    "added",
    "missing",
    "deleted",
    "replaced",
    "modified",
    "merged",
    "conflicted",
    "ignored",
    "obstructed",
    "external",
    "incomplete",
}


@pytest.mark.parametrize(
    ("status", "visible", "selectable", "checked", "policy"),
    (
        ("none", True, False, False, "不进入批次"),
        ("unversioned", True, True, False, "勾选后在目标新增"),
        ("normal", True, False, False, "不进入批次"),
        ("added", True, True, True, "在目标新增文件"),
        ("missing", True, True, True, "仅源分支；目标不变"),
        ("deleted", True, True, True, "同步删除到目标"),
        ("replaced", True, False, False, "不进入批次"),
        ("modified", True, True, True, "同步修改到目标"),
        ("merged", True, False, False, "不进入批次"),
        ("conflicted", True, False, False, "不进入批次"),
        ("ignored", False, False, False, "不进入批次"),
        ("obstructed", True, False, False, "不进入批次"),
        ("external", False, False, False, "不进入批次"),
        ("incomplete", True, False, False, "不进入批次"),
    ),
)
def test_every_native_status_has_an_explicit_policy(
    tmp_path, monkeypatch, status, visible, selectable, checked, policy
):
    branch = tmp_path / "develop"
    branch.mkdir()
    path = branch / f"{status}.xlsx"
    if status not in {"missing", "deleted", "none"}:
        path.write_bytes(b"fixture")
    versioned = status not in {"none", "unversioned", "ignored", "external"}
    record = sp.SvnStatusRecord(
        path=str(path),
        node_kind="file",
        node_status=status,
        text_status=status,
        prop_status="normal",
        versioned=versioned,
        conflicted=status == "conflicted",
        file_external=status == "external",
    )
    monkeypatch.setattr(bs, "scan_status", lambda *_args, **_kwargs: [record])
    items = bs.scan_changes(str(tmp_path), "develop", str(branch))
    assert bool(items) is visible
    if visible:
        item = items[0]
        assert item.selectable is selectable
        assert item.checked is checked
        assert bs._multi_branch_policy_text(item.node_status, item.reason) == policy
    else:
        assert policy == "不进入批次"


def test_native_status_enum_is_fully_accounted_for():
    assert set(sp.STATUS_NAMES.values()) == ALL_NATIVE_STATUSES


def test_workbench_root_enter_excludes_message_and_search_inputs():
    app = object.__new__(bs.BranchSubmitWorkbench)
    app.tk = tk
    app.ttk = ttk
    app.message = object()
    app._keyboard_search_entry = object()
    app.source_box = object()
    calls = []
    class Button:
        def instate(self, _state):
            return True
    app.submit_button = Button()
    app.preflight_button = Button()
    app._submit = lambda: calls.append("submit")
    app._preflight = lambda: calls.append("preflight")
    assert app._root_enter(type("Event", (), {"widget": app.message})()) is None
    assert app._root_enter(type("Event", (), {"widget": app._keyboard_search_entry})()) == "break"
    assert app._root_enter(type("Event", (), {"widget": app.source_box})()) == "break"
    assert calls == []
    assert app._root_enter(type("Event", (), {"widget": object()})()) == "break"
    assert calls == ["submit"]


@pytest.mark.parametrize(
    ("overrides", "visible", "selectable", "checked", "reason_fragment"),
    (
        ({"switched": True}, True, False, False, "switched"),
        ({"file_external": True}, False, False, False, ""),
        ({"wc_locked": True}, True, False, False, "Cleanup"),
        ({"prop_status": "modified"}, True, False, False, "属性修改"),
        ({"prop_status": "conflicted"}, True, False, False, "冲突"),
        ({"changelist": "ignore-on-commit"}, True, True, False, ""),
        ({"lock_owner": "tester"}, True, True, True, ""),
    ),
)
def test_status_flags_and_changelists_are_explicitly_handled(
    tmp_path, monkeypatch, overrides, visible, selectable, checked, reason_fragment
):
    branch = tmp_path / "develop"
    branch.mkdir()
    path = branch / "Flagged.xlsx"
    path.write_bytes(b"fixture")
    values = {
        "path": str(path),
        "node_kind": "file",
        "node_status": "modified",
        "text_status": "modified",
        "prop_status": "normal",
        "versioned": True,
    }
    values.update(overrides)
    record = sp.SvnStatusRecord(**values)
    monkeypatch.setattr(bs, "scan_status", lambda *_args, **_kwargs: [record])
    items = bs.scan_changes(str(tmp_path), "develop", str(branch))
    assert bool(items) is visible
    if visible:
        item = items[0]
        assert item.selectable is selectable
        assert item.checked is checked
        if reason_fragment:
            assert reason_fragment in item.reason


def test_added_node_property_state_is_part_of_the_add(tmp_path, monkeypatch):
    branch = tmp_path / "develop"
    branch.mkdir()
    path = branch / "Added.xlsx"
    path.write_bytes(b"fixture")
    record = sp.SvnStatusRecord(
        path=str(path),
        node_kind="file",
        node_status="added",
        text_status="added",
        prop_status="modified",
        versioned=True,
    )
    monkeypatch.setattr(bs, "scan_status", lambda *_args, **_kwargs: [record])
    item = bs.scan_changes(str(tmp_path), "develop", str(branch))[0]
    assert item.selectable and item.checked and not item.reason


@pytest.mark.parametrize("kind", ("dir", "symlink", "unknown", "none"))
def test_non_file_nodes_never_enter_a_batch(tmp_path, monkeypatch, kind):
    branch = tmp_path / "develop"
    branch.mkdir()
    path = branch / "NotARegularFile.xlsx"
    record = sp.SvnStatusRecord(
        path=str(path),
        node_kind=kind,
        node_status="modified",
        text_status="modified",
        prop_status="normal",
        versioned=True,
    )
    monkeypatch.setattr(bs, "scan_status", lambda *_args, **_kwargs: [record])
    item = bs.scan_changes(str(tmp_path), "develop", str(branch))[0]
    assert not item.selectable
    assert not item.checked
    assert "普通文件" in item.reason


def test_move_metadata_is_blocked_before_preflight(tmp_path):
    item = bs.SvnChangeItem(
        path=os.fspath(tmp_path / "develop" / "Moved.xlsx"),
        relative_path="Moved.xlsx",
        extension=".xlsx",
        node_kind="file",
        node_status="added",
        text_status="added",
        prop_status="normal",
        versioned=True,
        moved_from="Old.xlsx",
        checked=True,
        selectable=True,
    )
    batch = bs.BranchSubmitBatch(
        batch_id="move-policy",
        wc_root=os.fspath(tmp_path),
        source_branch="develop",
        target_branches=["release"],
        files=[],
        message="",
    )
    engine = bs.BranchSubmitEngine(os.fspath(tmp_path), allowed_branches=("develop", "release"))
    with pytest.raises(RuntimeError, match="Repair Move"):
        engine._source_snapshot(batch, item)
