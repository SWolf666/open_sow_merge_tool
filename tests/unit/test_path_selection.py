from __future__ import annotations

import os

import pytest

from sow_merge_tool.path_selection import (
    DirectoryStatus,
    PairStatus,
    PathSelectionError,
    map_relative_files,
    recheck_selection,
    snapshot_selection,
    validate_directory,
    validate_file_pair,
)


def _file(path, payload: bytes = b"xlsx-fixture"):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(payload)
    return path


def test_explicit_file_pair_accepts_different_paths_and_rejects_same_file(tmp_path):
    left = _file(tmp_path / "源" / "配置.xlsx", b"left")
    right = _file(tmp_path / "目标 with spaces" / "配置.xlsx", b"right")
    pair = validate_file_pair(left, right)
    assert pair.status is PairStatus.SAME_NAME_DIFFERENT_PATH
    assert pair.ready

    same = validate_file_pair(left, left)
    assert same.status is PairStatus.SAME_FILE
    assert not same.ready


def test_file_pair_reports_missing_and_incompatible_types(tmp_path):
    left = _file(tmp_path / "left.xlsx")
    assert validate_file_pair(left, tmp_path / "missing.xlsx").status is PairStatus.MISSING_RIGHT
    right = _file(tmp_path / "right.xlsm")
    pair = validate_file_pair(left, right)
    assert pair.status is PairStatus.TYPE_INCOMPATIBLE
    assert "类型" in pair.reason


def test_directory_mapping_uses_relative_path_not_basename(tmp_path):
    left = tmp_path / "left"
    right = tmp_path / "right"
    _file(left / "a" / "配置.xlsx", b"a-left")
    _file(left / "b" / "配置.xlsx", b"b-left")
    _file(right / "a" / "配置.xlsx", b"a-right")
    _file(right / "other" / "配置.xlsx", b"other-right")
    mappings = map_relative_files(left, right)
    statuses = {item.relative_path: item.status for item in mappings}
    assert statuses["a/配置.xlsx"] is PairStatus.SAME_NAME_DIFFERENT_PATH
    assert statuses["b/配置.xlsx"] is PairStatus.MISSING_RIGHT
    assert statuses["other/配置.xlsx"] is PairStatus.MISSING_LEFT
    assert all("配置.xlsx" in item.relative_path for item in mappings)


def test_directory_validation_allows_normal_directory_and_rejects_reparse(tmp_path, monkeypatch):
    normal = validate_directory(tmp_path / "普通目录")
    assert normal.status is DirectoryStatus.MISSING
    (tmp_path / "普通目录").mkdir()
    assert validate_directory(tmp_path / "普通目录").ready
    monkeypatch.setattr("sow_merge_tool.path_selection._is_reparse", lambda _path: True)
    result = validate_directory(tmp_path / "普通目录")
    assert result.status is DirectoryStatus.REPARSE
    assert not result.ready


def test_hash_recheck_is_read_only_and_fails_closed_on_drift(tmp_path):
    left = _file(tmp_path / "left.xlsx", b"one")
    right = _file(tmp_path / "right.xlsx", b"two")
    snapshot = snapshot_selection(left, right)
    assert snapshot.left.sha256 and snapshot.right.sha256
    _file(left, b"changed")
    with pytest.raises(PathSelectionError, match="发生变化"):
        recheck_selection(snapshot)
    assert os.path.isfile(left)
