import importlib.util
import subprocess
from pathlib import Path

_RUFF_CHANGED_SPEC = importlib.util.spec_from_file_location(
    "ruff_changed_gate",
    Path(__file__).parents[2] / "tools" / "ruff_changed.py",
)
assert _RUFF_CHANGED_SPEC is not None and _RUFF_CHANGED_SPEC.loader is not None
_RUFF_CHANGED = importlib.util.module_from_spec(_RUFF_CHANGED_SPEC)
_RUFF_CHANGED_SPEC.loader.exec_module(_RUFF_CHANGED)
diagnostics_on_changed_lines = _RUFF_CHANGED.diagnostics_on_changed_lines
changed_python_lines = _RUFF_CHANGED.changed_python_lines
map_lines_through_diff = _RUFF_CHANGED.map_lines_through_diff
merge_line_maps = _RUFF_CHANGED.merge_line_maps
parse_added_lines = _RUFF_CHANGED.parse_added_lines
run_gate = _RUFF_CHANGED.run_gate


def test_parse_added_lines_keeps_each_hunk_in_its_file():
    diff = """diff --git a/src/one.py b/src/one.py
--- a/src/one.py
+++ b/src/one.py
@@ -1,0 +2,2 @@
+new_line
+new_line_2
diff --git a/tests/two.py b/tests/two.py
--- a/tests/two.py
+++ b/tests/two.py
@@ -4 +4,3 @@
-old
+replacement
+added
 context
"""
    assert parse_added_lines(diff) == {
        "src/one.py": {2, 3},
        "tests/two.py": {4, 5},
    }


def test_merge_line_maps_deduplicates_committed_and_worktree_lines():
    assert merge_line_maps(
        {"src/a.py": {2, 3}},
        {"src/a.py": {3, 4}, "src/b.py": {1}},
    ) == {"src/a.py": {2, 3, 4}, "src/b.py": {1}}


def test_diagnostics_are_selected_only_when_range_hits_changed_line():
    diagnostics = [
        {
            "code": "E1",
            "filename": "src/a.py",
            "location": {"row": 4},
            "end_location": {"row": 4},
        },
        {
            "code": "E2",
            "filename": "src/a.py",
            "location": {"row": 8},
            "end_location": {"row": 9},
        },
        {
            "code": "E3",
            "filename": "src/other.py",
            "location": {"row": 1},
            "end_location": {"row": 1},
        },
    ]
    changed = {"src/a.py": {4, 6}, "src/other.py": {2}}
    assert [item["code"] for item in diagnostics_on_changed_lines(diagnostics, changed)] == ["E1"]


def _git(repo: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", *args],
        cwd=repo,
        check=True,
        text=True,
        encoding="utf-8",
        capture_output=True,
    )
    return result.stdout


def _git_repo(tmp_path: Path, name: str, content: str) -> tuple[Path, Path]:
    repo = tmp_path / name
    repo.mkdir()
    _git(repo, "init", "-q")
    _git(repo, "config", "user.email", "ruff-gate@example.invalid")
    _git(repo, "config", "user.name", "ruff-gate")
    source = repo / "sample.py"
    source.write_text(content, encoding="utf-8")
    _git(repo, "add", "sample.py")
    _git(repo, "commit", "-qm", "baseline")
    return repo, source


def test_real_git_changed_lines_cover_commit_only_and_dirty_only(tmp_path, monkeypatch):
    commit_repo, commit_source = _git_repo(tmp_path, "commit-only", "one = 1\n")
    commit_source.write_text("one = 1\ncommitted = 2\n", encoding="utf-8")
    _git(commit_repo, "add", "sample.py")
    _git(commit_repo, "commit", "-qm", "committed change")
    monkeypatch.chdir(commit_repo)
    assert changed_python_lines("HEAD^")["sample.py"] == {2}

    dirty_repo, dirty_source = _git_repo(tmp_path, "dirty-only", "one = 1\n")
    dirty_source.write_text("one = 1\ndirty = 2\n", encoding="utf-8")
    monkeypatch.chdir(dirty_repo)
    assert changed_python_lines("HEAD")["sample.py"] == {2}


def test_real_git_maps_commit_lines_after_a_dirty_prefix_insert(tmp_path, monkeypatch):
    repo, source = _git_repo(tmp_path, "commit-and-dirty", "one = 1\ntwo = 2\n")
    source.write_text("one = 1\ntwo = 2\ncommitted = 3\n", encoding="utf-8")
    _git(repo, "add", "sample.py")
    _git(repo, "commit", "-qm", "committed change")
    source.write_text(
        "dirty_prefix = 0\none = 1\ntwo = 2\ncommitted = 3\n",
        encoding="utf-8",
    )
    monkeypatch.chdir(repo)
    assert changed_python_lines("HEAD^")["sample.py"] == {1, 4}


def test_real_git_untracked_python_is_checked_as_all_current_lines(tmp_path, monkeypatch):
    repo, _source = _git_repo(tmp_path, "untracked", "one = 1\n")
    (repo / "new.py").write_text("first = 1\nsecond = 2\n", encoding="utf-8")
    monkeypatch.chdir(repo)
    assert changed_python_lines("HEAD")["new.py"] == {1, 2}


def test_real_git_deletion_only_has_no_current_added_lines(tmp_path, monkeypatch):
    repo, source = _git_repo(tmp_path, "deletion-only", "one = 1\n")
    source.unlink()
    _git(repo, "add", "-u")
    monkeypatch.chdir(repo)
    assert "sample.py" not in changed_python_lines("HEAD")


def test_deleted_only_python_has_no_added_line_diagnostics():
    diagnostic = {
        "code": "E1",
        "filename": "sample.py",
        "location": {"row": 1},
        "end_location": {"row": 1},
    }
    assert diagnostics_on_changed_lines([diagnostic], {"sample.py": set()}) == []


def test_run_gate_fails_closed_for_ruff_nonzero_and_bad_json(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(_RUFF_CHANGED, "changed_python_lines", lambda _base: {"sample.py": {1}})
    monkeypatch.setattr(_RUFF_CHANGED, "_run_git", lambda *_args: "deadbeef")

    class Result:
        returncode = 1
        stdout = "[]"

    monkeypatch.setattr(_RUFF_CHANGED.subprocess, "run", lambda *_args, **_kwargs: Result())
    exit_code, payload = run_gate("HEAD")
    assert exit_code != 0 and "非零退出码" in payload["error"]

    Result.returncode = 0
    Result.stdout = "not-json"
    exit_code, payload = run_gate("HEAD")
    assert exit_code != 0 and "JSON 解析失败" in payload["error"]


def test_run_gate_fails_closed_for_ruff_process_error(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(_RUFF_CHANGED, "changed_python_lines", lambda _base: {"sample.py": {1}})
    monkeypatch.setattr(_RUFF_CHANGED, "_run_git", lambda *_args: "deadbeef")

    def fail_process(*_args, **_kwargs):
        raise OSError("ruff missing")

    monkeypatch.setattr(_RUFF_CHANGED.subprocess, "run", fail_process)
    exit_code, payload = run_gate("HEAD")
    assert exit_code != 0 and "Ruff 子进程失败" in payload["error"]


def test_map_lines_through_diff_handles_insert_delete_and_replacement():
    diff = """diff --git a/sample.py b/sample.py
--- a/sample.py
+++ b/sample.py
@@ -0,0 +1 @@
+prefix = 0
@@ -3,1 +4,0 @@
-old = 1
@@ -5,1 +5,2 @@
-replace = 1
+replace = 2
+replace_extra = 3
"""
    assert map_lines_through_diff({"sample.py": {2, 3, 5}}, diff) == {
        "sample.py": {3, 5}
    }
