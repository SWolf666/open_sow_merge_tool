"""Ruff gate for lines changed by a commit and/or the current worktree.

The normal project gate intentionally keeps its historical scope.  This
companion gate makes production changes auditable without requiring the whole
legacy module to be clean first: it runs Ruff on the affected Python files and
fails only when a diagnostic intersects an added line.

Use a resolved commit hash for release checks, for example::

    python tools/ruff_changed.py --base "$(git rev-parse HEAD^)"

The default ``HEAD`` base checks the current worktree.  A release caller
should pass the resolved parent hash so the just-created commit and any local
follow-up edits are both covered.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from collections.abc import Iterable
from pathlib import Path


def _run_git(*args: str) -> str:
    result = subprocess.run(
        ["git", *args],
        check=True,
        text=True,
        encoding="utf-8",
        capture_output=True,
    )
    return result.stdout


def _normalize_path(value: str) -> str:
    path = str(value or "").replace("\\", "/")
    path = path.removeprefix("./")
    try:
        path = Path(path).resolve().relative_to(Path.cwd().resolve()).as_posix()
    except ValueError:
        path = Path(path).as_posix()
    return path


def parse_added_lines(diff_text: str) -> dict[str, set[int]]:
    """Parse added line numbers from a zero-context unified diff."""
    added: dict[str, set[int]] = {}
    current: str | None = None
    new_line: int | None = None
    for line in str(diff_text or "").splitlines():
        if line.startswith("+++ b/"):
            current = _normalize_path(line[6:])
            added.setdefault(current, set())
            new_line = None
            continue
        if line.startswith("@@"):
            if current is None:
                continue
            match = re.search(r"\+(\d+)(?:,(\d+))?", line)
            if match:
                new_line = int(match.group(1))
            continue
        if current is None or new_line is None or line.startswith("\\"):
            continue
        if line.startswith("+"):
            added[current].add(new_line)
            new_line += 1
        elif not line.startswith("-"):
            new_line += 1
    return added


def parse_diff_hunks(
    diff_text: str,
) -> dict[str, list[tuple[int, int, int, int]]]:
    """Return ``(old_start, old_count, new_start, new_count)`` per file."""
    hunks: dict[str, list[tuple[int, int, int, int]]] = {}
    current: str | None = None
    for line in str(diff_text or "").splitlines():
        if line.startswith("+++ b/"):
            current = _normalize_path(line[6:])
            hunks.setdefault(current, [])
            continue
        if current is None or not line.startswith("@@"):
            continue
        match = re.search(
            r"-(\d+)(?:,(\d+))?\s+\+(\d+)(?:,(\d+))?",
            line,
        )
        if match:
            hunks[current].append(
                (
                    int(match.group(1)),
                    int(match.group(2) or "1"),
                    int(match.group(3)),
                    int(match.group(4) or "1"),
                )
            )
    return hunks


def map_lines_through_diff(
    line_map: dict[str, Iterable[int]], diff_text: str
) -> dict[str, set[int]]:
    """Map line numbers from a committed file to its current worktree.

    The hunk's old coordinates refer to ``HEAD`` and its new coordinates refer
    to the worktree.  Insertion/deletion/replacement hunks therefore update
    the committed line set before it is merged with dirty additions.
    """
    hunks_by_file = parse_diff_hunks(diff_text)
    mapped: dict[str, set[int]] = {}
    for filename, lines in line_map.items():
        key = _normalize_path(filename)
        file_hunks = hunks_by_file.get(key, [])
        target: set[int] = set()
        for raw_line in lines:
            line = int(raw_line)
            if line <= 0:
                continue
            mapped_line = line
            for old_start, old_count, new_start, new_count in file_hunks:
                if old_count == 0:
                    # ``-N,0`` inserts after old line N, before N+1.
                    if line > old_start:
                        mapped_line += new_count
                    continue
                old_end = old_start + old_count - 1
                if line < old_start:
                    break
                if line <= old_end:
                    if new_count <= 0:
                        mapped_line = 0
                        break
                    relative = line - old_start
                    mapped_line = new_start + min(relative, new_count - 1)
                    break
                mapped_line += new_count - old_count
            if mapped_line > 0:
                target.add(mapped_line)
        mapped[key] = target
    return mapped


def merge_line_maps(*maps: dict[str, Iterable[int]]) -> dict[str, set[int]]:
    """Merge changed-line maps while preserving deterministic file order."""
    result: dict[str, set[int]] = {}
    for line_map in maps:
        for filename, lines in line_map.items():
            key = _normalize_path(filename)
            result.setdefault(key, set()).update(int(line) for line in lines)
    return result


def diagnostics_on_changed_lines(
    diagnostics: Iterable[dict], changed_lines: dict[str, set[int]]
) -> list[dict]:
    """Return Ruff diagnostics whose source range intersects added lines."""
    matched: list[dict] = []
    for diagnostic in diagnostics:
        filename = _normalize_path(str(diagnostic.get("filename", "")))
        lines = changed_lines.get(filename, set())
        location = diagnostic.get("location") or {}
        start = int(location.get("row") or 0)
        end_location = diagnostic.get("end_location") or {}
        end = int(end_location.get("row") or start)
        if not lines:
            continue
        if start <= 0 or any(start <= line <= max(start, end) for line in lines):
            matched.append(diagnostic)
    return matched


def _diff_for_base(base: str) -> str:
    return _run_git("diff", "--no-ext-diff", "--unified=0", f"{base}..HEAD", "--")


def _diff_for_worktree(base: str) -> str:
    return _run_git("diff", "--no-ext-diff", "--unified=0", base, "--")


def _untracked_python_lines() -> dict[str, set[int]]:
    result: dict[str, set[int]] = {}
    for raw_path in _run_git("ls-files", "--others", "--exclude-standard").splitlines():
        path = Path(raw_path)
        if path.suffix.lower() != ".py" or not path.is_file():
            continue
        key = _normalize_path(raw_path)
        result[key] = set(range(1, len(path.read_text(encoding="utf-8").splitlines()) + 1))
    return result


def changed_python_lines(base: str = "HEAD") -> dict[str, set[int]]:
    """Collect committed and worktree additions relative to ``base``."""
    committed = parse_added_lines(_diff_for_base(base))
    worktree_diff = _diff_for_worktree("HEAD")
    worktree = parse_added_lines(worktree_diff)
    committed_current = map_lines_through_diff(committed, worktree_diff)
    return merge_line_maps(committed_current, worktree, _untracked_python_lines())


def run_gate(base: str = "HEAD") -> tuple[int, dict]:
    changed = changed_python_lines(base)
    files = sorted(path for path in changed if Path(path).suffix.lower() == ".py")
    diagnostics: list[dict] = []
    all_diagnostics: list[dict] = []
    ruff_exit = 0
    error: str | None = None
    legacy_diagnostic_count = 0
    if files:
        try:
            ruff = subprocess.run(
                [sys.executable, "-m", "ruff", "check", "--output-format", "json", *files],
                check=False,
                text=True,
                encoding="utf-8",
                capture_output=True,
                timeout=120,
            )
            ruff_exit = ruff.returncode
        except (OSError, subprocess.SubprocessError) as exc:
            error = f"Ruff 子进程失败：{exc}"
        else:
            try:
                parsed = json.loads(ruff.stdout or "")
                if not isinstance(parsed, list):
                    raise TypeError("Ruff JSON 顶层不是数组")
                all_diagnostics = parsed
            except (json.JSONDecodeError, TypeError, ValueError) as exc:
                error = f"Ruff JSON 解析失败：{exc}"
            if error is None:
                diagnostics = diagnostics_on_changed_lines(all_diagnostics, changed)
                if ruff.returncode != 0:
                    # The production tree intentionally carries pre-existing
                    # legacy Ruff debt.  A non-zero exit containing only
                    # diagnostics outside the committed/worktree additions
                    # is recorded as baseline debt; an empty/invalid result,
                    # or any diagnostic on a changed line, remains fail-closed.
                    if all_diagnostics and not diagnostics:
                        legacy_diagnostic_count = len(all_diagnostics)
                    else:
                        error = f"Ruff 返回非零退出码：{ruff.returncode}"
    payload = {
        "base": base,
        "head": _run_git("rev-parse", "HEAD").strip(),
        "worktree_included": True,
        "files": files,
        "changed_line_counts": {path: len(changed[path]) for path in sorted(changed)},
        "ruff_exit": ruff_exit,
        "full_diagnostic_count": len(all_diagnostics) if files else 0,
        "legacy_diagnostic_count": legacy_diagnostic_count,
        "error": error,
        "changed_line_diagnostics": [
            {
                "code": item.get("code"),
                "filename": _normalize_path(str(item.get("filename", ""))),
                "line": (item.get("location") or {}).get("row"),
                "message": item.get("message"),
            }
            for item in diagnostics
        ],
    }
    return (2 if error else 1 if diagnostics else 0), payload


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--base",
        default="HEAD",
        help="resolved commit/ref before the change; defaults to HEAD for worktree-only checks",
    )
    args = parser.parse_args(argv)
    try:
        exit_code, payload = run_gate(args.base)
    except (OSError, subprocess.SubprocessError) as exc:
        print(getattr(exc, "stderr", None) or str(exc), file=sys.stderr)
        return 2
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
