"""Headless profile for source-delta reuse and target projection.

The default fixture is synthetic and disposable.  Set
``SOW_PROFILE_SOURCE_BEFORE``, ``SOW_PROFILE_SOURCE_AFTER`` and
``SOW_PROFILE_TARGET`` to profile approved World/Language copies without
opening Tk or Excel.  Only timing/count metadata is written to the ignored
performance directory; workbook paths and cell contents are not recorded.
"""

from __future__ import annotations

import json
import os
import shutil
import statistics
import tempfile
import time
from pathlib import Path

from openpyxl import Workbook

from sow_merge_tool import fast_branch_merge as fast_merge

ITERATIONS = 20
P95_LIMIT_MS = 1500.0


def _write_fixture(path: Path, *, changed: bool) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "Language"
    sheet.append(["id", "value", "note"])
    for row in range(1, 1201):
        sheet.append([f"key-{row}", row + (1 if changed and row % 97 == 0 else 0), "fixture"])
    book.save(path)
    book.close()


def _resolve_inputs(root: Path) -> tuple[Path, Path, Path, bool]:
    configured = tuple(
        os.environ.get(name, "").strip()
        for name in (
            "SOW_PROFILE_SOURCE_BEFORE",
            "SOW_PROFILE_SOURCE_AFTER",
            "SOW_PROFILE_TARGET",
        )
    )
    if all(configured):
        paths = tuple(Path(value).resolve() for value in configured)
        if not all(path.is_file() for path in paths):
            raise FileNotFoundError("profile workbook input does not exist")
        return paths[0], paths[1], paths[2], False
    before, after, target = (root / name for name in ("before.xlsx", "after.xlsx", "target.xlsx"))
    _write_fixture(before, changed=False)
    _write_fixture(after, changed=True)
    _write_fixture(target, changed=False)
    return before, after, target, True


def _profile(before: Path, after: Path, target: Path) -> dict[str, object]:
    cold_ms = []
    warm_ms = []
    incoming = None
    for index in range(ITERATIONS):
        if index == 0:
            fast_merge.cache_clear()
        started = time.perf_counter()
        delta = fast_merge.analyze_source(str(before), str(after), before_revision=1, source_revision=1)
        decision = fast_merge.analyze_target(delta, str(target), target_revision=1)
        elapsed = (time.perf_counter() - started) * 1000.0
        (cold_ms if index == 0 else warm_ms).append(elapsed)
        incoming = delta.incoming_count
        if decision.disposition not in {"direct", "already_applied", "confirmation_required"}:
            raise RuntimeError(f"unexpected profile decision: {decision.disposition}")
    samples = cold_ms + warm_ms
    return {
        "iterations": ITERATIONS,
        "cold_ms": cold_ms,
        "warm_ms": warm_ms,
        "median_ms": statistics.median(samples),
        "p95_ms": sorted(samples)[max(0, int(len(samples) * 0.95) - 1)],
        "max_ms": max(samples),
        "incoming_count": incoming,
        "cache": {
            "index": fast_merge._load_cached.cache_info()._asdict(),
            "content_digest": fast_merge._content_digest.cache_info()._asdict(),
        },
    }


def main() -> None:
    root = Path(tempfile.mkdtemp(prefix="sow-merge-profile-"))
    try:
        before, after, target, synthetic = _resolve_inputs(root)
        result = _profile(before, after, target)
        result["fixture"] = "synthetic" if synthetic else "approved-input"
        result["p95_limit_ms"] = P95_LIMIT_MS
        baseline_path = os.environ.get("SOW_PROFILE_BASELINE", "").strip()
        if baseline_path:
            baseline = json.loads(Path(baseline_path).read_text(encoding="utf-8"))
            baseline_p95 = float(baseline.get("p95_ms", 0.0))
            result["baseline_p95_ms"] = baseline_p95
            if baseline_p95 > 0 and result["p95_ms"] > baseline_p95 * 1.10:
                raise RuntimeError(
                    f"merge profile p95 regressed: {result['p95_ms']:.1f}ms > {baseline_p95 * 1.10:.1f}ms"
                )
        if result["p95_ms"] > P95_LIMIT_MS:
            raise RuntimeError(f"merge profile p95 exceeds {P95_LIMIT_MS:.0f}ms")
        output = Path("artifacts/performance/update96-merge.json")
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(result, ensure_ascii=False))
    finally:
        shutil.rmtree(root, ignore_errors=True)


if __name__ == "__main__":
    main()
