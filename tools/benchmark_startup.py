"""Headless startup profile for large Excel copies.

Set ``SOW_PROFILE_BUILDING`` to an approved workbook copy to profile it.  If
unset, a disposable 20-sheet synthetic workbook is used.  Only sizes,
counts, and timings are emitted; input paths and cell contents never appear in
the JSON result.  ``SOW_PROFILE_BUILDING_UI=1`` additionally runs one hidden
Tk startup probe and records the existing startup trace/heartbeat metrics.
"""

from __future__ import annotations

import json
import os
import shutil
import statistics
import tempfile
import threading
import time
from itertools import pairwise
from pathlib import Path

from openpyxl import Workbook, load_workbook

from sow_merge_tool import legacy_core as smt

ITERATIONS = 4
P95_LIMIT_MS = 1500.0


def _write_synthetic(path: Path) -> None:
    workbook = Workbook(write_only=True)
    for index in range(20):
        sheet = workbook.create_sheet(f"Sheet{index:02d}")
        sheet.append(["id", "value", "note"])
        for row in range(1, 1601):
            sheet.append([f"key-{row}", row, "profile"])
    workbook.save(path)
    workbook.close()


def _prepare_inputs(root: Path) -> tuple[Path, Path, str]:
    configured = os.environ.get("SOW_PROFILE_BUILDING", "").strip()
    source = Path(configured).resolve() if configured else None
    if source is not None and not source.is_file():
        raise FileNotFoundError("SOW_PROFILE_BUILDING does not exist")
    if source is None:
        source = root / "synthetic.xlsx"
        _write_synthetic(source)
        fixture = "synthetic"
    else:
        fixture = "approved-copy"
    left = root / "left.xlsx"
    right = root / "right.xlsx"
    shutil.copy2(source, left)
    shutil.copy2(source, right)
    return left, right, fixture


def _read_only_phase(left: Path, right: Path) -> dict[str, float | int]:
    started = time.perf_counter()
    left_wb = load_workbook(left, data_only=True, read_only=True, keep_links=False)
    right_wb = load_workbook(right, data_only=True, read_only=True, keep_links=False)
    opened_ms = (time.perf_counter() - started) * 1000.0
    try:
        catalog_started = time.perf_counter()
        left_names = list(left_wb.sheetnames)
        right_names = list(right_wb.sheetnames)
        catalog_ms = (time.perf_counter() - catalog_started) * 1000.0
        first_name = left_names[0] if left_names else None
        first_started = time.perf_counter()
        row_count = 0
        if first_name is not None:
            for _row in left_wb[first_name].iter_rows(max_row=120, values_only=True):
                row_count += 1
        first_sheet_ms = (time.perf_counter() - first_started) * 1000.0
        return {
            "read_only_open_ms": opened_ms,
            "sheet_catalog_ms": catalog_ms,
            "first_sheet_preview_ms": first_sheet_ms,
            "sheet_count": len(set(left_names) | set(right_names)),
            "preview_rows": row_count,
        }
    finally:
        left_wb.close()
        right_wb.close()


def _profile_inputs(left: Path, right: Path) -> dict[str, object]:
    samples = []
    for _index in range(ITERATIONS):
        samples.append(_read_only_phase(left, right))
    cold = samples[0]
    warm_values = [float(item["read_only_open_ms"]) for item in samples[1:]]
    all_values = [float(item["read_only_open_ms"]) for item in samples]
    return {
        "iterations": ITERATIONS,
        "cold": cold,
        "warm_open_ms": warm_values,
        "warm_open_median_ms": statistics.median(warm_values),
        "open_p95_ms": sorted(all_values)[max(0, int(len(all_values) * 0.95) - 1)],
    }


def _headless_heartbeat_probe(left: Path, right: Path) -> dict[str, float]:
    """Prove the expensive read is off the caller's cooperative event loop."""
    done = threading.Event()
    worker_error: list[Exception] = []

    def worker() -> None:
        try:
            _read_only_phase(left, right)
        except (OSError, RuntimeError, KeyError, ValueError, TypeError) as exc:  # pragma: no cover - surfaced below
            worker_error.append(exc)
        finally:
            done.set()

    thread = threading.Thread(target=worker, name="sow-startup-profile", daemon=True)
    thread.start()
    pulses = []
    while not done.wait(0.01):
        pulses.append(time.perf_counter())
    thread.join(timeout=2.0)
    if worker_error:
        raise worker_error[0]
    gaps = [
        (right - left) * 1000.0
        for left, right in pairwise(pulses)
    ]
    return {
        "samples": float(len(pulses)),
        "max_gap_ms": max(gaps, default=0.0),
    }


def _hidden_ui_probe(left: Path, right: Path) -> dict[str, object]:
    app = None
    try:
        app = smt.SowMergeApp(str(left), str(right))
        app.root.withdraw()
        probe_started = time.perf_counter()
        first_sheet = str(getattr(app, "selected_sheet", "") or "")
        first_ready_at = None
        deadline = time.monotonic() + 45.0
        while time.monotonic() < deadline:
            app.root.update()
            if first_ready_at is None:
                view = getattr(app, "sheet_views", {}).get(first_sheet)
                if view is not None and bool(getattr(view, "_data_ready", False)):
                    first_ready_at = time.perf_counter()
            if getattr(app, "_ui_heartbeat_samples", 0) >= 10:
                break
            time.sleep(0.01)
        durations = app._startup_trace.durations()
        return {
            "trace_ms": {key: round(value * 1000.0, 2) for key, value in durations.items()},
            "first_sheet": first_sheet,
            "first_sheet_ready": first_ready_at is not None,
            "first_sheet_ready_ms": (
                round((first_ready_at - probe_started) * 1000.0, 2)
                if first_ready_at is not None
                else None
            ),
            "edit_loaded_during_probe": bool(app._edit_workbooks_ready()),
            "heartbeat_samples": int(getattr(app, "_ui_heartbeat_samples", 0)),
            "heartbeat_max_gap_ms": round(
                float(getattr(app, "_ui_heartbeat_max_gap", 0.0)) * 1000.0,
                2,
            ),
        }
    finally:
        if app is not None:
            app._shutdown_root()


def main() -> None:
    root = Path(tempfile.mkdtemp(prefix="sow-startup-profile-"))
    try:
        threshold_override = os.environ.get("SOW_PROFILE_LARGE_THRESHOLD", "").strip()
        if threshold_override:
            smt._LARGE_SHEET_ROW_THRESHOLD = max(1, int(threshold_override))
        left, right, fixture = _prepare_inputs(root)
        result = {
            "fixture": fixture,
            "large_sheet_threshold": int(smt._LARGE_SHEET_ROW_THRESHOLD),
            "input_bytes": left.stat().st_size,
            "profile": _profile_inputs(left, right),
            "headless_heartbeat": _headless_heartbeat_probe(left, right),
        }
        if os.environ.get("SOW_PROFILE_BUILDING_UI", "").strip() == "1":
            result["hidden_ui"] = _hidden_ui_probe(left, right)
        p95_ms = float(result["profile"]["open_p95_ms"])
        result["p95_limit_ms"] = P95_LIMIT_MS
        if p95_ms > P95_LIMIT_MS:
            raise RuntimeError(f"read-only startup open p95 exceeds {P95_LIMIT_MS:.0f}ms")
        output = Path("artifacts/performance/building-startup.json")
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(result, ensure_ascii=False))
    finally:
        shutil.rmtree(root, ignore_errors=True)


if __name__ == "__main__":
    main()
