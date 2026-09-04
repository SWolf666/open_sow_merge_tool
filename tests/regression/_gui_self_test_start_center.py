"""Native smoke for the no-argument start centre and explicit path picker."""

from __future__ import annotations

import os
import shutil
import tempfile
import time
import tkinter as tk
from pathlib import Path

from openpyxl import Workbook

from sow_merge_tool.launch_center import (
    CompareSelectionDialog,
    ComparisonSessionManager,
    StartCenter,
)


def _labels(widget):
    result = []
    try:
        text = str(widget.cget("text"))
        if text:
            result.append(text)
    except (tk.TclError, TypeError):
        pass
    for child in widget.winfo_children():
        result.extend(_labels(child))
    return result


def _book(path: Path, payload: bytes):
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    workbook.active["A1"] = payload.decode("utf-8")
    workbook.save(path)
    workbook.close()


def main() -> None:
    if os.name == "nt":
        try:
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, OSError):
            pass
    root = tk.Tk()
    root.withdraw()
    root.tk.call("tk", "scaling", 1.0)
    center = StartCenter(initial_paths=())
    center.root.update_idletasks()
    labels = _labels(center.root)
    assert "Excel 文件比较 / 合并" in labels
    assert "多分支 SVN 提交" in labels
    center.root.destroy()

    temp_root = Path(tempfile.mkdtemp(prefix="sow_start_center_native_"))
    try:
        left = temp_root / "中文 路径" / "配置.xlsx"
        right = temp_root / "目标 路径" / "配置.xlsx"
        _book(left, b"left")
        _book(right, b"right")
        captured = []
        process = type("Process", (), {"poll": lambda self: None})()
        manager = ComparisonSessionManager(
            root,
            popen_factory=lambda _command, **_kwargs: process,
        )
        dialog = CompareSelectionDialog(
            root,
            initial_paths=(str(left), str(right)),
            on_confirm=captured.append,
            manager=manager,
        )
        root.update_idletasks()
        assert dialog.confirm_button.instate(["!disabled"])
        dialog.confirm_button.invoke()
        assert captured and captured[0].paths == (str(left), str(right))
        assert dialog.win.winfo_exists(), "打开比较后配对列表必须保留"

        left_dir = temp_root / "folder-left"
        right_dir = temp_root / "folder-right"
        _book(left_dir / "config" / "Alpha.xlsx", b"alpha-left")
        _book(right_dir / "config" / "Alpha.xlsx", b"alpha-right")
        folder_result = []
        folder_dialog = CompareSelectionDialog(
            root,
            initial_paths=(str(left_dir), str(right_dir)),
            on_confirm=folder_result.append,
            manager=ComparisonSessionManager(
                root,
                popen_factory=lambda _command, **_kwargs: type(
                    "Process", (), {"poll": lambda self: None}
                )(),
            ),
        )
        deadline = time.monotonic() + 3
        while time.monotonic() < deadline:
            root.update()
            rows = folder_dialog.mapping_tree.get_children()
            if rows:
                break
            time.sleep(0.02)
        assert rows
        folder_dialog.mapping_tree.selection_set(rows[0])
        folder_dialog._selection_changed()
        assert folder_dialog.confirm_button.instate(["!disabled"])
        folder_dialog.confirm_button.invoke()
        assert folder_result and folder_result[0].mapping.relative_path == "config/Alpha.xlsx"
        assert folder_dialog.win.winfo_exists()
        broken = temp_root / "broken.xlsx"
        broken.write_bytes(b"not an OOXML workbook")
        broken_dialog = CompareSelectionDialog(
            root,
            initial_paths=(str(broken), str(right)),
            manager=ComparisonSessionManager(
                root,
                popen_factory=lambda _command, **_kwargs: type(
                    "Process", (), {"poll": lambda self: None}
                )(),
            ),
        )
        root.update_idletasks()
        broken_dialog.confirm_button.invoke()
        assert broken_dialog.win.winfo_exists()
        assert "损坏" in broken_dialog.status_var.get()
        dialog._cancel()
        folder_dialog._cancel()
        broken_dialog._cancel()
    finally:
        try:
            root.destroy()
        except tk.TclError:
            pass
        shutil.rmtree(temp_root, ignore_errors=True)
    print("PASS: start centre and explicit path picker native smoke")


if __name__ == "__main__":
    main()
