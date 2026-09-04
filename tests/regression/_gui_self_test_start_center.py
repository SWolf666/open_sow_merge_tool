"""Native smoke for the no-argument start centre and explicit path picker."""

from __future__ import annotations

import os
import shutil
import tempfile
import tkinter as tk
from pathlib import Path

from sow_merge_tool.launch_center import CompareSelectionDialog, StartCenter


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
    path.write_bytes(payload)


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
        dialog = CompareSelectionDialog(root, initial_paths=(str(left), str(right)), on_confirm=captured.append)
        root.update_idletasks()
        assert dialog.confirm_button.instate(["!disabled"])
        dialog.confirm_button.invoke()
        assert captured and captured[0].paths == (str(left), str(right))

        left_dir = temp_root / "folder-left"
        right_dir = temp_root / "folder-right"
        _book(left_dir / "config" / "Alpha.xlsx", b"alpha-left")
        _book(right_dir / "config" / "Alpha.xlsx", b"alpha-right")
        folder_result = []
        folder_dialog = CompareSelectionDialog(
            root,
            initial_paths=(str(left_dir), str(right_dir)),
            on_confirm=folder_result.append,
        )
        root.update_idletasks()
        rows = folder_dialog.mapping_tree.get_children()
        assert rows
        folder_dialog.mapping_tree.selection_set(rows[0])
        folder_dialog._selection_changed()
        assert folder_dialog.confirm_button.instate(["!disabled"])
        folder_dialog.confirm_button.invoke()
        assert folder_result and folder_result[0].mapping.relative_path == "config/Alpha.xlsx"
    finally:
        try:
            root.destroy()
        except tk.TclError:
            pass
        shutil.rmtree(temp_root, ignore_errors=True)
    print("PASS: start centre and explicit path picker native smoke")


if __name__ == "__main__":
    main()
