"""Native smoke for persistent pairing-list and comparison-session lifecycle."""

from __future__ import annotations

import shutil
import tempfile
import tkinter as tk
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from sow_merge_tool import launch_center as lc


class _FakeProcess:
    def __init__(self):
        self.returncode = None
        self.terminated = False

    def poll(self):
        return self.returncode

    def terminate(self):
        self.terminated = True
        self.returncode = -15


class _FakeManager:
    def __init__(self, _root):
        self.process = _FakeProcess()
        self.last = None

    @property
    def active_session(self):
        return self.last if self.last is not None and self.last.active else None

    def start(self, left, right, *, callback=None):
        launch = lc.ComparisonLaunch.for_paths(left, right)
        session = lc.ComparisonSession("native-session", launch, process=self.process, state="比较中")
        self.last = session
        if callback:
            callback(session)
        return session

    def poll_now(self):
        if self.last is not None and self.process.returncode is not None:
            self.last.returncode = self.process.returncode
            self.last.state = "已关闭"
            self.last.closed_at = 1.0
        return self.last

    def close(self, *, terminate=False):
        if terminate:
            self.process.terminate()

    def resume(self, _root=None):
        return None


def _book(path: Path, value: bytes):
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    workbook.active["A1"] = value.decode("utf-8")
    workbook.save(path)
    workbook.close()


def main() -> None:
    root = tk.Tk()
    root.withdraw()
    temp_root = Path(tempfile.mkdtemp(prefix="sow_compare_sessions_native_"))
    try:
        left = temp_root / "Source 中文.xlsx"
        right = temp_root / "Target 中文.xlsx"
        _book(left, b"left")
        _book(right, b"right")
        with patch.object(lc, "ComparisonSessionManager", _FakeManager):
            center = lc.StartCenter(initial_paths=(str(left), str(right)))
            center._compare()
            dialog = center.compare_dialog
            assert dialog is not None
            root.update_idletasks()
            assert dialog.mapping_tree.get_children()
            assert center.root.state() == "withdrawn"
            dialog.confirm_button.invoke()
            root.update_idletasks()
            assert dialog.win.winfo_exists(), "启动子比较后列表不能销毁"
            assert dialog.manager.last is not None
            assert dialog.manager.last.state == "比较中"
            assert "--compare" not in dialog.manager.last.launch.command
            assert dialog.manager.last.launch.command[-2:] == (str(left), str(right))

            dialog.manager.process.returncode = 0
            dialog._update_session_row(dialog.manager.poll_now())
            root.update_idletasks()
            row = dialog.mapping_tree.get_children()[0]
            assert dialog.mapping_tree.set(row, "session") == "已关闭"
            dialog._cancel()
            root.update_idletasks()
            assert center.root.state() != "withdrawn"
            center.root.destroy()
    finally:
        try:
            root.destroy()
        except tk.TclError:
            pass
        shutil.rmtree(temp_root, ignore_errors=True)
    print("PASS: comparison session lifecycle native smoke")


if __name__ == "__main__":
    main()
