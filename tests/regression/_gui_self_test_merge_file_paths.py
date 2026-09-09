"""Headed smoke test for visible workbook paths and aligned compare cards."""

from __future__ import annotations

import ctypes
import os
import struct
import tempfile
import time
import tkinter as tk
import zlib
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

import sow_merge_tool as smt


def _write_book(path: str, value: str) -> None:
    workbook = Workbook()
    workbook.active["A1"] = value
    workbook.save(path)
    workbook.close()


def _wait_for_view(app, timeout: float = 12.0):
    app.nb.select(app._sheet_containers["Sheet"])
    deadline = time.time() + timeout
    while time.time() < deadline:
        app.root.update()
        view = app.sheet_views.get("Sheet")
        if view is not None:
            return view
        time.sleep(0.01)
    raise AssertionError("merge path GUI view was not created")


def _write_window_png(root, path: Path) -> None:
    """Capture the native window without adding an image dependency."""
    import win32gui
    import win32ui

    hwnd = root.winfo_id()
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = max(1, int(right - left))
    height = max(1, int(bottom - top))
    window_dc = win32gui.GetWindowDC(hwnd)
    source_dc = win32ui.CreateDCFromHandle(window_dc)
    capture_dc = source_dc.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(source_dc, width, height)
    capture_dc.SelectObject(bitmap)
    try:
        if not ctypes.windll.user32.PrintWindow(hwnd, capture_dc.GetSafeHdc(), 2):
            raise AssertionError("native PrintWindow failed")
        raw = bitmap.GetBitmapBits(True)
    finally:
        capture_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, window_dc)
        win32gui.DeleteObject(bitmap.GetHandle())

    # ``PrintWindow`` returns the compatible bitmap in display order through
    # ``GetBitmapBits(True)`` on the supported Windows drivers.  Emit a small
    # RGB PNG using only the standard library so the visual artefact is
    # reproducible in CI.
    rows = []
    stride = width * 4
    for row in range(height):
        pixels = raw[row * stride:(row + 1) * stride]
        rgb = bytearray(width * 3)
        for column in range(width):
            source = column * 4
            target = column * 3
            rgb[target:target + 3] = bytes(
                (pixels[source + 2], pixels[source + 1], pixels[source])
            )
        rows.append(b"\x00" + bytes(rgb))

    def chunk(kind: bytes, payload: bytes) -> bytes:
        return (
            struct.pack(">I", len(payload))
            + kind
            + payload
            + struct.pack(">I", zlib.crc32(kind + payload) & 0xFFFFFFFF)
        )

    png = (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(b"".join(rows), 6))
        + chunk(b"IEND", b"")
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(png)


def _assert_command_layout(app, width: int, height: int, scale: float) -> None:
    root = app.root
    root.state("normal")
    root.geometry(f"{width}x{height}+40+40")
    deadline = time.monotonic() + 0.35
    while time.monotonic() < deadline:
        root.update()
        time.sleep(0.02)
    root.update_idletasks()
    root_left = root.winfo_rootx()
    root_top = root.winfo_rooty()
    root_right = root_left + root.winfo_width()
    root_bottom = root_top + root.winfo_height()
    core = {
        "重算并刷新": app.recalc_btn,
        "上一差异": app.top_prev_btn,
        "下一差异": app.top_next_btn,
        "搜索": app.top_search_btn,
        "筛选": app.top_filter_btn,
        "应用": app.top_apply_btn,
        "保留": app.top_retain_btn,
        "撤销": app.top_undo_btn,
        "保存": app.primary_save_btn,
        "另存为": app.top_save_as_btn,
        "更多": app.more_menu,
    }
    if app.top_base_btn is not None:
        core["采用 Base"] = app.top_base_btn
    rectangles = []
    for name, widget in core.items():
        if not widget.winfo_ismapped():
            if widget is app.recalc_btn:
                assert width < 1080 and scale >= 1.9, (
                    f"仅允许高 DPI 窄窗口收纳比较组：{width}x{height} @ {scale}"
                )
                menu_labels = [
                    app.more_menu_model.entrycget(index, "label")
                    for index in range(app.more_menu_model.index("end") + 1)
                    if app.more_menu_model.type(index) == "command"
                ]
                assert "重算并刷新" in menu_labels, "高 DPI 窄窗口未提供重算入口"
                assert "比较设置" in menu_labels, "高 DPI 窄窗口未提供比较设置入口"
                invoked = []
                with patch.object(
                    app,
                    "recalc_and_refresh",
                    side_effect=lambda bucket=invoked: bucket.append("recalc"),
                ), patch.object(
                    app,
                    "_show_compare_settings",
                    side_effect=lambda bucket=invoked: bucket.append("settings"),
                ):
                    for index in range(app.more_menu_model.index("end") + 1):
                        if app.more_menu_model.type(index) != "command":
                            continue
                        if app.more_menu_model.entrycget(index, "label") in {
                            "重算并刷新",
                            "比较设置",
                        }:
                            app.more_menu_model.invoke(index)
                assert invoked == ["recalc", "settings"], (
                    f"更多菜单回调未实际执行：{invoked}"
                )
                continue
            raise AssertionError(f"{name} 未映射（{width}x{height}）")
        x = widget.winfo_rootx()
        y = widget.winfo_rooty()
        w = widget.winfo_width()
        h = widget.winfo_height()
        assert w > 0 and h > 0, f"{name} 无尺寸（{width}x{height}）"
        assert root_left <= x and x + w <= root_right, (
            f"{name} 横向越界：x={x} w={w} root=[{root_left},{root_right}]"
        )
        assert root_top <= y and y + h <= root_bottom, (
            f"{name} 纵向越界：y={y} h={h} root=[{root_top},{root_bottom}]"
        )
        rectangles.append((name, x, y, x + w, y + h))
    for index, (left_name, left_x0, left_y0, left_x1, left_y1) in enumerate(rectangles):
        for right_name, right_x0, right_y0, right_x1, right_y1 in rectangles[index + 1:]:
            overlaps = left_x0 < right_x1 and right_x0 < left_x1 and left_y0 < right_y1 and right_y0 < left_y1
            assert not overlaps, f"命令按钮重叠：{left_name} / {right_name}（{width}x{height}）"
    assert app.primary_save_btn.winfo_ismapped()
    assert app.top_save_as_btn.winfo_ismapped()
    assert app.more_menu.winfo_ismapped()
    assert app.top_apply_btn.winfo_ismapped()
    if scale == 1.0 and width == 1920:
        assert app.recalc_btn.winfo_ismapped(), "100%/1920 比较按钮不应收纳"
        assert app.top_save_as_btn.winfo_ismapped(), "100%/1920 另存为按钮不应收纳"


def _new_app_at_scaling(base_path: str, mine_path: str, *, scale: float, theirs_path: str | None = None):
    """Create the real window only after the requested Tk scaling is active."""
    root = tk.Tk()
    root.withdraw()
    root.tk.call("tk", "scaling", scale)
    smt._STARTUP_PROGRESS_ROOT = root
    if theirs_path is None:
        return smt.SowMergeApp(
            base_path,
            mine_path,
            raw_base=base_path,
            raw_mine=mine_path,
        )
    return smt.SowMergeApp(
        mine_path,
        theirs_path,
        merge_mode=True,
        base_path=base_path,
        raw_base=base_path,
        raw_mine=mine_path,
        raw_theirs=theirs_path,
    )


def main() -> None:
    with tempfile.TemporaryDirectory(prefix="sow-visible-paths-") as temp_root:
        base_dir = os.path.join(temp_root, "源分支 修改前")
        mine_dir = os.path.join(temp_root, "目标分支 修改后预览")
        theirs_dir = os.path.join(temp_root, "来源分支 修改后")
        os.makedirs(base_dir)
        os.makedirs(mine_dir)
        os.makedirs(theirs_dir)
        base_path = os.path.join(base_dir, "Language.xlsx")
        mine_path = os.path.join(mine_dir, "Language.xlsx")
        theirs_path = os.path.join(theirs_dir, "Language.xlsx")
        _write_book(base_path, "修改前")
        _write_book(mine_path, "修改后")
        _write_book(theirs_path, "来源修改")

        # Keep the path-card checks on a real two-way window, then create one
        # fresh window per scaling factor so the widgets are constructed under
        # the same DPI that the layout gate is validating.
        app = _new_app_at_scaling(base_path, mine_path, scale=1.0)
        try:
            view = _wait_for_view(app)
            app.root.update_idletasks()
            assert app.sheet_filter_diff_button.winfo_ismapped()
            assert app.sheet_filter_all_button.winfo_ismapped()
            assert app.sheet_filter_summary_var.get().startswith("有差异")
            assert app._sheet_filter_model.only_diff is True
            assert os.path.normpath(base_path) in view.path_file_label_a.cget("text")
            assert os.path.normpath(mine_path) in view.path_file_label_b.cget("text")
            assert not view.path_card_base.winfo_manager(), "two-way mode must not reserve a blank Base card"
            assert int(view.path_card_a.grid_info()["column"]) == 0
            assert int(view.path_card_b.grid_info()["column"]) == 1
            assert abs(view.path_card_a.winfo_width() - view.path_card_b.winfo_width()) <= 4
            view._copy_visible_pane_path("b")
            assert app.root.clipboard_get() == mine_path
            assert view.save_b_btn.cget("text") == "保存修改后文件"
        finally:
            app._shutdown_root()

        screenshots = Path("tmp/postinstall_acceptance")
        for scale in (1.0, 1.25, 1.5, 2.0, 2.666):
            app = _new_app_at_scaling(base_path, mine_path, scale=scale)
            try:
                _wait_for_view(app)
                for width, height in ((900, 620), (1366, 768), (1920, 1080)):
                    _assert_command_layout(app, width, height, scale)
                    if scale == 1.0:
                        _write_window_png(
                            app.root,
                            screenshots / f"ordinary_{width}_after.png",
                        )
            finally:
                app._shutdown_root()

        # The three-way layout has the widest semantic labels (Base/Mine/
        # Theirs and the Base direction command), so run the same Native gate
        # against a separately constructed three-way window at every DPI.
        for scale in (1.0, 1.25, 1.5, 2.0, 2.666):
            app = _new_app_at_scaling(
                base_path,
                mine_path,
                scale=scale,
                theirs_path=theirs_path,
            )
            try:
                _wait_for_view(app)
                for width, height in ((900, 620), (1366, 768), (1920, 1080)):
                    _assert_command_layout(app, width, height, scale)
                    if scale == 1.0:
                        _write_window_png(
                            app.root,
                            screenshots / f"threeway_{width}_after.png",
                        )
            finally:
                app._shutdown_root()
    print("PASS: visible workbook paths and aligned compare cards")


if __name__ == "__main__":
    main()
