"""The no-argument start centre and the explicit Excel path picker.

The centre is intentionally a thin Tk adapter around :mod:`path_selection`.
Explorer and TortoiseSVN can pass one path as a prefill, but the user must
still confirm both sides before a workbook is opened.  Directory selections
show an explicit relative-path matrix; no file is guessed from a basename and
no workbook is written while the picker is open.
"""

from __future__ import annotations

import os
from collections.abc import Iterable
from dataclasses import dataclass

from .path_selection import (
    PairStatus,
    PathSelectionError,
    RelativeMapping,
    SelectionSnapshot,
    compare_selection_from_paths,
    map_relative_files,
    recheck_selection,
    snapshot_selection,
    validate_directory,
    validate_file_pair,
)


@dataclass(frozen=True)
class StartCenterResult:
    mode: str | None
    paths: tuple[str, ...] = ()
    mapping: RelativeMapping | None = None


def _status_text(status: PairStatus) -> str:
    return {
        PairStatus.MATCHED: "匹配",
        PairStatus.SAME_NAME_DIFFERENT_PATH: "同名不同路径（已明确选择）",
        PairStatus.MISSING_LEFT: "左侧缺失",
        PairStatus.MISSING_RIGHT: "右侧缺失",
        PairStatus.SAME_FILE: "同一文件（阻断）",
        PairStatus.TYPE_INCOMPATIBLE: "类型不兼容（阻断）",
        PairStatus.HASH_CHANGED: "内容已变化（请重新选择）",
        PairStatus.UNSUPPORTED: "类型不支持（阻断）",
        PairStatus.INVALID: "无效（阻断）",
    }.get(status, str(status))


class CompareSelectionDialog:
    """Explicit two-side file/folder selection dialog.

    ``on_confirm`` receives ``(left, right)`` and is called only after a hash
    recheck.  The object is useful to native tests without requiring a
    workbook or changing the application's global picker function.
    """

    def __init__(self, parent, *, initial_paths: Iterable[str] = (), on_confirm=None, on_cancel=None):
        import tkinter as tk
        from tkinter import ttk

        self.tk, self.ttk = tk, ttk
        self.parent = parent
        self.on_confirm = on_confirm
        self.on_cancel = on_cancel
        self.result: StartCenterResult | None = None
        self.snapshot: SelectionSnapshot | None = None
        self.mapping_rows: dict[str, RelativeMapping] = {}
        self.left_var = tk.StringVar()
        self.right_var = tk.StringVar()
        self.status_var = tk.StringVar(value="请选择左右两侧的 Excel 文件或文件夹。")
        values = list(compare_selection_from_paths(initial_paths))
        if values:
            self.left_var.set(values[0])
        if len(values) > 1:
            self.right_var.set(values[1])

        self.win = tk.Toplevel(parent)
        self.win.title("Excel 文件比较 / 合并")
        self.win.geometry("820x570")
        self.win.minsize(720, 500)
        self.win.transient(parent)
        self._build()
        self.win.protocol("WM_DELETE_WINDOW", self._cancel)
        if values:
            self._validate()

    def _build(self):
        ttk = self.ttk
        outer = ttk.Frame(self.win, padding=14)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="选择比较两侧", font=("Segoe UI", 13, "bold")).pack(anchor="w")
        ttk.Label(
            outer,
            text="可以分别选择 Excel 文件或文件夹。路径不同不会按文件名猜测；文件夹只按相对路径映射，最终仍需明确选择文件。",
            wraplength=770,
        ).pack(anchor="w", pady=(3, 12))

        rows = ttk.LabelFrame(outer, text="比较输入", padding=10)
        rows.pack(fill="x")
        for row, (label, variable) in enumerate((("左侧 / Source", self.left_var), ("右侧 / Target", self.right_var))):
            ttk.Label(rows, text=label, width=18).grid(row=row, column=0, sticky="w", pady=4)
            entry = ttk.Entry(rows, textvariable=variable)
            entry.grid(row=row, column=1, sticky="ew", padx=(6, 6), pady=4)
            setattr(self, "left_entry" if row == 0 else "right_entry", entry)
            ttk.Button(rows, text="选择文件…", command=lambda v=variable: self._choose_file(v)).grid(row=row, column=2, pady=4)
            ttk.Button(rows, text="选择文件夹…", command=lambda v=variable: self._choose_directory(v)).grid(row=row, column=3, padx=(5, 0), pady=4)
        rows.columnconfigure(1, weight=1)
        self.left_var.trace_add("write", lambda *_args: self._schedule_validate())
        self.right_var.trace_add("write", lambda *_args: self._schedule_validate())

        status = ttk.LabelFrame(outer, text="路径与哈希检查", padding=8)
        status.pack(fill="x", pady=(10, 0))
        ttk.Label(status, textvariable=self.status_var, wraplength=760).pack(fill="x")

        mapping_box = ttk.LabelFrame(outer, text="相对路径映射（文件夹模式）", padding=8)
        mapping_box.pack(fill="both", expand=True, pady=(10, 0))
        holder = ttk.Frame(mapping_box)
        holder.pack(fill="both", expand=True)
        self.mapping_tree = ttk.Treeview(
            holder,
            columns=("relative", "left", "right", "status"),
            show="headings",
            selectmode="browse",
        )
        for key, title, width in (("relative", "相对路径", 250), ("left", "左侧", 190), ("right", "右侧", 190), ("status", "状态", 160)):
            self.mapping_tree.heading(key, text=title)
            self.mapping_tree.column(key, width=width, anchor="w", stretch=key in {"relative", "left", "right"})
        scroll = ttk.Scrollbar(holder, orient="vertical", command=self.mapping_tree.yview)
        self.mapping_tree.configure(yscrollcommand=scroll.set)
        self.mapping_tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.mapping_tree.bind("<<TreeviewSelect>>", lambda _event: self._selection_changed())
        self.mapping_tree.bind("<Double-1>", lambda _event: self._confirm())

        footer = ttk.Frame(outer)
        footer.pack(fill="x", pady=(10, 0))
        self.confirm_button = ttk.Button(footer, text="打开所选比较", command=self._confirm)
        self.confirm_button.pack(side="right")
        ttk.Button(footer, text="取消", command=self._cancel).pack(side="right", padx=(0, 8))

    def _choose_file(self, variable):
        from tkinter import filedialog

        value = filedialog.askopenfilename(
            parent=self.win,
            title="选择 Excel 文件",
            filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if value:
            variable.set(value)

    def _choose_directory(self, variable):
        from tkinter import filedialog

        value = filedialog.askdirectory(parent=self.win, title="选择比较文件夹")
        if value:
            variable.set(value)

    def _schedule_validate(self):
        try:
            after_id = getattr(self, "_validate_after", None)
            if after_id:
                self.win.after_cancel(after_id)
            self._validate_after = self.win.after(120, self._validate)
        except self.tk.TclError:
            pass

    def _clear_mapping(self):
        self.mapping_rows = {}
        for iid in self.mapping_tree.get_children():
            self.mapping_tree.delete(iid)

    def _validate(self):
        self._clear_mapping()
        self.snapshot = None
        left = self.left_var.get().strip()
        right = self.right_var.get().strip()
        self.confirm_button.state(["disabled"])
        if not left or not right:
            self.status_var.set("请分别选择左侧和右侧路径；右键传入的路径仅作为预填。")
            return
        left_probe = validate_directory(left) if os.path.isdir(left) else None
        right_probe = validate_directory(right) if os.path.isdir(right) else None
        if (left_probe is None) != (right_probe is None):
            self.status_var.set("两侧必须同时选择文件，或同时选择文件夹；类型不兼容。")
            return
        if left_probe and right_probe:
            try:
                if not left_probe.ready or not right_probe.ready:
                    detail = "; ".join(left_probe.issues + right_probe.issues)
                    self.status_var.set(f"文件夹不可用：{detail or '请重新选择'}")
                    return
                mappings = map_relative_files(left, right)
            except (OSError, ValueError, PathSelectionError) as exc:
                self.status_var.set(str(exc))
                return
            for index, mapping in enumerate(mappings):
                iid = f"mapping-{index}"
                self.mapping_rows[iid] = mapping
                self.mapping_tree.insert(
                    "", "end", iid=iid,
                    values=(mapping.relative_path, mapping.left_path or "—", mapping.right_path or "—", _status_text(mapping.status)),
                    tags=("ready",) if mapping.complete else ("blocked",),
                )
            self.mapping_tree.tag_configure("ready", foreground="#176B3A")
            self.mapping_tree.tag_configure("blocked", foreground="#A61B1B")
            self.status_var.set(f"已发现 {len(mappings)} 个相对路径；请选择一行完整匹配后打开。")
            self._selection_changed()
            return
        pair = validate_file_pair(left, right)
        if not pair.ready:
            self.status_var.set(f"{_status_text(pair.status)}：{pair.reason}")
            return
        try:
            self.snapshot = snapshot_selection(left, right)
        except PathSelectionError as exc:
            self.status_var.set(str(exc))
            return
        self.status_var.set(f"{_status_text(pair.status)}：哈希已记录；确认前会再次核对。")
        self.confirm_button.state(["!disabled"])

    def _selection_changed(self):
        selected = self.mapping_tree.selection()
        mapping = self.mapping_rows.get(selected[0]) if selected else None
        if mapping and mapping.complete and mapping.left_path and mapping.right_path:
            try:
                self.snapshot = snapshot_selection(mapping.left_path, mapping.right_path, relative_path=mapping.relative_path)
                self.confirm_button.state(["!disabled"])
                self.status_var.set(f"已选择 {mapping.relative_path}：{_status_text(mapping.status)}；确认前会再次核对哈希。")
            except PathSelectionError as exc:
                self.snapshot = None
                self.status_var.set(str(exc))
                self.confirm_button.state(["disabled"])
        else:
            self.snapshot = None
            self.confirm_button.state(["disabled"])

    def _confirm(self):
        if self.snapshot is None:
            self.status_var.set("请先选择状态为“匹配”的文件对。")
            return
        try:
            current = recheck_selection(self.snapshot)
        except PathSelectionError as exc:
            self.status_var.set(str(exc))
            return
        self.result = StartCenterResult("compare", (current.left.path, current.right.path), current.mapping)
        if callable(self.on_confirm):
            self.on_confirm(self.result)
        self.win.destroy()

    def _cancel(self):
        self.result = StartCenterResult(None)
        if callable(self.on_cancel):
            self.on_cancel()
        try:
            self.win.destroy()
        except self.tk.TclError:
            pass


class StartCenter:
    def __init__(self, *, initial_paths: Iterable[str] = (), result: dict | None = None):
        import tkinter as tk
        from tkinter import ttk

        self.tk, self.ttk = tk, ttk
        self.initial_paths = compare_selection_from_paths(initial_paths)
        self.result_box = result if result is not None else {}
        self.root = tk.Tk()
        self.root.title("sow_merge_tool · Excel 工作台")
        self.root.geometry("640x400")
        self.root.minsize(560, 340)
        self.root.protocol("WM_DELETE_WINDOW", self._cancel)
        self._build()

    def _build(self):
        ttk = self.ttk
        outer = ttk.Frame(self.root, padding=24)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="Excel 工作台", font=("Segoe UI", 20, "bold")).pack(anchor="w")
        ttk.Label(
            outer,
            text="选择比较/合并文件，或进入多分支 SVN 提交。右键传入的路径只会预填，不会自动修改工作簿。",
            wraplength=580,
        ).pack(anchor="w", pady=(5, 20))
        choices = ttk.LabelFrame(outer, text="开始一个安全操作", padding=16)
        choices.pack(fill="x")
        ttk.Button(
            choices,
            text="Excel 文件比较 / 合并",
            command=self._compare,
            width=30,
        ).pack(side="left", padx=(0, 12), ipady=6)
        ttk.Button(
            choices,
            text="多分支 SVN 提交",
            command=self._branch,
            width=24,
        ).pack(side="left", ipady=6)
        prefill = "\n".join(self.initial_paths) if self.initial_paths else "未从外部预填路径"
        box = ttk.LabelFrame(outer, text="当前预填路径", padding=10)
        box.pack(fill="both", expand=True, pady=(16, 0))
        text = self.tk.Text(box, height=5, wrap="none", state="normal")
        text.insert("1.0", prefill)
        text.configure(state="disabled")
        text.pack(fill="both", expand=True)
        ttk.Label(
            outer,
            text="比较前会检查文件类型、同一文件、相对路径、SVN 仓库边界并复核哈希。",
        ).pack(anchor="w", pady=(10, 0))

    def _compare(self):
        pending: dict[str, StartCenterResult] = {}
        dialog = CompareSelectionDialog(
            self.root,
            initial_paths=self.initial_paths,
            on_confirm=lambda result: pending.__setitem__("value", result),
        )
        self.root.wait_window(dialog.win)
        if pending.get("value") is not None:
            self._finish(pending["value"])

    def _branch(self):
        self._finish(StartCenterResult("branch", self.initial_paths))

    def _finish(self, result: StartCenterResult):
        self.result_box["value"] = result
        try:
            self.root.destroy()
        except self.tk.TclError:
            pass

    def _cancel(self):
        self._finish(StartCenterResult(None))


def launch_start_center(
    initial_paths: Iterable[str] = (),
    *,
    run_mainloop: bool = True,
) -> StartCenterResult:
    """Run the no-argument/prefilled center and return the chosen mode."""
    result: dict = {}
    center = StartCenter(initial_paths=initial_paths, result=result)
    if run_mainloop:
        center.root.mainloop()
    return result.get("value", StartCenterResult(None))


__all__ = ["CompareSelectionDialog", "StartCenter", "StartCenterResult", "launch_start_center"]
