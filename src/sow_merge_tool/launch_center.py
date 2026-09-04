"""The no-argument start centre and persistent Excel comparison sessions.

Explorer/TortoiseSVN launches only prefill a path.  The start centre keeps a
pairing list alive while a separately launched comparison window is open, so
the user can review another row after the child closes.  Directory indexing
is deliberately path-only and is performed in a worker; hashes and workbook
parsing happen only after an explicit row-open action.
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import time
from collections.abc import Callable, Iterable
from dataclasses import dataclass

from .path_selection import (
    PairStatus,
    PathSelectionError,
    RelativeMapping,
    SelectionSnapshot,
    compare_selection_from_paths,
    map_relative_files,
    override_mapping,
    recheck_selection,
    snapshot_selection,
    validate_directory,
    validate_excel_package,
    validate_file_pair,
)

COMPARISON_STATES = ("未打开", "启动中", "比较中", "已关闭", "启动失败")


@dataclass(frozen=True)
class StartCenterResult:
    mode: str | None
    paths: tuple[str, ...] = ()
    mapping: RelativeMapping | None = None


@dataclass(frozen=True)
class ComparisonLaunch:
    """A direct child command; paths are never routed through ``--compare``."""

    left_path: str
    right_path: str
    command: tuple[str, ...]

    @classmethod
    def for_paths(cls, left_path: str, right_path: str) -> ComparisonLaunch:
        command: list[str] = [sys.executable]
        if not getattr(sys, "frozen", False):
            command.extend(("-m", "sow_merge_tool"))
        # ``--source-target`` gives the child its explicit read-only Source /
        # writable Target role without recursively opening the picker.
        command.extend(("--source-target", os.fspath(left_path), os.fspath(right_path)))
        return cls(os.fspath(left_path), os.fspath(right_path), tuple(command))

    build_command = for_paths


@dataclass
class ComparisonSession:
    session_id: str
    launch: ComparisonLaunch
    state: str = "未打开"
    returncode: int | None = None
    error: str = ""
    process: object | None = None
    started_at: float = 0.0
    closed_at: float = 0.0

    @property
    def active(self) -> bool:
        process = self.process
        try:
            return bool(process is not None and process.poll() is None)
        except (AttributeError, OSError):
            return False

    @property
    def status(self) -> str:
        return self.state


class ComparisonSessionManager:
    """Own one child comparison process and poll it without blocking Tk."""

    def __init__(
        self,
        root=None,
        *,
        popen_factory: Callable | None = None,
        command_builder: Callable[[str, str], ComparisonLaunch] | None = None,
        poll_ms: int = 80,
    ):
        self.root = root
        self.popen_factory = popen_factory or subprocess.Popen
        self.command_builder = command_builder or ComparisonLaunch.for_paths
        self.poll_ms = max(20, int(poll_ms))
        self._sequence = 0
        self._active_session: ComparisonSession | None = None
        self._last_session: ComparisonSession | None = None
        self._callback: Callable[[ComparisonSession], None] | None = None
        self._after_id = None
        self._closed = False

    @property
    def active_session(self) -> ComparisonSession | None:
        session = self._active_session
        return session if session is not None and session.active else None

    @property
    def active(self) -> bool:
        return self.active_session is not None

    @property
    def last_session(self) -> ComparisonSession | None:
        return self._last_session

    def _emit(self, session: ComparisonSession) -> None:
        callback = self._callback
        if callback is not None:
            try:
                callback(session)
            except (AttributeError, RuntimeError, TypeError):
                pass

    def _spawn(self, command: list[str]):
        kwargs = {"close_fds": True}
        if os.name == "nt":
            kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        try:
            return self.popen_factory(command, **kwargs)
        except TypeError:
            # A tiny injected fake often accepts only ``command``.
            return self.popen_factory(command)

    def start(
        self,
        left_path: str,
        right_path: str,
        *,
        callback: Callable[[ComparisonSession], None] | None = None,
    ) -> ComparisonSession | None:
        """Start one child; reject a second live child instead of duplicating it."""
        active = self.active_session
        if active is not None:
            return None
        self._closed = False
        self._callback = callback
        self._sequence += 1
        launch = self.command_builder(os.fspath(left_path), os.fspath(right_path))
        session = ComparisonSession(
            f"compare-{self._sequence}", launch, state="启动中", started_at=time.monotonic()
        )
        self._active_session = session
        self._last_session = session
        self._emit(session)
        try:
            session.process = self._spawn(list(launch.command))
        except (OSError, RuntimeError, TypeError, ValueError, subprocess.SubprocessError) as exc:
            session.state = "启动失败"
            session.error = f"{type(exc).__name__}: {exc}"
            session.closed_at = time.monotonic()
            self._active_session = None
            self._emit(session)
            return session
        session.state = "比较中"
        self._emit(session)
        self._schedule_poll()
        return session

    launch = start

    def _schedule_poll(self) -> None:
        if self._closed or self.root is None or self._after_id is not None:
            return
        try:
            self._after_id = self.root.after(self.poll_ms, self._poll)
        except (AttributeError, RuntimeError):
            self._after_id = None

    def _poll(self) -> None:
        self._after_id = None
        session = self._active_session
        if session is None:
            return
        try:
            code = session.process.poll()
        except (AttributeError, OSError) as exc:
            code = -1
            session.error = f"比较进程状态读取失败：{exc}"
        if code is None:
            self._schedule_poll()
            return
        session.returncode = int(code)
        session.state = "已关闭"
        session.closed_at = time.monotonic()
        self._active_session = None
        self._emit(session)

    def poll_now(self) -> ComparisonSession | None:
        """Synchronous test hook; it never waits for the process."""
        self._poll()
        return self._last_session

    def close(self, *, terminate: bool = False) -> None:
        """Stop monitoring; by default leave the child alive when the list closes."""
        self._closed = True
        if self.root is not None and self._after_id is not None:
            try:
                self.root.after_cancel(self._after_id)
            except (AttributeError, RuntimeError):
                pass
        self._after_id = None
        if terminate and self._active_session is not None:
            try:
                self._active_session.process.terminate()
            except (AttributeError, OSError):
                pass

    def resume(self, root=None) -> None:
        """Resume polling an existing child when the pairing list is reopened."""
        if root is not None:
            self.root = root
        self._closed = False
        if self._active_session is not None:
            self._schedule_poll()


def _status_text(status: PairStatus) -> str:
    return {
        PairStatus.MATCHED: "匹配",
        PairStatus.SAME_NAME_DIFFERENT_PATH: "同名不同路径（已明确选择）",
        PairStatus.MISSING_LEFT: "左侧缺失（可指定左文件）",
        PairStatus.MISSING_RIGHT: "右侧缺失（可指定右文件）",
        PairStatus.SAME_FILE: "同一文件（阻断）",
        PairStatus.TYPE_INCOMPATIBLE: "类型不兼容（阻断）",
        PairStatus.HASH_CHANGED: "内容已变化（请重新选择）",
        PairStatus.UNSUPPORTED: "类型不支持（阻断）",
        PairStatus.INVALID: "无效（阻断）",
    }.get(status, str(status))


class CompareSelectionDialog:
    """Persistent pairing list that can launch exactly one comparison child."""

    def __init__(
        self,
        parent,
        *,
        initial_paths: Iterable[str] = (),
        on_confirm=None,
        on_cancel=None,
        manager: ComparisonSessionManager | None = None,
    ):
        import tkinter as tk
        from tkinter import ttk

        self.tk, self.ttk = tk, ttk
        self.parent = parent
        self.on_confirm = on_confirm
        self.on_cancel = on_cancel
        self.result: StartCenterResult | None = None
        self.snapshot: SelectionSnapshot | None = None
        self.mapping_rows: dict[str, RelativeMapping] = {}
        self._row_sessions: dict[str, str] = {}
        self._sessions: dict[str, ComparisonSession] = {}
        self._manual_overrides: dict[str, tuple[str | None, str | None]] = {}
        self._mapping_generation = 0
        self._mapping_cancel = threading.Event()
        self._mapping_polling = False
        self._mapping_queue: queue.Queue = queue.Queue()
        self._mapping_after = None
        self._directory_svn_label = "待确认"
        self._mapping_tooltip = None
        self._mapping_tooltip_after = None
        self.manager = manager or ComparisonSessionManager(parent)
        self._owns_manager = manager is None
        self.left_var = tk.StringVar()
        self.right_var = tk.StringVar()
        self.status_var = tk.StringVar(value="请选择左右两侧的 Excel 文件或文件夹。")
        values = list(compare_selection_from_paths(initial_paths))
        if values:
            self.left_var.set(values[0])
        if len(values) > 1:
            self.right_var.set(values[1])

        self.win = tk.Toplevel(parent)
        self.win.title("Excel 文件比较 / 合并 · 配对列表")
        self.win.geometry("1080x650")
        self.win.minsize(820, 540)
        try:
            if parent.winfo_viewable():
                self.win.transient(parent)
        except (AttributeError, self.tk.TclError):
            pass
        self._build()
        self.win.protocol("WM_DELETE_WINDOW", self._cancel)
        self.win.deiconify()
        self.win.lift()
        if values:
            self._validate()

    def _build(self):
        ttk = self.ttk
        outer = ttk.Frame(self.win, padding=14)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="Excel 比较配对列表", font=("Segoe UI", 14, "bold")).pack(anchor="w")
        ttk.Label(
            outer,
            text="Source 左侧只读，Target 右侧可保存。文件夹只建立相对路径索引；请双击或明确点击“打开比较”启动一个比较窗口。",
            wraplength=1000,
        ).pack(anchor="w", pady=(3, 10))

        rows = ttk.LabelFrame(outer, text="比较输入", padding=8)
        rows.pack(fill="x")
        for row, (label, variable) in enumerate((("Source（左侧）", self.left_var), ("Target（右侧）", self.right_var))):
            ttk.Label(rows, text=label, width=16).grid(row=row, column=0, sticky="w", pady=3)
            entry = ttk.Entry(rows, textvariable=variable)
            entry.grid(row=row, column=1, sticky="ew", padx=5, pady=3)
            setattr(self, "left_entry" if row == 0 else "right_entry", entry)
            ttk.Button(rows, text="选择文件…", command=lambda v=variable: self._choose_file(v)).grid(row=row, column=2, pady=3)
            ttk.Button(rows, text="选择文件夹…", command=lambda v=variable: self._choose_directory(v)).grid(row=row, column=3, padx=(5, 0), pady=3)
        rows.columnconfigure(1, weight=1)
        self.left_var.trace_add("write", lambda *_args: self._schedule_validate())
        self.right_var.trace_add("write", lambda *_args: self._schedule_validate())

        ttk.Label(outer, textvariable=self.status_var, wraplength=1000).pack(anchor="w", pady=(8, 4))
        mapping_box = ttk.LabelFrame(outer, text="Source / Target 映射（长路径可横向滚动、悬停和复制）", padding=7)
        mapping_box.pack(fill="both", expand=True)
        holder = ttk.Frame(mapping_box)
        holder.pack(fill="both", expand=True)
        columns = ("relative", "source", "target", "status", "svn", "session")
        self.mapping_tree = ttk.Treeview(holder, columns=columns, show="headings", selectmode="browse")
        for key, title, width in (
            ("relative", "相对路径", 210),
            ("source", "Source 完整路径", 300),
            ("target", "Target 完整路径", 300),
            ("status", "映射状态", 180),
            ("svn", "SVN 证据", 100),
            ("session", "比较会话", 100),
        ):
            self.mapping_tree.heading(key, text=title)
            self.mapping_tree.column(key, width=width, minwidth=60, anchor="w", stretch=key in {"source", "target", "relative"})
        yscroll = ttk.Scrollbar(holder, orient="vertical", command=self.mapping_tree.yview)
        xscroll = ttk.Scrollbar(holder, orient="horizontal", command=self.mapping_tree.xview)
        self.mapping_tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.mapping_tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        holder.columnconfigure(0, weight=1)
        holder.rowconfigure(0, weight=1)
        self.mapping_tree.tag_configure("ready", foreground="#176B3A")
        self.mapping_tree.tag_configure("blocked", foreground="#A61B1B")
        self.mapping_tree.tag_configure("running", foreground="#174A7E")
        self.mapping_tree.bind("<<TreeviewSelect>>", lambda _event: self._selection_changed())
        self.mapping_tree.bind("<Double-1>", lambda _event: self._open_selected())
        self.mapping_tree.bind("<Button-3>", self._row_menu)
        self.mapping_tree.bind("<Motion>", self._mapping_hover)
        self.mapping_tree.bind("<Leave>", lambda _event: self._hide_mapping_tooltip())

        tools = ttk.Frame(outer)
        tools.pack(fill="x", pady=(8, 0))
        self.refresh_button = ttk.Button(tools, text="刷新映射", command=self._validate)
        self.refresh_button.pack(side="left")
        self.cancel_scan_button = ttk.Button(tools, text="取消扫描", command=self._cancel_mapping)
        self.cancel_scan_button.pack(side="left", padx=(5, 0))
        self.choose_left_button = ttk.Button(tools, text="指定左文件…", command=lambda: self._choose_override("left"))
        self.choose_left_button.pack(side="left", padx=(18, 0))
        self.choose_right_button = ttk.Button(tools, text="指定右文件…", command=lambda: self._choose_override("right"))
        self.choose_right_button.pack(side="left", padx=(5, 0))
        self.session_status_var = self.tk.StringVar(value="当前没有打开的比较会话")
        ttk.Label(tools, textvariable=self.session_status_var).pack(side="left", padx=(15, 0))

        footer = ttk.Frame(outer)
        footer.pack(fill="x", pady=(8, 0))
        self.confirm_button = ttk.Button(footer, text="打开比较", command=self._open_selected)
        self.confirm_button.pack(side="right")
        ttk.Button(footer, text="关闭列表", command=self._cancel).pack(side="right", padx=(0, 8))

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
            if self._mapping_after:
                self.win.after_cancel(self._mapping_after)
            self._mapping_after = self.win.after(120, self._validate)
        except (AttributeError, self.tk.TclError):
            self._mapping_after = None

    def _clear_mapping(self, *, clear_overrides: bool = True):
        self.mapping_rows.clear()
        if clear_overrides:
            self._manual_overrides.clear()
        for iid in self.mapping_tree.get_children():
            self.mapping_tree.delete(iid)
        self.snapshot = None
        self.confirm_button.state(["disabled"])

    def _set_scanning(self, active: bool):
        self.refresh_button.state(["disabled"] if active else ["!disabled"])
        self.cancel_scan_button.state(["!disabled"] if active else ["disabled"])
        self.confirm_button.state(["disabled"] if active else ["!disabled"])

    def _svn_evidence(self, left: str | None, right: str | None) -> str:
        if not left or not right:
            return "待选择"
        return self._directory_svn_label

    def _render_mappings(self, mappings: Iterable[RelativeMapping]):
        self._clear_mapping(clear_overrides=False)
        for index, mapping in enumerate(mappings):
            iid = f"mapping-{index}"
            override = self._manual_overrides.get(mapping.relative_path)
            if override is not None:
                mapping = override_mapping(mapping, left_path=override[0], right_path=override[1])
            self.mapping_rows[iid] = mapping
            session_text = self._session_text(iid)
            self.mapping_tree.insert(
                "", "end", iid=iid,
                values=(
                    mapping.relative_path,
                    mapping.left_path or "—",
                    mapping.right_path or "—",
                    _status_text(mapping.status),
                    self._svn_evidence(mapping.left_path, mapping.right_path),
                    session_text,
                ),
                tags=("ready",) if mapping.complete else ("blocked",),
            )
        if self.mapping_rows:
            first = next(iter(self.mapping_rows))
            self.mapping_tree.selection_set(first)
            self.mapping_tree.focus(first)
        self.status_var.set(f"已加载 {len(self.mapping_rows)} 个配对；目录扫描只读取路径，打开前才核对 SHA256。")
        self._selection_changed()

    def _start_directory_mapping(self, left: str, right: str):
        self._mapping_generation += 1
        generation = self._mapping_generation
        self._mapping_cancel.set()
        self._mapping_cancel = threading.Event()
        cancel_event = self._mapping_cancel
        self._set_scanning(True)
        self.status_var.set("正在后台索引相对路径（不解析工作簿、不计算哈希）…")

        def worker():
            try:
                left_validation = validate_directory(left)
                right_validation = validate_directory(
                    right, expected_repository=left_validation.repository
                )
                if not left_validation.comparison_allowed or not right_validation.comparison_allowed:
                    raise PathSelectionError(
                        "文件夹不可用："
                        + "; ".join(left_validation.issues + right_validation.issues)
                    )
                mappings = map_relative_files(left, right, include_hash=False)
                error = None
            except (OSError, PathSelectionError, RuntimeError, TypeError, ValueError) as exc:
                mappings, error = [], str(exc)
            svn_label = "普通目录"
            if error is None and left_validation.repository and right_validation.repository:
                svn_label = (
                    "同仓库"
                    if left_validation.repository.key == right_validation.repository.key
                    else "仓库不同"
                )
                if left_validation.warning or right_validation.warning:
                    svn_label += "（有提示）"
            elif left_validation.warning or right_validation.warning:
                svn_label = "普通目录（有提示）"
            self._mapping_queue.put(
                (generation, mappings, error, cancel_event.is_set(), svn_label)
            )

        threading.Thread(target=worker, name="sow-compare-path-index", daemon=True).start()
        if not self._mapping_polling:
            self._mapping_polling = True
            self.win.after(40, self._poll_mapping)

    def _poll_mapping(self):
        try:
            while True:
                generation, mappings, error, cancelled, svn_label = self._mapping_queue.get_nowait()
                if generation != self._mapping_generation:
                    continue
                self._mapping_polling = False
                self._set_scanning(False)
                if cancelled:
                    self.status_var.set("目录扫描已取消；旧映射未复用，请重新刷新。")
                elif error:
                    self.status_var.set(str(error))
                else:
                    self._directory_svn_label = svn_label
                    self._render_mappings(mappings)
                return
        except queue.Empty:
            pass
        try:
            if self.win.winfo_exists():
                self.win.after(40, self._poll_mapping)
            else:
                self._mapping_polling = False
        except self.tk.TclError:
            self._mapping_polling = False

    def _cancel_mapping(self):
        self._mapping_generation += 1
        self._mapping_cancel.set()
        self._mapping_polling = False
        self._set_scanning(False)
        self.status_var.set("目录扫描已取消；请点击刷新映射重新索引。")

    def _validate(self):
        self._mapping_generation += 1
        self._mapping_cancel.set()
        self._clear_mapping()
        left, right = self.left_var.get().strip(), self.right_var.get().strip()
        if not left or not right:
            self.status_var.set("请分别选择 Source 和 Target；外部路径仅作为预填。")
            return
        left_is_dir, right_is_dir = os.path.isdir(left), os.path.isdir(right)
        if left_is_dir != right_is_dir:
            self.status_var.set("两侧必须同时为文件，或同时为文件夹；类型不兼容。")
            return
        if left_is_dir:
            self._start_directory_mapping(left, right)
            return
        self._directory_svn_label = "普通目录"
        pair = validate_file_pair(left, right, include_hash=False)
        if not pair.ready:
            self.status_var.set(f"{_status_text(pair.status)}：{pair.reason}")
            return
        mapping = RelativeMapping("", pair.left.path, pair.right.path, pair.status, pair.reason)
        self._render_mappings((mapping,))

    def _selected_entry(self) -> tuple[str, RelativeMapping] | None:
        selection = self.mapping_tree.selection()
        if not selection:
            return None
        iid = selection[0]
        mapping = self.mapping_rows.get(iid)
        return (iid, mapping) if mapping is not None else None

    def _selection_changed(self):
        entry = self._selected_entry()
        if entry and entry[1].complete and not self.manager.active_session:
            self.confirm_button.state(["!disabled"])
        else:
            self.confirm_button.state(["disabled"])

    def _choose_override(self, side: str):
        entry = self._selected_entry()
        if not entry:
            self.status_var.set("请先选择需要人工指定的映射行。")
            return
        from tkinter import filedialog

        value = filedialog.askopenfilename(
            parent=self.win,
            title="指定 Source 文件" if side == "left" else "指定 Target 文件",
            filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not value:
            return
        iid, mapping = entry
        pair = (value, mapping.right_path) if side == "left" else (mapping.left_path, value)
        updated = override_mapping(mapping, left_path=pair[0], right_path=pair[1])
        self._manual_overrides[mapping.relative_path] = pair
        self.mapping_rows[iid] = updated
        self.mapping_tree.item(
            iid,
            values=(
                updated.relative_path,
                updated.left_path or "—",
                updated.right_path or "—",
                _status_text(updated.status),
                self._svn_evidence(updated.left_path, updated.right_path),
                self._session_text(iid),
            ),
            tags=("ready",) if updated.complete else ("blocked",),
        )
        self._selection_changed()

    def _row_menu(self, event):
        entry = self.mapping_tree.identify_row(event.y)
        if not entry:
            return "break"
        self.mapping_tree.selection_set(entry)
        menu = self.tk.Menu(self.win, tearoff=False)
        menu.add_command(label="指定左文件…", command=lambda: self._choose_override("left"))
        menu.add_command(label="指定右文件…", command=lambda: self._choose_override("right"))
        mapping = self.mapping_rows.get(entry)
        if mapping:
            menu.add_command(label="复制 Source 路径", command=lambda: self._copy(mapping.left_path))
            menu.add_command(label="复制 Target 路径", command=lambda: self._copy(mapping.right_path))
        menu.tk_popup(event.x_root, event.y_root)
        return "break"

    def _hide_mapping_tooltip(self):
        after_id = self._mapping_tooltip_after
        self._mapping_tooltip_after = None
        if after_id:
            try:
                self.win.after_cancel(after_id)
            except self.tk.TclError:
                pass
        tooltip = self._mapping_tooltip
        self._mapping_tooltip = None
        if tooltip is not None:
            try:
                tooltip.destroy()
            except self.tk.TclError:
                pass

    def _mapping_hover(self, event):
        self._hide_mapping_tooltip()
        iid = self.mapping_tree.identify_row(event.y)
        column = self.mapping_tree.identify_column(event.x)
        mapping = self.mapping_rows.get(iid)
        if not mapping or column not in {"#1", "#2", "#3", "#4", "#5", "#6"}:
            return
        values = {
            "#1": mapping.relative_path,
            "#2": mapping.left_path or "—",
            "#3": mapping.right_path or "—",
            "#4": _status_text(mapping.status),
            "#5": self._directory_svn_label,
            "#6": self._session_text(iid),
        }
        detail = values.get(column, "")
        if not detail or detail == "—":
            return
        try:
            self._mapping_tooltip_after = self.win.after(
                350,
                lambda: self._show_mapping_tooltip(detail, event.x_root, event.y_root),
            )
        except self.tk.TclError:
            self._mapping_tooltip_after = None

    def _show_mapping_tooltip(self, detail: str, x: int, y: int):
        self._mapping_tooltip_after = None
        try:
            tooltip = self.tk.Toplevel(self.win)
            tooltip.wm_overrideredirect(True)
            tooltip.attributes("-topmost", True)
            self.tk.Label(
                tooltip,
                text=detail,
                anchor="w",
                justify="left",
                wraplength=1000,
                bg="#FFFDE7",
                fg="#263238",
                bd=1,
                relief="solid",
                padx=7,
                pady=5,
            ).pack(fill="both", expand=True)
            tooltip.geometry(f"+{max(0, int(x) + 12)}+{max(0, int(y) + 12)}")
            self._mapping_tooltip = tooltip
        except self.tk.TclError:
            self._mapping_tooltip = None

    def _copy(self, value: str | None):
        if value:
            try:
                self.win.clipboard_clear()
                self.win.clipboard_append(value)
                self.win.update_idletasks()
            except self.tk.TclError:
                pass

    def _session_text(self, iid: str) -> str:
        session_id = self._row_sessions.get(iid)
        session = self._sessions.get(session_id) if session_id else None
        return session.state if session is not None else "未打开"

    def _update_session_row(self, session: ComparisonSession):
        self._sessions[session.session_id] = session
        iid = next((key for key, value in self._row_sessions.items() if value == session.session_id), None)
        if iid and self.mapping_tree.exists(iid):
            mapping = self.mapping_rows[iid]
            self.mapping_tree.set(iid, "session", session.state)
            self.mapping_tree.item(iid, tags=("running",) if session.state in {"启动中", "比较中"} else ("ready",) if mapping.complete else ("blocked",))
        self.session_status_var.set(
            f"比较会话：{session.state}"
            + (f"（{session.error}）" if session.error else "")
        )
        self.status_var.set(
            {"启动中": "比较正在启动；配对列表保持可见。", "比较中": "比较进行中；当前列表不会销毁。", "已关闭": "比较窗口已关闭；可继续选择其他配对。", "启动失败": "比较启动失败；请检查文件路径后重试。"}.get(session.state, self.status_var.get())
        )
        self._selection_changed()

    def _open_selected(self):
        if self.manager.active_session is not None:
            self.status_var.set("已有一个比较窗口正在运行；关闭后才能打开下一项。")
            return
        entry = self._selected_entry()
        if not entry:
            self.status_var.set("请先选择一行配对。")
            return
        iid, mapping = entry
        if not mapping.complete or not mapping.left_path or not mapping.right_path:
            self.status_var.set(f"{_status_text(mapping.status)}：请逐行指定缺失的一侧文件。")
            return
        try:
            snapshot = snapshot_selection(mapping.left_path, mapping.right_path, relative_path=mapping.relative_path)
            current = recheck_selection(snapshot)
            validate_excel_package(current.left.path)
            validate_excel_package(current.right.path)
        except PathSelectionError as exc:
            self.status_var.set(str(exc))
            return
        self.snapshot = current
        session = self.manager.start(
            current.left.path,
            current.right.path,
            callback=self._update_session_row,
        )
        if session is None:
            self.status_var.set("已有一个比较窗口正在运行；请先关闭它。")
            return
        self._row_sessions[iid] = session.session_id
        self._sessions[session.session_id] = session
        self.mapping_tree.set(iid, "session", session.state)
        self._selection_changed()
        self.result = StartCenterResult("compare", (current.left.path, current.right.path), current.mapping)
        if callable(self.on_confirm):
            self.on_confirm(self.result)
        self.status_var.set("比较已启动；列表仍保留，比较窗口关闭后可继续选择。")

    def _cancel(self):
        self._cancel_mapping()
        self._hide_mapping_tooltip()
        # Parent list closure never kills a child comparison.
        self.manager.close(terminate=False)
        self.result = self.result or StartCenterResult(None)
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
        self.compare_dialog: CompareSelectionDialog | None = None
        self.root = tk.Tk()
        self.session_manager = ComparisonSessionManager(self.root)
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
        ttk.Button(choices, text="Excel 文件比较 / 合并", command=self._compare, width=30).pack(side="left", padx=(0, 12), ipady=6)
        ttk.Button(choices, text="多分支 SVN 提交", command=self._branch, width=24).pack(side="left", ipady=6)
        prefill = "\n".join(self.initial_paths) if self.initial_paths else "未从外部预填路径"
        box = ttk.LabelFrame(outer, text="当前预填路径", padding=10)
        box.pack(fill="both", expand=True, pady=(16, 0))
        text = self.tk.Text(box, height=5, wrap="none", state="normal")
        text.insert("1.0", prefill)
        text.configure(state="disabled")
        text.pack(fill="both", expand=True)
        ttk.Label(outer, text="比较前会检查文件类型、同一文件、相对路径、SVN 仓库边界并复核哈希。").pack(anchor="w", pady=(10, 0))

    def _compare(self):
        existing = self.compare_dialog
        if existing is not None:
            try:
                if existing.win.winfo_exists():
                    existing.win.deiconify()
                    existing.win.lift()
                    return
            except self.tk.TclError:
                pass
        self.root.withdraw()
        self.session_manager.resume(self.root)
        self.compare_dialog = CompareSelectionDialog(
            self.root,
            initial_paths=self.initial_paths,
            on_cancel=self._restore_center,
            manager=self.session_manager,
        )

    def _restore_center(self):
        self.compare_dialog = None
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        except self.tk.TclError:
            pass

    def _branch(self):
        self._finish(StartCenterResult("branch", self.initial_paths))

    def _finish(self, result: StartCenterResult):
        self.result_box["value"] = result
        try:
            self.root.destroy()
        except self.tk.TclError:
            pass

    def _cancel(self):
        dialog = self.compare_dialog
        if dialog is not None:
            dialog._cancel()
            return
        self.session_manager.close(terminate=False)
        self._finish(StartCenterResult(None))


def launch_start_center(
    initial_paths: Iterable[str] = (),
    *,
    run_mainloop: bool = True,
    open_compare: bool = False,
) -> StartCenterResult:
    """Run the no-argument/prefilled centre until the user closes it."""
    result: dict = {}
    center = StartCenter(initial_paths=initial_paths, result=result)
    if open_compare:
        center._compare()
    if run_mainloop:
        center.root.mainloop()
    return result.get("value", StartCenterResult(None))


__all__ = [
    "COMPARISON_STATES",
    "CompareSelectionDialog",
    "ComparisonLaunch",
    "ComparisonSession",
    "ComparisonSessionManager",
    "StartCenter",
    "StartCenterResult",
    "launch_start_center",
]
