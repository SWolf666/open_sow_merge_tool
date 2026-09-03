"""Shared Windows-style Tk UI primitives.

The application has two substantial Tk surfaces.  This module intentionally
contains only reusable chrome, scheduling and tracing helpers; domain logic
stays in the Excel and SVN modules.
"""

from __future__ import annotations

import queue
import threading
import time
import difflib
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from enum import Enum
from typing import Any


@dataclass(frozen=True)
class UiTheme:
    """Small set of tokens shared by both windows."""

    window_bg: str = "#F5F6F7"
    panel_bg: str = "#FFFFFF"
    border: str = "#D3D6DA"
    text: str = "#202124"
    secondary_text: str = "#5F6368"
    accent: str = "#0F6CBD"
    accent_active: str = "#0B5CAD"
    success: str = "#107C10"
    warning: str = "#986F0B"
    error: str = "#C42B1C"
    disabled: str = "#8A8F98"
    row_alt: str = "#F8F9FA"
    font_family: str = "Segoe UI"
    fallback_font_family: str = "Microsoft YaHei UI"


THEME = UiTheme()


class DifferenceKind(str, Enum):
    """Stable, presentation-independent difference categories."""

    CONFLICT = "conflict"
    MODIFIED = "modified"
    ADDED = "added"
    DELETED = "deleted"
    STRUCTURE = "structure"
    PROCESSED = "processed"

    @property
    def label(self) -> str:
        return {
            self.CONFLICT: "冲突",
            self.MODIFIED: "修改",
            self.ADDED: "增行",
            self.DELETED: "删行",
            self.STRUCTURE: "结构变化",
            self.PROCESSED: "已处理",
        }[self]


@dataclass(frozen=True)
class DifferenceItem:
    """Read-only row consumed by a difference browser.

    The model deliberately carries no workbook objects.  It can therefore be
    indexed, filtered and rendered on the UI thread without causing another
    workbook scan or keeping a mutable worksheet alive.
    """

    id: str
    sheet: str
    kind: DifferenceKind | str
    location: str = ""
    row: int | None = None
    column: int | None = None
    summary: str = ""
    role: str = ""
    processed: bool = False
    conflict: bool = False
    base_value: object = None
    mine_value: object = None
    theirs_value: object = None
    source_side: str = "mine"
    target_side: str = "theirs"
    base_label: str = "Base"
    mine_label: str = "Mine"
    theirs_label: str = "Theirs"
    payload: tuple[tuple[str, Any], ...] = ()

    def __post_init__(self) -> None:
        raw_kind = str(self.kind).lower()
        aliases = {"insert": "added", "add": "added", "delete": "deleted", "remove": "deleted", "column": "structure", "sheet": "structure"}
        kind = self.kind if isinstance(self.kind, DifferenceKind) else DifferenceKind(aliases.get(raw_kind, raw_kind))
        object.__setattr__(self, "kind", kind)

    @property
    def kind_label(self) -> str:
        return self.kind.label

    @property
    def display_location(self) -> str:
        if self.location:
            return self.location
        if self.row is None:
            return ""
        return f"第 {self.row} 行" if self.column is None else f"第 {self.row} 行 · 第 {self.column} 列"

    @property
    def status_label(self) -> str:
        return "已处理" if self.processed else ("待处理" if self.conflict or self.kind != DifferenceKind.PROCESSED else "已处理")

    @property
    def character_diff(self) -> tuple[tuple[str, str], ...]:
        """Character-level opcodes for the current values, suitable for Tk tags."""
        values = {"base": self.base_value, "mine": self.mine_value, "theirs": self.theirs_value}
        left = str(values.get(self.source_side) if values.get(self.source_side) is not None else "")
        right = str(values.get(self.target_side) if values.get(self.target_side) is not None else "")
        return tuple((tag, "".join(left[i1:i2]) + "→" + "".join(right[j1:j2])) for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, left, right).get_opcodes() if tag != "equal")


@dataclass(frozen=True)
class CommandState:
    """Centralized command availability for both workbenches."""

    busy: bool = False
    ready: bool = True
    has_selection: bool = False
    has_unsaved_changes: bool = False
    has_conflicts: bool = False
    can_save: bool = False
    can_undo: bool = False

    def enabled(self, command: str) -> bool:
        name = str(command).lower().replace(" ", "_")
        if self.busy:
            return name in {"cancel", "stop", "help"}
        rules = {
            "save": self.can_save and self.has_unsaved_changes,
            "save_as": self.can_save,
            "save_merged": self.can_save,
            "undo": self.can_undo,
            "apply": self.ready and self.has_selection and not self.has_conflicts,
            "apply_source": self.ready and self.has_selection,
            "retain": self.ready and self.has_selection,
            "base": self.ready and self.has_selection,
            "adopt": self.ready and self.has_selection,
            "compare": self.ready,
            "navigate": self.ready,
            "previous": self.ready,
            "next": self.ready,
            "search": self.ready,
            "filter": self.ready,
            "shortcut": not self.busy,
        }
        return bool(rules.get(name, self.ready))


def compact_path(path: str | None, limit: int = 72) -> str:
    """Keep role and filename visible while exposing the full path elsewhere."""

    value = str(path or "")
    if len(value) <= limit:
        return value
    head = max(12, limit // 2 - 2)
    tail = max(16, limit - head - 3)
    return f"{value[:head]}…{value[-tail:]}"


def set_widget_busy(widget: Any, busy: bool) -> None:
    """Best-effort shared busy-state helper for ttk and classic Tk widgets."""

    try:
        if hasattr(widget, "state"):
            widget.state(["disabled" if busy else "!disabled"])
        else:
            widget.configure(state="disabled" if busy else "normal")
    except Exception:
        pass


def configure_ttk_style(root, *, theme: UiTheme = THEME):
    """Apply the common style without assuming a particular Tk theme."""

    from tkinter import TclError

    style = root.tk.call("ttk::style", "theme", "use")
    # Vista/Windows themes ignore a few background options.  Configure them
    # anyway so the same code remains usable on a developer workstation with
    # clam or a minimal Tk installation.
    try:
        root.tk.call("ttk::style", "theme", "use", "vista")
    except TclError:
        pass
    style_obj = root.tk.call("ttk::style", "theme", "use") or style
    del style_obj
    ttk = root.ttk if hasattr(root, "ttk") else None
    if ttk is None:
        # Tk roots do not expose ttk; import lazily to keep the module cheap.
        from tkinter import ttk as ttk_module

        ttk = ttk_module
    style = ttk.Style(root)
    style.configure("App.TFrame", background=theme.window_bg)
    style.configure("Panel.TFrame", background=theme.panel_bg)
    style.configure("App.TLabel", background=theme.window_bg, foreground=theme.text, font=(theme.font_family, 9))
    style.configure("Muted.App.TLabel", background=theme.window_bg, foreground=theme.secondary_text, font=(theme.font_family, 9))
    style.configure("Title.App.TLabel", background=theme.window_bg, foreground=theme.text, font=(theme.font_family, 10, "bold"))
    style.configure("Primary.TButton", padding=(12, 5), font=(theme.font_family, 9, "bold"))
    style.configure("App.TButton", padding=(10, 5), font=(theme.font_family, 9))
    style.configure("Treeview", rowheight=25, font=(theme.font_family, 9), background=theme.panel_bg, fieldbackground=theme.panel_bg, foreground=theme.text)
    style.configure("Treeview.Heading", font=(theme.font_family, 9, "bold"))
    style.map("Treeview", background=[("selected", "#DCEBFA")], foreground=[("selected", theme.text)])
    style.configure("App.Horizontal.TProgressbar", troughcolor="#E5E7EA", background=theme.accent)
    return style


@dataclass
class UiTrace:
    """In-memory startup trace that can be emitted without leaking file data."""

    started_at: float = field(default_factory=time.perf_counter)
    marks: list[tuple[str, float]] = field(default_factory=list)

    def mark(self, name: str) -> None:
        self.marks.append((str(name), time.perf_counter()))

    def durations(self) -> dict[str, float]:
        previous = self.started_at
        result: dict[str, float] = {}
        for name, stamp in self.marks:
            result[name] = max(0.0, stamp - previous)
            previous = stamp
        return result


class UiTaskRunner:
    """Run cancellable work off the Tk thread and marshal results safely."""

    def __init__(self, root, *, poll_ms: int = 40):
        self.root = root
        self.poll_ms = max(20, int(poll_ms))
        self._queue: queue.Queue = queue.Queue()
        self._generation = 0
        self._closed = False
        self._polling = False
        self._cancel_events: dict[int, threading.Event] = {}
        self._callbacks: dict[int, Callable[[object, Exception | None, int], None]] = {}

    def close(self) -> None:
        self._closed = True
        for event in self._cancel_events.values():
            event.set()
        self._callbacks.clear()

    def cancel(self, generation: int | None = None) -> None:
        if generation is None:
            generation = self._generation
        event = self._cancel_events.get(generation)
        if event is not None:
            event.set()

    def submit(
        self,
        worker: Callable[[threading.Event], object],
        on_done: Callable[[object, Exception | None, int], None],
    ) -> int:
        self._generation += 1
        generation = self._generation
        cancel_event = threading.Event()
        self._cancel_events[generation] = cancel_event
        self._callbacks[generation] = on_done

        def run() -> None:
            value = None
            error: Exception | None = None
            try:
                value = worker(cancel_event)
            except Exception as exc:  # noqa: BLE001  # worker failures are returned to Tk
                error = exc
            self._queue.put((generation, value, error))

        threading.Thread(target=run, name=f"sow-ui-task-{generation}", daemon=True).start()
        self._ensure_polling()
        return generation

    def _ensure_polling(self) -> None:
        if self._polling or self._closed:
            return
        self._polling = True
        self.root.after(self.poll_ms, self._poll)

    def _poll(self) -> None:
        if self._closed:
            self._polling = False
            return
        try:
            while True:
                generation, value, error = self._queue.get_nowait()
                self._cancel_events.pop(generation, None)
                callback = self._callbacks.pop(generation, None)
                if generation == self._generation and callback is not None:
                    # All callbacks are invoked on Tk's event thread.
                    callback(value, error, generation)
        except queue.Empty:
            pass
        self.root.after(self.poll_ms, self._poll)

def batched(values: Iterable[object], size: int) -> Iterable[list[object]]:
    """Yield small lists so a Tk render can return to its event loop."""

    batch: list[object] = []
    for value in values:
        batch.append(value)
        if len(batch) >= size:
            yield batch
            batch = []
    if batch:
        yield batch
