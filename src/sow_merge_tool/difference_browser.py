"""Compact, cache-backed difference browser used by the Excel workspace."""

from __future__ import annotations

from collections import Counter
from dataclasses import replace
from collections.abc import Callable, Iterable
import tkinter as tk
from tkinter import ttk

from .ui_foundation import DifferenceItem, DifferenceKind, THEME, configure_ttk_style


class DifferenceIndex:
    """Stable index with incremental processed-state updates and cheap filters."""

    def __init__(self, items: Iterable[DifferenceItem] = ()) -> None:
        self._items: dict[str, DifferenceItem] = {}
        self.replace(items)

    def replace(self, items: Iterable[DifferenceItem]) -> None:
        self._items = {item.id: item for item in items}

    def update(self, item_id: str, *, processed: bool = True) -> None:
        item = self._items.get(str(item_id))
        if item is not None:
            self._items[item.id] = replace(item, processed=processed)

    def all(self) -> tuple[DifferenceItem, ...]:
        return tuple(self._items.values())

    def counts(self) -> Counter:
        counts = Counter(item.kind for item in self._items.values())
        counts[DifferenceKind.PROCESSED] = sum(item.processed for item in self._items.values())
        counts["total"] = len(self._items)
        return counts

    def filter(self, query: str = "", *, only_unprocessed: bool = False, only_conflicts: bool = False) -> list[DifferenceItem]:
        needle = str(query or "").strip().casefold()
        result = []
        for item in self._items.values():
            if only_unprocessed and item.processed:
                continue
            if only_conflicts and not item.conflict:
                continue
            haystack = f"{item.sheet} {item.kind_label} {item.display_location} {item.summary} {item.role}".casefold()
            if needle and needle not in haystack:
                continue
            result.append(item)
        return result


class DifferenceBrowser:
    """Tk dock: navigation list only; mutations remain explicit toolbar actions."""

    def __init__(self, parent, *, on_select: Callable[[DifferenceItem], None] | None = None, height: int = 150, on_resize: Callable[[int, bool], None] | None = None):
        self.parent = parent
        self.on_select = on_select
        self.on_resize = on_resize
        self.index = DifferenceIndex()
        self.collapsed = False
        self.frame = ttk.Frame(parent, style="Difference.Dock.TFrame")
        style = configure_ttk_style(parent.winfo_toplevel())
        style.configure("Difference.Dock.TFrame", background=THEME.panel_bg, relief="solid", borderwidth=1)
        style.configure("Difference.Header.TFrame", background="#F3F6F9")
        style.configure("Difference.Header.TLabel", background="#F3F6F9", foreground=THEME.text, font=(THEME.font_family, 9, "bold"))
        style.configure("Difference.Muted.TLabel", background="#F3F6F9", foreground=THEME.secondary_text)
        self._height = max(80, int(height))
        self.frame.configure(height=self._height)
        self.frame.pack_propagate(False)
        header = ttk.Frame(self.frame, style="Difference.Header.TFrame", padding=(8, 5))
        header.pack(fill="x")
        ttk.Label(header, text="差异浏览器", style="Difference.Header.TLabel").pack(side="left")
        self.summary_var = tk.StringVar(value="全部 0")
        ttk.Label(header, textvariable=self.summary_var, style="Difference.Muted.TLabel").pack(side="left", padx=(12, 0))
        self.toggle_button = ttk.Button(header, text="收起", width=6, command=self.toggle)
        self.toggle_button.pack(side="right")
        self.grip = ttk.Separator(self.frame, orient="horizontal")
        self.grip.pack(fill="x", side="top")
        self.grip.bind("<B1-Motion>", self._resize)
        self.grip.bind("<Button-1>", lambda event: setattr(self, "_resize_origin", (event.y_root, self._height)))
        self.query_var = tk.StringVar()
        self._query_after = None
        entry = ttk.Entry(header, textvariable=self.query_var, width=22)
        self.query_entry = entry
        entry.pack(side="right", padx=(8, 0))
        entry.insert(0, "")
        entry.bind("<Escape>", lambda _e: self._clear_query())
        self.query_var.trace_add("write", lambda *_: self._schedule_render())
        ttk.Label(header, text="搜索", style="Difference.Muted.TLabel").pack(side="right")
        self.unprocessed_var = tk.BooleanVar(value=False)
        self.conflicts_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(header, text="只看未处理", variable=self.unprocessed_var, command=self.render).pack(side="left", padx=(16, 0))
        ttk.Checkbutton(header, text="只看冲突", variable=self.conflicts_var, command=self.render).pack(side="left", padx=(8, 0))
        body = ttk.Frame(self.frame, padding=(6, 0, 6, 6))
        body.pack(fill="both", expand=True)
        self.tabs = ttk.Notebook(body)
        self.list_tab = ttk.Frame(self.tabs)
        self.current_tab = ttk.Frame(self.tabs)
        self.tabs.add(self.list_tab, text="差异列表")
        self.tabs.add(self.current_tab, text="当前差异")
        self.tabs.pack(fill="both", expand=True)
        self.current_var = tk.StringVar(value="选择差异后显示完整值和处理方向。")
        ttk.Label(self.current_tab, textvariable=self.current_var, anchor="nw", justify="left", wraplength=900, padding=12).pack(fill="both", expand=True)
        self.tree = ttk.Treeview(self.list_tab, columns=("sheet", "kind", "location", "summary", "role", "status"), show="headings", height=4, selectmode="browse")
        headings = {"sheet": "Sheet", "kind": "类型", "location": "位置", "summary": "摘要", "role": "角色", "status": "处理状态"}
        widths = {"sheet": 110, "kind": 72, "location": 115, "summary": 340, "role": 110, "status": 72}
        for key, title in headings.items():
            self.tree.heading(key, text=title)
            self.tree.column(key, width=widths[key], anchor="w", stretch=key in {"summary", "sheet"})
        self.tree.tag_configure("processed", foreground=THEME.secondary_text)
        self.tree.tag_configure("conflict", foreground=THEME.error, background="#FFF3F3")
        scroll = ttk.Scrollbar(self.list_tab, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._selected)
        self.tree.bind("<Return>", self._activate)
        self.tree.bind("<Double-Button-1>", self._activate)
        self._rendered: dict[str, DifferenceItem] = {}
        self._page = 0
        self.page_var = tk.StringVar(value="第 1 页")
        pager = ttk.Frame(self.frame, padding=(6, 0, 6, 4))
        pager.pack(side="bottom", fill="x")
        ttk.Button(pager, text="上一页", width=7, command=self._previous_page).pack(side="right")
        ttk.Label(pager, textvariable=self.page_var, style="Difference.Muted.TLabel").pack(side="right", padx=6)
        ttk.Button(pager, text="下一页", width=7, command=self._next_page).pack(side="right")

    def _clear_query(self):
        self.query_var.set("")

    def toggle_unprocessed_filter(self):
        self.unprocessed_var.set(not self.unprocessed_var.get())
        self._page = 0
        self.render()

    def toggle(self):
        self.collapsed = not self.collapsed
        if self.collapsed:
            self.frame.configure(height=34)
            self.toggle_button.configure(text="展开")
        else:
            self.frame.configure(height=self._height)
            self.toggle_button.configure(text="收起")
        if self.on_resize:
            self.on_resize(self._height, self.collapsed)

    def set_layout(self, *, height: int | None = None, collapsed: bool | None = None, notify: bool = True) -> None:
        """Apply a bounded persisted layout without synthesizing a drag event."""
        if height is not None:
            self._height = max(90, min(460, int(height)))
        if collapsed is not None:
            self.collapsed = bool(collapsed)
        self.frame.configure(height=34 if self.collapsed else self._height)
        self.toggle_button.configure(text="展开" if self.collapsed else "收起")
        if notify and self.on_resize:
            self.on_resize(self._height, self.collapsed)

    def _resize(self, event):
        origin = getattr(self, "_resize_origin", (event.y_root, self._height))
        self._height = max(90, min(460, int(origin[1] + origin[0] - event.y_root)))
        if not self.collapsed:
            self.frame.configure(height=self._height)
        if self.on_resize:
            self.on_resize(self._height, self.collapsed)

    def _schedule_render(self):
        self._page = 0
        if self._query_after is not None:
            try: self.frame.after_cancel(self._query_after)
            except Exception: pass
        self._query_after = self.frame.after(120, self.render)

    def _next_page(self):
        items = self.index.filter(self.query_var.get(), only_unprocessed=self.unprocessed_var.get(), only_conflicts=self.conflicts_var.get())
        if (self._page + 1) * 500 < len(items):
            self._page += 1
            self.render()

    def _previous_page(self):
        if self._page > 0:
            self._page -= 1
            self.render()

    def set_items(self, items: Iterable[DifferenceItem]) -> None:
        selected = self.tree.selection()[0] if self.tree.selection() else None
        page = self._page
        self.index.replace(items)
        self._page = page
        self.render()
        if selected and selected in self._rendered:
            self.tree.selection_set(selected)
            self.tree.focus(selected)

    def mark_processed(self, item_id: str, processed: bool = True) -> None:
        selected = self.tree.selection()[0] if self.tree.selection() else None
        self.index.update(item_id, processed=processed)
        self.render()
        if selected and selected in self._rendered:
            self.tree.selection_set(selected)
            self.tree.focus(selected)

    def render(self) -> None:
        if not hasattr(self, "tree"):
            return
        counts = self.index.counts()
        self.summary_var.set(
            f"全部 {counts['total']} · 冲突 {counts[DifferenceKind.CONFLICT]} · 修改 {counts[DifferenceKind.MODIFIED]} · "
            f"增行 {counts[DifferenceKind.ADDED]} · 删行 {counts[DifferenceKind.DELETED]} · 已处理 {counts[DifferenceKind.PROCESSED]}"
        )
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self._rendered = {}
        # The index retains every item; the visible first page is bounded so a
        # very large workbook never blocks the Tk event loop.
        filtered = self.index.filter(self.query_var.get(), only_unprocessed=self.unprocessed_var.get(), only_conflicts=self.conflicts_var.get())
        page_count = max(1, (len(filtered) + 499) // 500)
        self._page = min(self._page, page_count - 1)
        self.page_var.set(f"第 {self._page + 1} 页 / {page_count}")
        for item in filtered[self._page * 500:(self._page + 1) * 500]:
            iid = f"diff-{item.id}"
            self._rendered[iid] = item
            self.tree.insert("", "end", iid=iid, values=(item.sheet, item.kind_label, item.display_location, item.summary, item.role, item.status_label), tags=(("processed" if item.processed else "conflict" if item.conflict else ""),))

    def _selected(self, _event=None):
        selection = self.tree.selection()
        if selection:
            item = self._rendered.get(selection[0])
            if item is not None:
                diffs = "；".join(f"{tag}: {value}" for tag, value in item.character_diff) or "无字符级差异"
                self.current_var.set(
                    f"{item.sheet} · {item.kind_label} · {item.display_location}\n\n{item.summary}\n"
                    f"{item.base_label}：{item.base_value!s}\n"
                    f"{item.mine_label}：{item.mine_value!s}\n"
                    f"{item.theirs_label}：{item.theirs_value!s}\n"
                    f"处理方向：{item.role}\n字符级差异：{diffs}\n状态：{item.status_label}"
                )
                if self.on_select:
                    self.on_select(item)

    def _activate(self, _event=None):
        self._selected()
        return "break"
