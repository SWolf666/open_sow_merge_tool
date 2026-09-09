"""Cache-only Sheet navigation filtering for the Excel comparison workspace."""

from __future__ import annotations

from collections import Counter
from collections.abc import Iterable
from dataclasses import dataclass


@dataclass(frozen=True)
class SheetFilterStatus:
    name: str
    order: int
    phase: str = "unknown"  # unknown/loading/ready/failed
    has_diff: bool | None = None
    count: int = 0
    error: str = ""

    @property
    def confirmed(self) -> bool:
        return self.phase == "ready" and self.has_diff is not None

    @property
    def wait_like(self) -> bool:
        return self.phase in {"unknown", "loading"}


class SheetFilterModel:
    """Stable, incremental Sheet filter state.

    The model never opens a workbook or computes a difference.  Unknown and
    loading sheets remain visible while ``only_diff`` is active so a late
    background result cannot make an inaccessible tab disappear.
    """

    def __init__(self, sheet_names: Iterable[str] = ()):
        self.mode = "diff"
        self.user_overrode_mode = False
        self.generation = 0
        self.current_sheet: str | None = None
        self.auto_select_pending = True
        self._names: tuple[str, ...] = ()
        self._statuses: dict[str, SheetFilterStatus] = {}
        self.reset(sheet_names)

    @property
    def only_diff(self) -> bool:
        return self.mode == "diff"

    @property
    def all_names(self) -> tuple[str, ...]:
        return self._names

    @property
    def statuses(self) -> dict[str, SheetFilterStatus]:
        return dict(self._statuses)

    def reset(self, sheet_names: Iterable[str]) -> None:
        ordered: list[str] = []
        for raw in sheet_names:
            name = str(raw)
            if name and name not in ordered:
                ordered.append(name)
        self._names = tuple(ordered)
        self._statuses = {
            name: SheetFilterStatus(name=name, order=index)
            for index, name in enumerate(self._names)
        }
        self.mode = "diff"
        self.user_overrode_mode = False
        self.generation += 1
        self.current_sheet = None
        self.auto_select_pending = True

    def sync_names(self, sheet_names: Iterable[str]) -> None:
        """Keep state for existing names while adding/removing catalog names."""
        ordered: list[str] = []
        for raw in sheet_names:
            name = str(raw)
            if name and name not in ordered:
                ordered.append(name)
        old = self._statuses
        self._names = tuple(ordered)
        self._statuses = {
            name: SheetFilterStatus(
                name=name,
                order=index,
                phase=old[name].phase,
                has_diff=old[name].has_diff,
                count=old[name].count,
                error=old[name].error,
            )
            if name in old
            else SheetFilterStatus(name=name, order=index)
            for index, name in enumerate(self._names)
        }
        if self.current_sheet not in self._statuses:
            self.current_sheet = None

    def set_mode(self, mode: str) -> bool:
        value = "all" if str(mode).lower() in {"all", "全部", "0"} else "diff"
        changed = value != self.mode
        self.mode = value
        self.user_overrode_mode = True
        return changed

    def toggle(self) -> bool:
        return self.set_mode("all" if self.only_diff else "diff")

    def mark_loading(self, name: str) -> None:
        self._set(name, phase="loading", has_diff=None, error="")

    def mark_failed(self, name: str, error: str = "") -> None:
        self._set(name, phase="failed", has_diff=None, error=str(error or "后台计算失败"))

    def publish(self, name: str, has_diff: bool, *, count: int = 0, confirmed: bool = True) -> None:
        self._set(
            name,
            phase="ready" if confirmed else "loading",
            has_diff=bool(has_diff) if confirmed else None,
            count=max(0, int(count)),
            error="",
        )

    def set_count(self, name: str, count: int) -> None:
        status = self._statuses.get(str(name))
        if status is not None:
            self._statuses[status.name] = SheetFilterStatus(
                name=status.name,
                order=status.order,
                phase=status.phase,
                has_diff=status.has_diff,
                count=max(0, int(count)),
                error=status.error,
            )

    def _set(self, name: str, **changes) -> None:
        key = str(name)
        status = self._statuses.get(key)
        if status is None:
            return
        values = {
            "name": status.name,
            "order": status.order,
            "phase": status.phase,
            "has_diff": status.has_diff,
            "count": status.count,
            "error": status.error,
        }
        values.update(changes)
        self._statuses[key] = SheetFilterStatus(**values)

    def update_counts_from_items(self, items: Iterable[object]) -> None:
        counts: Counter[str] = Counter()
        for item in items:
            sheet = getattr(item, "sheet", None)
            if sheet is not None:
                counts[str(sheet)] += 1
        for name in self._names:
            self.set_count(name, counts.get(name, 0))

    def status(self, name: str) -> SheetFilterStatus:
        return self._statuses[str(name)]

    def visible_names(self) -> tuple[str, ...]:
        if self.mode == "all":
            return self._names
        return tuple(
            name
            for name in self._names
            if not (
                self._statuses[name].phase == "ready"
                and self._statuses[name].has_diff is False
            )
        )

    def confirmed_diff_names(self) -> tuple[str, ...]:
        return tuple(
            name
            for name in self._names
            if self._statuses[name].phase == "ready"
            and self._statuses[name].has_diff is True
        )

    def summary(self) -> str:
        diff_count = len(self.confirmed_diff_names())
        total = len(self._names)
        pending = sum(1 for status in self._statuses.values() if status.wait_like)
        failed = sum(1 for status in self._statuses.values() if status.phase == "failed")
        suffix = f" · 待确认 {pending}" if pending else ""
        if failed:
            suffix += f" · 失败 {failed}"
        return f"有差异 {diff_count}/{total}{suffix}"

    def preferred_sheet(self, current: str | None = None) -> str | None:
        visible = self.visible_names()
        current_name = str(current or self.current_sheet or "")
        if current_name and current_name in visible:
            status = self._statuses[current_name]
            if self.mode == "all" or status.has_diff is not False:
                self.current_sheet = current_name
                return current_name
        for name in self.confirmed_diff_names():
            if name in visible:
                self.current_sheet = name
                return name
        if visible:
            self.current_sheet = visible[0]
            return visible[0]
        return None

    def should_show_empty(self) -> bool:
        return self.mode == "diff" and not self.confirmed_diff_names() and not any(
            status.wait_like for status in self._statuses.values()
        )


__all__ = ["SheetFilterModel", "SheetFilterStatus"]
