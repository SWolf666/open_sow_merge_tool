from __future__ import annotations

from dataclasses import dataclass

from sow_merge_tool.sheet_filter import SheetFilterModel


@dataclass
class _Item:
    sheet: str


def test_default_filter_keeps_unknown_then_hides_confirmed_clean_sheets():
    model = SheetFilterModel(("Clean", "Changed", "Pending"))
    assert model.visible_names() == ("Clean", "Changed", "Pending")
    model.mark_loading("Pending")
    model.publish("Clean", False, confirmed=True)
    model.publish("Changed", True, count=2, confirmed=True)
    assert model.visible_names() == ("Changed", "Pending")
    assert model.summary().startswith("有差异 1/3")
    assert model.status("Pending").wait_like


def test_filter_toggle_preserves_current_and_incremental_publish_does_not_reset_mode():
    model = SheetFilterModel(("A", "B", "C"))
    model.publish("A", True, count=1)
    model.publish("B", False)
    model.publish("C", True, count=3)
    assert model.preferred_sheet("C") == "C"
    model.set_mode("all")
    assert model.preferred_sheet("B") == "B"
    model.publish("A", False)
    assert model.mode == "all"
    assert model.preferred_sheet("B") == "B"
    model.set_mode("diff")
    assert model.visible_names() == ("C",)
    assert model.preferred_sheet("B") == "C"


def test_reset_is_new_comparison_boundary_and_empty_state_is_accessible():
    model = SheetFilterModel(("First", "Second"))
    model.publish("First", False)
    model.publish("Second", False)
    assert model.should_show_empty()
    assert model.visible_names() == ()
    model.set_mode("all")
    assert model.visible_names() == ("First", "Second")
    model.reset(("New", "Changed"))
    assert model.mode == "diff"
    assert model.user_overrode_mode is False
    assert model.visible_names() == ("New", "Changed")
    model.publish("New", False)
    model.publish("Changed", True)
    assert model.preferred_sheet() == "Changed"


def test_counts_consume_difference_items_without_rescanning():
    model = SheetFilterModel(("A", "B"))
    model.mark_loading("A")
    model.mark_loading("B")
    model.update_counts_from_items((_Item("A"), _Item("A"), _Item("B")))
    assert model.status("A").count == 2
    assert model.status("B").count == 1
    assert model.status("A").phase == "loading"
