from threading import Event
from typing import ClassVar

from sow_merge_tool.difference_browser import DifferenceIndex
from sow_merge_tool.legacy_core import (
    MergeLaunchContext,
    MergeScenario,
    RolePresentation,
    SheetView,
    SowMergeApp,
    _augment_three_way_row_pairs_with_base_deletions,
)
from sow_merge_tool.ui_foundation import DifferenceItem, DifferenceKind


def _items():
    return (
        DifferenceItem("a", "Data", DifferenceKind.MODIFIED, row=2, column=3, summary="价格", role="Mine"),
        DifferenceItem("b", "Data", "conflict", row=4, column=1, summary="名称", role="Theirs", conflict=True),
        DifferenceItem("c", "Meta", "structure", summary="新增列"),
    )


def test_difference_index_filters_without_mutating_source():
    index = DifferenceIndex(_items())
    assert len(index.filter()) == 3
    assert [item.id for item in index.filter("名称")] == ["b"]
    assert [item.id for item in index.filter(only_conflicts=True)] == ["b"]
    index.update("b")
    assert [item.id for item in index.filter(only_unprocessed=True)] == ["a", "c"]
    assert index.counts()[DifferenceKind.PROCESSED] == 1


def test_difference_item_maps_display_location_and_alias_kind():
    item = DifferenceItem("x", "Data", "insert", row=7, column=2)
    assert item.kind is DifferenceKind.ADDED
    assert item.display_location == "第 7 行 · 第 2 列"


def test_role_presentation_formats_real_directions():
    app = type("App", (), {"merge_mode": True, "has_base": True, "launch_context": MergeLaunchContext(MergeScenario.CROSS_BRANCH_MERGE)})()
    view = object.__new__(SheetView)
    view.role_presentation = RolePresentation.for_app(app)
    assert view._format_direction_action("BASE2A", "row") == "应用 Source Before → Target Working 行"
    assert view._format_direction_action("B2A", "global") == "应用 Source After → Target Working 全局"
    app.merge_mode = False
    app.has_base = False
    app.launch_context = None
    view.role_presentation = RolePresentation.for_app(app)
    assert view._format_direction_action("A2B", "region") == "应用 Base → Mine 区域"


def test_role_presentation_distinguishes_ordinary_three_way_and_cross_branch():
    ordinary = type(
        "App",
        (),
        {
            "merge_mode": True,
            "has_base": True,
            "launch_context": MergeLaunchContext(MergeScenario.UPDATE_CONFLICT),
        },
    )()
    ordinary_roles = RolePresentation.for_app(ordinary)
    assert ordinary_roles.action_label("B2A") == "应用 Theirs → Mine 行"
    cross = type(
        "App",
        (),
        {
            "merge_mode": True,
            "has_base": True,
            "launch_context": MergeLaunchContext(MergeScenario.CROSS_BRANCH_MERGE),
        },
    )()
    cross_roles = RolePresentation.for_app(cross)
    assert cross_roles.action_label("BASE2A") == "应用 Source Before → Target Working 行"
    assert cross_roles.action_label("MINE2A") == "保留 Target Working 行"


def test_difference_cache_covers_two_way_and_three_way_row_shapes():
    app = object.__new__(SowMergeApp)
    app.merge_mode = True
    app.has_base = True
    app.launch_context = MergeLaunchContext(MergeScenario.UPDATE_CONFLICT)
    app._difference_processed = set()
    app.merge_conflict_cells_by_sheet = {}
    two_way_added = {
        "sheet": "Data",
        "row_pairs": [(None, 3)],
        "pair_diff_cols": {0: {-1}},
        "column_comparison_cache": None,
    }
    two_way_deleted = dict(two_way_added, row_pairs=[(3, None)])
    common_insert = dict(two_way_added, row_pairs=[(3, 3)], pair_diff_cols={}, pair_base_diff_cols={0: {-1}}, pair_base_row_override={0: None})
    assert app._difference_items_from_cache(two_way_added)[0].kind is DifferenceKind.ADDED
    assert app._difference_items_from_cache(two_way_deleted)[0].kind is DifferenceKind.DELETED
    assert app._difference_items_from_cache(common_insert)[0].kind is DifferenceKind.ADDED


def test_common_base_deletion_is_projected_from_production_pair_shape():
    # The (None, None) row is produced by the Base-row augmentation, not a
    # hand-authored browser fixture.  This mirrors Base 1..10 with Mine and
    # Theirs deleting row 5 while both independently add row 11.
    pairs, overrides = _augment_three_way_row_pairs_with_base_deletions(
        [(1, 1), (2, 2), (3, 3), (4, 4), (5, 5), (6, 6), (7, 7), (8, 8), (9, 9), (10, 10)],
        {1: 1, 2: 2, 3: 3, 4: 4, 5: 6, 6: 7, 7: 8, 8: 9, 9: 10},
        {1: 1, 2: 2, 3: 3, 4: 4, 5: 6, 6: 7, 7: 8, 8: 9, 9: 10},
        10,
    )
    assert pairs[4] == (None, None)
    assert overrides == {4: 5}
    app = object.__new__(SowMergeApp)
    app.merge_mode = True
    app.has_base = True
    app._difference_processed = set()
    app.merge_conflict_cells_by_sheet = {}
    cache = {
        "sheet": "Data",
        "row_pairs": pairs + [(10, 10)],
        "pair_diff_cols": {10: set()},
        "pair_base_diff_cols": {4: {-1}, 10: {-1}},
        "pair_base_row_override": {4: 5},
        "pair_parts_a": {10: ("row-11",)},
        "pair_parts_b": {10: ("row-11",)},
        "pair_parts_base": {4: ("row-5",)},
        "column_comparison_cache": None,
    }
    items = app._difference_items_from_cache(cache)
    assert any(item.kind is DifferenceKind.DELETED and "共同删除" in item.summary for item in items)
    assert any(item.kind is DifferenceKind.ADDED and "共同新增" in item.summary for item in items)


def test_two_way_item_values_keep_base_mine_semantics():
    app = object.__new__(SowMergeApp)
    app.merge_mode = False
    app.has_base = False
    app.launch_context = None
    app._difference_processed = set()
    app.merge_conflict_cells_by_sheet = {}
    cache = {
        "sheet": "Data", "row_pairs": [(1, 1)], "pair_diff_cols": {0: {2}},
        "pair_parts_a": {0: ("base", "A-value")},
        "pair_parts_b": {0: ("mine", "B-value")},
        "pair_parts_base": {}, "column_comparison_cache": None,
    }
    item = app._difference_items_from_cache(cache)[0]
    assert item.base_value == "A-value" and item.mine_value == "B-value" and item.theirs_value is None


def test_two_way_item_values_accept_scalar_cells_and_keep_direction():
    class View:
        sheet = "Data"
        pair_diff_cols: ClassVar = {0: {1}}
        pair_base_diff_cols: ClassVar = {}
        row_pairs: ClassVar = [(1, 1)]
        pair_parts_a: ClassVar = {0: "Base text"}
        pair_parts_b: ClassVar = {0: 17}
        pair_parts_base: ClassVar = {}
        role_presentation = RolePresentation("Base", "Base", "Mine", "base", "base", "mine")

        def _is_three_way_enabled(self):
            return False

        def _active_column_projection(self):
            return type("Projection", (), {"physical_col": lambda _self, _side, _col: 1})()

    app = object.__new__(SowMergeApp)
    app.merge_mode = False
    app.has_base = False
    app._difference_processed = set()
    app.merge_conflict_cells_by_sheet = {}
    item = app._difference_items_from_view(View())[0]
    assert (item.base_value, item.mine_value, item.theirs_value) == ("Base text", 17, None)
    assert (item.source_side, item.target_side) == ("base", "mine")
    assert SowMergeApp._value_from_parts(0, 1, 1) == 0
    assert SowMergeApp._value_from_parts(None, 1, 1) is None


def test_structure_processed_state_uses_real_difference_item_id():
    app = object.__new__(SowMergeApp)
    app._difference_processed = set()
    app._diff_browser = None
    app.difference_items = (
        DifferenceItem(
            "Data:pair:-1:side:structure:kind:structure:row:0:col:2",
            "Data",
            DifferenceKind.STRUCTURE,
            column=2,
        ),
    )
    app.mark_structure_difference_processed("Data", 2, True)
    assert app._difference_processed == {
        "Data:pair:-1:side:structure:kind:structure:row:0:col:2"
    }


def test_row_processed_fallback_returns_real_ids_and_empty_miss():
    app = object.__new__(SowMergeApp)
    app._difference_processed = set()
    app._diff_browser = None
    app.difference_items = (
        DifferenceItem(
            "Data:pair:3:side:mine:kind:modified:row:4:col:2",
            "Data",
            DifferenceKind.MODIFIED,
            row=4,
            column=2,
        ),
        DifferenceItem(
            "Data:pair:3:side:theirs:kind:conflict:row:4:col:3",
            "Data",
            DifferenceKind.CONFLICT,
            row=4,
            column=3,
            conflict=True,
        ),
    )
    app.refresh_difference_browser = lambda: None
    marked = app.mark_difference_row_processed("Data", 4, True)
    assert marked == tuple(item.id for item in app.difference_items)
    assert set(marked) <= app._difference_processed
    assert app.mark_difference_row_processed("Data", 99, True) == ()


def test_command_state_selection_is_immediately_actionable():
    app = object.__new__(SowMergeApp)
    app.merge_mode = False
    app.has_base = False
    app.raw_base = None
    app.raw_mine = None
    app.raw_theirs = None
    app.diff_base_mine_mode = False
    app._is_closing = False
    app._interactive_action_event = Event()
    app.modified_a = False
    app.modified_b = False
    app.merged_path = None
    app.undo_stack = []
    app.merge_conflict_cells_by_sheet = {}
    app.selected_sheet = "Data"
    app.sheet_views = {
        "Data": type("View", (), {"_data_ready": True, "selected_pair_idx": 3})()
    }
    state = app.command_state()
    assert state.ready and state.has_selection
    assert state.enabled("apply_source")


def test_current_difference_character_diff_uses_source_and_target_roles():
    item = DifferenceItem(
        "x",
        "Data",
        DifferenceKind.MODIFIED,
        base_value="same",
        mine_value=12,
        theirs_value="changed",
        source_side="theirs",
        target_side="mine",
    )
    assert item.character_diff
    assert any("changed" in value and "12" in value for _tag, value in item.character_diff)


def test_open_view_projection_uses_physical_mapping_for_values():
    class Projection:
        def physical_col(self, side, logical):
            return {"A": {1: 2}, "B": {1: 1}}[side][logical]
    class View:
        sheet = "Data"
        pair_diff_cols: ClassVar = {0: {1}}
        pair_base_diff_cols: ClassVar = {}
        row_pairs: ClassVar = [(1, 1)]
        pair_parts_a: ClassVar = {0: ("padding", "Base-physical-2")}
        pair_parts_b: ClassVar = {0: ("Mine-physical-1",)}
        pair_parts_base: ClassVar = {}
        role_presentation = RolePresentation("Base", "Base", "Mine", "base", "base", "mine")
        def _is_three_way_enabled(self): return False
        def _active_column_projection(self): return Projection()
    app = object.__new__(SowMergeApp)
    app.merge_mode = False; app.has_base = False; app._difference_processed = set(); app.merge_conflict_cells_by_sheet = {}
    item = app._difference_items_from_view(View())[0]
    assert item.base_value == "Base-physical-2" and item.mine_value == "Mine-physical-1" and item.theirs_value is None
    assert (item.base_label, item.mine_label, item.theirs_label) == ("Base", "Mine", "Theirs")


def test_valid_projection_missing_side_does_not_fallback_to_adjacent_column():
    class Projection:
        def physical_col(self, _side, logical):
            return None if int(logical) == 1 else int(logical)

    class View:
        def _active_column_projection(self):
            return Projection()

    assert SowMergeApp._logical_value_from_view(
        View(), 0, 1, "B", ("adjacent", "target")
    ) is None
    assert SowMergeApp._logical_value_from_view(
        View(), 0, 2, "B", ("adjacent", "target")
    ) == "target"
    assert SowMergeApp._logical_value_from_cache(
        {"column_comparison_cache": Projection()},
        0,
        1,
        "B",
        ("adjacent", "target"),
    ) is None


def test_missing_projection_information_keeps_logical_fallback():
    class View:
        def _active_column_projection(self):
            return None

    assert SowMergeApp._logical_value_from_view(
        View(), 0, 2, "B", ("first", "second")
    ) == "second"


def test_row_number_click_is_navigation_only():
    class Info:
        def configure(self, **_kwargs):
            return None
    class Event:
        x = 1
        y = 1
    view = object.__new__(SheetView)
    view.left_ln = object()
    view.base_ln = object()
    view.right_ln = object()
    view.info = Info()
    view._hover_ln_line_left = None
    view._hover_ln_line_mid = None
    view._hover_ln_line_right = None
    view._pair_idx_for_line = lambda _line: 0
    view._clear_row_header_hover = lambda _widget: None
    view._select_line = lambda _line: None
    view._copy_selected_row = lambda *_args, **_kwargs: (_ for _ in ()).throw(AssertionError("browse click wrote"))
    assert view._on_row_header_click(view.left_ln, Event(), "B2A") == "break"
