from sow_merge_tool.vertical_layout import (
    MIN_C_HEIGHT,
    MIN_HOVER_HEIGHT,
    MIN_LOWER_HEIGHT,
    MIN_NAV_HEIGHT,
    VerticalLayout,
)


def test_vertical_layout_clamps_saved_sizes_to_window_budget():
    layout = VerticalLayout(lower_height=9999, hover_height=1, nav_height=9999)
    normalized = layout.normalized(620)

    assert normalized.lower_height <= 470
    assert normalized.lower_height >= MIN_LOWER_HEIGHT
    assert normalized.hover_height >= MIN_HOVER_HEIGHT
    assert normalized.nav_height >= MIN_NAV_HEIGHT
    assert normalized.lower_height - normalized.hover_height >= MIN_C_HEIGHT + 8


def test_vertical_layout_round_trips_settings_and_three_way_defaults():
    default = VerticalLayout.default(three_way=True)
    restored = VerticalLayout.from_mapping(default.to_mapping(), three_way=False)

    assert restored == default
    assert default.lower_height > VerticalLayout.default().lower_height
    assert default.hover_height > VerticalLayout.default().hover_height


def test_vertical_layout_invalid_settings_fail_closed_to_defaults():
    default = VerticalLayout.default()

    assert VerticalLayout.from_mapping(None) == default
    assert VerticalLayout.from_mapping({"lower_height": "bad"}) == default

