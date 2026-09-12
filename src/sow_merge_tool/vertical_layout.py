"""Pure layout state for vertically resizable comparison modules."""

from __future__ import annotations

from dataclasses import dataclass

GRIP_HEIGHT = 8
MIN_LOWER_HEIGHT = 132
MIN_C_HEIGHT = 64
MIN_HOVER_HEIGHT = 48
MIN_NAV_HEIGHT = 30
MAX_NAV_HEIGHT = 160


def _clamp(value: int, minimum: int, maximum: int) -> int:
    maximum = max(minimum, int(maximum))
    return max(int(minimum), min(int(value), maximum))


@dataclass(frozen=True)
class VerticalLayout:
    """Persisted vertical module sizes in logical pixels."""

    lower_height: int = 220
    hover_height: int = 84
    nav_height: int = 34

    @classmethod
    def default(cls, *, three_way: bool = False) -> VerticalLayout:
        return cls(
            lower_height=248 if three_way else 220,
            hover_height=124 if three_way else 84,
            nav_height=34,
        )

    @classmethod
    def from_mapping(cls, value: object, *, three_way: bool = False) -> VerticalLayout:
        default = cls.default(three_way=three_way)
        if not isinstance(value, dict):
            return default
        try:
            return cls(
                lower_height=int(value.get("lower_height", default.lower_height)),
                hover_height=int(value.get("hover_height", default.hover_height)),
                nav_height=int(value.get("nav_height", default.nav_height)),
            )
        except (TypeError, ValueError):
            return default

    def normalized(self, available_height: int, *, three_way: bool = False) -> VerticalLayout:
        """Clamp saved/user sizes to the current window/DPI budget."""
        available = max(MIN_LOWER_HEIGHT + MIN_C_HEIGHT + GRIP_HEIGHT, int(available_height or 0))
        max_lower = max(MIN_LOWER_HEIGHT, available - 150)
        lower = _clamp(self.lower_height, MIN_LOWER_HEIGHT, max_lower)
        max_hover = max(MIN_HOVER_HEIGHT, lower - MIN_C_HEIGHT - GRIP_HEIGHT)
        hover = _clamp(self.hover_height, MIN_HOVER_HEIGHT, max_hover)
        nav = _clamp(self.nav_height, MIN_NAV_HEIGHT, MAX_NAV_HEIGHT)
        return VerticalLayout(lower, hover, nav)

    def to_mapping(self) -> dict[str, int]:
        return {
            "lower_height": int(self.lower_height),
            "hover_height": int(self.hover_height),
            "nav_height": int(self.nav_height),
        }
