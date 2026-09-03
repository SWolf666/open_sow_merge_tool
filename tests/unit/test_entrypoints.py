from __future__ import annotations

import sow_merge_tool as public


def test_public_entrypoint_and_version() -> None:
    assert public.APP_NAME == "sow_merge_tool"
    assert public.APP_VERSION.endswith("update91")
    assert callable(public.run_entrypoint)
