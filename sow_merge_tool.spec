# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

project_root = Path(SPECPATH).resolve()
src_root = project_root / "src"

a = Analysis(
    [str(src_root / "sow_merge_tool" / "__main__.py")],
    pathex=[str(src_root)],
    binaries=[],
    datas=[],
    hiddenimports=[
        "sow_merge_tool.legacy_core",
        "sow_merge_tool.branch_submit",
        "sow_merge_tool.svn_status_provider",
        "sow_merge_tool.ui_foundation",
        "sow_merge_tool.path_selection",
        "sow_merge_tool.launch_center",
    ],
    hookspath=[],
    hooksconfig={},
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="sow_merge_tool",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
