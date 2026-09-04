"""Repeatable local benchmark for the branch workbench UI path.

This benchmark creates only a temporary synthetic SVN database and writes
timings to the ignored artifacts directory.  It never opens a business
working copy or records its paths.
"""

from __future__ import annotations

import json
import os
import shutil
import sqlite3
import tempfile
import time
import tkinter as tk
from pathlib import Path
from unittest.mock import patch

from sow_merge_tool import branch_submit as bs


def _fixture() -> str:
    root = tempfile.mkdtemp(prefix="sow-ui-benchmark-")
    os.makedirs(os.path.join(root, ".svn"))
    with sqlite3.connect(os.path.join(root, ".svn", "wc.db")) as conn:
        conn.executescript(
            """
            create table REPOSITORY (id integer primary key, root text, uuid text);
            create table NODES (
              wc_id integer, local_relpath text, op_depth integer, parent_relpath text,
              repos_id integer, repos_path text, revision integer, presence text,
              moved_here integer, moved_to text, kind text, properties blob, depth text,
              checksum text, symlink_target text, changed_revision integer,
              changed_date integer, changed_author text, translated_size integer,
              last_mod_time integer, dav_cache blob, file_external integer,
              inherited_props blob
            );
            insert into REPOSITORY values (1, 'file:///benchmark', 'ui-benchmark');
            """
        )
        for index in range(50):
            name = "develop" if index == 0 else f"feature_{index:02d}"
            os.makedirs(os.path.join(root, name))
            conn.execute(
                """insert into NODES
                (wc_id,local_relpath,op_depth,parent_relpath,repos_id,repos_path,revision,presence,moved_here,moved_to,kind,changed_revision,changed_date,changed_author,file_external)
                values (1,?,0,'',1,?,1,'normal',0,'','dir',1,?, 'benchmark',0)""",
                (name, "sheets/" + name, (index + 1) * 1_000_000),
            )
    return root


def main() -> None:
    root_path = _fixture()
    try:
        discovery_ms = []
        for _ in range(5):
            started = time.perf_counter()
            candidates = bs.discover_branch_candidates(root_path)
            discovery_ms.append((time.perf_counter() - started) * 1000)

        items = [
            bs.SvnChangeItem(
                path=os.path.join(root_path, "develop", f"配置_{index:04d}.xlsx"),
                relative_path=f"配置_{index:04d}.xlsx",
                extension=".xlsx",
                node_kind="file",
                node_status="modified" if index % 5 else "unversioned",
                text_status="modified",
                prop_status="normal",
                versioned=index % 5 != 0,
                checked=index % 5 != 0,
                selectable=True,
            )
            for index in range(1000)
        ]
        ui_root = tk.Tk()
        ui_root.withdraw()
        try:
            context = bs.BranchContext(root_path, "develop", os.path.join(root_path, "develop"), candidates=candidates)
            started = time.perf_counter()
            with patch.object(bs, "scan_changes", lambda *_args, **_kwargs: items):
                app = bs.BranchSubmitWorkbench(ui_root, context)
                ui_root.deiconify()
                for _ in range(300):
                    ui_root.update()
                    if len(app.items) == len(items):
                        break
                    time.sleep(0.005)
                render_ms = (time.perf_counter() - started) * 1000
                app._quick_check("none")
                app._quick_check("all")
                app.target_vars["feature_01"].set(True)
                app._target_selection["feature_01"] = True
                app.target_search_var.set("feature_02")
                for _ in range(20):
                    ui_root.update()
                    time.sleep(0.002)
                assert "feature_01" in app._selected_targets(), "隐藏搜索结果不应丢失分支选择"
                app.target_search_var.set("")
                for _ in range(20):
                    ui_root.update()
                    time.sleep(0.002)
                assert len(app.tree.get_children()) == len(items)
            payload = {
                "branch_count": len(candidates),
                "item_count": len(app.items),
                "discovery_ms": discovery_ms,
                "workbench_ready_ms": render_ms,
            }
        finally:
            ui_root.destroy()
        output = Path("artifacts/performance/update94-ui.json")
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(payload, ensure_ascii=False))
    finally:
        shutil.rmtree(root_path, ignore_errors=True)


if __name__ == "__main__":
    main()
