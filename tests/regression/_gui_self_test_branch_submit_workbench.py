"""Headed smoke test for the dense multi-branch submit workbench."""

from __future__ import annotations

import os
import shutil
import sqlite3
import tempfile
import time
import tkinter as tk
from unittest.mock import patch

from sow_merge_tool import branch_submit as bs


def _create_fixture():
    root = tempfile.mkdtemp(prefix="branch-submit-gui-")
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
            create table ACTUAL_NODE (
              wc_id integer, local_relpath text, parent_relpath text, properties blob,
              conflict_old text, conflict_new text, conflict_working text,
              prop_reject text, changelist text, text_mod text,
              tree_conflict_data text, conflict_data blob, older_checksum text,
              left_checksum text, right_checksum text
            );
            insert into REPOSITORY values (1, 'file:///repo', 'gui-fixture');
            """
        )
        for index in range(32):
            name = (
                "develop"
                if index == 0
                else "master"
                if index == 1
                else "feature_超长中文分支_配置同步_2026_第三季度_策划验收"
                if index == 2
                else f"feature_{index:02d}"
            )
            folder = os.path.join(root, name)
            os.makedirs(folder)
            open(os.path.join(folder, "seed.xlsx"), "wb").close()
            conn.execute(
                """insert into NODES
                (wc_id,local_relpath,op_depth,parent_relpath,repos_id,repos_path,revision,presence,moved_here,moved_to,kind,changed_revision,changed_author,file_external)
                values (1,?,0,'',1,?,1,'normal',0,'','dir',1,'tester',0)""",
                (name, "sheets/" + name),
            )
    return root


def main():
    fixture = _create_fixture()
    root = tk.Tk()
    root.withdraw()
    # Construct this Native surface under the Windows high-DPI scaling used by
    # the installation smoke test, then exercise the lower scale variants.
    root.tk.call("tk", "scaling", 2.666)
    items = [
        bs.SvnChangeItem(
            path=os.path.join(fixture, "develop", f"配置_{index:03d}.xlsx"),
            relative_path=f"配置_{index:03d}.xlsx", extension=".xlsx", node_kind="file",
            node_status="modified" if index % 4 else "unversioned",
            text_status="modified", prop_status="normal", versioned=index % 4 != 0,
            checked=index % 4 != 0, selectable=True,
        )
        for index in range(200)
    ]
    try:
        context = bs.BranchContext(fixture, "develop", os.path.join(fixture, "develop"))
        with patch.object(bs, "scan_changes", lambda *_args, **_kwargs: items):
            app = bs.BranchSubmitWorkbench(root, context)
            root.deiconify()
            deadline = time.time() + 5
            while len(app.items) != 200 and time.time() < deadline:
                root.update(); time.sleep(0.02)
            root.update_idletasks()
            assert len(app.items) == 200
            assert len(app.target_vars) == 31
            assert "master" in app.target_vars and not app.target_vars["master"].get()
            assert len(app.tree.get_children()) == 200
            assert int(root.winfo_width()) >= 900 and int(root.winfo_height()) >= 620
            assert set(app.tree["columns"]) == {"check", "path", "handling", "extension", "status", "property", "lock", "switched", "changelist"}
            assert "预检查" in app.preflight_button.cget("text")
            assert app.preflight_button.instate(["disabled"]), "without targets, multi-branch preflight is not applicable"
            assert app.submit_button.instate(["!disabled"]), "without targets, native single-branch submit must be available"
            assert app.submit_button.cget("text") == "SVN 单分支提交"
            for scale in (1.0, 1.25, 1.5, 2.0):
                root.tk.call("tk", "scaling", scale)
                for width, height in ((900, 620), (1366, 768), (1920, 1080)):
                    root.state("normal")
                    root.geometry(f"{width}x{height}+40+40")
                    root.update_idletasks()
                    root.update()
                    root_bottom = root.winfo_rooty() + root.winfo_height()
                    for button in (app.preflight_button, app.submit_button):
                        assert button.winfo_ismapped()
                        assert button.winfo_rooty() + button.winfo_height() <= root_bottom
            root.tk.call("tk", "scaling", 1.0)
            root.geometry("1120x760+40+40")
            root.update_idletasks()
            root.update()
            direct_calls = []
            app.engine._tortoise = lambda command, paths, **kwargs: direct_calls.append((command, list(paths), kwargs)) or 0
            app._submit_single_branch()
            deadline = time.time() + 5
            while (app._commit_active or not direct_calls) and time.time() < deadline:
                root.update(); time.sleep(0.02)
            assert direct_calls and direct_calls[0][0] == "commit"
            assert direct_calls[0][1] and direct_calls[0][2].get("message") is None
            root.geometry("900x620")
            root.update()
            time.sleep(0.03)
            root.update()
            root_bottom = root.winfo_rooty() + root.winfo_height()
            for button in (app.preflight_button, app.submit_button):
                assert button.winfo_ismapped(), f"footer button is not mapped: {button.cget('text')}"
                assert button.winfo_rooty() + button.winfo_height() <= root_bottom, (
                    f"footer button is outside the client area: {button.cget('text')} "
                    f"button_y={button.winfo_rooty()} button_h={button.winfo_height()} "
                    f"root_y={root.winfo_rooty()} root_h={root.winfo_height()} "
                    f"footer_y={app.footer_host.winfo_rooty()} footer_h={app.footer_host.winfo_height()} "
                    f"footer_req={app.footer_host.winfo_reqheight()} "
                    f"outer_y={app.footer_host.master.winfo_rooty()} outer_h={app.footer_host.master.winfo_height()} "
                    f"outer_req={app.footer_host.master.winfo_reqheight()}"
                )
            long_name = next(
                name for name in app.target_vars if "超长中文分支" in name
            )
            long_iid = next(
                iid for iid, name in app._target_rows.items() if name == long_name
            )
            assert app.target_xscrollbar.winfo_ismapped(), "目标分支横向滚动条不可见"
            app.target_tree.see(long_iid)
            root.update_idletasks()
            app.target_tree.xview_moveto(1.0)
            assert app.target_tree.xview()[0] > 0, "长分支名无法横向浏览"
            long_bbox = app.target_tree.bbox(long_iid)
            assert long_bbox, "长分支名行未渲染"

            class HoverEvent:
                x = 12
                y = long_bbox[1] + max(1, long_bbox[3] // 2)
                x_root = 80
                y_root = 120

            app._target_tree_hover(HoverEvent())
            deadline = time.monotonic() + 0.6
            while time.monotonic() < deadline and app._branch_tooltip is None:
                root.update()
                time.sleep(0.02)
            assert app._branch_tooltip is not None, "长分支名悬停全文提示未出现"
            tooltip_labels = [
                child
                for child in app._branch_tooltip.winfo_children()
                if child.winfo_class() == "Label"
            ]
            assert tooltip_labels and long_name in tooltip_labels[0].cget("text")
            app._copy_to_clipboard(long_name)
            assert root.clipboard_get() == long_name
            app._hide_branch_tooltip()
            long_scope = os.path.join(
                fixture,
                "develop",
                "配置表_中文路径_第一季度_最终验收_含有很长目录名",
            )
            app.scope_var.set(long_scope)
            root.update_idletasks()
            assert app.scope_scrollbar.winfo_ismapped(), "扫描路径横向滚动条不可见"
            app.scope_entry.xview_moveto(1.0)
            assert app.scope_entry.xview()[0] > 0, "中文扫描路径无法横向浏览"
            app._copy_to_clipboard(long_scope)
            assert root.clipboard_get() == long_scope
            app.scope_var.set(context.scope_path)
            target_name = next(name for name in app.target_vars if name != "master")
            app.target_vars[target_name].set(True)
            app._target_selection[target_name] = True
            app._refresh_primary_button()
            assert app.preflight_button.instate(["!disabled"]), "empty commit message must not block preflight"
            assert app.submit_button.instate(["disabled"]), "empty commit message must still block commit"
            assert app.submit_button.cget("text") == "② 开始提交"
            app.message.insert("1.0", "GUI 门禁测试")
            app._refresh_primary_button()
            assert app.preflight_button.instate(["!disabled"])
            assert app.submit_button.instate(["disabled"]), "selection alone must not enable submit"
            for item in app.items:
                item.checked = item.relative_path == "配置_001.xlsx"
            action = bs.BatchFileAction(
                branch=target_name,
                relative_path="配置_001.xlsx",
                operation="modify",
                state="confirmation_required",
                reason="目标分支同一单元格已有独立修改",
            )
            plan = bs.FilePlan(
                relative_path="配置_001.xlsx",
                operation="modify",
                actions={target_name: action},
                target_summaries={target_name: {"confirmation": 1}},
                target_details={target_name: [{
                    "kind": "confirmation", "sheet": "Data", "key": "1001",
                    "field": "value", "before": "旧", "source": "新", "target": "目标值",
                    "reason": action.reason,
                }]},
            )
            app.current_batch = bs.BranchSubmitBatch(
                batch_id="gui-manual",
                wc_root=fixture,
                source_branch="develop",
                target_branches=[target_name],
                files=[plan],
                message="GUI 门禁测试",
                scope_path=os.path.join(fixture, "develop"),
                source_status="ready",
                target_status={target_name: "confirmation_required"},
            )
            app._approved_preflight_signature = app._request_signature()
            app._render_target_statuses()
            root.update_idletasks()
            matrix_probe = {}

            def inspect_matrix():
                windows = [child for child in root.winfo_children() if isinstance(child, tk.Toplevel)]
                window = next(child for child in windows if "预检查结果" in child.title())

                def descendants(widget):
                    result = []
                    for child in widget.winfo_children():
                        result.append(child)
                        result.extend(descendants(child))
                    return result

                widgets = descendants(window)
                trees = [widget for widget in widgets if widget.winfo_class() == "Treeview"]
                buttons = [str(widget.cget("text")) for widget in widgets if widget.winfo_class() == "TButton"]
                matrix_probe["columns"] = tuple(trees[0]["columns"])
                matrix_probe["buttons"] = buttons
                matrix_probe["width"] = window.winfo_width()
                window.destroy()

            root.after(120, inspect_matrix)
            app._matrix_dialog(app.current_batch)
            assert matrix_probe["columns"] == ("branch", "file", "operation", "state", "reason")
            assert "查看目标修改点" in matrix_probe["buttons"]
            assert matrix_probe["width"] >= 860
            assert app.confirmation_alert.winfo_manager() == "pack"
            assert "内容重叠" in app.confirmation_alert_var.get()
            assert app.submit_button.instate(["disabled"]), "confirmation items must keep submit gated"
            app._open_confirmation_dialog()
            root.update_idletasks()
            dialog_tree = app._confirmation_dialog_tree
            assert dialog_tree is not None
            assert tuple(dialog_tree["columns"]) == ("branch", "file", "state", "reason")
            dialog_rows = dialog_tree.get_children()
            assert len(dialog_rows) == 1 and dialog_tree.set(dialog_rows[0], "state") == "待确认"
            app._set_confirmation_dialog_row(target_name, "配置_001.xlsx", "completed", "已确认采用源修改")
            root.update_idletasks()
            assert len(dialog_tree.get_children()) == 1, "completed row must remain in the current dialog"
            assert dialog_tree.set(dialog_rows[0], "state") == "已确认"
            app._confirmation_dialog.grab_release()
            app._confirmation_dialog.destroy()
            app._confirmation_dialog = None
            app._confirmation_dialog_tree = None
            app._confirmation_dialog_button = None
            app._confirmation_exclude_button = None
            app._confirmation_detail = None
            app._confirmation_dialog_summary_var = None
            app._confirmation_dialog_rows = {}
            action.state = "ready"
            action.confirmed = True
            app.current_batch.target_status[target_name] = "ready"
            app._render_target_statuses()
            app._refresh_primary_button()
            root.update_idletasks()
            assert not app.confirmation_alert.winfo_manager()
            assert app.submit_button.instate(["!disabled"]), "confirmed items may pass the gate"
            signature = app._approved_preflight_signature
            app.message.delete("1.0", tk.END)
            app._message_changed()
            assert app._approved_preflight_signature == signature, "commit message edits must preserve preflight"
            assert app.preflight_button.instate(["!disabled"])
            assert app.submit_button.instate(["disabled"]), "blank message must gate only final commit"
            app.message.insert("1.0", "修改后的提交说明")
            app._message_changed()
            assert app._approved_preflight_signature == signature
            assert app.submit_button.instate(["!disabled"]), "refilling message must not require preflight again"
            hold = float(os.environ.get("SOW_GUI_TEST_HOLD", "0") or 0)
            deadline = time.time() + hold
            while time.time() < deadline:
                root.update(); time.sleep(0.03)
        print("PASS: branch-submit workbench 32 branches / 200 files")
    finally:
        try: root.destroy()
        except tk.TclError: pass
        shutil.rmtree(fixture, ignore_errors=True)


if __name__ == "__main__":
    main()
