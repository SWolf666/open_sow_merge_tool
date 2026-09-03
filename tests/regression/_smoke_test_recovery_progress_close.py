"""Ensure recovery progress windows reject early close and finish idempotently."""

from __future__ import annotations

import time
import tkinter as tk

from sow_merge_tool import branch_submit as bs


class _Engine:
    def restore_uncommitted(self, batch):
        time.sleep(0.8)
        return batch


def main() -> None:
    root = tk.Tk()
    root.withdraw()
    close_attempt = {"called": False, "exists_during_worker": False}

    def attempt_early_close(progress):
        def invoke_protocol():
            close_attempt["called"] = True
            try:
                command = progress.protocol("WM_DELETE_WINDOW")
                root.tk.call(command)
                close_attempt["exists_during_worker"] = bool(progress.winfo_exists())
            except tk.TclError:
                close_attempt["exists_during_worker"] = False

        root.after(30, invoke_protocol)

    batch = bs.BranchSubmitBatch(
        batch_id="recovery-progress-smoke",
        wc_root=".",
        source_branch="develop",
        target_branches=["release"],
        files=[],
        message="recovery smoke",
    )
    try:
        state = bs._restore_batch_with_progress(
            root,
            _Engine(),
            batch,
            _on_progress=attempt_early_close,
        )
        assert close_attempt["called"]
        assert close_attempt["exists_during_worker"], "early WM_DELETE must not destroy progress window"
        assert state.get("done") is True and state.get("result") is batch
        # The completion cleanup is idempotent even after the worker has
        # already destroyed the progress window.
        print("SMOKE_RECOVERY_PROGRESS_CLOSE_OK")
    finally:
        try:
            root.destroy()
        except tk.TclError:
            pass


if __name__ == "__main__":
    main()
