"""Fail-closed SVN status discovery for the multi-branch submit workbench.

The normal path uses ``svn status --xml``.  Game-planning workstations often
only install TortoiseSVN, so the fallback loads TortoiseSVN's bundled
Subversion runtime in a disposable child process.  A native ABI failure must
never take down the UI or be mistaken for a clean working copy.
"""

from __future__ import annotations

import ctypes
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import uuid
import xml.etree.ElementTree as ET
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable


STATUS_NAMES = {
    1: "none",
    2: "unversioned",
    3: "normal",
    4: "added",
    5: "missing",
    6: "deleted",
    7: "replaced",
    8: "modified",
    9: "merged",
    10: "conflicted",
    11: "ignored",
    12: "obstructed",
    13: "external",
    14: "incomplete",
}
NODE_KIND_NAMES = {0: "none", 1: "file", 2: "dir", 3: "unknown", 4: "symlink"}
INTERESTING_STATUS = {
    "unversioned", "added", "missing", "deleted", "replaced",
    "modified", "merged", "conflicted", "obstructed", "incomplete",
}


@dataclass
class SvnStatusRecord:
    path: str
    node_kind: str = "unknown"
    node_status: str = "none"
    text_status: str = "none"
    prop_status: str = "none"
    versioned: bool = False
    conflicted: bool = False
    switched: bool = False
    file_external: bool = False
    wc_locked: bool = False
    lock_owner: str = ""
    changelist: str = ""
    moved_from: str = ""
    moved_to: str = ""
    revision: int | None = None
    # ``revision`` is the working-copy BASE revision.  The repository fields
    # are populated only when ``svn status --verbose`` (or the native
    # TortoiseSVN ABI) can prove a newer server-side node.  Keeping them
    # optional preserves compatibility with old cached/test records while
    # allowing the submit engine to enforce per-file freshness when evidence
    # is available.
    repository_revision: int | None = None
    repository_node_status: str = ""
    repository_text_status: str = ""
    repository_prop_status: str = ""
    repository_changed_revision: int | None = None
    # Friendly aliases used by newer callers; both names are serialized so
    # cached status adapters can migrate without losing evidence.
    remote_revision: int | None = None
    remote_node_status: str = ""
    remote_text_status: str = ""
    remote_prop_status: str = ""
    repos_root_url: str = ""
    repos_uuid: str = ""
    repos_relpath: str = ""

    @classmethod
    def from_dict(cls, value: dict) -> "SvnStatusRecord":
        fields = cls.__dataclass_fields__
        return cls(**{key: value[key] for key in fields if key in value})


class SvnStatusError(RuntimeError):
    pass


def _decode(value: bytes | None) -> str:
    return value.decode("utf-8", errors="replace") if value else ""


def _find_svn_cli() -> str | None:
    candidate = shutil.which("svn")
    if candidate:
        return candidate
    for root in (
        os.environ.get("ProgramFiles", r"C:\Program Files"),
        os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
    ):
        candidate = os.path.join(root, "TortoiseSVN", "bin", "svn.exe")
        if os.path.isfile(candidate):
            return candidate
    return None


def _find_tortoise_bin() -> str | None:
    for root in (
        os.environ.get("ProgramFiles", r"C:\Program Files"),
        os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
    ):
        candidate = os.path.join(root, "TortoiseSVN", "bin")
        if all(os.path.isfile(os.path.join(candidate, name)) for name in (
            "libapr_tsvn.dll", "libsvn_tsvn.dll",
        )):
            return candidate
    return None


def _parse_cli_status(xml_text: str, requested_path: str) -> list[SvnStatusRecord]:
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as exc:
        raise SvnStatusError(f"SVN 状态 XML 无效：{exc}") from exc
    requested = os.path.abspath(requested_path)
    records: list[SvnStatusRecord] = []
    changelist_by_entry: dict[int, str] = {}
    for changelist in root.findall(".//changelist"):
        name = str(changelist.get("name") or "")
        for entry in changelist.findall(".//entry"):
            changelist_by_entry[id(entry)] = name
    for entry in root.findall(".//entry"):
        raw_path = str(entry.get("path") or "")
        path = raw_path if os.path.isabs(raw_path) else os.path.join(requested, raw_path)
        wc = entry.find("wc-status")
        if wc is None:
            continue
        lock = wc.find("lock")
        item = str(wc.get("item") or "none")
        props = str(wc.get("props") or "none")
        repos = wc.find("repos-status")
        if repos is None:
            repos = entry.find("repos-status")
        commit = wc.find("commit")
        remote_item = str(repos.get("item") or "") if repos is not None else ""
        remote_props = str(repos.get("props") or "") if repos is not None else ""
        remote_revision = None
        if repos is not None and str(repos.get("revision") or "").isdigit():
            remote_revision = int(repos.get("revision"))
        changed_revision = None
        if commit is not None and str(commit.get("revision") or "").isdigit():
            changed_revision = int(commit.get("revision"))
        record = SvnStatusRecord(
            path=os.path.abspath(path),
            node_kind="dir" if os.path.isdir(path) else "file",
            node_status=item,
            text_status=item if item in {"modified", "conflicted", "normal"} else "none",
            prop_status=props,
            versioned=item not in {"none", "unversioned", "ignored"},
            conflicted=item == "conflicted" or props == "conflicted" or wc.get("tree-conflicted") == "true",
            switched=wc.get("switched") == "true",
            file_external=item == "external",
            wc_locked=wc.get("wc-locked") == "true",
            lock_owner=str(lock.findtext("owner") or "") if lock is not None else "",
            changelist=str(changelist_by_entry.get(id(entry)) or entry.get("changelist") or wc.get("changelist") or ""),
            moved_from=str(wc.get("moved-from") or ""),
            moved_to=str(wc.get("moved-to") or ""),
            revision=int(wc.get("revision")) if str(wc.get("revision") or "").isdigit() else None,
            repository_revision=remote_revision,
            repository_node_status=remote_item,
            repository_text_status=remote_item,
            repository_prop_status=remote_props,
            repository_changed_revision=changed_revision,
            remote_revision=remote_revision,
            remote_node_status=remote_item,
            remote_text_status=remote_item,
            remote_prop_status=remote_props,
        )
        records.append(record)
    return records


def _communicate_cancelable(command: list[str], *, timeout: float, cancel_event: threading.Event | None, capture: bool) -> tuple[int, str, str]:
    stdout_file = tempfile.TemporaryFile() if capture else None
    stderr_file = tempfile.TemporaryFile() if capture else None
    try:
        process = subprocess.Popen(
            command,
            stdout=stdout_file if capture else subprocess.DEVNULL,
            stderr=stderr_file if capture else subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        deadline = time.monotonic() + timeout
        while process.poll() is None:
            if cancel_event is not None and cancel_event.is_set():
                process.terminate()
                try:
                    process.wait(timeout=2)
                except subprocess.TimeoutExpired:
                    process.kill()
                raise SvnStatusError("SVN 状态扫描已取消")
            if time.monotonic() >= deadline:
                process.kill()
                raise SvnStatusError("SVN 状态扫描超时")
            time.sleep(0.05)
        stdout = stderr = ""
        if capture:
            stdout_file.seek(0); stderr_file.seek(0)
            stdout = stdout_file.read().decode("utf-8", errors="replace")
            stderr = stderr_file.read().decode("utf-8", errors="replace")
        return int(process.returncode or 0), stdout, stderr
    finally:
        if stdout_file: stdout_file.close()
        if stderr_file: stderr_file.close()


def _status_via_cli(
    path: str,
    cancel_event: threading.Event | None = None,
    *,
    remote: bool = False,
    show_updates: bool = False,
) -> list[SvnStatusRecord]:
    svn = _find_svn_cli()
    if not svn:
        raise SvnStatusError("svn-cli-unavailable")
    try:
        command = [svn, "status", "--xml", "--verbose", "--depth", "infinity", "--ignore-externals"]
        if remote or show_updates:
            # ``--show-updates`` is read-only but asks the repository for the
            # exact selected-path freshness.  It never updates the working
            # copy; callers decide whether an explicit Update is allowed.
            command.append("--show-updates")
        command.append(path)
        returncode, stdout, stderr = _communicate_cancelable(
            command,
            timeout=120,
            cancel_event=cancel_event,
            capture=True,
        )
    except SvnStatusError:
        raise
    except Exception as exc:
        raise SvnStatusError(f"SVN 状态扫描启动失败：{exc}") from exc
    if returncode != 0:
        detail = (stderr or stdout or "").strip()
        raise SvnStatusError(f"SVN 状态扫描失败：{detail or returncode}")
    return _parse_cli_status(stdout, path)


class _SvnOptRevisionValue(ctypes.Union):
    _fields_ = [("number", ctypes.c_long), ("date", ctypes.c_longlong)]


class _SvnOptRevision(ctypes.Structure):
    _fields_ = [("kind", ctypes.c_int), ("value", _SvnOptRevisionValue)]


class _SvnClientStatus(ctypes.Structure):
    # Subversion 1.14 svn_client_status_t through moved_to_abspath.  The object
    # is allocated by libsvn, so forward-compatible tail fields are harmless.
    _fields_ = [
        ("kind", ctypes.c_int),
        ("local_abspath", ctypes.c_char_p),
        ("filesize", ctypes.c_longlong),
        ("versioned", ctypes.c_int),
        ("conflicted", ctypes.c_int),
        ("node_status", ctypes.c_int),
        ("text_status", ctypes.c_int),
        ("prop_status", ctypes.c_int),
        ("wc_is_locked", ctypes.c_int),
        ("copied", ctypes.c_int),
        ("repos_root_url", ctypes.c_char_p),
        ("repos_uuid", ctypes.c_char_p),
        ("repos_relpath", ctypes.c_char_p),
        # svn_revnum_t is C ``long`` (32-bit on Windows), while file sizes,
        # apr_time_t and the revision option union remain 64-bit.
        ("revision", ctypes.c_long),
        ("changed_rev", ctypes.c_long),
        ("changed_date", ctypes.c_longlong),
        ("changed_author", ctypes.c_char_p),
        ("switched", ctypes.c_int),
        ("file_external", ctypes.c_int),
        ("lock", ctypes.c_void_p),
        ("changelist", ctypes.c_char_p),
        ("depth", ctypes.c_int),
        ("ood_kind", ctypes.c_int),
        ("repos_node_status", ctypes.c_int),
        ("repos_text_status", ctypes.c_int),
        ("repos_prop_status", ctypes.c_int),
        ("repos_lock", ctypes.c_void_p),
        ("ood_changed_rev", ctypes.c_long),
        ("ood_changed_date", ctypes.c_longlong),
        ("ood_changed_author", ctypes.c_char_p),
        ("backwards_compatibility_baton", ctypes.c_void_p),
        ("moved_from_abspath", ctypes.c_char_p),
        ("moved_to_abspath", ctypes.c_char_p),
    ]


def _native_error(svn, pointer, operation: str) -> None:
    if not pointer:
        return
    message = ""
    try:
        svn.svn_err_best_message.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_size_t]
        svn.svn_err_best_message.restype = ctypes.c_char_p
        buffer = ctypes.create_string_buffer(2048)
        svn.svn_err_best_message(pointer, buffer, len(buffer))
        message = buffer.value.decode("utf-8", errors="replace")
        svn.svn_error_clear.argtypes = [ctypes.c_void_p]
        svn.svn_error_clear(pointer)
    finally:
        raise SvnStatusError(f"{operation}：{message or 'native SVN error'}")


def query_tortoise_status_in_child(path: str, *, remote: bool = False) -> list[SvnStatusRecord]:
    """Execute only inside the isolated helper process."""
    bin_dir = _find_tortoise_bin()
    if not bin_dir or not hasattr(os, "add_dll_directory") or not hasattr(ctypes, "WinDLL"):
        raise SvnStatusError("TortoiseSVN 状态运行库不可用")
    dll_cookie = os.add_dll_directory(bin_dir)
    apr_ready = False
    pool = ctypes.c_void_p()
    callbacks = []
    try:
        apr = ctypes.WinDLL(os.path.join(bin_dir, "libapr_tsvn.dll"))
        svn = ctypes.WinDLL(os.path.join(bin_dir, "libsvn_tsvn.dll"))
        apr.apr_initialize.argtypes = []
        apr.apr_initialize.restype = ctypes.c_int
        if int(apr.apr_initialize()):
            raise SvnStatusError("APR 初始化失败")
        apr_ready = True
        apr.apr_pool_create_ex.argtypes = [ctypes.POINTER(ctypes.c_void_p), ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p]
        apr.apr_pool_create_ex.restype = ctypes.c_int
        if int(apr.apr_pool_create_ex(ctypes.byref(pool), None, None, None)):
            raise SvnStatusError("APR 内存池创建失败")
        svn.svn_dso_initialize2.argtypes = []
        svn.svn_dso_initialize2.restype = ctypes.c_void_p
        _native_error(svn, svn.svn_dso_initialize2(), "SVN DSO 初始化失败")
        if hasattr(svn, "svn_wc_initialize"):
            svn.svn_wc_initialize.argtypes = [ctypes.c_void_p]
            svn.svn_wc_initialize.restype = ctypes.c_void_p
            _native_error(svn, svn.svn_wc_initialize(pool), "SVN 工作副本初始化失败")
        context = ctypes.c_void_p()
        svn.svn_client_create_context2.argtypes = [ctypes.POINTER(ctypes.c_void_p), ctypes.c_void_p, ctypes.c_void_p]
        svn.svn_client_create_context2.restype = ctypes.c_void_p
        _native_error(svn, svn.svn_client_create_context2(ctypes.byref(context), None, pool), "SVN 客户端初始化失败")

        records: list[SvnStatusRecord] = []
        # Subversion callbacks use the C calling convention.  On Win64 the
        # register ABI is shared, but CFUNCTYPE is still the correct and less
        # surprising declaration (and remains correct for 32-bit builds).
        callback_type = ctypes.CFUNCTYPE(
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_char_p,
            ctypes.POINTER(_SvnClientStatus),
            ctypes.c_void_p,
        )

        def receive(_baton, callback_path, status_ptr, _scratch):
            try:
                if not status_ptr:
                    return None
                status = status_ptr.contents
                node_status = STATUS_NAMES.get(int(status.node_status), f"unknown-{status.node_status}")
                remote_node_status = STATUS_NAMES.get(int(status.repos_node_status), "none")
                remote_text_status = STATUS_NAMES.get(int(status.repos_text_status), "none")
                remote_prop_status = STATUS_NAMES.get(int(status.repos_prop_status), "none")
                if node_status in {"normal", "ignored", "external", "none"} and not remote and not (
                    status.conflicted or status.switched or status.file_external
                    or STATUS_NAMES.get(int(status.prop_status)) not in {"none", "normal"}
                    or remote_node_status not in {"none", "normal"}
                    or remote_text_status not in {"none", "normal"}
                    or remote_prop_status not in {"none", "normal"}
                ):
                    return None
                raw_path = _decode(status.local_abspath) or _decode(callback_path)
                records.append(SvnStatusRecord(
                    path=os.path.abspath(raw_path),
                    node_kind=NODE_KIND_NAMES.get(int(status.kind), "unknown"),
                    node_status=node_status,
                    text_status=STATUS_NAMES.get(int(status.text_status), "unknown"),
                    prop_status=STATUS_NAMES.get(int(status.prop_status), "unknown"),
                    versioned=bool(status.versioned),
                    conflicted=bool(status.conflicted),
                    switched=bool(status.switched),
                    file_external=bool(status.file_external),
                    wc_locked=bool(status.wc_is_locked),
                    changelist=_decode(status.changelist),
                    moved_from=_decode(status.moved_from_abspath),
                    moved_to=_decode(status.moved_to_abspath),
                    revision=int(status.revision) if int(status.revision) >= 0 else None,
                    repository_revision=(
                        int(status.ood_changed_rev) if int(status.ood_changed_rev) >= 0 else None
                    ),
                    repository_node_status=remote_node_status,
                    repository_text_status=remote_text_status,
                    repository_prop_status=remote_prop_status,
                    repository_changed_revision=(
                        int(status.ood_changed_rev) if int(status.ood_changed_rev) >= 0 else None
                    ),
                    remote_revision=(
                        int(status.ood_changed_rev) if int(status.ood_changed_rev) >= 0 else None
                    ),
                    remote_node_status=remote_node_status,
                    remote_text_status=remote_text_status,
                    remote_prop_status=remote_prop_status,
                    repos_root_url=_decode(status.repos_root_url),
                    repos_uuid=_decode(status.repos_uuid),
                    repos_relpath=_decode(status.repos_relpath),
                ))
            except Exception:
                # Callback exceptions cannot cross a C ABI.  Add an explicit
                # poison record so the parent fails closed.
                records.append(SvnStatusRecord(path=os.path.abspath(path), node_status="status-callback-failed"))
            return None

        callback = callback_type(receive)
        callbacks.append(callback)
        revision = _SvnOptRevision()
        revision.kind = 0  # svn_opt_revision_unspecified
        result_revision = ctypes.c_long(-1)
        svn.svn_client_status6.argtypes = [
            ctypes.POINTER(ctypes.c_long), ctypes.c_void_p, ctypes.c_char_p,
            ctypes.POINTER(_SvnOptRevision), ctypes.c_int,
            ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
            ctypes.c_int, ctypes.c_void_p, callback_type, ctypes.c_void_p, ctypes.c_void_p,
        ]
        svn.svn_client_status6.restype = ctypes.c_void_p
        error = svn.svn_client_status6(
            ctypes.byref(result_revision), context, os.path.abspath(path).encode("utf-8"),
            ctypes.byref(revision), 3,  # svn_depth_infinity
            1 if remote else 0,
            1 if remote else 0,
            1,
            0,
            1,
            0,
            None, callback, None, pool,
        )
        _native_error(svn, error, "TortoiseSVN 状态扫描失败")
        poison = next((item for item in records if item.node_status == "status-callback-failed"), None)
        if poison:
            raise SvnStatusError("TortoiseSVN 状态回调解析失败")
        return records
    finally:
        callbacks.clear()
        if pool:
            try:
                apr.apr_pool_destroy.argtypes = [ctypes.c_void_p]
                apr.apr_pool_destroy(pool)
            except Exception:
                pass
        if apr_ready:
            try:
                apr.apr_terminate()
            except Exception:
                pass
        dll_cookie.close()


def internal_status_entrypoint(argv: list[str]) -> int:
    if len(argv) not in {2, 3}:
        return 2
    path, output_path = argv
    remote = len(argv) == 3 and argv[2] in {"--remote", "--show-updates"}
    payload: dict
    try:
        payload = {"ok": True, "items": [asdict(item) for item in query_tortoise_status_in_child(path, remote=remote)]}
    except Exception as exc:
        payload = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
    try:
        tmp = output_path + f".tmp-{os.getpid()}"
        with open(tmp, "w", encoding="utf-8") as stream:
            json.dump(payload, stream, ensure_ascii=False)
        os.replace(tmp, output_path)
    except OSError:
        return 3
    return 0 if payload.get("ok") else 1


def _status_via_tortoise_child(
    path: str,
    host_script: str | None = None,
    cancel_event: threading.Event | None = None,
    *,
    remote: bool = False,
) -> list[SvnStatusRecord]:
    if not _find_tortoise_bin():
        raise SvnStatusError("未找到 svn.exe 或 TortoiseSVN 状态运行库")
    result_path = os.path.join(tempfile.gettempdir(), f"sow_svn_status_{os.getpid()}_{uuid.uuid4().hex}.json")
    command = [sys.executable]
    if not getattr(sys, "frozen", False):
        # The implementation moved under src/; using the module entry point
        # avoids reconstructing a stale root script path and works both from
        # an editable checkout and an installed package.
        command.extend(("-m", "sow_merge_tool"))
    command.extend(("--internal-svn-status-query", os.path.abspath(path), result_path))
    if remote:
        command.append("--remote")
    try:
        returncode, _stdout, _stderr = _communicate_cancelable(
            command,
            timeout=180,
            cancel_event=cancel_event,
            capture=False,
        )
        try:
            with open(result_path, "r", encoding="utf-8") as stream:
                payload = json.load(stream)
        except (OSError, ValueError) as exc:
            raise SvnStatusError(f"TortoiseSVN 状态子进程没有返回有效结果（退出码 {returncode}）") from exc
        if returncode != 0 or not payload.get("ok"):
            raise SvnStatusError(str(payload.get("error") or f"状态子进程退出码 {returncode}"))
        return [SvnStatusRecord.from_dict(item) for item in payload.get("items", [])]
    finally:
        try:
            os.remove(result_path)
        except OSError:
            pass


def scan_status(
    path: str,
    *,
    host_script: str | None = None,
    cancel_event: threading.Event | None = None,
    remote: bool = False,
    show_updates: bool = False,
) -> list[SvnStatusRecord]:
    """Return interesting recursive WC statuses, or raise instead of guessing."""
    absolute = os.path.abspath(path)
    if not os.path.exists(absolute):
        # Missing versioned files are scanned through their nearest existing
        # parent; the caller filters the exact path afterwards.
        probe = os.path.dirname(absolute)
        while probe and not os.path.exists(probe):
            parent = os.path.dirname(probe)
            if parent == probe:
                break
            probe = parent
        absolute = probe
    if not absolute or not os.path.exists(absolute):
        raise SvnStatusError(f"状态扫描路径不存在：{path}")
    if _find_svn_cli():
        return _status_via_cli(
            absolute,
            cancel_event=cancel_event,
            remote=remote,
            show_updates=show_updates,
        )
    return _status_via_tortoise_child(
        absolute,
        host_script=host_script,
        cancel_event=cancel_event,
        remote=remote or show_updates,
    )


def records_by_path(records: Iterable[SvnStatusRecord]) -> dict[str, SvnStatusRecord]:
    return {os.path.normcase(os.path.abspath(item.path)): item for item in records}


def record_for_path(path: str, records: Iterable[SvnStatusRecord]) -> SvnStatusRecord | None:
    return records_by_path(records).get(os.path.normcase(os.path.abspath(path)))
