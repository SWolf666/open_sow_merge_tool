"""Read-only, structured Subversion history for audit and reconciliation.

The commit-message history stored by TortoiseSVN is useful as an input
convenience, but it is not repository evidence.  This module talks to the
Subversion command line client (including the client bundled with
TortoiseSVN) and parses ``svn log --xml -v`` into immutable records.  It never
updates, commits, reverts, or interprets a log message as an instruction.
"""

from __future__ import annotations

import ctypes
import json
import os
import re
import sys
import tempfile
import threading
import uuid
import xml.etree.ElementTree as ET
from collections.abc import Callable, Iterable
from dataclasses import asdict, dataclass
from typing import ClassVar

from .svn_status_provider import (
    _communicate_cancelable,
    _find_svn_cli,
    _find_tortoise_bin,
    _native_error,
)

_URL_RE = re.compile(r"^[A-Za-z][A-Za-z0-9+.-]*://")


class SvnLogError(RuntimeError):
    """Raised when repository history cannot be read or verified."""


@dataclass(frozen=True)
class SvnLogPath:
    """One repository path changed by a revision."""

    path: str
    action: str = ""
    node_kind: str = ""
    text_modified: bool | None = None
    properties_modified: bool | None = None
    copyfrom_path: str = ""
    copyfrom_revision: int | None = None


# Compatibility spelling for callers that describe the path collection as
# "changed paths" rather than SVN's ``<path>`` element.
SvnChangedPath = SvnLogPath


class _SvnOptRevisionValue(ctypes.Union):
    _fields_: ClassVar = [("number", ctypes.c_long), ("date", ctypes.c_longlong)]


class _SvnOptRevision(ctypes.Structure):
    _fields_ = [("kind", ctypes.c_int), ("value", _SvnOptRevisionValue)]


class _SvnLogChangedPath(ctypes.Structure):
    _fields_ = [
        ("action", ctypes.c_char),
        ("copyfrom_path", ctypes.c_char_p),
        ("copyfrom_rev", ctypes.c_long),
    ]


@dataclass(frozen=True)
class SvnLogEntry:
    """A revision and its structured metadata from the repository."""

    revision: int
    author: str = ""
    date: str = ""
    message: str = ""
    changed_paths: tuple[SvnLogPath, ...] = ()
    mergeinfo: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, object]:
        """Return a JSON-safe representation for a batch journal."""
        return asdict(self)


def _local_name(tag: str) -> str:
    return str(tag).rsplit("}", 1)[-1]


def _child_text(node: ET.Element, name: str) -> str:
    for child in list(node):
        if _local_name(child.tag) == name:
            return str(child.text or "").strip()
    return ""


def _optional_bool(value: str | None) -> bool | None:
    if value is None or str(value).strip() == "":
        return None
    normalized = str(value).strip().lower()
    if normalized in {"true", "1", "yes"}:
        return True
    if normalized in {"false", "0", "no"}:
        return False
    return None


def _optional_revision(value: str | None) -> int | None:
    try:
        return int(str(value)) if value is not None and str(value).strip() else None
    except (TypeError, ValueError):
        return None


def _mergeinfo_values(entry: ET.Element) -> tuple[str, ...]:
    values: list[str] = []
    for node in entry.iter():
        name = _local_name(node.tag)
        property_name = str(node.get("name") or "").strip().lower()
        if name.lower() == "mergeinfo" or property_name == "svn:mergeinfo":
            text = "".join(node.itertext()).strip()
            if text:
                values.append(text)
        elif name.lower() == "path" and "mergeinfo" in node.attrib:
            text = str(node.get("mergeinfo") or "").strip()
            if text:
                values.append(text)
    return tuple(dict.fromkeys(values))


def parse_svn_log_xml(xml_text: str) -> list[SvnLogEntry]:
    """Parse ``svn log --xml -v`` output without making any filesystem call."""
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as exc:
        raise SvnLogError(f"SVN 日志 XML 无效：{exc}") from exc

    result: list[SvnLogEntry] = []
    for entry in root.iter():
        if _local_name(entry.tag) != "logentry":
            continue
        revision = _optional_revision(entry.get("revision"))
        if revision is None:
            raise SvnLogError("SVN 日志缺少有效 revision")
        changed: list[SvnLogPath] = []
        for path_node in entry.iter():
            if _local_name(path_node.tag) != "path":
                continue
            path = str("".join(path_node.itertext()) or "").strip()
            if not path:
                continue
            changed.append(
                SvnLogPath(
                    path=path,
                    action=str(path_node.get("action") or ""),
                    node_kind=str(path_node.get("kind") or ""),
                    text_modified=_optional_bool(path_node.get("text-mods")),
                    properties_modified=_optional_bool(path_node.get("prop-mods")),
                    copyfrom_path=str(path_node.get("copyfrom-path") or ""),
                    copyfrom_revision=_optional_revision(path_node.get("copyfrom-rev")),
                )
            )
        result.append(
            SvnLogEntry(
                revision=revision,
                author=_child_text(entry, "author"),
                date=_child_text(entry, "date"),
                message=_child_text(entry, "msg"),
                changed_paths=tuple(changed),
                mergeinfo=_mergeinfo_values(entry),
            )
        )
    return result


def _native_decode(value: bytes | None) -> str:
    return value.decode("utf-8", errors="replace") if value else ""


def _native_hash_items(apr, hash_pointer: int | None):
    if not hash_pointer:
        return
    apr.apr_hash_first.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    apr.apr_hash_first.restype = ctypes.c_void_p
    apr.apr_hash_next.argtypes = [ctypes.c_void_p]
    apr.apr_hash_next.restype = ctypes.c_void_p
    apr.apr_hash_this.argtypes = [
        ctypes.c_void_p,
        ctypes.POINTER(ctypes.c_void_p),
        ctypes.POINTER(ctypes.c_ssize_t),
        ctypes.POINTER(ctypes.c_void_p),
    ]
    apr.apr_hash_this.restype = None
    index = apr.apr_hash_first(None, hash_pointer)
    while index:
        key = ctypes.c_void_p()
        key_length = ctypes.c_ssize_t()
        value = ctypes.c_void_p()
        apr.apr_hash_this(index, ctypes.byref(key), ctypes.byref(key_length), ctypes.byref(value))
        yield key, int(key_length.value), value
        index = apr.apr_hash_next(index)


def _native_changed_paths(apr, pointer: int | None) -> list[SvnLogPath]:
    result: list[SvnLogPath] = []
    for key, key_length, value in _native_hash_items(apr, pointer):
        if not key or not value:
            continue
        path = ctypes.string_at(key, key_length).decode("utf-8", errors="replace")
        try:
            changed = ctypes.cast(value, ctypes.POINTER(_SvnLogChangedPath)).contents
            action = changed.action.decode("ascii", errors="replace") if changed.action else ""
            copyfrom_path = _native_decode(changed.copyfrom_path)
            copyfrom_revision = int(changed.copyfrom_rev) if int(changed.copyfrom_rev) >= 0 else None
            result.append(
                SvnLogPath(
                    path=path,
                    action=action,
                    copyfrom_path=copyfrom_path,
                    copyfrom_revision=copyfrom_revision,
                )
            )
        except (ValueError, OSError):
            continue
    return result


def _apr_array_with_strings(apr, pool: int, values: Iterable[bytes]):
    values = list(values)
    apr.apr_array_make.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int]
    apr.apr_array_make.restype = ctypes.c_void_p
    apr.apr_array_push.argtypes = [ctypes.c_void_p]
    apr.apr_array_push.restype = ctypes.c_void_p
    array = apr.apr_array_make(pool, max(1, len(values)), ctypes.sizeof(ctypes.c_void_p))
    keepalive: list[bytes] = []
    for value in values:
        raw = bytes(value)
        keepalive.append(raw)
        slot = apr.apr_array_push(array)
        ctypes.cast(slot, ctypes.POINTER(ctypes.c_char_p))[0] = raw
    return array, keepalive


def query_tortoise_log_in_child(
    path: str,
    *,
    limit: int = 25,
    start_revision: int | None = None,
    end_revision: int | None = None,
) -> list[SvnLogEntry]:
    """Read logs through TortoiseSVN's bundled libsvn in an isolated child.

    The C callback and APR hash traversal stay in the helper process.  A bad
    ABI/runtime cannot bring down the workbench; the parent receives a clear
    ``SvnLogError`` instead.
    """
    bin_dir = _find_tortoise_bin()
    if not bin_dir or not hasattr(os, "add_dll_directory") or not hasattr(ctypes, "WinDLL"):
        raise SvnLogError("TortoiseSVN 日志运行库不可用")
    dll_cookie = os.add_dll_directory(bin_dir)
    apr_ready = False
    pool = ctypes.c_void_p()
    callbacks = []
    keepalive: list[bytes] = []
    stage = "load-runtime"
    try:
        apr = ctypes.WinDLL(os.path.join(bin_dir, "libapr_tsvn.dll"))
        svn = ctypes.WinDLL(os.path.join(bin_dir, "libsvn_tsvn.dll"))
        apr.apr_initialize.argtypes = []
        apr.apr_initialize.restype = ctypes.c_int
        if int(apr.apr_initialize()):
            raise SvnLogError("APR 初始化失败")
        apr_ready = True
        apr.apr_pool_create_ex.argtypes = [
            ctypes.POINTER(ctypes.c_void_p),
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_void_p,
        ]
        apr.apr_pool_create_ex.restype = ctypes.c_int
        if int(apr.apr_pool_create_ex(ctypes.byref(pool), None, None, None)):
            raise SvnLogError("APR 内存池创建失败")
        stage = "create-client-context"
        _native_error(svn, svn.svn_dso_initialize2(), "SVN DSO 初始化失败")
        context = ctypes.c_void_p()
        svn.svn_client_create_context2.argtypes = [
            ctypes.POINTER(ctypes.c_void_p),
            ctypes.c_void_p,
            ctypes.c_void_p,
        ]
        svn.svn_client_create_context2.restype = ctypes.c_void_p
        _native_error(
            svn,
            svn.svn_client_create_context2(ctypes.byref(context), None, pool),
            "SVN 客户端初始化失败",
        )

        stage = "build-target-array"
        targets, target_keepalive = _apr_array_with_strings(
            apr,
            pool,
            [_target_for_svn(path).encode("utf-8")],
        )
        keepalive.extend(target_keepalive)
        peg_revision = _SvnOptRevision()
        peg_revision.kind = 7  # svn_opt_revision_head
        start = _SvnOptRevision()
        start.kind = 1
        start.value.number = 1 if start_revision is None else int(start_revision)
        end = _SvnOptRevision()
        if end_revision is None:
            end.kind = 7  # svn_opt_revision_head
        else:
            end.kind = 1
            end.value.number = int(end_revision)
        records: list[SvnLogEntry] = []
        callback_type = ctypes.CFUNCTYPE(
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_long,
            ctypes.c_char_p,
            ctypes.c_char_p,
            ctypes.c_char_p,
            ctypes.c_void_p,
        )

        def receive(_baton, changed_paths, revision, author, date, message, _scratch):
            try:
                records.append(
                    SvnLogEntry(
                        revision=int(revision),
                        author=_native_decode(author),
                        date=_native_decode(date),
                        message=_native_decode(message),
                        changed_paths=_native_changed_paths(apr, changed_paths),
                    )
                )
            except (
                ctypes.ArgumentError,
                AttributeError,
                IndexError,
                OSError,
                RuntimeError,
                TypeError,
                ValueError,
            ) as exc:
                # Callback exceptions cannot cross the C ABI.  The sentinel
                # revision is converted to a normal provider error below.
                records.append(
                    SvnLogEntry(revision=-1, message=f"callback-error: {exc}")
                )

        callback = callback_type(receive)
        callbacks.append(callback)
        stage = "call-client-log3"
        svn.svn_client_log3.argtypes = [
            ctypes.c_void_p,
            ctypes.POINTER(_SvnOptRevision),
            ctypes.POINTER(_SvnOptRevision),
            ctypes.POINTER(_SvnOptRevision),
            ctypes.c_int,
            ctypes.c_int,
            ctypes.c_int,
            callback_type,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_void_p,
        ]
        svn.svn_client_log3.restype = ctypes.c_void_p
        error = svn.svn_client_log3(
            targets,
            ctypes.byref(peg_revision),
            ctypes.byref(start),
            ctypes.byref(end),
            max(1, min(500, int(limit))),
            0,
            0,
            callback,
            None,
            context,
            pool,
        )
        _native_error(svn, error, "TortoiseSVN 日志读取失败")
        callback_errors = [item for item in records if item.revision < 0]
        if callback_errors:
            raise SvnLogError(callback_errors[0].message)
        records.sort(key=lambda item: item.revision, reverse=True)
        return records[: max(1, min(500, int(limit)))]
    except SvnLogError:
        raise
    except (AttributeError, ctypes.ArgumentError, IndexError, OSError, RuntimeError, TypeError, ValueError) as exc:
        raise SvnLogError(f"TortoiseSVN 日志读取失败（阶段 {stage}）：{exc}") from exc
    finally:
        callbacks.clear()
        if pool:
            try:
                apr.apr_pool_destroy.argtypes = [ctypes.c_void_p]
                apr.apr_pool_destroy(pool)
            except (AttributeError, OSError):
                pass
        if apr_ready:
            try:
                apr.apr_terminate()
            except (AttributeError, OSError):
                pass
        dll_cookie.close()


def _target_for_svn(path: str) -> str:
    value = os.fspath(path)
    if _URL_RE.match(value) or value.startswith("^"):
        return value
    return os.path.abspath(value)


def _entry_from_dict(value: dict) -> SvnLogEntry:
    changed = []
    for raw in value.get("changed_paths", ()):
        if isinstance(raw, dict):
            fields = SvnLogPath.__dataclass_fields__
            changed.append(SvnLogPath(**{key: raw[key] for key in fields if key in raw}))
    fields = SvnLogEntry.__dataclass_fields__
    payload = {key: value[key] for key in fields if key in value}
    payload["changed_paths"] = tuple(changed)
    payload["mergeinfo"] = tuple(str(item) for item in payload.get("mergeinfo", ()) if item)
    return SvnLogEntry(**payload)


def _log_via_tortoise_child(
    path: str,
    *,
    limit: int,
    start_revision: int | None,
    end_revision: int | None,
    cancel_event: threading.Event | None,
) -> list[SvnLogEntry]:
    """Ask the isolated TortoiseSVN-runtime helper for structured JSON."""
    result_path = os.path.join(
        tempfile.gettempdir(),
        f"sow_svn_log_{os.getpid()}_{uuid.uuid4().hex}.json",
    )
    command = [sys.executable]
    if not getattr(sys, "frozen", False):
        command.extend(("-m", "sow_merge_tool"))
    command.extend(
        (
            "--internal-svn-log-query",
            _target_for_svn(path),
            result_path,
            "--limit",
            str(limit),
        )
    )
    if start_revision is not None:
        command.extend(("--start-revision", str(int(start_revision))))
    if end_revision is not None:
        command.extend(("--end-revision", str(int(end_revision))))
    try:
        returncode, _stdout, _stderr = _communicate_cancelable(
            command,
            timeout=240,
            cancel_event=cancel_event,
            capture=False,
        )
        try:
            with open(result_path, "r", encoding="utf-8") as stream:
                payload = json.load(stream)
        except (OSError, ValueError) as exc:
            raise SvnLogError(
                f"TortoiseSVN 日志子进程没有返回有效结果（退出码 {returncode}）"
            ) from exc
        if returncode != 0 or not payload.get("ok"):
            raise SvnLogError(str(payload.get("error") or f"日志子进程退出码 {returncode}"))
        return [_entry_from_dict(item) for item in payload.get("entries", ()) if isinstance(item, dict)]
    except SvnLogError:
        raise
    except (AttributeError, OSError, RuntimeError, TypeError, ValueError) as exc:
        raise SvnLogError(f"TortoiseSVN 日志子进程启动失败：{exc}") from exc
    finally:
        try:
            os.remove(result_path)
        except OSError:
            pass


def internal_log_entrypoint(argv: list[str]) -> int:
    """Entry point used only by the hidden TortoiseSVN runtime subprocess."""
    if len(argv) < 2:
        return 2
    path, output_path = argv[:2]
    options = {str(argv[index]): argv[index + 1] for index in range(2, len(argv) - 1, 2)}
    try:
        entries = query_tortoise_log_in_child(
            path,
            limit=int(options.get("--limit", 25)),
            start_revision=(
                int(options["--start-revision"])
                if options.get("--start-revision") is not None else None
            ),
            end_revision=(
                int(options["--end-revision"])
                if options.get("--end-revision") is not None else None
            ),
        )
        payload = {"ok": True, "entries": [entry.to_dict() for entry in entries]}
    except (AttributeError, OSError, RuntimeError, TypeError, ValueError, SvnLogError) as exc:
        payload = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
    try:
        temporary = output_path + f".tmp-{os.getpid()}"
        with open(temporary, "w", encoding="utf-8") as stream:
            json.dump(payload, stream, ensure_ascii=False)
        os.replace(temporary, output_path)
    except OSError:
        return 3
    return 0 if payload.get("ok") else 1


def _runner_result(value) -> tuple[int, str, str]:
    if isinstance(value, tuple) and len(value) == 3:
        return int(value[0]), str(value[1] or ""), str(value[2] or "")
    return (
        int(getattr(value, "returncode", 1)),
        str(getattr(value, "stdout", "") or ""),
        str(getattr(value, "stderr", "") or ""),
    )


def _run_log_command(
    command: list[str],
    *,
    runner: Callable | None,
    cancel_event: threading.Event | None,
) -> tuple[int, str, str]:
    if runner is None:
        try:
            return _communicate_cancelable(
                command,
                timeout=180,
                cancel_event=cancel_event,
                capture=True,
            )
        except Exception as exc:
            if isinstance(exc, SvnLogError):
                raise
            raise SvnLogError(f"SVN 日志读取启动失败：{exc}") from exc
    try:
        try:
            value = runner(command, timeout=180, cancel_event=cancel_event)
        except TypeError:
            value = runner(command)
        return _runner_result(value)
    except Exception as exc:
        if isinstance(exc, SvnLogError):
            raise
        raise SvnLogError(f"SVN 日志读取启动失败：{exc}") from exc


def read_svn_log(
    path: str,
    *,
    limit: int = 25,
    start_revision: int | None = None,
    end_revision: int | None = None,
    cancel_event: threading.Event | None = None,
    runner: Callable | None = None,
) -> list[SvnLogEntry]:
    """Read repository history for one path using a read-only SVN command.

    ``runner`` is intentionally injectable for headless tests and a host's
    TortoiseSVN runtime adapter.  Its result may be a CompletedProcess-like
    object or ``(returncode, stdout, stderr)``.  No caller should use the log
    message to select or synthesize a content operation.
    """
    try:
        safe_limit = max(1, min(500, int(limit)))
    except (TypeError, ValueError):
        raise SvnLogError("SVN 日志条数必须是正整数") from None
    svn = _find_svn_cli() if runner is None else "injected-svn"
    if runner is None and not svn:
        return _log_via_tortoise_child(
            path,
            limit=safe_limit,
            start_revision=start_revision,
            end_revision=end_revision,
            cancel_event=cancel_event,
        )
    command = [str(svn), "log", "--xml", "-v", "--limit", str(safe_limit)]
    if start_revision is not None or end_revision is not None:
        start = "1" if start_revision is None else str(int(start_revision))
        end = "HEAD" if end_revision is None else str(int(end_revision))
        command.extend(("-r", f"{start}:{end}"))
    command.append(_target_for_svn(path))
    returncode, stdout, stderr = _run_log_command(
        command,
        runner=runner,
        cancel_event=cancel_event,
    )
    if returncode != 0:
        detail = (stderr or stdout or "").strip()
        raise SvnLogError(f"SVN 日志读取失败：{detail or f'退出码 {returncode}'}")
    return parse_svn_log_xml(stdout)


class SvnLogProvider:
    """Small callable provider used by the workbench and injected tests."""

    def __init__(self, *, runner: Callable | None = None):
        self.runner = runner

    def read(self, path: str, **kwargs) -> list[SvnLogEntry]:
        return read_svn_log(path, runner=self.runner, **kwargs)

    __call__ = read


read_structured_svn_log = read_svn_log


__all__ = [
    "SvnChangedPath",
    "SvnLogEntry",
    "SvnLogError",
    "SvnLogPath",
    "SvnLogProvider",
    "parse_svn_log_xml",
    "read_structured_svn_log",
    "read_svn_log",
]
