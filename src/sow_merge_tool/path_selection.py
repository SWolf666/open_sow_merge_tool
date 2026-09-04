"""Read-only path selection and pairing rules for the comparison start centre.

The old picker deliberately accepted two files only.  This module keeps the
selection contract independent from Tk so Explorer/TortoiseSVN launches and
headless tests use exactly the same safety checks.  Nothing in this module
opens or saves an Excel workbook; the only file operation is hashing a path
when the caller explicitly asks for a re-check.
"""

from __future__ import annotations

import hashlib
import os
import sqlite3
import stat
import zipfile
from collections.abc import Iterable
from dataclasses import dataclass
from enum import Enum
from pathlib import Path

SUPPORTED_EXCEL_EXTENSIONS = (".xlsx", ".xlsm")


class PathSelectionError(ValueError):
    """A path cannot safely participate in a comparison selection."""


class PairStatus(str, Enum):
    MATCHED = "matched"
    MISSING_LEFT = "missing-left"
    MISSING_RIGHT = "missing-right"
    SAME_FILE = "same-file"
    SAME_NAME_DIFFERENT_PATH = "same-name-different-path"
    TYPE_INCOMPATIBLE = "type-incompatible"
    HASH_CHANGED = "hash-changed"
    UNSUPPORTED = "unsupported"
    INVALID = "invalid"


class DirectoryStatus(str, Enum):
    NORMAL = "normal"
    SVN_VALID = "svn-valid"
    SVN_MISMATCH = "svn-mismatch"
    SVN_LAYOUT_INVALID = "svn-layout-invalid"
    SWITCHED = "switched"
    EXTERNAL = "external"
    REPARSE = "reparse-point"
    MISSING = "missing"
    NOT_DIRECTORY = "not-directory"


@dataclass(frozen=True)
class RepositoryIdentity:
    root_url: str
    uuid: str

    @property
    def key(self) -> tuple[str, str]:
        return self.root_url.rstrip("/"), self.uuid


@dataclass(frozen=True)
class PathProbe:
    path: str
    exists: bool
    kind: str
    extension: str
    name: str
    sha256: str | None = None
    repository: RepositoryIdentity | None = None
    relative_path: str = ""
    switched: bool = False
    external: bool = False
    reparse_point: bool = False


@dataclass(frozen=True)
class FilePair:
    left: PathProbe
    right: PathProbe
    status: PairStatus
    reason: str = ""
    relative_path: str = ""

    @property
    def ready(self) -> bool:
        return self.status in {PairStatus.MATCHED, PairStatus.SAME_NAME_DIFFERENT_PATH}


@dataclass(frozen=True)
class DirectoryValidation:
    path: str
    status: DirectoryStatus
    repository: RepositoryIdentity | None = None
    top_level: tuple[str, ...] = ()
    issues: tuple[str, ...] = ()
    reparse_paths: tuple[str, ...] = ()

    @property
    def ready(self) -> bool:
        return self.status in {DirectoryStatus.NORMAL, DirectoryStatus.SVN_VALID}

    @property
    def comparison_allowed(self) -> bool:
        """Compare-only policy: repository warnings do not block safe files."""
        return self.status not in {
            DirectoryStatus.MISSING,
            DirectoryStatus.NOT_DIRECTORY,
            DirectoryStatus.REPARSE,
        }

    @property
    def warning(self) -> str:
        return "; ".join(self.issues)


@dataclass(frozen=True)
class RelativeMapping:
    relative_path: str
    left_path: str | None
    right_path: str | None
    status: PairStatus
    reason: str = ""

    @property
    def complete(self) -> bool:
        return self.status in {PairStatus.MATCHED, PairStatus.SAME_NAME_DIFFERENT_PATH}


@dataclass(frozen=True)
class SelectionSnapshot:
    """The user-approved pair and its hashes at the moment of confirmation."""

    left: PathProbe
    right: PathProbe
    mapping: RelativeMapping | None = None
    selected_at: float = 0.0


def normalize_path(path: str | os.PathLike[str]) -> str:
    value = os.path.abspath(os.path.expanduser(os.fspath(path)))
    return os.path.normcase(os.path.normpath(value))


def _within(path: str, parent: str) -> bool:
    try:
        return normalize_path(os.path.commonpath((path, parent))) == normalize_path(parent)
    except ValueError:
        return False


def sha256_path(path: str | os.PathLike[str]) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def validate_excel_package(path: str | os.PathLike[str]) -> None:
    """Validate the OOXML container at the final open boundary.

    Directory indexing intentionally does not call this function.  It is a
    cheap ZIP-member check used only after the user explicitly opens one row.
    """
    value = os.fspath(path)
    try:
        with zipfile.ZipFile(value, "r") as package:
            names = set(package.namelist())
            required = {"[Content_Types].xml", "xl/workbook.xml"}
            if not required.issubset(names):
                raise PathSelectionError(f"Excel 文件损坏或不是有效 OOXML 工作簿：{value}")
            bad = package.testzip()
            if bad:
                raise PathSelectionError(f"Excel 文件损坏（ZIP 成员无法读取）：{bad}")
    except PathSelectionError:
        raise
    except (OSError, zipfile.BadZipFile, RuntimeError) as exc:
        raise PathSelectionError(f"Excel 文件损坏或不可读：{value}") from exc


def _is_reparse(path: str) -> bool:
    try:
        attributes = getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400)
        return bool(os.lstat(path).st_file_attributes & attributes)
    except (AttributeError, OSError):
        try:
            return bool(os.path.islink(path))
        except OSError:
            return False


def _repository_identity(path: str) -> RepositoryIdentity | None:
    """Read the repository root/UUID without invoking SVN or changing state."""
    probe = os.path.abspath(path if os.path.isdir(path) else os.path.dirname(path))
    while probe:
        db = os.path.join(probe, ".svn", "wc.db")
        if os.path.isfile(db):
            try:
                with sqlite3.connect(f"file:{db}?mode=ro", uri=True) as conn:
                    row = conn.execute(
                        "select root, uuid from REPOSITORY order by id limit 1"
                    ).fetchone()
                if row and row[0] and row[1]:
                    return RepositoryIdentity(str(row[0]).rstrip("/"), str(row[1]))
            except sqlite3.Error:
                return None
        parent = os.path.dirname(probe)
        if parent == probe:
            break
        probe = parent
    return None


def inspect_path(path: str | os.PathLike[str], *, include_hash: bool = True) -> PathProbe:
    value = os.path.abspath(os.path.expanduser(os.fspath(path)))
    exists = os.path.exists(value)
    reparse = _is_reparse(value) if exists else False
    if not exists:
        kind = "missing"
    elif reparse:
        kind = "reparse"
    elif os.path.isfile(value):
        kind = "file"
    elif os.path.isdir(value):
        kind = "directory"
    else:
        kind = "other"
    extension = Path(value).suffix.lower() if kind == "file" else ""
    try:
        digest = sha256_path(value) if include_hash and kind == "file" else None
    except OSError as exc:
        raise PathSelectionError(f"文件不可读：{value}") from exc
    return PathProbe(
        path=value,
        exists=exists,
        kind=kind,
        extension=extension,
        name=os.path.basename(value),
        sha256=digest,
        repository=_repository_identity(value) if exists else None,
        reparse_point=reparse,
    )


def _svn_node_flags(directory: str) -> tuple[set[str], set[str]]:
    """Return switched and external relative paths from the read-only wc DB."""
    switched: set[str] = set()
    external: set[str] = set()
    probe = os.path.abspath(directory)
    db = ""
    wc_root = probe
    while probe:
        candidate = os.path.join(probe, ".svn", "wc.db")
        if os.path.isfile(candidate):
            db = candidate
            wc_root = probe
            break
        parent = os.path.dirname(probe)
        if parent == probe:
            break
        probe = parent
    if not db:
        return switched, external
    try:
        with sqlite3.connect(f"file:{db}?mode=ro", uri=True) as conn:
            columns = {row[1] for row in conn.execute("pragma table_info(NODES)")}
            if "switched" in columns:
                for row in conn.execute("select local_relpath from NODES where switched=1"):
                    rel = str(row[0] or "")
                    switched.add(os.path.join(wc_root, *rel.split("/")) if rel else wc_root)
            if "file_external" in columns:
                for row in conn.execute("select local_relpath from NODES where file_external=1"):
                    rel = str(row[0] or "")
                    external.add(os.path.join(wc_root, *rel.split("/")) if rel else wc_root)
    except sqlite3.Error:
        # An unreadable database is reported by validate_directory as an
        # invalid SVN selection rather than guessed to be a normal directory.
        raise PathSelectionError(f"无法读取 SVN 工作副本数据库：{db}")
    return switched, external


def validate_directory(
    path: str | os.PathLike[str],
    *,
    expected_repository: RepositoryIdentity | None = None,
    expected_top_level: Iterable[str] | None = None,
    reject_reparse: bool = True,
    status_scanner=None,
) -> DirectoryValidation:
    value = os.path.abspath(os.path.expanduser(os.fspath(path)))
    if not os.path.exists(value):
        return DirectoryValidation(value, DirectoryStatus.MISSING, issues=("目录不存在",))
    if not os.path.isdir(value):
        return DirectoryValidation(value, DirectoryStatus.NOT_DIRECTORY, issues=("不是文件夹",))
    if reject_reparse and _is_reparse(value):
        return DirectoryValidation(value, DirectoryStatus.REPARSE, issues=("目录是 reparse/symlink 节点",))

    repo = _repository_identity(value)
    issues: list[str] = []
    status = DirectoryStatus.NORMAL if repo is None else DirectoryStatus.SVN_VALID
    top_level = tuple(sorted(
        entry.name for entry in os.scandir(value)
        if entry.is_dir(follow_symlinks=False) and not entry.name.startswith(".")
    ))
    if expected_repository and repo and repo.key != expected_repository.key:
        status = DirectoryStatus.SVN_MISMATCH
        issues.append("仓库 URL 或 UUID 与另一侧不一致")
    elif expected_repository and repo is None:
        status = DirectoryStatus.SVN_MISMATCH
        issues.append("另一侧是 SVN 工作副本，本侧不是同一仓库")
    if expected_top_level:
        expected = {str(item) for item in expected_top_level if str(item)}
        if not expected.issubset(set(top_level)):
            status = DirectoryStatus.SVN_LAYOUT_INVALID
            issues.append("缺少已验证的 SVN 顶层分支目录")
    try:
        switched, external = _svn_node_flags(value)
    except PathSelectionError as exc:
        switched, external = set(), set()
        status = DirectoryStatus.SVN_LAYOUT_INVALID
        issues.append(str(exc))
    switched = {path for path in switched if _within(path, value)}
    external = {path for path in external if _within(path, value)}
    if callable(status_scanner):
        try:
            for record in status_scanner(value):
                if getattr(record, "switched", False):
                    switched.add(os.fspath(getattr(record, "path", "")))
                if getattr(record, "file_external", False) or getattr(record, "node_status", "") == "external":
                    external.add(os.fspath(getattr(record, "path", "")))
        except (OSError, RuntimeError, TypeError, ValueError) as exc:
            status = DirectoryStatus.SVN_LAYOUT_INVALID
            issues.append(f"SVN 状态扫描失败：{exc}")
    if switched:
        issues.append("工作副本包含 switched 节点")
    if external:
        issues.append("工作副本包含 svn:externals 节点")
    reparse_paths: list[str] = []
    for root, dirs, files in os.walk(value, followlinks=False):
        for name in (*dirs, *files):
            candidate = os.path.join(root, name)
            if _is_reparse(candidate):
                reparse_paths.append(candidate)
        dirs[:] = [name for name in dirs if not _is_reparse(os.path.join(root, name))]
    if reparse_paths and reject_reparse:
        issues.append("目录中包含 reparse/symlink 节点")
    return DirectoryValidation(
        value, status, repo, top_level, tuple(dict.fromkeys(issues)), tuple(reparse_paths)
    )


def _pair_reason(left: PathProbe, right: PathProbe) -> tuple[PairStatus, str]:
    if not left.exists:
        return PairStatus.MISSING_LEFT, "左侧路径不存在"
    if not right.exists:
        return PairStatus.MISSING_RIGHT, "右侧路径不存在"
    if normalize_path(left.path) == normalize_path(right.path):
        return PairStatus.SAME_FILE, "两侧选择了同一个文件"
    if left.kind != "file" or right.kind != "file":
        return PairStatus.TYPE_INCOMPATIBLE, "两侧必须都是普通 Excel 文件"
    if left.extension not in SUPPORTED_EXCEL_EXTENSIONS or right.extension not in SUPPORTED_EXCEL_EXTENSIONS:
        return PairStatus.UNSUPPORTED, "只支持 .xlsx 或 .xlsm 文件"
    if left.extension != right.extension:
        return PairStatus.TYPE_INCOMPATIBLE, "两侧 Excel 类型不兼容（.xlsx/.xlsm 不一致）"
    if left.name.lower() == right.name.lower():
        return PairStatus.SAME_NAME_DIFFERENT_PATH, "文件名相同但路径不同，已由用户明确选择"
    return PairStatus.MATCHED, "已匹配；路径不同不会按文件名自动猜测"


def validate_file_pair(
    left_path: str | os.PathLike[str],
    right_path: str | os.PathLike[str],
    *,
    expected_repository: RepositoryIdentity | None = None,
    include_hash: bool = True,
    strict_repository: bool = False,
) -> FilePair:
    left = inspect_path(left_path, include_hash=include_hash)
    right = inspect_path(right_path, include_hash=include_hash)
    status, reason = _pair_reason(left, right)
    if status in {PairStatus.MATCHED, PairStatus.SAME_NAME_DIFFERENT_PATH}:
        if strict_repository and expected_repository and left.repository and right.repository and left.repository.key != expected_repository.key:
            status, reason = PairStatus.TYPE_INCOMPATIBLE, "两侧不在同一 SVN 仓库"
        elif strict_repository and expected_repository and (left.repository or right.repository) != expected_repository:
            status, reason = PairStatus.TYPE_INCOMPATIBLE, "路径不属于预期 SVN 仓库"
    return FilePair(left, right, status, reason)


def _iter_excel_files(directory: str, *, excluded_paths: Iterable[str] = ()) -> dict[str, str]:
    result: dict[str, str] = {}
    excluded = tuple(os.path.abspath(os.fspath(path)) for path in excluded_paths if path)
    for root, dirs, files in os.walk(directory, followlinks=False):
        if any(_within(root, path) for path in excluded):
            dirs[:] = []
            continue
        dirs[:] = [name for name in dirs if name != ".svn" and not _is_reparse(os.path.join(root, name))]
        dirs[:] = [
            name for name in dirs
            if not any(_within(os.path.join(root, name), path) for path in excluded)
        ]
        for name in files:
            if Path(name).suffix.lower() not in SUPPORTED_EXCEL_EXTENSIONS:
                continue
            path = os.path.join(root, name)
            if _is_reparse(path):
                continue
            if any(_within(path, excluded_path) for excluded_path in excluded):
                continue
            relative = os.path.relpath(path, directory).replace("\\", "/")
            result[relative.casefold()] = path
    return result


def map_relative_files(
    left_directory: str | os.PathLike[str],
    right_directory: str | os.PathLike[str],
    *,
    relative_paths: Iterable[str] | None = None,
    include_hash: bool = False,
) -> list[RelativeMapping]:
    """Map only equal relative paths; never select by basename alone."""
    left_validation = validate_directory(left_directory)
    right_validation = validate_directory(
        right_directory, expected_repository=left_validation.repository
    )
    if left_validation.repository is None and right_validation.repository is not None:
        right_validation = DirectoryValidation(
            right_validation.path,
            DirectoryStatus.SVN_MISMATCH,
            right_validation.repository,
            right_validation.top_level,
            right_validation.issues + ("另一侧不是 SVN 工作副本，不能自动按仓库映射",),
            right_validation.reparse_paths,
        )
    if not left_validation.comparison_allowed or not right_validation.comparison_allowed:
        raise PathSelectionError(
            "目录不能用于比较：" + "; ".join(left_validation.issues + right_validation.issues)
        )
    _left_switched, left_external = _svn_node_flags(left_validation.path)
    _right_switched, right_external = _svn_node_flags(right_validation.path)
    left_files = _iter_excel_files(left_validation.path, excluded_paths=left_external)
    right_files = _iter_excel_files(right_validation.path, excluded_paths=right_external)
    keys = (
        {str(value).replace("\\", "/").casefold() for value in relative_paths}
        if relative_paths is not None
        else set(left_files) | set(right_files)
    )
    result: list[RelativeMapping] = []
    for key in sorted(keys):
        left = left_files.get(key)
        right = right_files.get(key)
        rel = (os.path.relpath(left or right or key, left_validation.path if left else right_validation.path)).replace("\\", "/")
        if left and right:
            pair = validate_file_pair(left, right, include_hash=include_hash)
            result.append(RelativeMapping(rel, left, right, pair.status, pair.reason))
        elif left:
            result.append(RelativeMapping(rel, left, None, PairStatus.MISSING_RIGHT, "右侧缺少相同相对路径文件"))
        else:
            result.append(RelativeMapping(rel, None, right, PairStatus.MISSING_LEFT, "左侧缺少相同相对路径文件"))
    return result


def override_mapping(
    mapping: RelativeMapping,
    *,
    left_path: str | os.PathLike[str] | None = None,
    right_path: str | os.PathLike[str] | None = None,
) -> RelativeMapping:
    """Apply an explicit per-row file choice without basename inference."""
    left = os.fspath(left_path) if left_path is not None else mapping.left_path
    right = os.fspath(right_path) if right_path is not None else mapping.right_path
    if not left or not right:
        status = PairStatus.MISSING_LEFT if not left else PairStatus.MISSING_RIGHT
        return RelativeMapping(mapping.relative_path, left, right, status, "仍缺少一侧明确文件")
    pair = validate_file_pair(left, right, include_hash=False)
    return RelativeMapping(mapping.relative_path, pair.left.path, pair.right.path, pair.status, pair.reason)


def snapshot_selection(
    left_path: str | os.PathLike[str],
    right_path: str | os.PathLike[str],
    *,
    relative_path: str = "",
) -> SelectionSnapshot:
    pair = validate_file_pair(left_path, right_path)
    if not pair.ready:
        raise PathSelectionError(pair.reason)
    mapping = RelativeMapping(relative_path, pair.left.path, pair.right.path, pair.status, pair.reason) if relative_path else None
    return SelectionSnapshot(pair.left, pair.right, mapping)


def recheck_selection(snapshot: SelectionSnapshot) -> SelectionSnapshot:
    """Re-read existence/type/hash before opening; fail closed on any drift."""
    current = validate_file_pair(snapshot.left.path, snapshot.right.path)
    if not current.ready:
        raise PathSelectionError(f"选择已失效：{current.reason}")
    for label, before, after in (("左侧", snapshot.left, current.left), ("右侧", snapshot.right, current.right)):
        if before.sha256 and before.sha256 != after.sha256:
            raise PathSelectionError(f"{label}文件在确认后发生变化，请重新选择")
    return SelectionSnapshot(current.left, current.right, snapshot.mapping, snapshot.selected_at)


def compare_selection_from_paths(paths: Iterable[str | os.PathLike[str]]) -> tuple[str, ...]:
    """Normalize explicit Explorer prefill paths without selecting anything."""
    result: list[str] = []
    seen: set[str] = set()
    for raw in paths:
        value = os.path.abspath(os.path.expanduser(os.fspath(raw)))
        key = normalize_path(value)
        if value and key not in seen:
            result.append(value)
            seen.add(key)
    return tuple(result)


__all__ = [
    "SUPPORTED_EXCEL_EXTENSIONS",
    "DirectoryStatus",
    "DirectoryValidation",
    "FilePair",
    "PairStatus",
    "PathProbe",
    "PathSelectionError",
    "RelativeMapping",
    "RepositoryIdentity",
    "SelectionSnapshot",
    "compare_selection_from_paths",
    "inspect_path",
    "map_relative_files",
    "normalize_path",
    "override_mapping",
    "recheck_selection",
    "sha256_path",
    "snapshot_selection",
    "validate_directory",
    "validate_excel_package",
    "validate_file_pair",
]
