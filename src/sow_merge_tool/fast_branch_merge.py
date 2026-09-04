"""Fast, conservative source-change projection for multi-branch submission.

The source delta is calculated once from ``source-before -> source-after`` and
then replayed onto each target workbook.  Unrelated target content is always
retained: a modified target file is never replaced by the source workbook.

Analysis is read-only.  Candidate materialization is explicit and writes a
copy derived from the target workbook into the batch artifact directory.
"""

from __future__ import annotations

import hashlib
import os
import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from functools import lru_cache

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
M = "{" + MAIN_NS + "}"
_EXCEL_ESCAPE_RE = re.compile(r"_x([0-9A-Fa-f]{4})_")


@dataclass(frozen=True)
class FastCell:
    value: object
    formula: str = ""
    style: str = ""
    cell_type: str = ""


@dataclass(frozen=True)
class FastRow:
    number: int
    key: str
    cells: dict[int, FastCell]
    style: str = ""


@dataclass
class FastSheet:
    name: str
    member: str
    rows: dict[int, FastRow] = field(default_factory=dict)
    key_rows: dict[str, int] = field(default_factory=dict)
    duplicate_keys: set[str] = field(default_factory=set)
    headers: tuple[str, ...] = ()
    changed_members: set[str] = field(default_factory=set)


@dataclass
class FastWorkbook:
    path: str
    signature: tuple[str, str, int | None]
    members: dict[str, tuple[int, int]]
    payloads: dict[str, bytes]
    sheets: dict[str, FastSheet]
    shared_strings: tuple[str, ...]


@dataclass
class FastSourceDelta:
    source_path: str
    before_path: str
    changed_sheets: dict[str, dict] = field(default_factory=dict)
    unsupported_reason: str = ""
    incoming_count: int = 0
    changed_count: int = 0
    before_revision: int | None = None
    source_revision: int | None = None


@dataclass
class FastTargetDecision:
    disposition: str
    reason: str = ""
    summary: dict = field(default_factory=dict)
    details: list[dict] = field(default_factory=list)


@lru_cache(maxsize=256)
def _content_digest(path: str, size: int, modified_ns: int) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _signature(path: str, revision: int | None = None) -> tuple[str, str, int | None]:
    info = os.stat(path)
    absolute = os.path.normcase(os.path.abspath(path))
    digest = _content_digest(absolute, int(info.st_size), int(info.st_mtime_ns))
    return absolute, digest, None if revision is None else int(revision)


def _text(node: ET.Element | None) -> str:
    if node is None:
        return ""
    return "".join(part.text or "" for part in node.iter(M + "t"))


def _normalize_excel_text(value: object) -> object:
    """Normalize equivalent text emitted by different Excel versions.

    Excel may persist carriage returns either as a literal CRLF pair or as the
    OOXML escape ``_x000D_`` followed by LF.  Treating those representations as
    different creates a large false source delta after an otherwise harmless
    save in another Office build.
    """
    if not isinstance(value, str):
        return value
    value = _EXCEL_ESCAPE_RE.sub(lambda match: chr(int(match.group(1), 16)), value)
    return value.replace("\r\n", "\n").replace("\r", "\n")


def _parse_xml(payload: bytes) -> ET.Element:
    # openpyxl can persist one logical CRLF as CR+CRLF.  XML parsers normalize
    # both carriage returns to LF before the application sees them, creating a
    # false blank line.  Preserve that producer-specific sequence as one CRLF
    # without collapsing intentional consecutive newlines in user text.
    return ET.fromstring(payload.replace(b"\r\r\n", b"&#13;&#10;"))


def _shared_strings(payload: bytes | None) -> tuple[str, ...]:
    if not payload:
        return ()
    root = _parse_xml(payload)
    return tuple(_text(item) for item in root.findall(M + "si"))


def _cell_value(cell: ET.Element, strings: tuple[str, ...]) -> FastCell:
    formula = (cell.findtext(M + "f") or "").strip()
    value_node = cell.find(M + "v")
    value = value_node.text if value_node is not None else None
    cell_type = cell.attrib.get("t", "")
    if cell_type == "s" and value is not None:
        try:
            value = strings[int(value)]
        except (ValueError, IndexError):
            value = None
    elif cell_type == "inlineStr":
        value = _text(cell.find(M + "is"))
    elif value is not None and cell_type not in {"str", "e"}:
        # Preserve numeric text without an expensive type coercion.  The
        # semantic comparison only needs stable equality for the same OOXML
        # value representation.
        value = value.strip()
    return FastCell(value=value, formula=formula, style=cell.attrib.get("s", ""), cell_type=cell_type)


def _column_number(reference: str) -> int:
    letters = "".join(char for char in reference if char.isalpha())
    result = 0
    for char in letters.upper():
        result = result * 26 + ord(char) - 64
    return result


def _sheet_rows(payload: bytes, strings: tuple[str, ...]) -> dict[int, FastRow]:
    root = _parse_xml(payload)
    rows: dict[int, FastRow] = {}
    for row in root.findall(".//" + M + "row"):
        try:
            number = int(row.attrib.get("r", "0"))
        except ValueError:
            continue
        cells: dict[int, FastCell] = {}
        for cell in row.findall(M + "c"):
            ref = cell.attrib.get("r", "")
            if not ref:
                continue
            column = _column_number(ref)
            if column:
                cells[column] = _cell_value(cell, strings)
        key_cell = cells.get(1)
        key = str(_normalize_excel_text(key_cell.value)).strip() if key_cell and key_cell.value not in (None, "") else ""
        rows[number] = FastRow(number, key, cells, row.attrib.get("s", ""))
    return rows


def _sheet_map(payloads: dict[str, bytes]) -> dict[str, str]:
    workbook = ET.fromstring(payloads["xl/workbook.xml"])
    rels = ET.fromstring(payloads["xl/_rels/workbook.xml.rels"])
    targets = {
        item.attrib.get("Id", ""): item.attrib.get("Target", "")
        for item in rels.findall("{" + PACKAGE_REL_NS + "}Relationship")
    }
    result: dict[str, str] = {}
    for sheet in workbook.findall(".//" + M + "sheet"):
        name = sheet.attrib.get("name", "")
        target = targets.get(sheet.attrib.get("{" + REL_NS + "}id", ""), "")
        target = target.lstrip("/")
        target = target if target.startswith("xl/") else os.path.normpath(os.path.join("xl", target)).replace("\\", "/")
        if name and target in payloads:
            result[name] = target
    return result


@lru_cache(maxsize=32)
def _load_cached(signature: tuple[str, str, int | None]) -> FastWorkbook:
    path = signature[0]
    with zipfile.ZipFile(path, "r") as package:
        names = [name for name in package.namelist() if not name.endswith("/")]
        payloads = {name: package.read(name) for name in names}
    strings = _shared_strings(payloads.get("xl/sharedStrings.xml"))
    sheets: dict[str, FastSheet] = {}
    for name, member in _sheet_map(payloads).items():
        rows = _sheet_rows(payloads[member], strings)
        keys: dict[str, int] = {}
        duplicates: set[str] = set()
        key_counts: dict[str, int] = {}
        for number, row in rows.items():
            if number > 1 and row.key and row.key not in {"string", "唯一值", "type"}:
                key_counts[row.key] = key_counts.get(row.key, 0) + 1
        for number, row in rows.items():
            if number <= 1 or not row.key or row.key in {"string", "唯一值", "type"}:
                continue
            if key_counts.get(row.key, 0) > 1:
                duplicates.add(row.key)
                continue
            if row.key in keys:
                duplicates.add(row.key)
            else:
                keys[row.key] = number
        headers = tuple(
            str(_normalize_excel_text(rows.get(1, FastRow(1, "", {})).cells[column].value or ""))
            for column in sorted(rows.get(1, FastRow(1, "", {})).cells)
        )
        sheets[name] = FastSheet(name, member, rows, keys, duplicates, headers)
    members = {name: (info.file_size, info.CRC) for name, info in _zip_infos(path)}
    return FastWorkbook(path, signature, members, payloads, sheets, strings)


def _zip_infos(path: str):
    with zipfile.ZipFile(path, "r") as package:
        return [(info.filename, info) for info in package.infolist() if not info.filename.endswith("/")]


def load_workbook_index(path: str, *, revision: int | None = None) -> FastWorkbook:
    """Load one immutable index keyed by bytes, mtime and optional SVN rev."""
    return _load_cached(_signature(path, revision))


def _cell_equal(left: FastCell | None, right: FastCell | None) -> bool:
    def fingerprint(cell: FastCell | None) -> tuple:
        if cell is None or (not cell.formula and cell.value in (None, "")):
            return ("blank",)
        if cell.formula:
            return ("formula", _normalize_excel_text(cell.formula))
        if cell.cell_type in {"s", "inlineStr", "str"}:
            return ("text", _normalize_excel_text(cell.value))
        if cell.cell_type in {"", "n"}:
            return ("number", cell.value)
        return (cell.cell_type, cell.value)
    return fingerprint(left) == fingerprint(right)


def _sheet_semantic_equal(left: FastSheet, right: FastSheet) -> bool:
    """Compare only cell values/formulas, ignoring save-time presentation churn."""
    if set(left.rows) != set(right.rows):
        return False
    for row_number in left.rows:
        left_row, right_row = left.rows[row_number], right.rows[row_number]
        columns = set(left_row.cells) | set(right_row.cells)
        if any(not _cell_equal(left_row.cells.get(column), right_row.cells.get(column)) for column in columns):
            return False
    return True


def _cell_display(cell: FastCell | None) -> str:
    if cell is None:
        return "（空）"
    if cell.formula:
        return "=" + cell.formula.lstrip("=")
    if cell.value in (None, ""):
        return "（空）"
    return str(cell.value)


def _header(sheet: FastSheet, column: int) -> str:
    if 0 < column <= len(sheet.headers):
        return str(sheet.headers[column - 1] or "")
    return ""


def _mapped_column(source: FastSheet, target: FastSheet, source_column: int) -> int | None:
    source_header = _header(source, source_column)
    if source_header:
        matches = [column for column, header in enumerate(target.headers, 1) if header == source_header]
        return matches[0] if len(matches) == 1 else None
    if source.headers == target.headers and source_column <= len(target.headers):
        return source_column
    return None


def _row_matches_source_fields(
    source_sheet: FastSheet,
    source_row: FastRow,
    target_sheet: FastSheet,
    target_row: FastRow,
) -> bool:
    columns = set(source_row.cells) | set(range(1, len(source_sheet.headers) + 1))
    for source_column in columns:
        target_column = _mapped_column(source_sheet, target_sheet, source_column)
        if target_column is None:
            return False
        if not _cell_equal(source_row.cells.get(source_column), target_row.cells.get(target_column)):
            return False
    return True


def _detail(
    *,
    sheet: str,
    key: str,
    kind: str,
    apply_kind: str | None = None,
    column: int | None = None,
    field_name: str = "",
    before: FastCell | None = None,
    source: FastCell | None = None,
    target: FastCell | None = None,
    reason: str = "",
) -> dict:
    result = {"sheet": sheet, "key": key, "kind": kind}
    if apply_kind:
        result["apply_kind"] = apply_kind
    if column is not None:
        result["column"] = column
    if field_name:
        result["field"] = field_name
    if kind == "confirmation":
        result.update({
            "before": _cell_display(before),
            "source": _cell_display(source),
            "target": _cell_display(target),
            "reason": reason or "目标同一位置存在独立修改",
        })
    elif reason:
        result["reason"] = reason
    return result


def _schema_compatible(before: FastSheet, after: FastSheet, target: FastSheet | None = None) -> bool:
    before_headers = [value for value in before.headers if value]
    after_headers = [value for value in after.headers if value]
    if len(before_headers) != len(set(before_headers)) or len(after_headers) != len(set(after_headers)):
        return False
    if before.headers != after.headers:
        return False
    if target is not None:
        # Ignore formatted blank columns; only compare non-empty header tokens.
        target_values = [value for value in target.headers if value]
        if len(target_values) != len(set(target_values)):
            return False
        source_headers = set(after_headers)
        target_headers = set(target_values)
        if not source_headers.issubset(target_headers):
            return False
    return True


def _xml_sections(payload: bytes, names: set[str]) -> tuple[bytes, ...]:
    """Return only explicitly protected OOXML sections in document order."""
    root = ET.fromstring(payload)
    return tuple(
        ET.tostring(node, encoding="utf-8")
        for node in root.iter()
        if node.tag.rsplit("}", 1)[-1] in names
    )


def _node_signature(
    node: ET.Element,
    *,
    ignored_tags: set[str] | None = None,
    ignored_attrs: set[str] | None = None,
) -> tuple:
    ignored_tags = ignored_tags or set()
    ignored_attrs = ignored_attrs or set()
    local = node.tag.rsplit("}", 1)[-1]
    attributes = tuple(sorted(
        (key.rsplit("}", 1)[-1], value)
        for key, value in node.attrib.items()
        if key.rsplit("}", 1)[-1] not in ignored_attrs
    ))
    children = tuple(
        _node_signature(child, ignored_tags=ignored_tags, ignored_attrs=ignored_attrs)
        for child in node
        if child.tag.rsplit("}", 1)[-1] not in ignored_tags
    )
    return local, attributes, (node.text or "").strip(), children


def _protected_sheet_signature(payload: bytes) -> tuple:
    root = ET.fromstring(payload)
    sections: list[tuple] = []
    for node in root:
        name = node.tag.rsplit("}", 1)[-1]
        if name == "mergeCells":
            sections.append((name, tuple(sorted(
                child.attrib.get("ref", "") for child in node
                if child.tag.rsplit("}", 1)[-1] == "mergeCell"
            ))))
        elif name == "dataValidations":
            sections.append((name, _node_signature(node, ignored_attrs={"count", "uid"})))
        elif name == "tableParts":
            # Table definitions themselves are protected package members; the
            # relationship id is allowed to be regenerated during save.
            sections.append((name, len(list(node))))
        elif name == "sheetProtection":
            sections.append((name, _node_signature(node, ignored_attrs={"uid"})))
        elif name == "autoFilter":
            # Excel and WPS attach different revision/extension metadata to an
            # otherwise identical filter range.  Preserve the range and any
            # real filterColumn criteria, but ignore producer-specific extras.
            sections.append((name, _node_signature(
                node,
                ignored_tags={"extLst"},
                ignored_attrs={"uid", "filterBottomFollowUsedRange"},
            )))
    return tuple(sections)


def _source_structure_issue(before: FastWorkbook, after: FastWorkbook) -> str:
    """Reject workbook edits that cannot be represented as record changes.

    Cell values and formulas are handled by the source delta.  This audit is
    fail-closed for structures that change configuration topology or workbook
    execution, while producer-specific presentation metadata is ignored.
    """
    before_members = set(before.members)
    after_members = set(after.members)

    # Only structures that affect cell topology or executable workbook
    # semantics are protected.  Selection, view, column-width, phonetic,
    # comments/VML, conditional-format IDs, hyperlinks and package metadata
    # routinely churn when the same workbook is saved by another Excel build;
    # they are not configuration changes and are intentionally ignored.
    protected_prefixes = (
        "xl/tables/", "xl/pivot", "xl/embeddings/", "xl/activeX/",
        "xl/ctrlProps/", "customXml/",
    )
    protected_members = {"xl/vbaProject.bin", "xl/connections.xml"}
    for member in sorted(before_members | after_members):
        if (
            member in protected_members or member.startswith(protected_prefixes)
        ) and before.payloads.get(member) != after.payloads.get(member):
            return f"工作簿包含无法按记录同步的结构变化：{member}"

    for name in sorted(after.sheets):
        left = before.sheets.get(name)
        right = after.sheets.get(name)
        if left is None or right is None:
            continue
        if _protected_sheet_signature(before.payloads[left.member]) != _protected_sheet_signature(
            after.payloads[right.member]
        ):
            return f"{name} 的合并单元格、数据校验、表格或保护结构发生变化"

    workbook_member = "xl/workbook.xml"
    if (
        workbook_member in before.payloads
        and workbook_member in after.payloads
        and _xml_sections(before.payloads[workbook_member], {"definedNames"}) != _xml_sections(
            after.payloads[workbook_member], {"definedNames"}
        )
    ):
        return "工作簿结构、工作表属性或定义名称发生变化"
    return ""


def analyze_source(
    before_path: str,
    source_path: str,
    *,
    before_revision: int | None = None,
    source_revision: int | None = None,
) -> FastSourceDelta:
    before = load_workbook_index(before_path, revision=before_revision)
    after = load_workbook_index(source_path, revision=source_revision)
    delta = FastSourceDelta(
        source_path,
        before_path,
        before_revision=before_revision,
        source_revision=source_revision,
    )
    if set(before.sheets) != set(after.sheets):
        delta.unsupported_reason = "工作表结构发生变化"
        return delta
    structure_issue = _source_structure_issue(before, after)
    if structure_issue:
        delta.unsupported_reason = structure_issue
        return delta
    for name in sorted(after.sheets):
        left, right = before.sheets[name], after.sheets[name]
        if before.payloads.get(left.member) == after.payloads.get(right.member):
            continue
        if _sheet_semantic_equal(left, right):
            continue
        if left.duplicate_keys or right.duplicate_keys or not _schema_compatible(left, right):
            delta.unsupported_reason = f"{name} 的唯一键或字段结构无法证明稳定"
            return delta
        added = sorted(set(right.key_rows) - set(left.key_rows))
        deleted = sorted(set(left.key_rows) - set(right.key_rows))
        changed: list[dict] = []
        for key in sorted(set(left.key_rows) & set(right.key_rows)):
            old = left.rows[left.key_rows[key]]
            new = right.rows[right.key_rows[key]]
            columns = sorted(set(old.cells) | set(new.cells))
            for column in columns:
                if not _cell_equal(old.cells.get(column), new.cells.get(column)):
                    changed.append({"key": key, "before_row": old.number, "after_row": new.number, "column": column})
        if added or deleted or changed:
            delta.changed_sheets[name] = {"added": added, "deleted": deleted, "changed": changed}
            delta.incoming_count += len(added) + len(deleted) + len(changed)
    delta.changed_count = delta.incoming_count
    return delta


def analyze_target(
    delta: FastSourceDelta,
    target_path: str,
    *,
    target_revision: int | None = None,
) -> FastTargetDecision:
    if delta.unsupported_reason:
        return FastTargetDecision("unsupported", delta.unsupported_reason)
    before = load_workbook_index(delta.before_path, revision=delta.before_revision)
    source = load_workbook_index(delta.source_path, revision=delta.source_revision)
    # Classification is read-only.  Candidate writing later reuses the target
    # package and replaces only worksheet payloads that receive source changes.
    target = load_workbook_index(target_path, revision=target_revision)
    details: list[dict] = []
    for sheet_name, changes in delta.changed_sheets.items():
        source_sheet = source.sheets[sheet_name]
        before_sheet = before.sheets[sheet_name]
        target_sheet = target.sheets.get(sheet_name)
        if target_sheet is None or not _schema_compatible(before_sheet, source_sheet, target_sheet):
            return FastTargetDecision("unsupported", f"目标缺少可可靠映射的工作表或字段：{sheet_name}")
        if target_sheet.duplicate_keys:
            return FastTargetDecision("unsupported", f"目标工作表存在重复唯一键：{sheet_name}")
        added_rows = sorted(source_sheet.key_rows[key] for key in changes["added"])
        if added_rows:
            expected_tail = list(range(max(before_sheet.rows or {1}) + 1, max(source_sheet.rows or {1}) + 1))
            if added_rows != expected_tail:
                return FastTargetDecision("unsupported", f"{sheet_name} 的新增记录不是连续表尾追加，无法证明行序安全")
        for key in changes["added"]:
            if key in target_sheet.key_rows:
                source_row = source_sheet.rows[source_sheet.key_rows[key]]
                target_row = target_sheet.rows[target_sheet.key_rows[key]]
                if _row_matches_source_fields(source_sheet, source_row, target_sheet, target_row):
                    details.append(_detail(sheet=sheet_name, key=key, kind="already"))
                else:
                    columns = set(source_row.cells) | set(range(1, len(source_sheet.headers) + 1))
                    for source_column in sorted(columns):
                        target_column = _mapped_column(source_sheet, target_sheet, source_column)
                        if target_column is None:
                            return FastTargetDecision(
                                "unsupported",
                                f"{sheet_name} 的新增记录字段无法唯一映射：{_header(source_sheet, source_column) or source_column}",
                            )
                        source_cell = source_row.cells.get(source_column)
                        target_cell = target_row.cells.get(target_column)
                        if _cell_equal(source_cell, target_cell):
                            continue
                        details.append(_detail(
                            sheet=sheet_name,
                            key=key,
                            kind="confirmation",
                            apply_kind="change",
                            column=source_column,
                            field_name=_header(source_sheet, source_column),
                            source=source_cell,
                            target=target_cell,
                            reason="源分支新增了该记录，但目标已存在同一唯一键",
                        ))
                continue
            details.append(_detail(sheet=sheet_name, key=key, kind="add"))
        for key in changes["deleted"]:
            target_row = target_sheet.key_rows.get(key)
            if target_row is None:
                details.append(_detail(sheet=sheet_name, key=key, kind="already"))
            elif not _row_matches_source_fields(
                before_sheet,
                before_sheet.rows[before_sheet.key_rows[key]],
                target_sheet,
                target_sheet.rows[target_row],
            ):
                details.append(_detail(
                    sheet=sheet_name,
                    key=key,
                    kind="confirmation",
                    apply_kind="delete",
                    before=before_sheet.rows[before_sheet.key_rows[key]].cells.get(1),
                    target=target_sheet.rows[target_row].cells.get(1),
                    reason="源分支删除了该记录，但目标记录已有独立修改",
                ))
            else:
                details.append(_detail(sheet=sheet_name, key=key, kind="delete"))
        missing_target_keys: set[str] = set()
        for change in changes["changed"]:
            key = change["key"]
            target_row = target_sheet.key_rows.get(key)
            if target_row is None:
                if key in missing_target_keys:
                    continue
                missing_target_keys.add(key)
                details.append(_detail(
                    sheet=sheet_name,
                    key=key,
                    kind="confirmation",
                    apply_kind="restore_record",
                    reason="目标不存在对应记录；确认后将按源分支结果恢复该记录",
                ))
                continue
            old = before_sheet.rows[change["before_row"]].cells.get(change["column"])
            new = source_sheet.rows[change["after_row"]].cells.get(change["column"])
            target_column = _mapped_column(source_sheet, target_sheet, change["column"])
            if target_column is None:
                return FastTargetDecision("unsupported", f"{sheet_name} 的字段无法唯一映射：{_header(source_sheet, change['column']) or change['column']}")
            mine = target_sheet.rows[target_row].cells.get(target_column)
            if _cell_equal(mine, new):
                kind = "already"
            elif _cell_equal(mine, old):
                kind = "change"
            else:
                kind = "confirmation"
            details.append(_detail(
                sheet=sheet_name,
                key=key,
                kind=kind,
                apply_kind="change" if kind == "confirmation" else None,
                column=change["column"],
                field_name=_header(source_sheet, change["column"]),
                before=old,
                source=new,
                target=mine,
            ))
    confirmation = [item for item in details if item["kind"] == "confirmation"]
    direct = [item for item in details if item["kind"] in {"add", "delete", "change"}]
    already = [item for item in details if item["kind"] == "already"]
    if confirmation:
        return FastTargetDecision(
            "confirmation_required",
            confirmation[0].get("reason", "目标同一位置存在独立修改"),
            {"direct": len(direct), "already": len(already), "confirmation": len(confirmation)},
            details,
        )
    if not direct:
        return FastTargetDecision("already_applied", "目标已经包含源分支变更", {"already": len(already)}, details)
    return FastTargetDecision(
        "direct",
        f"可直接同步 {len(direct)} 项，保留目标分支其他内容",
        {"direct": len(direct), "already": len(already)},
        details,
    )


def cache_clear() -> None:
    _load_cached.cache_clear()
    _content_digest.cache_clear()


def apply_source_change_plan(
    before_path: str,
    source_path: str,
    target_path: str,
    output_path: str,
    decision: FastTargetDecision,
    *,
    confirmed: bool = False,
    source_revision: int | None = None,
    target_revision: int | None = None,
) -> None:
    """Apply the source delta to a target-derived candidate.

    This is intentionally called only after the source commit has been
    reconciled and the target has passed a fresh SVN status check.  The fast
    analyzer remains read-only; this function is the first workbook write.

    A conflict plan can only be materialized after explicit confirmation.  It
    still writes only the locations changed on the source branch.
    """
    allowed = decision.disposition == "direct" or (
        decision.disposition == "confirmation_required" and confirmed
    )
    if not allowed:
        raise ValueError(f"不能物化 {decision.disposition} 动作")
    source = load_workbook_index(source_path, revision=source_revision)
    target = load_workbook_index(target_path, revision=target_revision)
    changed_payloads: dict[str, bytes] = {}
    by_sheet: dict[str, list[dict]] = {}
    for detail in decision.details:
        kind = detail.get("kind")
        if kind == "confirmation":
            kind = detail.get("apply_kind") if confirmed else None
        if kind in {"add", "delete", "change", "restore_record"}:
            by_sheet.setdefault(str(detail["sheet"]), []).append({**detail, "resolved_kind": kind})

    # Parse and serialize each changed worksheet once.  The previous writer did
    # this once per cell, which made large Language workbooks quadratic.
    for sheet, details in by_sheet.items():
        source_ws = source.sheets[sheet]
        target_ws = target.sheets[sheet]
        root = ET.fromstring(target.payloads[target_ws.member])
        sheet_data = root.find(M + "sheetData")
        if sheet_data is None:
            raise ValueError(f"目标工作表缺少 sheetData：{sheet}")
        next_row = max(target_ws.rows or {1}) + 1
        restored_rows: dict[str, int] = {}
        for detail in details:
            kind = str(detail["resolved_kind"])
            key = str(detail["key"])
            source_row = source_ws.key_rows.get(key)
            target_row = restored_rows.get(key, target_ws.key_rows.get(key))
            if kind in {"add", "restore_record"}:
                if target_row is None:
                    _append_row_xml(sheet_data, source_ws, source_row, target_ws, next_row)
                    restored_rows[key] = next_row
                    next_row += 1
            elif kind == "delete":
                if target_row is not None:
                    _remove_row_xml(sheet_data, target_row)
            elif kind == "change" and target_row is not None:
                source_column = int(detail["column"])
                target_column = _target_column_for_source(source_ws, target_ws, source_column)
                if target_column is None:
                    raise ValueError(f"目标字段无法映射：{sheet}/{key}/{source_column}")
                source_cell = source_ws.rows[source_row].cells.get(source_column)
                _set_cell_xml(sheet_data, target_row, target_column, source_cell or FastCell(None))
        _update_dimension(root, sheet_data)
        changed_payloads[target_ws.member] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    _safe_rewrite_package(target_path, output_path, changed_payloads)


def _target_column_for_source(source_ws: FastSheet, target_ws: FastSheet, source_column: int) -> int | None:
    return _mapped_column(source_ws, target_ws, source_column)


def _append_row_xml(sheet_data: ET.Element, source_ws: FastSheet, source_row: int | None, target_ws: FastSheet, target_row: int) -> None:
    if source_row is None:
        raise ValueError("源新增记录缺少唯一键行")
    template_number = max(target_ws.key_rows.values(), default=max(target_ws.rows or {1}))
    template_row = target_ws.rows.get(template_number)
    row_attributes = {"r": str(target_row)}
    if template_row is not None and template_row.style and template_row.style != "0":
        row_attributes["s"] = template_row.style
        row_attributes["customFormat"] = "1"
    row = ET.Element(M + "row", row_attributes)
    for column, source_cell in source_ws.rows[source_row].cells.items():
        target_column = _target_column_for_source(source_ws, target_ws, column)
        if not target_column:
            continue
        cell = ET.SubElement(row, M + "c", {"r": f"{_column_name(target_column)}{target_row}"})
        template_cell = template_row.cells.get(target_column) if template_row is not None else None
        if template_cell is not None and template_cell.style and template_cell.style != "0":
            cell.attrib["s"] = template_cell.style
        _write_cell_payload(cell, source_cell)
    sheet_data.append(row)


def _remove_row_xml(sheet_data: ET.Element, row_number: int) -> None:
    for row in list(sheet_data.findall(M + "row")):
        if row.attrib.get("r") == str(row_number):
            sheet_data.remove(row)
            return


def _update_dimension(root: ET.Element, sheet_data: ET.Element) -> None:
    max_row = 1
    max_column = 1
    for row in sheet_data.findall(M + "row"):
        try:
            max_row = max(max_row, int(row.attrib.get("r", "1")))
        except ValueError:
            pass
        for cell in row.findall(M + "c"):
            max_column = max(max_column, _column_number(cell.attrib.get("r", "A1")))
    dimension = root.find(M + "dimension")
    if dimension is not None:
        dimension.attrib["ref"] = f"A1:{_column_name(max_column)}{max_row}"


def _set_cell_xml(sheet_data: ET.Element, row_number: int, column: int, source_cell: FastCell) -> None:
    row = next((item for item in sheet_data.findall(M + "row") if item.attrib.get("r") == str(row_number)), None)
    if row is None:
        row = ET.Element(M + "row", {"r": str(row_number)})
        sheet_data.append(row)
    reference = f"{_column_name(column)}{row_number}"
    cell = next((item for item in row.findall(M + "c") if item.attrib.get("r") == reference), None)
    if cell is None:
        cell = ET.SubElement(row, M + "c", {"r": reference})
    _write_cell_payload(cell, source_cell)


def _write_cell_payload(cell: ET.Element, source_cell: FastCell) -> None:
    for child in list(cell):
        if child.tag in {M + "f", M + "v", M + "is"}:
            cell.remove(child)
    if source_cell.formula:
        cell.attrib.pop("t", None)
        ET.SubElement(cell, M + "f").text = source_cell.formula
        return
    value = source_cell.value
    if value is None or value == "":
        cell.attrib.pop("t", None)
        return
    if source_cell.cell_type in {"s", "inlineStr", "str"} or isinstance(value, str) and not _looks_numeric(value):
        cell.attrib["t"] = "inlineStr"
        inline = ET.SubElement(cell, M + "is")
        ET.SubElement(inline, M + "t").text = str(value)
    elif source_cell.cell_type in {"b", "e", "d"}:
        cell.attrib["t"] = source_cell.cell_type
        ET.SubElement(cell, M + "v").text = str(value)
    else:
        cell.attrib.pop("t", None)
        ET.SubElement(cell, M + "v").text = str(value)


def _looks_numeric(value: str) -> bool:
    try:
        float(value)
        return True
    except (TypeError, ValueError):
        return False


def _column_name(column: int) -> str:
    value = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        value = chr(65 + remainder) + value
    return value


def _safe_rewrite_package(source_path: str, output_path: str, changed_payloads: dict[str, bytes]) -> None:
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    temporary = output_path + f".tmp-source-change-{os.getpid()}"
    try:
        with zipfile.ZipFile(source_path, "r") as source, zipfile.ZipFile(temporary, "w") as destination:
            for info in source.infolist():
                payload = changed_payloads.get(info.filename, source.read(info.filename))
                destination.writestr(info, payload)
        os.replace(temporary, output_path)
    finally:
        try:
            if os.path.exists(temporary):
                os.remove(temporary)
        except OSError:
            pass
