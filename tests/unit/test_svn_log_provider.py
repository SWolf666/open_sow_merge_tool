from __future__ import annotations

from types import SimpleNamespace

import pytest

from sow_merge_tool import svn_log_provider as log_provider_module
from sow_merge_tool.branch_submit import BranchSubmitBatch, BranchSubmitEngine
from sow_merge_tool.svn_log_provider import (
    SvnLogError,
    SvnLogProvider,
    parse_svn_log_xml,
    read_svn_log,
)

LOG_XML = """
<log>
  <logentry revision="42">
    <author>alice</author>
    <date>2026-09-04T06:00:00.000000Z</date>
    <paths>
      <path action="A" kind="file" text-mods="true" prop-mods="false"
            copyfrom-path="/trunk/config.xlsx" copyfrom-rev="40">/branches/release/config.xlsx</path>
    </paths>
    <property name="svn:mergeinfo">/trunk:40-42</property>
    <msg>sync configuration</msg>
  </logentry>
</log>
"""


def test_parse_structured_log_keeps_revision_identity_and_copyfrom():
    entries = parse_svn_log_xml(LOG_XML)
    assert len(entries) == 1
    entry = entries[0]
    assert (entry.revision, entry.author, entry.message) == (42, "alice", "sync configuration")
    changed = entry.changed_paths[0]
    assert changed.path == "/branches/release/config.xlsx"
    assert changed.action == "A"
    assert changed.copyfrom_path == "/trunk/config.xlsx"
    assert changed.copyfrom_revision == 40
    assert changed.text_modified is True
    assert changed.properties_modified is False
    assert entry.mergeinfo == ("/trunk:40-42",)
    assert entry.to_dict()["changed_paths"][0]["copyfrom_revision"] == 40


def test_read_svn_log_uses_read_only_xml_verbose_command(tmp_path):
    commands = []

    def runner(command):
        commands.append(command)
        return SimpleNamespace(returncode=0, stdout=LOG_XML, stderr="")

    target = tmp_path / "release"
    entries = read_svn_log(
        str(target),
        limit=999,
        start_revision=40,
        end_revision=42,
        runner=runner,
    )
    assert [entry.revision for entry in entries] == [42]
    assert commands == [[
        "injected-svn",
        "log",
        "--xml",
        "-v",
        "--limit",
        "500",
        "-r",
        "40:42",
        str(target.resolve()),
    ]]


def test_provider_is_callable_and_missing_or_bad_log_is_unknown():
    provider = SvnLogProvider(runner=lambda _command: (0, LOG_XML, ""))
    assert provider("https://svn.example/repo/trunk")[0].revision == 42

    with pytest.raises(SvnLogError, match="XML 无效"):
        parse_svn_log_xml("<log>")

    with pytest.raises(SvnLogError, match="读取失败"):
        read_svn_log("branch", runner=lambda _command: (1, "", "authorization failed"))

    with pytest.raises(SvnLogError, match="启动失败"):
        read_svn_log("branch", runner=lambda _command: (_ for _ in ()).throw(OSError("runner down")))


def test_engine_records_log_evidence_but_skips_no_target_batches(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "local"))
    calls = []

    def provider(path, **_kwargs):
        calls.append(path)
        return parse_svn_log_xml(LOG_XML)

    batch = BranchSubmitBatch(
        batch_id="log-audit",
        wc_root=str(tmp_path),
        source_branch="develop",
        target_branches=["release"],
        files=[],
        message="audit",
    )
    engine = BranchSubmitEngine(str(tmp_path), log_provider=provider)
    evidence = engine._record_structured_log(
        batch,
        branch="develop",
        path=str(tmp_path / "develop"),
        phase="source-reconcile",
    )
    assert evidence is not None
    assert evidence["state"] == "available"
    assert evidence["revisions"] == [42]
    assert batch.svn_log["develop:source-reconcile"]["entries"][0]["revision"] == 42
    assert batch.journal[-1]["kind"] == "svn-log-audit"
    assert calls == [str(tmp_path / "develop")]

    no_target = BranchSubmitBatch(
        batch_id="no-target-log-audit",
        wc_root=str(tmp_path),
        source_branch="develop",
        target_branches=[],
        files=[],
        message="native commit",
    )
    assert engine._record_structured_log(
        no_target,
        branch="develop",
        path=str(tmp_path / "develop"),
        phase="source-reconcile",
    ) is None
    assert no_target.svn_log == {}
    assert calls == [str(tmp_path / "develop")]


def test_engine_marks_provider_failure_unknown_without_changing_action_state(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "local"))
    batch = BranchSubmitBatch(
        batch_id="log-unknown",
        wc_root=str(tmp_path),
        source_branch="develop",
        target_branches=["release"],
        files=[],
        message="audit",
    )
    engine = BranchSubmitEngine(
        str(tmp_path),
        log_provider=lambda _path, **_kwargs: (_ for _ in ()).throw(SvnLogError("server unavailable")),
    )
    evidence = engine._record_structured_log(
        batch,
        branch="release",
        path=str(tmp_path / "release"),
        phase="target-reconcile",
    )
    assert evidence is not None
    assert evidence["state"] == "unknown"
    assert "server unavailable" in evidence["reason"]
    assert batch.journal[-1]["kind"] == "svn-log-audit"


def test_read_svn_log_uses_hidden_tortoise_runtime_fallback_without_svn_cli(monkeypatch):
    calls = {}

    def fallback(path, **kwargs):
        calls["path"] = path
        calls.update(kwargs)
        return [parse_svn_log_xml(LOG_XML)[0]]

    monkeypatch.setattr(log_provider_module, "_find_svn_cli", lambda: None)
    monkeypatch.setattr(log_provider_module, "_log_via_tortoise_child", fallback)
    entries = log_provider_module.read_svn_log(
        "C:/wc/develop/config.xlsx",
        limit=12,
        start_revision=10,
        end_revision=42,
    )
    assert entries[0].revision == 42
    assert calls == {
        "path": "C:/wc/develop/config.xlsx",
        "limit": 12,
        "start_revision": 10,
        "end_revision": 42,
        "cancel_event": None,
    }
