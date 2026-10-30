"""Registration must not damage private or concurrently edited Codex settings."""

from __future__ import annotations

import importlib.util
import json
import tomllib
from contextlib import contextmanager
from pathlib import Path

import pytest

from axmath_mcp import __version__

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("register_codex", ROOT / "scripts/register_codex.py")
registration = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(registration)


@pytest.fixture
def entry(tmp_path):
    return registration.build_server_config(tmp_path / "checkout with 中文", "C:\\Tools\\uv.exe")


def test_generated_toml_round_trips_spaces_unicode_and_current_version(entry):
    parsed = tomllib.loads(registration.render_config_snippet(entry))
    assert parsed["mcp_servers"]["axmath"] == entry
    assert entry["env"]["AXMATH_MCP_VERSION"] == __version__
    assert entry["args"][-3:] == ["python", "-m", "axmath_mcp.server"]


def test_new_config_does_not_require_an_existing_codex_directory(tmp_path, entry):
    path = tmp_path / "new-codex-home/config.toml"
    result = registration.install_config(path, entry)
    assert result["changed"] is True and result["backup"] is None
    assert tomllib.loads(path.read_text(encoding="utf-8"))["mcp_servers"]["axmath"] == entry


def test_append_preserves_comments_bom_and_unrelated_settings_bytes(tmp_path, entry):
    path = tmp_path / "config.toml"
    original = b'\xef\xbb\xbf# private settings\r\nmodel = "example"\r\n[mcp_servers.other]\r\ncommand = "other"\r\n'
    path.write_bytes(original)
    result = registration.install_config(path, entry)
    assert path.read_bytes().startswith(original)
    assert Path(result["backup"]).read_bytes() == original
    parsed = tomllib.loads(path.read_text(encoding="utf-8-sig"))
    assert parsed["model"] == "example"
    assert parsed["mcp_servers"]["other"] == {"command": "other"}


def test_identical_install_is_noop_without_backup_or_write(tmp_path, entry):
    path = tmp_path / "config.toml"
    registration.install_config(path, entry)
    original = path.read_bytes()
    before = path.stat().st_mtime_ns
    result = registration.install_config(path, entry)
    assert result["status"] == "already_configured" and result["changed"] is False
    assert result["backup"] is None
    assert path.read_bytes() == original and path.stat().st_mtime_ns == before
    assert list(tmp_path.glob("*.axmath-backup-*")) == []


def test_conflicting_entry_is_preserved_without_backup(tmp_path, entry):
    path = tmp_path / "config.toml"
    original = b'[mcp_servers.axmath]\ncommand = "another-server"\n'
    path.write_bytes(original)
    with pytest.raises(ValueError, match="different axmath entry"):
        registration.install_config(path, entry)
    assert path.read_bytes() == original
    assert list(tmp_path.glob("*.axmath-backup-*")) == []


@pytest.mark.parametrize("contents", [b"[bad", b'mcp_servers = "not-a-table"\n'])
def test_invalid_existing_config_is_not_rewritten(tmp_path, entry, contents):
    path = tmp_path / "config.toml"
    path.write_bytes(contents)
    with pytest.raises((ValueError, TypeError)):
        registration.install_config(path, entry)
    assert path.read_bytes() == contents


@pytest.mark.parametrize("change", ["edit", "delete", "create_empty"])
def test_concurrent_edit_delete_or_create_is_not_overwritten(tmp_path, entry, monkeypatch, change):
    path = tmp_path / "config.toml"
    if change != "create_empty":
        path.write_bytes(b'# initial\nmodel = "one"\n')
    actual_temporary = registration.tempfile.NamedTemporaryFile

    @contextmanager
    def racing_temporary(**kwargs):
        with actual_temporary(**kwargs) as temporary:
            yield temporary
        if change == "edit":
            path.write_bytes(b'# changed concurrently\nmodel = "two"\n')
        elif change == "delete":
            path.unlink()
        else:
            path.write_bytes(b"")

    monkeypatch.setattr(registration.tempfile, "NamedTemporaryFile", racing_temporary)
    with pytest.raises(RuntimeError, match="changed concurrently"):
        registration.install_config(path, entry)
    if change == "delete":
        assert not path.exists()
    else:
        expected = b"" if change == "create_empty" else b'# changed concurrently\nmodel = "two"\n'
        assert path.read_bytes() == expected
    assert not list(tmp_path.glob("axmath-config-*"))


@pytest.mark.parametrize("value", ["", "relative-folder"])
def test_relative_codex_home_is_rejected(value):
    with pytest.raises(ValueError, match="absolute"):
        registration.resolve_config_directory({"CODEX_HOME": value})


def test_codex_home_override_and_default_are_explicit(tmp_path):
    custom = tmp_path / "custom"
    assert registration.resolve_config_directory({"CODEX_HOME": str(custom)}) == custom
    assert registration.resolve_config_directory({}, home=tmp_path) == tmp_path / ".codex"


def test_public_examples_have_portable_placeholders():
    codex = tomllib.loads((ROOT / "config/codex.toml").read_text(encoding="utf-8"))
    generic = json.loads((ROOT / "config/mcp.json").read_text(encoding="utf-8"))
    for example in (codex["mcp_servers"]["axmath"], generic["mcpServers"]["axmath"]):
        assert example["command"] == "uv"
        assert example["args"][1] == r"C:\path\to\axmath-mcp"
        assert example["args"][-3:] == ["python", "-m", "axmath_mcp.server"]
