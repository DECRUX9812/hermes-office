"""Settings come from plugin.yaml config_schema (Hermes Settings > Plugins), never from a plugin UI."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

try:
    import yaml

    def _load_yaml(text: str) -> Any:
        return yaml.safe_load(text)
except ImportError:  # the Hermes venv ships ruamel.yaml instead of PyYAML
    from ruamel.yaml import YAML

    def _load_yaml(text: str) -> Any:
        return YAML(typ="safe").load(text)

from conftest import REPO, call
from hermes_office import service, settings


class FakeCtx:
    def __init__(self, values: dict[str, Any]):
        self.values = values

    def get_config(self, key: str, default: Any = None) -> Any:
        return self.values.get(key, default)


@pytest.fixture(autouse=True)
def _clean_settings():
    settings._reset_for_tests()
    yield
    settings._reset_for_tests()


def test_manifest_declares_every_setting_with_its_default() -> None:
    manifest = _load_yaml((REPO / "plugin.yaml").read_text(encoding="utf-8"))
    schema = manifest["config_schema"]
    assert set(schema) == set(settings.DEFAULTS)
    types = {"list": list, "str": str, "int": int}
    for key, spec in schema.items():
        assert spec["default"] == settings.DEFAULTS[key], key
        assert isinstance(spec["default"], types[spec["type"]]), key
        assert spec.get("label") and spec.get("description"), key
    assert manifest["name"] == "office-desk"
    assert manifest["provides_hooks"] == ["pre_llm_call", "post_llm_call"]


def test_defaults_without_hermes_context(hermes_home) -> None:
    assert settings.default_author() == "Hermes"
    assert settings.recalc_timeout() == 120.0
    assert settings.folders() == []


def test_values_from_context_and_bad_types_fall_back(hermes_home) -> None:
    settings.bind(FakeCtx({"default_author": "Legal Team", "recalc_timeout": 30}))
    assert settings.default_author() == "Legal Team"
    assert settings.recalc_timeout() == 30.0
    settings.bind(FakeCtx({"default_author": 7, "recalc_timeout": "soon", "folders": "not-a-list"}))
    assert settings.default_author() == "Hermes"
    assert settings.recalc_timeout() == 120.0
    assert settings.folders() == []


def test_configured_folders_are_put_on_the_desk(hermes_home, tmp_path: Path) -> None:
    keep = tmp_path / "Configured"
    keep.mkdir()
    settings.bind(FakeCtx({"folders": [str(keep), str(tmp_path / "missing"), "relative/path"]}))
    desk = service.desk()
    assert [f["path"] for f in desk.list_folders()] == [str(keep.resolve())]
    service.desk()  # idempotent
    assert len(service.desk().list_folders()) == 1


def test_default_author_setting_is_used_for_tracked_changes(files) -> None:
    settings.bind(FakeCtx({"default_author": "Legal Team"}))
    target = files["folder"] / "Services Agreement.docx"
    out = call("docx_tracked_edit", {"path": str(target), "find": "within 30 days", "replace": "within 45 days"})
    assert out["ok"] is True, out
    import docx
    authors = {r.author for r in docx.Document(str(target)).revisions}
    assert authors == {"Legal Team"}, authors
