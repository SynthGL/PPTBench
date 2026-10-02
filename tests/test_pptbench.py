from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest

from pptbench import fixtures
from pptbench.adapters import run_adapter
from pptbench.fixtures import materialize
from pptbench.report import generate
from pptbench.scoring import score


def test_real_python_pptx_template_edit_and_collateral_scoring(tmp_path: Path) -> None:
    source = materialize(tmp_path)["mixed-60"]
    output = tmp_path / "edited.pptx"
    result = run_adapter("python-pptx", "template-mutation", source, output, 60)
    assert result["outcome"] == "success"

    checks = score("template-mutation", source, output)
    outcomes = {check["name"]: check for check in checks}
    for name in (
        "table-cell-1-1-updated",
        "table-cell-2-2-updated",
        "bullet-paragraph-1-updated",
    ):
        assert outcomes[name]["outcome"] == "success"
    # python-pptx re-serializes the two edited slides and drops the unreachable
    # opaque part; only the opaque loss is a semantic failure.
    assert {check["name"] for check in checks if check["outcome"] == "failure"} == {
        "opaque-parts-preserved",
        "raw-untouched-part-equality",
    }
    assert outcomes["raw-untouched-part-equality"]["category"] == "byte-only"
    assert outcomes["raw-untouched-part-equality"]["scored"] is False


def test_wrong_edit_and_collateral_change_fail(tmp_path: Path) -> None:
    source = materialize(tmp_path)["mixed-60"]
    wrong = tmp_path / "wrong.pptx"
    wrong.write_bytes(source.read_bytes())
    no_op_checks = score("template-mutation", source, wrong)
    assert any(
        check["name"] == "table-cell-1-1-updated" and check["outcome"] == "failure"
        for check in no_op_checks
    )
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(wrong, "w") as changed:
        for item in original.infolist():
            data = original.read(item.filename)
            changed.writestr(
                item,
                data.replace(b"Synthetic mixed", b"Collateral changed")
                if item.filename == "ppt/slides/slide1.xml"
                else data,
            )
    assert any(
        check["name"] == "only-declared-text-nodes-changed" and check["outcome"] == "failure"
        for check in score("template-mutation", source, wrong)
    )


def test_missing_result_unavailable_and_timeout_are_honest(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = materialize(tmp_path)["mixed-60"]
    missing = run_adapter("does-not-exist", "feature-matrix", source, tmp_path / "out.pptx", 1)
    assert missing["outcome"] == "unsupported"
    monkeypatch.delenv("PPTBENCH_EXTERNAL_COMMAND", raising=False)
    assert (
        run_adapter("external-command", "feature-matrix", source, tmp_path / "out.pptx", 1)[
            "outcome"
        ]
        == "unavailable"
    )


def test_report_escapes_html_and_retains_raw_hash(tmp_path: Path) -> None:
    run = tmp_path / "run.json"
    run.write_text(
        json.dumps(
            {
                "schema_version": 2,
                "fixtures": [],
                "results": [
                    {
                        "adapter": "<x>",
                        "lane": "chart-data",
                        "fixture": "charts",
                        "outcome": "failure",
                        "reason": "<script>alert(1)</script>",
                        "checks": [],
                        "samples": [],
                    }
                ],
            }
        ),
        encoding="utf-8",
    )
    dashboard, _ = generate(run)
    text = dashboard.read_text(encoding="utf-8")
    assert "&lt;script&gt;" in text and "<script>alert(1)</script>" not in text


def test_frozen_manifest_and_package_resource_are_checked(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    fixtures.verify_frozen_data()
    data = fixtures.manifest()
    assert data["schema_version"] == 2
    assert data["corpus"]["recipe"] == "pptbench.fixtures.FROZEN_RECIPE"

    monkeypatch.setitem(data["corpus"], "recipe_sha256", "bad")

    def tampered_manifest() -> fixtures.Manifest:
        return data

    monkeypatch.setattr(fixtures, "manifest", tampered_manifest)
    with pytest.raises(ValueError, match="recipe digest"):
        fixtures.verify_frozen_data()
