from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

from pptbench import adapters, runner
from pptbench.adapters import python_pptx_info, run_adapter
from pptbench.models import SCHEMA_VERSION, AdapterInfo, Result, Sample
from pptbench.report import generate


def _external_script(path: Path, body: str) -> str:
    path.write_text(body, encoding="utf-8")
    return f"{sys.executable} {path} {{input}} {{output}} {{lane}}"


def test_timeout_retains_partial_stdout_and_never_reuses_candidate(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    script = _external_script(
        tmp_path / "timeout.py",
        "import sys, time\nprint('partial-output', flush=True)\ntime.sleep(10)\n",
    )
    monkeypatch.setenv("PPTBENCH_EXTERNAL_COMMAND", script)
    source = tmp_path / "source.pptx"
    source.write_bytes(b"source")
    output = tmp_path / "candidate.pptx"
    result = run_adapter("external-command", "feature-matrix", source, output, 0.05)
    assert result["outcome"] == "timeout"
    assert result["details"]["stdout_bytes"] > 0
    assert (output.with_suffix(".pptx.adapter") / "stdout.bin").read_bytes().startswith(b"partial")
    output.write_bytes(b"stale")
    reused = run_adapter("external-command", "feature-matrix", source, output, 0.05)
    assert reused["outcome"] == "failure"
    assert "new" in str(reused["reason"])


def test_external_identity_hashes_binary_and_configuration(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    script = _external_script(tmp_path / "ok.py", "")
    monkeypatch.setenv("PPTBENCH_EXTERNAL_COMMAND", script)
    info = adapters.external_command_info()
    assert info.available
    assert info.identity is not None and "config:" in info.identity
    assert str(tmp_path) not in info.identity


def test_package_identity_covers_multiple_modules() -> None:
    info = python_pptx_info()
    assert info.available
    assert info.identity is not None and info.identity.startswith("modules:")


def test_runner_retains_warmups_and_early_measured_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    fixtures = {
        "mixed-60": tmp_path / "mixed-60.pptx",
        "charts": tmp_path / "charts.pptx",
    }
    for fixture in fixtures.values():
        fixture.write_bytes(b"fixture")

    def fake_manifest() -> dict[str, object]:
        return {"corpus": {"fixtures": []}}

    def fake_materialize(_: Path) -> dict[str, Path]:
        return fixtures

    def empty_inventory(_: Path) -> list[dict[str, str]]:
        return []

    def one_adapter() -> list[AdapterInfo]:
        return [AdapterInfo("python-pptx", "test", "test", True)]

    monkeypatch.setattr(runner, "manifest", fake_manifest)
    monkeypatch.setattr(runner, "materialize", fake_materialize)
    monkeypatch.setattr(runner, "fixture_inventory", empty_inventory)
    monkeypatch.setattr(runner, "available_adapters", one_adapter)
    calls = iter(
        [
            {
                "outcome": "success",
                "elapsed_ms": 1,
                "peak_rss_bytes": 2,
                "output_sha256": "x",
            },
            {"outcome": "failure", "reason": "first measured failed"},
            {
                "outcome": "success",
                "elapsed_ms": 1,
                "peak_rss_bytes": 2,
                "output_sha256": "x",
            },
        ]
    )

    def fake_run(*_: object) -> dict[str, object]:
        return next(calls)

    monkeypatch.setattr(runner, "run_adapter", fake_run)

    def no_checks(*_: object) -> list[dict[str, object]]:
        return []

    monkeypatch.setattr(runner, "score", no_checks)
    output = runner.benchmark(
        tmp_path / "run",
        lanes=["template-mutation"],
        adapters=["python-pptx"],
        warmups=1,
        iterations=2,
    )
    report = json.loads(output.read_text(encoding="utf-8"))
    samples = report["results"][0]["samples"]
    assert [sample["phase"] for sample in samples] == ["warmup", "measured", "measured"]
    assert samples[1]["outcome"] == "failure"
    assert report["results"][0]["outcome"] == "failure"


def test_runner_rejects_nonempty_output_before_materialization(tmp_path: Path) -> None:
    output = tmp_path / "run"
    output.mkdir()
    (output / "existing.txt").write_text("user artifact", encoding="utf-8")
    with pytest.raises(ValueError, match="new or empty"):
        runner.benchmark(output, lanes=["template-mutation"], adapters=["python-pptx"])


def test_report_detects_output_tampering_and_raw_edits(tmp_path: Path) -> None:
    fixtures = tmp_path / "fixtures"
    raw = tmp_path / "raw" / "adapter" / "lane"
    fixtures.mkdir(parents=True)
    raw.mkdir(parents=True)
    fixture = fixtures / "mixed-60.pptx"
    output = raw / "measured-1.pptx"
    fixture.write_bytes(b"fixture")
    output.write_bytes(b"output")
    report = tmp_path / "run.json"
    report.write_text(
        json.dumps(
            {
                "schema_version": SCHEMA_VERSION,
                "fixtures": [{"id": "mixed-60", "path": fixture.name, "sha256": _hash(fixture)}],
                "results": [
                    {
                        "adapter": "adapter",
                        "lane": "lane",
                        "fixture": "mixed-60",
                        "input_sha256": _hash(fixture),
                        "outcome": "success",
                        "samples": [
                            {
                                "output_path": "raw/adapter/lane/measured-1.pptx",
                                "output_sha256": _hash(output),
                            }
                        ],
                    }
                ],
            }
        ),
        encoding="utf-8",
    )
    generate(report)
    output.write_bytes(b"edited")
    _, _ = generate(report)
    manifest = json.loads(
        (tmp_path / "report" / "report-manifest.json").read_text(encoding="utf-8")
    )
    assert manifest["integrity"]["outcome"] == "tampered"
    report.write_text(report.read_text(encoding="utf-8") + "\n", encoding="utf-8")
    generate(report)
    manifest = json.loads(
        (tmp_path / "report" / "report-manifest.json").read_text(encoding="utf-8")
    )
    assert "raw report changed" in ";".join(manifest["integrity"]["problems"])


def test_report_detects_retained_failed_candidate_tampering(tmp_path: Path) -> None:
    fixtures = tmp_path / "fixtures"
    raw = tmp_path / "raw" / "adapter" / "lane"
    fixtures.mkdir(parents=True)
    raw.mkdir(parents=True)
    fixture = fixtures / "mixed-60.pptx"
    candidate = raw / "measured-1.pptx"
    fixture.write_bytes(b"fixture")
    candidate.write_bytes(b"candidate")
    report = tmp_path / "run.json"
    report.write_text(
        json.dumps(
            {
                "schema_version": SCHEMA_VERSION,
                "fixtures": [{"id": "mixed-60", "path": fixture.name, "sha256": _hash(fixture)}],
                "results": [
                    {
                        "adapter": "adapter",
                        "lane": "lane",
                        "fixture": "mixed-60",
                        "input_sha256": _hash(fixture),
                        "outcome": "failure",
                        "samples": [
                            {
                                "outcome": "failure",
                                "output_path": None,
                                "output_sha256": None,
                                "checks": [
                                    {
                                        "name": "scored check",
                                        "outcome": "failure",
                                        "scored": True,
                                    }
                                ],
                                "details": {
                                    "candidate_path": "raw/adapter/lane/measured-1.pptx",
                                    "candidate_sha256": _hash(candidate),
                                },
                            }
                        ],
                    }
                ],
            }
        ),
        encoding="utf-8",
    )
    generate(report)
    manifest = json.loads(
        (tmp_path / "report" / "report-manifest.json").read_text(encoding="utf-8")
    )
    assert manifest["integrity"]["outcome"] == "success"

    candidate.write_bytes(b"edited candidate")
    generate(report)
    manifest = json.loads(
        (tmp_path / "report" / "report-manifest.json").read_text(encoding="utf-8")
    )
    assert manifest["integrity"]["outcome"] == "tampered"
    assert "candidate adapter/lane hash does not match" in manifest["integrity"]["problems"]


def test_optional_render_uses_a_run_owned_profile(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "raw" / "adapter" / "lane" / "measured-1.pptx"
    source.parent.mkdir(parents=True)
    source.write_bytes(b"output")
    sample = Sample(
        phase="measured",
        iteration=1,
        outcome="success",
        reason=None,
        elapsed_ms=1,
        peak_rss_bytes=1,
        output_path="raw/adapter/lane/measured-1.pptx",
        output_sha256=_hash(source),
    )
    result = Result("lane", "mixed-60", "adapter", "success", None, None, [sample])
    captured: list[list[str]] = []

    def available_office() -> AdapterInfo:
        return AdapterInfo("libreoffice-render", "test", "test", True)

    def capture(
        command: list[str],
        _: Path,
        __: float,
        ___: float,
        ____: Path | None,
        _____: Path,
    ) -> dict[str, object]:
        captured.append(command)
        return {"outcome": "success"}

    monkeypatch.setattr(runner, "libreoffice_info", available_office)
    monkeypatch.setattr(runner, "_optional_command", capture)
    runner._render_evidence(tmp_path, [result], 10, float("inf"))
    assert captured[0][1] == (
        f"-env:UserInstallation={(tmp_path / 'render-profiles' / 'adapter' / 'lane' / 'measured-1').as_uri()}"
    )


def test_optional_timeout_and_quoted_validator_argv(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "input.pptx"
    source.write_bytes(b"input")
    sleeper = tmp_path / "sleeper.py"
    sleeper.write_text("import time\ntime.sleep(10)\n", encoding="utf-8")
    timed = runner._optional_command(
        [sys.executable, str(sleeper)],
        tmp_path,
        0.01,
        float("inf"),
        None,
        source,
    )
    assert timed["outcome"] == "timeout"

    quoted = tmp_path / "validator with spaces.py"
    quoted.write_text("", encoding="utf-8")
    sample = Sample(
        phase="measured",
        iteration=1,
        outcome="success",
        reason=None,
        elapsed_ms=1,
        peak_rss_bytes=1,
        output_path="input.pptx",
        output_sha256=_hash(source),
    )
    result = Result("lane", "mixed-60", "adapter", "success", None, None, [sample])
    captured: list[list[str]] = []

    def capture(
        command: list[str],
        _: Path,
        __: float,
        ___: float,
        ____: Path | None,
        _____: Path,
    ) -> dict[str, object]:
        captured.append(command)
        return {"outcome": "success"}

    monkeypatch.setenv(
        "PPTBENCH_OPENXML_VALIDATOR", f'{sys.executable} "{quoted}" --input {{input}}'
    )
    monkeypatch.setattr(runner, "_optional_command", capture)
    runner._openxml_evidence(tmp_path, [result], 10, float("inf"))
    assert captured[0][1] == str(quoted)
    assert captured[0][-1] == str(source)


def _hash(path: Path) -> str:
    import hashlib

    return hashlib.sha256(path.read_bytes()).hexdigest()
