"""Run controlled editing lanes and retain raw evidence for every outcome."""

from __future__ import annotations

import os
import re
import shlex
import shutil
import signal
import subprocess
import zipfile
from datetime import UTC, datetime
from pathlib import Path
from time import perf_counter
from typing import Any

from .adapters import available_adapters, libreoffice_info, run_adapter
from .fixtures import Manifest, fixture_inventory, manifest, materialize
from .models import SCHEMA_VERSION, Result, Sample
from .scoring import score
from .util import environment_identity, json_dump, relative, sha256_file

DEFAULT_LANES = ("feature-matrix", "template-mutation", "chart-data")
DEFAULT_ADAPTERS = ("python-pptx", "wolfppt-wheel")
_SAFE_REVISION = re.compile(r"^[0-9a-f]{7,64}$")


def benchmark(
    output_root: Path,
    *,
    lanes: list[str] | None = None,
    adapters: list[str] | None = None,
    fixtures: list[str] | None = None,
    iterations: int = 1,
    warmups: int = 0,
    timeout_seconds: float = 60,
    overall_timeout_seconds: float = 600,
    reuse_fixtures: Path | None = None,
    render: bool = False,
    openxml_validate: bool = False,
    wolfppt_artifact: Path | None = None,
    wolfppt_source_revision: str | None = None,
    wolfppt_build_mode: str | None = None,
) -> Path:
    """Create one immutable benchmark run below a new or empty output root."""
    if iterations < 1 or warmups < 0 or timeout_seconds <= 0 or overall_timeout_seconds <= 0:
        raise ValueError("iterations >= 1, warmups >= 0, and positive timeouts are required")
    selected_lanes = tuple(lanes or DEFAULT_LANES)
    selected_adapters = tuple(adapters or DEFAULT_ADAPTERS)
    if not selected_lanes or not selected_adapters:
        raise ValueError("at least one lane and adapter are required")
    unknown_lanes = set(selected_lanes) - set(DEFAULT_LANES)
    known_adapters = {item.name for item in available_adapters()}
    unknown_adapters = set(selected_adapters) - known_adapters
    if unknown_lanes or unknown_adapters:
        problems = [
            *(f"unknown lane: {item}" for item in sorted(unknown_lanes)),
            *(f"unknown adapter: {item}" for item in sorted(unknown_adapters)),
        ]
        raise ValueError("; ".join(problems))
    fixture_manifest = manifest()
    known_fixtures = {str(item["id"]) for item in fixture_manifest["corpus"]["fixtures"]}
    if fixtures and (unknown_fixtures := set(fixtures) - known_fixtures):
        raise ValueError(f"unknown fixtures: {', '.join(sorted(unknown_fixtures))}")
    run_root = output_root.resolve()
    _require_new_or_empty_directory(run_root)
    wolfppt_receipt = _wolfppt_receipt(
        wolfppt_artifact, wolfppt_source_revision, wolfppt_build_mode
    )
    run_root.mkdir(parents=True, exist_ok=False)
    fixture_paths = _prepare_fixtures(run_root, reuse_fixtures, fixture_manifest)

    deadline = perf_counter() + overall_timeout_seconds
    results: list[Result] = []
    for lane in selected_lanes:
        fixture_id = "charts" if lane == "chart-data" else "mixed-60"
        if fixtures and fixture_id not in fixtures:
            continue
        source = fixture_paths[fixture_id]
        for adapter in selected_adapters:
            samples: list[Sample] = []
            for phase, count in (("warmup", warmups), ("measured", iterations)):
                for iteration in range(1, count + 1):
                    target = run_root / "raw" / adapter / lane / f"{phase}-{iteration}.pptx"
                    remaining = deadline - perf_counter()
                    if remaining <= 0:
                        raw: dict[str, Any] = {
                            "outcome": "timeout",
                            "reason": "overall run budget exhausted before child start",
                        }
                    else:
                        raw = run_adapter(
                            adapter,
                            lane,
                            source,
                            target,
                            min(timeout_seconds, remaining),
                        )
                    samples.append(
                        _sample_from_raw(phase, iteration, raw, lane, source, target, run_root)
                    )
            checks = [check for sample in samples for check in sample.checks]
            failures = [sample for sample in samples if sample.outcome != "success"]
            results.append(
                Result(
                    lane=lane,
                    fixture=fixture_id,
                    adapter=adapter,
                    outcome="success" if not failures else failures[0].outcome,
                    reason=None if not failures else failures[0].reason,
                    input_sha256=sha256_file(source),
                    samples=samples,
                    checks=checks,
                )
            )
    report = {
        "schema_version": SCHEMA_VERSION,
        "created_at_utc": datetime.now(UTC).isoformat(),
        "status": _overall_status(results),
        "configuration": {
            "lanes": list(selected_lanes),
            "adapters": list(selected_adapters),
            "fixtures": fixtures or sorted(fixture_paths),
            "iterations": iterations,
            "warmups": warmups,
            "timeout_seconds": timeout_seconds,
            "overall_timeout_seconds": overall_timeout_seconds,
            "cache_mode": "validated-fixture-reuse" if reuse_fixtures else "fresh-fixtures",
            "cold_subprocess": True,
            "warmup_disclosure": "warmups are retained but do not remove interpreter startup",
            "wolfppt_receipt": wolfppt_receipt,
        },
        "environment": environment_identity(),
        "adapters": [item.to_dict() for item in available_adapters()],
        "manifest": fixture_manifest,
        "fixtures": fixture_inventory(run_root),
        "results": [item.to_dict() for item in results],
        "optional_evidence": {
            "render": _render_evidence(run_root, results, timeout_seconds, deadline)
            if render
            else [],
            "openxml_validation": _openxml_evidence(run_root, results, timeout_seconds, deadline)
            if openxml_validate
            else [],
        },
    }
    json_dump(run_root / "run.json", report)
    return run_root / "run.json"


def _require_new_or_empty_directory(path: Path) -> None:
    if path.exists() and (not path.is_dir() or path.is_symlink() or any(path.iterdir())):
        raise ValueError("output root must be a new or empty non-symlink directory")


def _prepare_fixtures(
    run_root: Path, reuse_fixtures: Path | None, fixture_manifest: Manifest
) -> dict[str, Path]:
    if reuse_fixtures is None:
        return materialize(run_root)
    source_root = reuse_fixtures.resolve()
    if (source_root / "fixtures").is_dir():
        source_root = source_root / "fixtures"
    expected = {str(item["id"]): item for item in fixture_manifest["corpus"]["fixtures"]}
    destination = run_root / "fixtures"
    destination.mkdir()
    for fixture_id, expected_data in expected.items():
        source = source_root / f"{fixture_id}.pptx"
        if (
            not source.is_file()
            or source.is_symlink()
            or not zipfile.is_zipfile(source)
            or sha256_file(source) != expected_data["sha256"]
        ):
            raise ValueError(f"fixture reuse validation failed: {fixture_id}")
        shutil.copyfile(source, destination / source.name)
    return materialize(run_root)


def _sample_from_raw(
    phase: str,
    iteration: int,
    raw: dict[str, Any],
    lane: str,
    source: Path,
    target: Path,
    run_root: Path,
) -> Sample:
    outcome = str(raw["outcome"])
    reason = raw.get("reason")
    checks: list[dict[str, Any]] = []
    if outcome == "success":
        try:
            checks = score(lane, source, target)
        except (KeyError, OSError, TypeError, ValueError, zipfile.BadZipFile) as exc:
            outcome = "failure"
            reason = f"scoring error: {type(exc).__name__}"
        else:
            if any(_is_score_failure(check) for check in checks):
                outcome = "failure"
                reason = "one or more scored semantic or preservation checks failed"
    details = _relative_details(dict(raw.get("details", {})), target, run_root)
    if target.is_file():
        details["candidate_path"] = relative(target, run_root)
        details["candidate_sha256"] = sha256_file(target)
    return Sample(
        phase=phase,
        iteration=iteration,
        outcome=outcome,
        reason=str(reason) if reason else None,
        elapsed_ms=_number_or_none(raw.get("elapsed_ms")),
        peak_rss_bytes=_integer_or_none(raw.get("peak_rss_bytes")),
        output_path=relative(target, run_root)
        if outcome == "success" and target.is_file()
        else None,
        output_sha256=raw.get("output_sha256") if outcome == "success" else None,
        checks=checks,
        details=details,
    )


def _is_score_failure(check: dict[str, Any]) -> bool:
    return check.get("outcome") != "success" and bool(check.get("scored", True))


def _number_or_none(value: Any) -> float | None:
    return float(value) if isinstance(value, int | float) else None


def _integer_or_none(value: Any) -> int | None:
    return int(value) if isinstance(value, int) else None


def _relative_details(details: dict[str, Any], target: Path, run_root: Path) -> dict[str, Any]:
    artifact_dir = details.pop("artifact_dir", None)
    if isinstance(artifact_dir, str):
        directory = target.parent / artifact_dir
        if directory.is_dir():
            details["artifact_dir"] = relative(directory, run_root)
            for key in ("stdout_artifact", "stderr_artifact"):
                name = details.get(key)
                if isinstance(name, str) and (directory / name).is_file():
                    details[key] = relative(directory / name, run_root)
    return details


def _wolfppt_receipt(
    artifact: Path | None, source_revision: str | None, build_mode: str | None
) -> dict[str, Any]:
    if source_revision is not None and not _SAFE_REVISION.fullmatch(source_revision):
        raise ValueError("WolfPPT source revision must be a hexadecimal revision")
    if artifact is None:
        if build_mode is not None:
            raise ValueError("WolfPPT build mode requires an artifact path")
        return {
            "build_mode": "unknown",
            "artifact": None,
            "source_revision": source_revision,
        }
    if not artifact.is_file() or artifact.is_symlink():
        raise ValueError("WolfPPT artifact must be a regular file")
    if build_mode not in {None, "provided-release-wheel", "source-build-wheel"}:
        raise ValueError("unknown WolfPPT build mode")
    return {
        "build_mode": build_mode or "unknown",
        "artifact": {"name": artifact.name, "sha256": sha256_file(artifact)},
        "source_revision": source_revision,
    }


def _overall_status(results: list[Result]) -> str:
    if not results:
        return "failure"
    outcomes = {result.outcome for result in results}
    if "failure" in outcomes or "timeout" in outcomes:
        return "failure"
    if outcomes <= {"unavailable", "unsupported"}:
        return "unavailable"
    return "success"


def _render_evidence(
    run_root: Path, results: list[Result], timeout_seconds: float, deadline: float
) -> list[dict[str, Any]]:
    info = libreoffice_info()
    if not info.available:
        return [{"outcome": "unavailable", "reason": info.reason, "recovery": info.recovery}]
    evidence: list[dict[str, Any]] = []
    for result in results:
        for sample in result.samples:
            if sample.output_path is None:
                continue
            source = run_root / sample.output_path
            target_dir = run_root / "render" / result.adapter / result.lane
            target_dir.mkdir(parents=True, exist_ok=True)
            output = target_dir / f"{source.stem}.pdf"
            profile = (
                run_root
                / "render-profiles"
                / result.adapter
                / result.lane
                / f"{sample.phase}-{sample.iteration}"
            )
            command = [
                "soffice",
                f"-env:UserInstallation={profile.resolve().as_uri()}",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(target_dir),
                str(source),
            ]
            evidence.append(
                _optional_command(command, run_root, timeout_seconds, deadline, output, source)
            )
    return evidence


def _openxml_evidence(
    run_root: Path, results: list[Result], timeout_seconds: float, deadline: float
) -> list[dict[str, Any]]:
    template = os.environ.get("PPTBENCH_OPENXML_VALIDATOR")
    if not template:
        return [
            {
                "outcome": "unavailable",
                "reason": "PPTBENCH_OPENXML_VALIDATOR is not configured",
            }
        ]
    try:
        command_template = shlex.split(template)
    except ValueError:
        return [{"outcome": "failure", "reason": "PPTBENCH_OPENXML_VALIDATOR is malformed"}]
    if not command_template:
        return [{"outcome": "failure", "reason": "PPTBENCH_OPENXML_VALIDATOR is empty"}]
    evidence: list[dict[str, Any]] = []
    for result in results:
        for sample in result.samples:
            if sample.output_path is None:
                continue
            source = run_root / sample.output_path
            try:
                command = [part.format(input=str(source)) for part in command_template]
            except (IndexError, KeyError, ValueError):
                return [
                    {
                        "outcome": "failure",
                        "reason": "PPTBENCH_OPENXML_VALIDATOR has an invalid template field",
                    }
                ]
            evidence.append(
                _optional_command(command, run_root, timeout_seconds, deadline, None, source)
            )
    return evidence


def _optional_command(
    command: list[str],
    run_root: Path,
    timeout_seconds: float,
    deadline: float,
    expected: Path | None,
    source: Path,
) -> dict[str, Any]:
    if expected is not None and expected.exists():
        return _optional_receipt(
            "failure",
            "optional evidence output path already exists",
            command,
            run_root,
            source,
            expected,
            0,
            0,
            None,
        )
    remaining = deadline - perf_counter()
    if remaining <= 0:
        return _optional_receipt(
            "timeout",
            "overall run budget exhausted before optional evidence",
            command,
            run_root,
            source,
            expected,
            0,
            0,
            None,
        )
    executable = shutil.which(command[0])
    if executable is None:
        return _optional_receipt(
            "unavailable",
            f"executable unavailable: {Path(command[0]).name}",
            command,
            run_root,
            source,
            expected,
            0,
            0,
            None,
        )
    started = perf_counter()
    process = subprocess.Popen(
        command,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        start_new_session=os.name == "posix",
    )
    try:
        stdout, stderr = process.communicate(timeout=min(timeout_seconds, remaining))
    except subprocess.TimeoutExpired:
        _terminate_optional_process_group(process)
        stdout, stderr = process.communicate()
        return _optional_receipt(
            "timeout",
            f"exceeded {min(timeout_seconds, remaining):g}s",
            command,
            run_root,
            source,
            expected,
            len(stdout),
            len(stderr),
            (perf_counter() - started) * 1000,
        )
    outcome = (
        "success"
        if process.returncode == 0 and (expected is None or expected.is_file())
        else "failure"
    )
    reason = None if outcome == "success" else f"child exited {process.returncode}"
    return _optional_receipt(
        outcome,
        reason,
        command,
        run_root,
        source,
        expected,
        len(stdout),
        len(stderr),
        (perf_counter() - started) * 1000,
    )


def _terminate_optional_process_group(process: subprocess.Popen[bytes]) -> None:
    if process.poll() is not None:
        return
    try:
        if os.name == "posix":
            os.killpg(process.pid, signal.SIGTERM)
        else:
            process.terminate()
        process.wait(timeout=1)
    except (OSError, subprocess.TimeoutExpired):
        if process.poll() is None:
            if os.name == "posix":
                try:
                    os.killpg(process.pid, signal.SIGKILL)
                except OSError:
                    pass
            else:
                process.kill()


def _optional_receipt(
    outcome: str,
    reason: str | None,
    command: list[str],
    run_root: Path,
    source: Path,
    expected: Path | None,
    stdout_bytes: int,
    stderr_bytes: int,
    elapsed_ms: float | None,
) -> dict[str, Any]:
    receipt: dict[str, Any] = {
        "outcome": outcome,
        "command": _redact_optional_command(command, run_root),
        "source_path": relative(source, run_root),
        "source_sha256": sha256_file(source),
        "stdout_bytes": stdout_bytes,
        "stderr_bytes": stderr_bytes,
        "elapsed_ms": elapsed_ms,
    }
    if reason is not None:
        receipt["reason"] = reason
    if expected is not None:
        receipt["output_path"] = relative(expected, run_root)
        if expected.is_file():
            receipt["output_sha256"] = sha256_file(expected)
    return receipt


def _redact_optional_command(command: list[str], run_root: Path) -> list[str]:
    redacted: list[str] = []
    root_text = str(run_root)
    for item in command:
        if item.startswith("-env:UserInstallation="):
            redacted.append("-env:UserInstallation=<run-owned-profile>")
        elif Path(item).is_absolute():
            redacted.append(Path(item).name)
        else:
            redacted.append(item.replace(root_text, "<run-root>"))
    return redacted
