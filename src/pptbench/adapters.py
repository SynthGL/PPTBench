"""Adapter discovery and isolated child-process execution."""

from __future__ import annotations

import hashlib
import importlib.metadata
import importlib.util
import json
import os
import re
import shlex
import shutil
import subprocess
import sys
from importlib.machinery import ModuleSpec
from pathlib import Path
from string import Formatter
from time import perf_counter
from typing import Any

from .models import AdapterInfo
from .util import sha256_file

_SAFE_SEGMENT = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")
_SAFE_TEMPLATE_FIELDS = frozenset({"input", "output", "lane"})
_PRIVATE_PATH = re.compile(r"(?:~|/(?:Users|home|var/folders)/)[^\s'\"]+")


def available_adapters() -> list[AdapterInfo]:
    return [
        python_pptx_info(),
        wolfppt_info(),
        external_command_info(),
        libreoffice_info(),
    ]


def python_pptx_info() -> AdapterInfo:
    return _package_info("python-pptx", "pptx", "python-pptx")


def wolfppt_info() -> AdapterInfo:
    return _package_info("wolfppt", "wolfppt", "wolfppt-wheel")


def _package_info(distribution: str, module: str, adapter: str) -> AdapterInfo:
    try:
        version = importlib.metadata.version(distribution)
    except importlib.metadata.PackageNotFoundError:
        return AdapterInfo(
            adapter,
            None,
            None,
            False,
            f"installed {distribution} distribution not found",
            f"install {distribution} into the benchmark environment",
        )
    spec = importlib.util.find_spec(module)
    identity = _package_identity(spec)
    if identity is None:
        return AdapterInfo(
            adapter,
            version,
            None,
            False,
            f"installed {distribution} package files are not inspectable",
            "reinstall a regular wheel or package distribution",
        )
    return AdapterInfo(adapter, version, identity, True)


def external_command_info() -> AdapterInfo:
    configured = os.environ.get("PPTBENCH_EXTERNAL_COMMAND")
    if not configured:
        return AdapterInfo(
            "external-command",
            None,
            None,
            False,
            "PPTBENCH_EXTERNAL_COMMAND is not configured",
            "set a command template containing {input}, {output}, and {lane}",
        )
    try:
        template = _external_template(configured)
    except ValueError as exc:
        return AdapterInfo(
            "external-command",
            None,
            None,
            False,
            str(exc),
            "use only {input}, {output}, and {lane} placeholders",
        )
    executable = shutil.which(template[0])
    if executable is None:
        return AdapterInfo(
            "external-command",
            None,
            Path(template[0]).name,
            False,
            f"configured executable is unavailable: {Path(template[0]).name}",
            "install the executable or update PPTBENCH_EXTERNAL_COMMAND",
        )
    binary = Path(executable).resolve()
    identity = _file_identity(binary)
    version = _executable_version(binary)
    config_hash = hashlib.sha256(configured.encode("utf-8")).hexdigest()
    return AdapterInfo(
        "external-command",
        version,
        f"{identity};config:{config_hash}",
        True,
    )


def libreoffice_info() -> AdapterInfo:
    executable = shutil.which("soffice")
    if executable is None:
        return AdapterInfo(
            "libreoffice-render",
            None,
            None,
            False,
            "soffice is not on PATH",
            "install LibreOffice or put soffice on PATH",
        )
    binary = Path(executable).resolve()
    return AdapterInfo(
        "libreoffice-render", _executable_version(binary), _file_identity(binary), True
    )


def run_adapter(
    adapter: str,
    lane: str,
    source: Path,
    output: Path,
    timeout_seconds: float,
) -> dict[str, Any]:
    """Run exactly one adapter child with fresh monitor-owned resource accounting."""
    if not _SAFE_SEGMENT.fullmatch(adapter) or not _SAFE_SEGMENT.fullmatch(lane):
        return {"outcome": "unsupported", "reason": "unsafe adapter or lane identifier"}
    if timeout_seconds <= 0:
        return {"outcome": "failure", "reason": "timeout must be positive"}
    info = {item.name: item for item in available_adapters()}.get(adapter)
    if info is None:
        return {"outcome": "unsupported", "reason": f"unknown adapter: {adapter}"}
    if not info.available:
        return {
            "outcome": "unavailable",
            "reason": info.reason,
            "recovery": info.recovery,
        }
    if not source.is_file() or source.is_symlink():
        return {"outcome": "failure", "reason": "source must be a regular file"}
    if output.exists() or output.is_symlink():
        return {
            "outcome": "failure",
            "reason": "output must be a new regular-file path",
        }

    try:
        command = _adapter_command(adapter, lane, source, output)
    except ValueError as exc:
        return {"outcome": "unavailable", "reason": str(exc)}
    output.parent.mkdir(parents=True, exist_ok=True)
    artifact_dir = output.with_suffix(output.suffix + ".adapter")
    if artifact_dir.exists():
        return {"outcome": "failure", "reason": "adapter artifact path already exists"}
    artifact_dir.mkdir()
    receipt_path = artifact_dir / "receipt.json"
    stdout_path = artifact_dir / "stdout.bin"
    stderr_path = artifact_dir / "stderr.bin"
    monitor = [
        sys.executable,
        "-m",
        "pptbench.monitor",
        "--receipt",
        str(receipt_path),
        "--stdout",
        str(stdout_path),
        "--stderr",
        str(stderr_path),
        "--timeout",
        str(timeout_seconds),
        "--",
        *command,
    ]
    started = perf_counter()
    try:
        completed = subprocess.run(
            monitor,
            stdin=subprocess.DEVNULL,
            capture_output=True,
            timeout=timeout_seconds + 5,
            check=False,
        )
    except subprocess.TimeoutExpired:
        return {
            "outcome": "timeout",
            "reason": f"monitor exceeded {timeout_seconds + 5:g}s",
            "elapsed_ms": (perf_counter() - started) * 1000,
            "details": _artifact_details(artifact_dir, stdout_path, stderr_path, command),
        }
    if completed.returncode != 0 or not receipt_path.is_file():
        return {
            "outcome": "failure",
            "reason": "adapter monitor did not produce a receipt",
            "elapsed_ms": (perf_counter() - started) * 1000,
            "details": _artifact_details(artifact_dir, stdout_path, stderr_path, command),
        }
    try:
        receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
        outcome = str(receipt["outcome"])
        elapsed_ms = float(receipt["elapsed_ms"])
        peak_rss_bytes = int(receipt["peak_rss_bytes"])
    except (KeyError, TypeError, ValueError, json.JSONDecodeError):
        return {
            "outcome": "failure",
            "reason": "adapter monitor receipt is malformed",
            "elapsed_ms": (perf_counter() - started) * 1000,
            "details": _artifact_details(artifact_dir, stdout_path, stderr_path, command),
        }
    details = _artifact_details(artifact_dir, stdout_path, stderr_path, command)
    if outcome == "timeout":
        return {
            "outcome": "timeout",
            "reason": f"exceeded {timeout_seconds:g}s",
            "elapsed_ms": elapsed_ms,
            "peak_rss_bytes": peak_rss_bytes,
            "details": details,
        }
    if outcome != "success":
        return {
            "outcome": "failure",
            "reason": f"child exited {receipt.get('return_code', 'unknown')}",
            "elapsed_ms": elapsed_ms,
            "peak_rss_bytes": peak_rss_bytes,
            "details": details,
        }
    if not output.is_file() or output.is_symlink():
        return {
            "outcome": "failure",
            "reason": "adapter exited without producing a regular output file",
            "elapsed_ms": elapsed_ms,
            "peak_rss_bytes": peak_rss_bytes,
            "details": details,
        }
    return {
        "outcome": "success",
        "elapsed_ms": elapsed_ms,
        "peak_rss_bytes": peak_rss_bytes,
        "output_sha256": sha256_file(output),
        "details": details,
    }


def _adapter_command(adapter: str, lane: str, source: Path, output: Path) -> list[str]:
    if adapter == "external-command":
        return [
            part.format(input=str(source), output=str(output), lane=lane)
            for part in _external_template(os.environ["PPTBENCH_EXTERNAL_COMMAND"])
        ]
    return [
        sys.executable,
        "-m",
        "pptbench.worker",
        adapter,
        lane,
        str(source),
        str(output),
    ]


def _external_template(configured: str) -> list[str]:
    try:
        command = shlex.split(configured)
    except ValueError as exc:
        raise ValueError("external command is not valid shell-tokenized argv") from exc
    if not command:
        raise ValueError("external command is empty")
    fields: set[str] = set()
    formatter = Formatter()
    for part in command:
        try:
            parsed = list(formatter.parse(part))
        except ValueError as exc:
            raise ValueError("external command template is malformed") from exc
        for _, field, spec, conversion in parsed:
            if field is None:
                continue
            if field not in _SAFE_TEMPLATE_FIELDS or spec or conversion:
                raise ValueError("external command uses an unsafe template field")
            fields.add(field)
    if fields != _SAFE_TEMPLATE_FIELDS:
        raise ValueError("external command must contain {input}, {output}, and {lane}")
    return command


def _package_identity(spec: ModuleSpec | None) -> str | None:
    if spec is None or not spec.submodule_search_locations:
        return _module_identity(spec)
    roots = [Path(location) for location in spec.submodule_search_locations]
    entries = [
        (path.relative_to(root), path)
        for root in roots
        for path in root.rglob("*")
        if path.is_file() and path.suffix.lower() in {".py", ".so", ".dylib", ".pyd"}
    ]
    if not entries:
        return None
    digest = hashlib.sha256()
    for relative_path, path in sorted(entries):
        digest.update(relative_path.as_posix().encode())
        digest.update(bytes.fromhex(sha256_file(path)))
    return f"modules:{len(entries)}:{digest.hexdigest()}"


def _module_identity(spec: ModuleSpec | None) -> str | None:
    if spec is None or spec.origin is None:
        return None
    origin = Path(spec.origin)
    return _file_identity(origin) if origin.is_file() else None


def _file_identity(path: Path) -> str:
    return f"{path.name}:{sha256_file(path)}"


def _executable_version(executable: Path) -> str | None:
    try:
        completed = subprocess.run(
            [str(executable), "--version"],
            stdin=subprocess.DEVNULL,
            capture_output=True,
            timeout=3,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return None
    line = (completed.stdout or completed.stderr).decode("utf-8", errors="replace").splitlines()
    return _sanitize_text(line[0])[:200] if line else None


def _artifact_details(
    artifact_dir: Path, stdout_path: Path, stderr_path: Path, command: list[str]
) -> dict[str, Any]:
    return {
        "command": _redacted_command(command),
        "artifact_dir": artifact_dir.name,
        "stdout_artifact": stdout_path.name,
        "stderr_artifact": stderr_path.name,
        "stdout_bytes": stdout_path.stat().st_size if stdout_path.exists() else 0,
        "stderr_bytes": stderr_path.stat().st_size if stderr_path.exists() else 0,
        "diagnostic": _diagnostic(stderr_path),
    }


def _diagnostic(path: Path) -> str | None:
    if not path.is_file() or path.stat().st_size == 0:
        return None
    return _sanitize_text(path.read_bytes()[:4096].decode("utf-8", errors="replace"))[:1000]


def _sanitize_text(value: str) -> str:
    return _PRIVATE_PATH.sub("<private-path>", value).replace("\x00", "")


def _redacted_command(command: list[str]) -> list[str]:
    redacted: list[str] = []
    for part in command:
        redacted.append(Path(part).name if Path(part).is_absolute() else _sanitize_text(part))
    return redacted
