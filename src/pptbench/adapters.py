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
from collections.abc import Callable
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
_SCRIPT_EXEC = re.compile(r'exec\s+"([^"]+)"')
_PY_VERSION = re.compile(r'#define\s+PY_VERSION\s+"([^"]+)"')

CONTAINER_IMAGES = {
    "apache-poi": "pptbench/apache-poi:5.5.1",
    "open-xml-sdk": "pptbench/open-xml-sdk:3.5.1",
}
LIBREOFFICE_SCRIPT = Path("libreoffice") / "pptbench_uno.py"
AUTOMIZER_HELPER = Path("pptx-automizer") / "pptbench-automizer.js"


def available_adapters() -> list[AdapterInfo]:
    return [factory() for factory in _ADAPTER_INFO.values()]


def adapter_home() -> Path | None:
    """Directory holding helper programs: PPTBENCH_ADAPTER_HOME or a source checkout's."""
    configured = os.environ.get("PPTBENCH_ADAPTER_HOME")
    home = Path(configured) if configured else Path(__file__).resolve().parents[2] / "adapters"
    return home.resolve() if home.is_dir() else None


def docker_command() -> list[str]:
    """Docker CLI prefix, honoring PPTBENCH_DOCKER_CONTEXT for remote engines."""
    context = os.environ.get("PPTBENCH_DOCKER_CONTEXT")
    return ["docker", "--context", context] if context else ["docker"]


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


def libreoffice_uno_info() -> AdapterInfo:
    """LibreOffice Impress driven through a Python-UNO script in a fresh profile."""
    executable = shutil.which("soffice")
    if executable is None:
        return AdapterInfo(
            "libreoffice",
            None,
            None,
            False,
            "soffice is not on PATH",
            "install LibreOffice or put soffice on PATH",
        )
    script = _helper_file(LIBREOFFICE_SCRIPT)
    if script is None:
        return _missing_helper("libreoffice", LIBREOFFICE_SCRIPT)
    binary = Path(executable).resolve()
    banner = _executable_version(binary)
    if banner is None:
        return AdapterInfo(
            "libreoffice",
            None,
            _file_identity(binary),
            False,
            "soffice --version produced no version banner",
            "repair the LibreOffice installation",
        )
    parts = banner.split()
    version = parts[1] if len(parts) > 1 else banner
    build = parts[2] if len(parts) > 2 else "unknown"
    python = _libreoffice_python_version(binary)
    return AdapterInfo(
        "libreoffice",
        version,
        f"libreoffice:{version};build:{build};runtime:LibreOffice embedded Python {python};"
        f"{_file_identity(binary)};script:{sha256_file(script)}",
        True,
    )


def pptx_automizer_info() -> AdapterInfo:
    """pptx-automizer from the helper's npm lockfile, run by the local Node.js."""
    node = shutil.which("node")
    if node is None:
        return AdapterInfo(
            "pptx-automizer",
            None,
            None,
            False,
            "node is not on PATH",
            "install Node.js 22 or newer",
        )
    helper = _helper_file(AUTOMIZER_HELPER)
    if helper is None:
        return _missing_helper("pptx-automizer", AUTOMIZER_HELPER)
    package = helper.parent / "node_modules" / "pptx-automizer" / "package.json"
    lockfile = helper.parent / "package-lock.json"
    if not package.is_file() or not lockfile.is_file():
        return AdapterInfo(
            "pptx-automizer",
            None,
            None,
            False,
            "pptx-automizer is not installed next to the helper",
            f"run npm ci in {AUTOMIZER_HELPER.parent.as_posix()} under the adapter home",
        )
    try:
        version = str(json.loads(package.read_text(encoding="utf-8"))["version"])
    except (KeyError, OSError, ValueError):
        return AdapterInfo(
            "pptx-automizer",
            None,
            None,
            False,
            "installed pptx-automizer package.json is unreadable",
            f"run npm ci in {AUTOMIZER_HELPER.parent.as_posix()} under the adapter home",
        )
    runtime = _executable_version(Path(node).resolve()) or "unknown"
    return AdapterInfo(
        "pptx-automizer",
        version,
        f"pptx-automizer:{version};runtime:Node.js {runtime};"
        f"lockfile:{sha256_file(lockfile)};helper:{sha256_file(helper)}",
        True,
    )


def apache_poi_info() -> AdapterInfo:
    return _container_info("apache-poi")


def open_xml_sdk_info() -> AdapterInfo:
    return _container_info("open-xml-sdk")


def _container_info(adapter: str) -> AdapterInfo:
    """Identity of a locally built helper image on the configured Docker engine."""
    tag = CONTAINER_IMAGES[adapter]
    build = f"docker build -t {tag} adapters/{adapter} (honor PPTBENCH_DOCKER_CONTEXT)"
    if shutil.which("docker") is None:
        return AdapterInfo(adapter, None, None, False, "docker is not on PATH", build)
    try:
        completed = subprocess.run(
            [*docker_command(), "image", "inspect", tag],
            stdin=subprocess.DEVNULL,
            capture_output=True,
            timeout=30,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return AdapterInfo(
            adapter, None, None, False, "docker image inspect did not complete", build
        )
    if completed.returncode != 0:
        return AdapterInfo(adapter, None, None, False, f"helper image {tag} is not built", build)
    try:
        image = json.loads(completed.stdout)[0]
        labels = image["Config"]["Labels"] or {}
        library = str(labels["org.pptbench.library"])
        version = str(labels["org.pptbench.library-version"])
        runtime = str(labels["org.pptbench.runtime"])
        platform = f"{image['Os']}/{image['Architecture']}"
        image_id = str(image["Id"])
    except (IndexError, KeyError, TypeError, ValueError):
        return AdapterInfo(
            adapter,
            None,
            None,
            False,
            f"helper image {tag} lacks PPTBench labels",
            build,
        )
    if labels.get("org.pptbench.adapter") != adapter:
        return AdapterInfo(
            adapter,
            None,
            None,
            False,
            f"helper image {tag} belongs to another adapter",
            build,
        )
    return AdapterInfo(
        adapter,
        version,
        f"{library}:{version};runtime:{runtime};platform:{platform};image:{image_id}",
        True,
    )


def _helper_file(relative_path: Path) -> Path | None:
    home = adapter_home()
    if home is None:
        return None
    path = home / relative_path
    return path if path.is_file() and not path.is_symlink() else None


def _missing_helper(adapter: str, relative_path: Path) -> AdapterInfo:
    return AdapterInfo(
        adapter,
        None,
        None,
        False,
        f"helper program {relative_path.as_posix()} is not available",
        "run from a PPTBench source checkout or set PPTBENCH_ADAPTER_HOME to its adapters/",
    )


def _libreoffice_python_version(binary: Path) -> str:
    """Exact version of the Python embedded in the LibreOffice installation, if found."""
    program = binary
    if binary.read_bytes()[:2] == b"#!":
        # Package-manager wrappers (e.g. Homebrew casks) exec the real binary.
        match = _SCRIPT_EXEC.search(binary.read_text(encoding="utf-8", errors="replace"))
        if match is not None:
            program = Path(match.group(1)).resolve()
    install = program.parent.parent
    patterns = (
        "Frameworks/LibreOfficePython.framework/Versions/*/include/python*/patchlevel.h",
        "program/python-core-*/include/python*/patchlevel.h",
    )
    for pattern in patterns:
        for header in sorted(install.glob(pattern)):
            match = _PY_VERSION.search(header.read_text(encoding="utf-8", errors="replace"))
            if match is not None:
                return match.group(1)
    return "unresolved"


_ADAPTER_INFO: dict[str, Callable[[], AdapterInfo]] = {
    "python-pptx": python_pptx_info,
    "wolfppt-wheel": wolfppt_info,
    "apache-poi": apache_poi_info,
    "libreoffice": libreoffice_uno_info,
    "pptx-automizer": pptx_automizer_info,
    "open-xml-sdk": open_xml_sdk_info,
    "external-command": external_command_info,
    "libreoffice-render": libreoffice_info,
}


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
    factory = _ADAPTER_INFO.get(adapter)
    if factory is None:
        return {"outcome": "unsupported", "reason": f"unknown adapter: {adapter}"}
    info = factory()
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
    if adapter == "pptx-automizer":
        node = shutil.which("node")
        helper = _helper_file(AUTOMIZER_HELPER)
        if node is None or helper is None:
            raise ValueError("pptx-automizer helper or node became unavailable")
        return [node, str(helper), lane, str(source), str(output)]
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
