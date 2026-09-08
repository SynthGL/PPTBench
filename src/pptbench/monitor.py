"""Execute one adapter child and write an isolated resource-use receipt."""

from __future__ import annotations

import argparse
import json
import os
import resource
import signal
import subprocess
import sys
from pathlib import Path
from time import perf_counter
from typing import Any


def _peak_rss_bytes() -> int:
    """Return this monitor's child peak RSS in normalized bytes."""
    rss = resource.getrusage(resource.RUSAGE_CHILDREN).ru_maxrss
    return int(rss if sys.platform == "darwin" else rss * 1024)


def _terminate_descendants(process: subprocess.Popen[bytes]) -> None:
    """Terminate the adapter process group before collecting its final output."""
    if process.poll() is not None:
        return
    if os.name == "posix":
        os.killpg(process.pid, signal.SIGTERM)
        try:
            process.wait(timeout=1)
        except subprocess.TimeoutExpired:
            os.killpg(process.pid, signal.SIGKILL)
    else:
        process.terminate()
        try:
            process.wait(timeout=1)
        except subprocess.TimeoutExpired:
            process.kill()


def execute(
    command: list[str],
    receipt_path: Path,
    stdout_path: Path,
    stderr_path: Path,
    timeout_seconds: float,
) -> dict[str, Any]:
    """Run one command; this fresh monitor owns all accumulated child rusage."""
    receipt_path.parent.mkdir(parents=True, exist_ok=True)
    started = perf_counter()
    with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
        process = subprocess.Popen(
            command,
            stdin=subprocess.DEVNULL,
            stdout=stdout,
            stderr=stderr,
            start_new_session=os.name == "posix",
        )
        try:
            return_code = process.wait(timeout=timeout_seconds)
            outcome = "success" if return_code == 0 else "failure"
        except subprocess.TimeoutExpired:
            _terminate_descendants(process)
            return_code = process.wait()
            outcome = "timeout"
    elapsed_ms = (perf_counter() - started) * 1000
    return {
        "outcome": outcome,
        "return_code": return_code,
        "elapsed_ms": elapsed_ms,
        "peak_rss_bytes": _peak_rss_bytes(),
        "stdout_bytes": stdout_path.stat().st_size,
        "stderr_bytes": stderr_path.stat().st_size,
    }


def main() -> int:
    parser = argparse.ArgumentParser(prog="pptbench-monitor")
    parser.add_argument("--receipt", type=Path, required=True)
    parser.add_argument("--stdout", type=Path, required=True)
    parser.add_argument("--stderr", type=Path, required=True)
    parser.add_argument("--timeout", type=float, required=True)
    parser.add_argument("command", nargs=argparse.REMAINDER)
    args = parser.parse_args()
    if args.timeout <= 0 or not args.command:
        return 2
    command = args.command[1:] if args.command[0] == "--" else args.command
    if not command:
        return 2
    try:
        receipt = execute(command, args.receipt, args.stdout, args.stderr, args.timeout)
    except (OSError, ValueError) as exc:
        receipt = {
            "outcome": "failure",
            "reason": f"monitor error: {type(exc).__name__}",
        }
    args.receipt.write_text(json.dumps(receipt, sort_keys=True), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
