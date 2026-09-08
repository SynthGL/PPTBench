"""Typed, JSON-serializable records used by PPTBench."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any

SCHEMA_VERSION = 2
OUTCOMES = frozenset({"success", "failure", "unsupported", "unavailable", "timeout"})


@dataclass(frozen=True)
class AdapterInfo:
    name: str
    version: str | None
    identity: str | None
    available: bool
    reason: str | None = None
    recovery: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class Sample:
    phase: str
    iteration: int
    outcome: str
    reason: str | None
    elapsed_ms: float | None
    peak_rss_bytes: int | None
    output_path: str | None
    output_sha256: str | None
    checks: list[dict[str, Any]] = field(default_factory=list)
    details: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class Result:
    lane: str
    fixture: str
    adapter: str
    outcome: str
    reason: str | None
    input_sha256: str | None
    samples: list[Sample]
    checks: list[dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)
