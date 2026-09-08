"""HTML dashboard and SVG heatmap generated from a frozen raw report."""

from __future__ import annotations

import html
import json
from pathlib import Path, PurePosixPath
from typing import Any

from .models import SCHEMA_VERSION
from .util import json_dump, sha256_file


def generate(run_json: Path, output_dir: Path | None = None) -> tuple[Path, Path]:
    raw = json.loads(run_json.read_text(encoding="utf-8"))
    if raw.get("schema_version") != SCHEMA_VERSION:
        raise ValueError("unsupported report schema")
    target = output_dir or run_json.parent / "report"
    target.mkdir(parents=True, exist_ok=True)
    integrity = _verify_integrity(raw, run_json, target / "report-manifest.json")
    heatmap = target / "heatmap.svg"
    heatmap.write_text(_svg(raw), encoding="utf-8")
    dashboard = target / "index.html"
    dashboard.write_text(_html(raw, run_json, heatmap, integrity), encoding="utf-8")
    json_dump(
        target / "report-manifest.json",
        {
            "schema_version": SCHEMA_VERSION,
            "raw_sha256": sha256_file(run_json),
            "html": dashboard.name,
            "svg": heatmap.name,
            "integrity": integrity,
        },
    )
    return dashboard, heatmap


def _verify_integrity(
    raw: dict[str, Any], run_json: Path, previous_manifest_path: Path
) -> dict[str, Any]:
    run_root = run_json.parent.resolve()
    problems: list[str] = []
    previous_hash = _previous_raw_hash(previous_manifest_path)
    current_hash = sha256_file(run_json)
    if previous_hash is not None and previous_hash != current_hash:
        problems.append("raw report changed since the previous report generation")
    for fixture in raw.get("fixtures", []):
        fixture_id = str(fixture.get("id", "unknown"))
        fixture_path = run_root / "fixtures" / str(fixture.get("path", ""))
        _verify_hash(fixture_path, fixture.get("sha256"), f"fixture {fixture_id}", problems)
    for result in raw.get("results", []):
        label = f"{result.get('adapter', 'unknown')}/{result.get('lane', 'unknown')}"
        input_hash = result.get("input_sha256")
        fixture_name = str(result.get("fixture", ""))
        _verify_hash(
            run_root / "fixtures" / f"{fixture_name}.pptx",
            input_hash,
            f"input {label}",
            problems,
        )
        for sample in result.get("samples", []):
            output_path = sample.get("output_path")
            output_hash = sample.get("output_sha256")
            if output_path is not None or output_hash is not None:
                artifact_path = _run_relative_path(
                    run_root, output_path, f"output {label}", problems
                )
                if artifact_path is not None:
                    _verify_hash(artifact_path, output_hash, f"output {label}", problems)
            details = sample.get("details", {})
            candidate_path = details.get("candidate_path") if isinstance(details, dict) else None
            candidate_hash = details.get("candidate_sha256") if isinstance(details, dict) else None
            if (candidate_path is None and candidate_hash is None) or (
                candidate_path == output_path and candidate_hash == output_hash
            ):
                continue
            artifact_path = _run_relative_path(
                run_root, candidate_path, f"candidate {label}", problems
            )
            if artifact_path is not None:
                _verify_hash(artifact_path, candidate_hash, f"candidate {label}", problems)
    return {"outcome": "success" if not problems else "tampered", "problems": problems}


def _previous_raw_hash(path: Path) -> str | None:
    if not path.is_file():
        return None
    try:
        previous = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    value = previous.get("raw_sha256")
    return value if isinstance(value, str) else None


def _run_relative_path(root: Path, value: object, label: str, problems: list[str]) -> Path | None:
    if not isinstance(value, str):
        problems.append(f"{label} path is missing")
        return None
    path = PurePosixPath(value)
    if path.is_absolute() or ".." in path.parts:
        problems.append(f"{label} path is unsafe")
        return None
    return root.joinpath(*path.parts)


def _verify_hash(path: Path, expected: object, label: str, problems: list[str]) -> None:
    if not isinstance(expected, str) or len(expected) != 64:
        problems.append(f"{label} expected hash is malformed")
    elif not path.is_file():
        problems.append(f"{label} artifact is missing")
    elif sha256_file(path) != expected:
        problems.append(f"{label} hash does not match")


def _html(raw: dict[str, Any], run_json: Path, heatmap: Path, integrity: dict[str, Any]) -> str:
    rows = raw.get("results", [])
    title = html.escape(f"PPTBench report: {run_json.name}")
    detail_rows = "\n".join(_result_row(row) for row in rows)
    adapters = sorted({str(row["adapter"]) for row in rows})
    lanes = sorted({str(row["lane"]) for row in rows})
    integrity_text = html.escape("; ".join(integrity["problems"]) or "hashes verified")
    return f"""<!doctype html>
<html lang='en'><head><meta charset='utf-8'><title>{title}</title>
<style>body{{font:14px system-ui;margin:2rem;color:#18212b}}table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #ccd;padding:.45rem;text-align:left;vertical-align:top}}pre{{white-space:pre-wrap;max-width:70rem}}.success{{color:#146c43}}.failure,.timeout,.tampered{{color:#b42318}}.unavailable,.unsupported{{color:#7a5700}}</style></head>
<body><h1>{title}</h1><p>Raw evidence: <code>{html.escape(run_json.name)}</code>, SHA-256 <code>{sha256_file(run_json)}</code>. Integrity: <strong class='{html.escape(str(integrity["outcome"]))}'>{html.escape(str(integrity["outcome"]))}</strong>; {integrity_text}.</p>
<p>{_select("adapter", adapters, "adapter")} {_select("lane", lanes, "lane")}</p>
<object type='image/svg+xml' data='{html.escape(heatmap.name)}' aria-label='PPTBench status heatmap'></object>
<table><thead><tr><th>adapter</th><th>lane</th><th>fixture</th><th>outcome</th><th>iterations, reasons, and feature results</th><th>result reason</th></tr></thead><tbody>{detail_rows}</tbody></table>
<script>for(const id of ['adapter','lane'])document.getElementById(id).onchange=()=>{{for(const row of document.querySelectorAll('tbody tr'))row.hidden=['adapter','lane'].some(key=>document.getElementById(key).value&&row.dataset[key]!==document.getElementById(key).value)}};</script>
</body></html>"""


def _select(label: str, values: list[str], identifier: str) -> str:
    options = "".join(f"<option>{html.escape(value)}</option>" for value in values)
    return f"<label>{label} <select id='{identifier}'><option value=''>all</option>{options}</select></label>"


def _result_row(row: dict[str, Any]) -> str:
    adapter = html.escape(str(row.get("adapter", "")))
    lane = html.escape(str(row.get("lane", "")))
    samples = html.escape(json.dumps(row.get("samples", []), indent=2, sort_keys=True))
    return (
        f"<tr data-adapter='{adapter}' data-lane='{lane}'><td>{adapter}</td><td>{lane}</td>"
        f"<td>{html.escape(str(row.get('fixture', '')))}</td><td>{html.escape(str(row.get('outcome', '')))}</td>"
        f"<td><details><summary>all retained samples</summary><pre>{samples}</pre></details></td>"
        f"<td>{html.escape(str(row.get('reason') or ''))}</td></tr>"
    )


def _svg(raw: dict[str, Any]) -> str:
    rows = raw.get("results", [])
    adapters = sorted({str(row["adapter"]) for row in rows})
    lanes = sorted({str(row["lane"]) for row in rows})
    left, cell = 190, 170
    width = max(360, left + cell * len(lanes) + 20)
    height = max(120, 80 + 54 * len(adapters))
    colors = {
        "success": "#2d8a56",
        "failure": "#c83d34",
        "timeout": "#bb6b00",
        "unavailable": "#8b8b8b",
        "unsupported": "#6b5ca5",
    }
    states = {(str(row["adapter"]), str(row["lane"])): str(row["outcome"]) for row in rows}
    body = [
        f"<svg xmlns='http://www.w3.org/2000/svg' width='{width}' height='{height}' role='img' aria-label='PPTBench outcome heatmap'>",
        "<style>text{font:12px system-ui;fill:#18212b}.cell{fill:#fff;font-weight:600}</style>",
    ]
    for column, lane in enumerate(lanes):
        x = left + column * cell + 70
        body.append(f"<text x='{x}' y='34' text-anchor='middle'>{html.escape(lane)}</text>")
    for row_index, adapter in enumerate(adapters):
        y = 48 + row_index * 54
        body.append(f"<text x='8' y='{y + 28}'>{html.escape(adapter)}</text>")
        for column, lane in enumerate(lanes):
            outcome = states.get((adapter, lane), "unsupported")
            x = left + column * cell
            body.append(
                f"<rect x='{x}' y='{y}' width='150' height='40' rx='4' fill='{colors[outcome]}'/>"
                f"<text class='cell' x='{x + 75}' y='{y + 25}' text-anchor='middle'>{html.escape(outcome)}</text>"
            )
    body.append("</svg>")
    return "".join(body)
