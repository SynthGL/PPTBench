# PPTBench

PPTBench is a standalone local benchmark for evidence about editing an existing `.pptx`. It keeps invocation and scoring separate, records semantic and preservation results for every attempted run, and preserves run-relative raw artifacts. It does not claim that byte equality, a selected check count, or a passing synthetic task proves broad Office or rendering fidelity.

## Quick start

Requires Python 3.11+ and `uv`.

```sh
uv sync --extra dev
uv run pytest
uv run ruff check .
uv run mypy src
uv build
python3 -m venv /tmp/pptbench-wheel
/tmp/pptbench-wheel/bin/pip install dist/pptbench-*.whl
/tmp/pptbench-wheel/bin/pptbench benchmark --output /tmp/pptbench-run --adapter python-pptx --iterations 1 --timeout 60 --overall-timeout 600
/tmp/pptbench-wheel/bin/pptbench report /tmp/pptbench-run/run.json
```

All example output locations are fresh directories under `/tmp`. `--output` must name a new or empty non-symlink directory; PPTBench never overwrites an existing run, so use a new directory per run rather than deleting old ones.

## Fixtures

The wheel ships two project-authored synthetic fixtures as immutable package resources (`src/pptbench/data/fixtures/`):

- `mixed-60`: a deterministic 60-slide template. Each slide has a named bold title; slide 5 has a named 4x4 table (`PPTBenchTable`), slide 7 a named three-paragraph bullet body (`PPTBenchBullets`), every tenth slide an additional table, other selected slides body text boxes, slide 2 a small PNG picture, slide 1 a speaker note, and the package an opaque unknown part (`ppt/unknown/pptbench.xml`).
- `charts`: three slides with one chart each, a clustered-column category chart (`Revenue`), an XY scatter-with-lines chart (`Trend`), and a bubble chart (`Pipeline`), each linked to an embedded XLSX workbook.

Every run verifies each fixture's SHA-256 against the frozen manifest (`src/pptbench/data/manifest.json`, corpus kind `project-authored synthetic`, `not_ground_truth: true`) and copies the exact packaged bytes under the run root. Runs never regenerate fixtures at benchmark time and never import another project's corpus. These are synthetic benchmark inputs, not Office-authored or enterprise ground truth.

Regeneration is byte-reproducible by construction: the generator dependency is pinned (`python-pptx==1.0.2`), ZIP members are written in sorted order with fixed entry timestamps and attributes, `dcterms:created`/`dcterms:modified` core-property timestamps are pinned, and the same normalization is applied recursively to embedded packages, so wall-clock timestamps inside the nested chart workbooks cannot leak into frozen bytes.

## Lanes and scoring

Three lanes run by default; `--lane` and `--fixture` select subsets. `mixed-60` backs `feature-matrix` and `template-mutation`; `charts` backs `chart-data`.

### `feature-matrix`

A read/write round trip over `mixed-60` with 19 declared feature checks:

`slide-count`, `slide-order-and-identifiers`, `slide-relationship-graph`, `shape-tree-semantics`, `text-content-and-order`, `text-run-formatting`, `table-cell-grid-and-text`, `table-cell-formatting`, `bullet-paragraph-properties`, `layout-geometry`, `slide-layout-semantics`, `notes-content`, `notes-relationships`, `media-part-presence`, `media-exact-bytes`, `theme-semantics`, `master-semantics`, `relationship-semantics`, `opaque-package-part`.

Feature checks compare XML semantic signatures, so equivalent re-serialization passes while content or structure changes fail. A declared check whose feature is absent from the fixture reports `unscored` rather than pass or fail. A `package-readable` precondition and one byte-only observation (below) are recorded alongside the 19.

### `template-mutation`

Three exact targeted edits in `mixed-60`: table cell (1,1) becomes `UPDATED-TABLE-A` and cell (2,2) becomes `UPDATED-TABLE-B` in `PPTBenchTable` (slide 5), and paragraph 1 of `PPTBenchBullets` (slide 7) becomes `UPDATED-BULLET`.

- Exact targeted-edit checks (`table-cell-1-1-updated`, `table-cell-2-2-updated`, `bullet-paragraph-1-updated`) compare the exact text at the exact locations.
- `only-declared-text-nodes-changed` compares whole-package semantic signatures of the expected result (input plus exactly those edits) against the adapter output, so wrong values, edits on the wrong slide or shape, and unrelated semantic collateral (rewritten text, dropped or injected shapes, changed notes) fail, while equivalent re-serialization passes.
- Preservation checks hold notes semantics, media bytes, relationship graphs, and opaque parts unchanged.
- A byte-only observation compares raw bytes of untouched parts, excluding the two edited slides.

### `chart-data`

Replaces the data of all three charts in `charts`: category names and values, XY point pairs, and bubble x/y/size triples.

- Per-chart exact checks (`category-`, `xy-`, and `bubble-chart-exact-series-and-points`) verify chart kind, series identity, and the full target point mapping; a generic "some chart XML changed" signal does not pass.
- `chart-topology-and-relations` and `chart-slide-structure` preserve chart/workbook wiring and slide structure.
- `related-workbook-values` requires every chart formula to resolve to the same embedded workbook cells holding exactly the target values; `related-workbook-structure-and-formatting` preserves workbook structure and formatting outside the edited cells.
- `chart-formatting-preserved`, the shared preservation checks, and a byte-only observation, each excluding `ppt/charts/` and `ppt/embeddings/`.

### Check outcomes

Every check declares `category` (`semantic`, `feature`, `preservation`, or `byte-only`) and `scored`. Semantic, feature, and preservation failures are score-affecting. Byte-only observations (`raw-untouched-part-equality`, `scored: false`) are recorded for information: two packages that differ byte-for-byte but agree semantically pass, byte equality is never required, and a byte difference alone is never called corruption. Scoring first re-verifies the input is the exact frozen artifact (`fixture-source-frozen`) and never scores from a stale or partial candidate.

### Baseline damage is legitimate output

`python-pptx` is the baseline adapter, not an assumed-correct oracle. If it, or any adapter, rewrites unrelated text, drops notes or media, or damages an embedded workbook, the affected checks fail and the run records that failure. That is the benchmark producing evidence about the adapter. It is not a PPTBench defect, and it does not mean PPTBench's own test suite should fail: the suite pins scorer and runner behavior against constructed outputs (exact edits, wrong-slide and notes mutations, malformed charts, tampered fixtures), never the premise that a particular adapter passes every lane.

## Adapters

Actually implemented integrations:

- `python-pptx` (default): invoked through its public API in an isolated child process.
- `wolfppt-wheel` (default): present only when an installed `wolfppt` distribution is importable through its public `wolfppt.Presentation` API; it never resolves a checkout path. When it is not installed, its rows record `unavailable` with reason and recovery text.

Generic bridge, not an SDK integration:

- `external-command`: requires `PPTBENCH_EXTERNAL_COMMAND`, a shell-tokenized argv template containing exactly `{input}`, `{output}`, and `{lane}`; the template is not run through a shell. Its identity is the executable file hash, the executable version when available, and a configuration hash of the template. PPTBench knows nothing about the tool behind the template: this adapter confers no SDK-specific identity, integration, or compatibility claim. It exists for wrappers that genuinely edit an existing deck.

Generation-only, unscored:

- PptxGenJS generates new decks and does not open and edit an existing `.pptx` in the sense measured here. It is deliberately not registered as an adapter and never appears in scored results.

The run also records `libreoffice-render` tool availability for the optional render evidence below.

## WolfPPT artifact provenance

A supplied WolfPPT artifact can be recorded without guessing its build origin:

```sh
pptbench benchmark --output /tmp/pptbench-wolfppt --adapter wolfppt-wheel \
  --wolfppt-artifact /path/to/wolfppt.whl \
  --wolfppt-source-revision 5929d925536c3aa2e8ce49399c5d220bbf57df2a \
  --wolfppt-build-mode provided-release-wheel
```

The receipt records the observed artifact filename and SHA-256, the optional source revision (7-64 hexadecimal characters), and the stated build mode (`provided-release-wheel` or `source-build-wheel`). Omit what cannot be verified: build mode is then `unknown`, and PPTBench never infers a release build from a wheel filename.

## Observed run evidence (2026-09-08)

A local wheel-comparison run is retained in this repository:

- Report dashboard: [evidence/2026-09-08/wheel-comparison/report/index.html](evidence/2026-09-08/wheel-comparison/report/index.html)
- Outcome heatmap: [evidence/2026-09-08/wheel-comparison/report/heatmap.svg](evidence/2026-09-08/wheel-comparison/report/heatmap.svg)
- Raw run report: [evidence/2026-09-08/wheel-comparison/run.json](evidence/2026-09-08/wheel-comparison/run.json)
- Post-run verification record: [evidence/2026-09-08/verification.json](evidence/2026-09-08/verification.json)

Observed lane outcomes:

| Adapter | `feature-matrix` | `template-mutation` | `chart-data` |
|---|---|---|---|
| `python-pptx` 1.0.2 | failure | failure | success |
| `wolfppt-wheel` (installed WolfPPT 0.1.0) | success | success | failure |

`python-pptx` failed `feature-matrix` at `opaque-package-part`: the package's opaque unknown part is lost on round trip. It failed `template-mutation` at the collateral checks `only-declared-text-nodes-changed` and `opaque-parts-preserved`. WolfPPT carried the requested values in the chart cache, so the per-chart exact checks passed, but `related-workbook-values`, `related-workbook-structure-and-formatting`, and `chart-formatting-preserved` failed. No engine or adapter fixes are part of this evidence record.

The WolfPPT candidate was the installed 0.1.0 wheel built from source revision `5929d925536c3aa2e8ce49399c5d220bbf57df2a` (receipt `provided-release-wheel`), imported through the installed package rather than a source checkout. Each adapter and lane ran one warmup plus two measured iterations with a 20 second child timeout and a 120 second overall timeout on Python 3.14, macOS arm64. Timings include process startup and were taken with OS cache state uncontrolled. Every successful scored candidate passed the optional LibreOffice PDF render smoke; the Open XML validator was unavailable (`PPTBENCH_OPENXML_VALIDATOR` not configured). This small synthetic snapshot is not evidence of engine-wide rendering quality or speed superiority for either library, and no speed ranking is claimed.

## Identity and privacy

Adapter identity for installed packages hashes the installed package's module files and native binaries (`modules:<count>:<digest>`), not only `__init__.py`. Environment identity records Python version, implementation, platform, executable basename, and PID, never a private interpreter path. Command metadata reduces absolute paths to basenames and replaces private path prefixes (`~`, `/Users/...`, `/home/...`, `/var/folders/...`) with `<private-path>`; receipts persist run-relative artifact names, sanitized argv, byte counts, and bounded sanitized diagnostics, not raw tracebacks.

## Timing and execution

Every adapter attempt, including every warmup and measured iteration, runs as a new subprocess. Warmups are retained in the report but **do not remove interpreter startup**; PPTBench does not report a fabricated warm-process speed. A fresh monitor process records the adapter child's own POSIX resource usage (`RUSAGE_CHILDREN` of a monitor that has no earlier children), normalized to bytes (macOS reports bytes; Linux reports KiB). `--timeout` caps each child and `--overall-timeout` caps the full run; a timeout terminates the child process group, retains elapsed time and bounded artifact metadata, and is never scored from a stale candidate.

## Run layout and evidence paths

All receipt paths are relative to the run root:

| Path | Contents |
|---|---|
| `run.json` | Immutable raw report (schema version 2): configuration, environment and adapter identities, manifest, fixture inventory, every result with all retained samples and checks, optional evidence |
| `fixtures/mixed-60.pptx`, `fixtures/charts.pptx` | Hash-verified copies of the packaged fixture bytes |
| `raw/<adapter>/<lane>/warmup-<N>.pptx` | Warmup candidate outputs |
| `raw/<adapter>/<lane>/measured-<N>.pptx` | Measured candidate outputs |
| `raw/<adapter>/<lane>/<phase>-<N>.pptx.adapter/receipt.json` | Monitor receipt (outcome, elapsed, peak RSS, return code) |
| `raw/<adapter>/<lane>/<phase>-<N>.pptx.adapter/stdout.bin`, `stderr.bin` | Raw child output |
| `render/<adapter>/<lane>/<stem>.pdf` | Optional render evidence (only with `--render`) |
| `report/index.html`, `report/heatmap.svg`, `report/report-manifest.json` | Generated by `pptbench report` |

Cite observed runs by these run-relative paths together with the candidate SHA-256 values recorded in `run.json`. One retained local run is summarized under [Observed run evidence (2026-09-08)](#observed-run-evidence-2026-09-08); rates and comparisons are stated only from evidence linked from actual run roots.

## Reports and optional evidence

`pptbench report RUN_JSON [--output DIR]` creates `index.html`, `heatmap.svg`, and `report-manifest.json` under `<run root>/report` (or `DIR`). The dashboard is filterable by adapter and lane; each row embeds all retained warmup/measured samples with per-iteration reasons and checks. Links are report-relative. Regeneration verifies fixture, input, and output hashes for every retained candidate, including retained failed candidates, and detects a raw report changed since an earlier report manifest; missing, unsafe, or mismatched artifacts are reported as tampering rather than presented clean.

Optional rendering invokes local LibreOffice (`soffice` on PATH) only when `--render` is passed, converting each successful candidate to PDF. Optional Open XML validation invokes the configured `PPTBENCH_OPENXML_VALIDATOR` command only when `--openxml-validate` is passed. Missing tools are visible as `unavailable`, never green. These are smoke signals only: a successful render or validator exit does not prove visual equivalence in PowerPoint or LibreOffice, and PPTBench itself performs no Open XML schema validation when no validator is configured.

## Boundaries and maintenance

PPTBench is MIT licensed; see [LICENSE](LICENSE). `run.json` schema version 2 contains outcomes, wall time, independently attributable peak RSS, artifact hashes, sanitized command metadata, and raw stdout/stderr artifact paths. Observed runs are retained under `evidence/`; no numeric results are stated in this documentation outside a retained run root. The fixture generator is pinned to `python-pptx==1.0.2`; maintainers making an intentional corpus change must review it, run `uv run python -c 'from pptbench.fixtures import regenerate_frozen_data; regenerate_frozen_data()'`, and commit the resulting `src/pptbench/data/fixtures/*.pptx` plus manifest hashes together. See [PROVENANCE.md](PROVENANCE.md) for source and license review and the no-runtime-coupling rule, and [docs/METHODOLOGY.md](docs/METHODOLOGY.md) for the evidence design.
