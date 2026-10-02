# PPTBench methodology

Evidence design dated 2026-09-08.

## Execution evidence

PPTBench starts a fresh monitor process (`pptbench.monitor`) for every warmup and measured adapter invocation. The monitor starts one adapter child in its own process group, waits with the configured per-child timeout, and records that monitor's `RUSAGE_CHILDREN.ru_maxrss`. This is genuine per-invocation accounting: the monitor has no earlier benchmark children whose cumulative maximum could contaminate the receipt. Values are normalized to bytes (macOS reports bytes; Linux reports KiB).

Every invocation is retained, including warmups, failures, unavailable adapters, unsupported adapters, and timeouts. Warmups are clearly labelled and never presented as eliminating interpreter startup: every attempt is a cold subprocess scenario. The whole run has a separate explicit budget in addition to the per-child timeout. A timeout kills the child process group, retains elapsed time and raw stdout/stderr artifact files, and never permits a partial or old output to be scored.

Runs use a new or empty output root. Candidate paths and monitor artifact paths are unique per attempt, preventing overwrites and stale-candidate reuse. Receipts persist run-relative artifact names, sanitized argv, byte counts, and bounded sanitized diagnostics; they do not persist private absolute interpreter paths, scratch paths, or raw tracebacks.

## Scoring evidence

Scoring has independent dimensions, each recorded per check:

1. Exact targeted edits: the precise new values at the precise declared locations (table cells, a bullet paragraph, chart series and points).
2. Field-specific Open XML semantics: XML signature comparison where equivalent ZIP/XML serialization is allowed. Parts are identified by the relationship types that reach them and their content (media and opaque parts by content hash), never by part name, relationship id, ZIP order, XML prefix, attribute order, explicitly written schema defaults, or save-time metadata.
3. Unrelated-change preservation: whole-package semantic equality against the expected edited result, plus notes, media, relationship-graph, opaque-part, chart-topology, and embedded-workbook preservation outside the edited targets.
4. Byte-only observations for selected untouched parts, recorded unscored.
5. Optional local rendering and Open XML validator evidence, only when the tools are present and requested.

Every check explicitly declares `category` (`semantic`, `feature`, `preservation`, or `byte-only`) and whether it is score-affecting. Byte-only observations are retained but do not masquerade as semantic success, and a byte difference alone is never labeled corruption. A missing optional tool is visible as `unavailable`, never green. Scoring verifies the input is the exact frozen fixture before any comparison, and no selected check count is called broad Office fidelity.

Observed adapter damage, including from the baseline `python-pptx` adapter, is legitimate benchmark output: a failed semantic or preservation check is evidence about the adapter under test, not a scoring malfunction, and it does not imply that PPTBench's own tests should fail.

## Fixture and provenance evidence

Fixtures are immutable packaged bytes, generated ahead of time by `pptbench.fixtures` under the pinned `python-pptx==1.0.2` dependency and normalized deterministically: sorted ZIP members, fixed entry timestamps and attributes, pinned `dcterms:created`/`dcterms:modified`, applied recursively into embedded packages so nested XLSX workbook timestamps cannot vary between regenerations. The frozen manifest records the semantic recipe digest and per-fixture SHA-256; runs verify and copy those exact bytes. Validated fixture reuse (`--reuse-fixtures`) accepts only exact named regular artifacts whose hashes match the frozen manifest, and the reuse remains a new, independent run.

Adapter identity hashes installed package modules and native binaries; external identity includes executable hash, version when available, and command-template configuration hash. WolfPPT build mode, wheel artifact hash, and source revision are explicit receipts; a release build is recorded only when its provided artifact and provenance were observed, otherwise build mode is `unknown`.

`pptbench report` verifies fixture, input, and output hashes during every generation, for every retained candidate including failed ones. Its manifest records the raw report hash. If a later generation sees a changed raw report, missing artifact, unsafe relative path, or hash mismatch, the dashboard and manifest report tampering rather than silently presenting a clean result.

## Evidence paths

Raw evidence is run-relative and complete per attempt: `run.json` at the root; verified `fixtures/` copies; `raw/<adapter>/<lane>/<phase>-<N>.pptx` candidates, each with a `.adapter/` sidecar holding `receipt.json`, `stdout.bin`, and `stderr.bin`; optional `render/<adapter>/<lane>/` PDFs; and `report/` dashboards with `report-manifest.json`. Reports and comparisons must link these run-relative paths and the SHA-256 values recorded in `run.json` from actual runs; the retained 2026-09-08 run is linked from the README, and observed outputs beyond a retained run root are not pre-stated in this documentation.

## Limitations

- Fixtures are project-authored synthetic decks, not Office-authored or enterprise ground truth; results generalize only to the measured scenarios, and the retained 2026-09-08 run is a small synthetic snapshot, not evidence of engine-wide rendering quality or speed superiority for any adapter.
- Byte-only observations are information, not verdicts; semantic equality with different bytes is success.
- Optional render and validator evidence are smoke signals: they do not prove visual equivalence in PowerPoint or LibreOffice, and PPTBench performs no Open XML schema validation itself when no validator is configured.
- Timings are cold-subprocess wall times under the recorded environment identity; they include interpreter and process startup, were taken with OS cache state uncontrolled, are not warm-process throughput numbers, and support no broad performance claim about any library.
