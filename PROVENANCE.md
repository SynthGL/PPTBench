# Provenance

PPTBench is standalone. It does not import a sibling checkout, embed WolfPPT's runtime, or consult any external scoring service.

## Reviewed source material

Useful patterns were independently reimplemented from the MIT-licensed WolfPPT repository at commit `5929d925536c3aa2e8ce49399c5d220bbf57df2a`:

| Source path | SHA-256 at review | Use in PPTBench |
|---|---|---|
| `LICENSE` | `24accf5eaddfee8d5ee2b5ff2dd212d62cd29a62d194fb0c5bc681b2c7a79730` | License terms and attribution review |
| `src/wolfppt/adapters/python_pptx.py` | `36f805950220d2a9a2184f1fac5b6ab3b18b3f4680c0f88121328a46562ca63f` | Public python-pptx adapter boundary |
| `src/wolfppt/benchmark.py` | `eb12c578aa62d18a6552ec4a4526fabd017b50314478956a2e79158cc3deb323` | Repeated-run and unavailable-tool reporting pattern |
| `src/wolfppt/benchmark_models.py` | `e712fa8cb30f80254e6e414d94c601302b5a4059ef676cd73b9494f94b32fe73` | Per-sample timing statistics pattern |
| `src/wolfppt/corpus.py` | `cc12139007894d44c195722ba520f7370d0c22c9eb8a0d21d64d68f26b244d8a` | Synthetic fixture construction approach |

WolfPPT is MIT licensed. PPTBench contains no copied WolfPPT source files or source fixtures.

## Fixture provenance

PPTBench's decks are project-authored synthetic material, generated only by `pptbench.fixtures` under the pinned generator dependency (`python-pptx==1.0.2`). Generation normalizes every package deterministically (sorted ZIP members, fixed entry timestamps and attributes, pinned `dcterms:created`/`dcterms:modified`) and applies the same normalization recursively to embedded packages, including the XLSX workbooks embedded in the chart fixture, so nested timestamps cannot perturb frozen bytes. The reviewed outputs are immutable package resources in `src/pptbench/data/fixtures/`; their semantic recipe digest and exact package hashes are frozen in `src/pptbench/data/manifest.json`; benchmark runs only verify and copy those bytes. They are not Office-authored or enterprise ground truth.

## Adapter provenance

- `python-pptx` is invoked through its public API.
- `wolfppt-wheel` is optional: it imports only an installed `wolfppt` distribution and calls its documented public `Presentation` API. It never resolves a checkout path. Supplied artifacts are recorded through the explicit `--wolfppt-artifact`, `--wolfppt-source-revision`, and `--wolfppt-build-mode` receipts; an unproven build is recorded as `unknown`.
- `external-command` executes only an explicitly supplied argv template and reports `unavailable` when omitted. It is a generic bridge: PPTBench makes no SDK-specific identity or compatibility claim about the tool behind it.
- PptxGenJS is generation-only: it produces new decks rather than editing an existing `.pptx`, so it is deliberately not registered as an existing-deck editing candidate and is never scored.

## Ownership and updates

PPTBench maintainers own its corpus, scorer, report, and adapter contracts. Updating a frozen task or expected semantic value requires changing its manifest digest and recording the reason in a reviewed change. An intentional fixture update must run `uv run python -c 'from pptbench.fixtures import regenerate_frozen_data; regenerate_frozen_data()'` with the pinned generator dependency, then commit the generated `src/pptbench/data/fixtures/*.pptx` and manifest together. Do not silently replace fixture bytes or score data.
