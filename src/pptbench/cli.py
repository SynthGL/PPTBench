"""PPTBench command line interface."""

from __future__ import annotations

import argparse
from pathlib import Path

from .report import generate
from .runner import benchmark


def main() -> int:
    parser = argparse.ArgumentParser(prog="pptbench")
    commands = parser.add_subparsers(dest="command", required=True)
    run = commands.add_parser("benchmark", help="run controlled editing evidence lanes")
    run.add_argument("--output", type=Path, required=True, help="new or empty run root")
    run.add_argument(
        "--lane",
        action="append",
        dest="lanes",
        choices=["feature-matrix", "template-mutation", "chart-data"],
    )
    run.add_argument("--adapter", action="append", dest="adapters")
    run.add_argument("--fixture", action="append", dest="fixtures", choices=["mixed-60", "charts"])
    run.add_argument("--iterations", type=int, default=1)
    run.add_argument("--warmups", type=int, default=0)
    run.add_argument("--timeout", type=float, default=60, help="per-child seconds")
    run.add_argument("--overall-timeout", type=float, default=600, help="whole-run seconds")
    run.add_argument(
        "--reuse-fixtures",
        type=Path,
        metavar="DIR",
        help="validated fixture directory or prior run root; output remains a new run",
    )
    run.add_argument("--wolfppt-artifact", type=Path, help="observed WolfPPT wheel artifact")
    run.add_argument("--wolfppt-source-revision", help="source revision for the supplied wheel")
    run.add_argument(
        "--wolfppt-build-mode",
        choices=["provided-release-wheel", "source-build-wheel"],
        help="observed build provenance; omitted means unknown",
    )
    run.add_argument(
        "--render",
        action="store_true",
        help="invoke configured local LibreOffice evidence lane",
    )
    run.add_argument(
        "--openxml-validate",
        action="store_true",
        help="invoke PPTBENCH_OPENXML_VALIDATOR when configured",
    )
    report = commands.add_parser("report", help="render an HTML dashboard and SVG heatmap")
    report.add_argument("run_json", type=Path)
    report.add_argument("--output", type=Path)
    args = parser.parse_args()
    if args.command == "benchmark":
        path = benchmark(
            args.output,
            lanes=args.lanes,
            adapters=args.adapters,
            fixtures=args.fixtures,
            iterations=args.iterations,
            warmups=args.warmups,
            timeout_seconds=args.timeout,
            overall_timeout_seconds=args.overall_timeout,
            reuse_fixtures=args.reuse_fixtures,
            wolfppt_artifact=args.wolfppt_artifact,
            wolfppt_source_revision=args.wolfppt_source_revision,
            wolfppt_build_mode=args.wolfppt_build_mode,
            render=args.render,
            openxml_validate=args.openxml_validate,
        )
        print(path)
        return 0
    dashboard, heatmap = generate(args.run_json, args.output)
    print(dashboard)
    print(heatmap)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
