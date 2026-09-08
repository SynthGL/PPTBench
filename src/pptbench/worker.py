"""Child entry point. Keeps adapter work outside the benchmark coordinator."""

from __future__ import annotations

import importlib
import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any, Protocol, cast


class PresentationFactory(Protocol):
    def __call__(self, path: str) -> PresentationDocument: ...


class PresentationDocument(Protocol):
    slides: Any

    def save(self, path: str) -> None: ...


class CategoryData(Protocol):
    categories: list[str]

    def add_series(self, name: str, values: tuple[int, ...]) -> object: ...


class XYSeries(Protocol):
    def add_data_point(self, x: int, y: int) -> object: ...


class BubbleSeries(Protocol):
    def add_data_point(self, x: int, y: int, size: int) -> object: ...


class XYData(Protocol):
    def add_series(self, name: str) -> XYSeries: ...


class BubbleData(Protocol):
    def add_series(self, name: str) -> BubbleSeries: ...


def _presentation(adapter: str) -> PresentationFactory:
    module_name = {"python-pptx": "pptx", "wolfppt-wheel": "wolfppt"}.get(adapter)
    if module_name is None:
        raise ValueError(f"worker does not implement adapter {adapter}")
    module = importlib.import_module(module_name)
    factory = getattr(module, "Presentation", None)
    if not callable(factory):
        raise TypeError(f"{module_name}.Presentation is unavailable")
    return cast(PresentationFactory, factory)


def edit(adapter: str, lane: str, source: Path, output: Path) -> None:
    presentation = _presentation(adapter)(str(source))
    if lane == "template-mutation":
        _mutate_template(presentation)
    elif lane == "chart-data":
        _replace_chart_data(presentation)
    elif lane != "feature-matrix":
        raise ValueError(f"unsupported lane: {lane}")
    presentation.save(str(output))


def _mutate_template(presentation: PresentationDocument) -> None:
    table_shape = _shape_named(presentation.slides[4], "PPTBenchTable")
    table_shape.table.cell(1, 1).text = "UPDATED-TABLE-A"
    table_shape.table.cell(2, 2).text = "UPDATED-TABLE-B"
    bullets = _shape_named(presentation.slides[6], "PPTBenchBullets")
    bullets.text_frame.paragraphs[1].text = "UPDATED-BULLET"


def _replace_chart_data(presentation: PresentationDocument) -> None:
    category_factory, xy_factory, bubble_factory = _chart_factories()
    category = category_factory()
    category.categories = ["East", "42", "84"]
    category.add_series("Revenue", (42, 84, 126))
    presentation.slides[0].shapes[0].chart.replace_data(category)

    xy = xy_factory()
    xy_series = xy.add_series("Trend")
    for x, y in ((5, 13), (21, 34), (55, 89)):
        xy_series.add_data_point(x, y)
    presentation.slides[1].shapes[0].chart.replace_data(xy)

    bubble = bubble_factory()
    bubble_series = bubble.add_series("Pipeline")
    for x, y, size in ((5, 13, 21), (34, 55, 89), (8, 13, 21)):
        bubble_series.add_data_point(x, y, size)
    presentation.slides[2].shapes[0].chart.replace_data(bubble)


def _chart_factories() -> tuple[
    Callable[[], CategoryData], Callable[[], XYData], Callable[[], BubbleData]
]:
    module = importlib.import_module("pptx.chart.data")
    return (
        cast(Callable[[], CategoryData], module.CategoryChartData),
        cast(Callable[[], XYData], module.XyChartData),
        cast(Callable[[], BubbleData], module.BubbleChartData),
    )


def _shape_named(slide: Any, name: str) -> Any:
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    raise LookupError(f"required shape is missing: {name}")


def main() -> int:
    if len(sys.argv) != 5:
        print(
            "usage: python -m pptbench.worker ADAPTER LANE INPUT OUTPUT",
            file=sys.stderr,
        )
        return 2
    try:
        edit(sys.argv[1], sys.argv[2], Path(sys.argv[3]), Path(sys.argv[4]))
    except (
        AttributeError,
        ImportError,
        KeyError,
        LookupError,
        OSError,
        TypeError,
        ValueError,
    ) as exc:
        print(f"{type(exc).__name__}: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
