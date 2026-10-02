"""Exact Open XML semantic scorers with separate byte-preservation observations.

Scored checks observe only what a user of the file can observe. Parts are resolved
through the name-free :class:`~pptbench.semantics.PackageModel`: slides by
presentation order, other parts by the relationship types that reach them, media by
content hash. Part names, relationship ids, ZIP order, XML prefixes, attribute order,
explicitly written schema defaults, and save-time metadata never decide a scored
outcome; byte equality is recorded only as an unscored observation.
"""

from __future__ import annotations

import copy
import re
import zipfile
from collections import Counter
from collections.abc import Callable
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, TypeVar
from xml.etree import ElementTree as ET

from .fixtures import FROZEN_RECIPE, verify_fixture
from .semantics import ROOT, PackageModel, short_type
from .util import package_bytes, package_parts

Check = dict[str, Any]
T = TypeVar("T")

_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_CELL_RANGE = re.compile(r"^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$")
_CHART_OWNED = frozenset({"chart", "package"})


def score(lane: str, source: Path, output: Path) -> list[Check]:
    fixture_ids = {
        "feature-matrix": "mixed-60",
        "template-mutation": "mixed-60",
        "chart-data": "charts",
    }
    if lane not in fixture_ids:
        return [_check("lane-known", False, f"unknown lane {lane}")]
    try:
        verify_fixture(fixture_ids[lane], source)
    except (OSError, ValueError) as exc:
        return [_check("fixture-source-frozen", False, str(exc))]
    try:
        before = package_parts(source)
        after = package_parts(output)
    except (OSError, ValueError, ET.ParseError, zipfile.BadZipFile) as exc:
        return [_check("package-readable", False, str(exc))]
    if lane == "feature-matrix":
        try:
            return _feature_matrix(before, after)
        except (KeyError, ValueError, ET.ParseError) as exc:
            return [_check("package-readable", False, str(exc))]
    if lane == "template-mutation":
        try:
            return _template_mutation(before, after)
        except (KeyError, ValueError, ET.ParseError) as exc:
            return [_check("template-package-semantics", False, str(exc))]
    if lane == "chart-data":
        return _chart_data(before, after)
    raise AssertionError("known lane dispatch is incomplete")


def _feature_matrix(before: dict[str, bytes], after: dict[str, bytes]) -> list[Check]:
    source = PackageModel(before)
    result = PackageModel(after)
    checks = [_check("package-readable", True)]
    contracts: list[tuple[str, Callable[[PackageModel], object]]] = [
        ("slide-count", lambda model: len(_slides(model))),
        ("slide-order-and-identifiers", _slide_identifiers),
        (
            "slide-relationship-graph",
            lambda model: {number: model.outgoing(key) for number, key in _numbered_slides(model)},
        ),
        ("shape-tree-semantics", _slide_shape_trees),
        ("text-content-and-order", _slide_text),
        ("text-run-formatting", _text_formatting),
        ("table-cell-grid-and-text", _table_grid),
        ("table-cell-formatting", _table_formatting),
        ("bullet-paragraph-properties", _paragraph_properties),
        ("layout-geometry", _shape_geometry),
        ("slide-layout-semantics", lambda model: _nodes(model, "slideLayout")),
        (
            "notes-content",
            lambda model: {key: model.signature(key) for key in model.keys_of_type("notesSlide")},
        ),
        (
            "notes-relationships",
            lambda model: {key: model.outgoing(key) for key in model.keys_of_type("notesSlide")},
        ),
        (
            "media-part-presence",
            lambda model: sorted({(owner, rel_type) for owner, rel_type, _ in model.media()}),
        ),
        ("media-exact-bytes", lambda model: model.media()),
        ("theme-semantics", lambda model: _nodes(model, "theme")),
        ("master-semantics", lambda model: _nodes(model, "slideMaster")),
        ("relationship-semantics", _relationship_graph),
        (
            "opaque-package-part",
            lambda model: _retained(source.opaque_parts(), model.opaque_parts()),
        ),
    ]
    for name, extractor in contracts:
        expected = extractor(source)
        if _is_absent_feature(expected):
            checks.append(_unscored(name, "fixture does not contain this declared feature"))
        else:
            checks.append(_equal(name, expected, extractor(result), category="feature"))
    checks.append(_byte_observation("raw-untouched-part-equality", before, after))
    return checks


def _template_mutation(before: dict[str, bytes], after: dict[str, bytes]) -> list[Check]:
    source = PackageModel(before)
    result = PackageModel(after)
    expected = PackageModel(before)
    source_slides = _slides(source)
    table_part = _slide_name(source, source_slides, 5)
    bullet_part = _slide_name(source, source_slides, 7)
    table = copy.deepcopy(source.root(table_part))
    bullet = copy.deepcopy(source.root(bullet_part))
    for row, column, value in [(1, 1, "UPDATED-TABLE-A"), (2, 2, "UPDATED-TABLE-B")]:
        _set_single_text(_table_cell(table, "PPTBenchTable", row, column), value)
    _set_single_text(_shape_paragraph(bullet, "PPTBenchBullets", 1), "UPDATED-BULLET")
    expected.replace_root(table_part, table)
    expected.replace_root(bullet_part, bullet)

    result_slides = _slides(result)
    table_root = result.root(_slide_name(result, result_slides, 5))
    bullet_root = result.root(_slide_name(result, result_slides, 7))
    locations = [
        (
            "table-cell-1-1-updated",
            "UPDATED-TABLE-A",
            _single_text(_table_cell(table_root, "PPTBenchTable", 1, 1)),
        ),
        (
            "table-cell-2-2-updated",
            "UPDATED-TABLE-B",
            _single_text(_table_cell(table_root, "PPTBenchTable", 2, 2)),
        ),
        (
            "bullet-paragraph-1-updated",
            "UPDATED-BULLET",
            _single_text(_shape_paragraph(bullet_root, "PPTBenchBullets", 1)),
        ),
    ]
    checks = [
        _check(name, actual == wanted, "wrong target value") for name, wanted, actual in locations
    ]
    checks.append(
        _equal(
            "only-declared-text-nodes-changed",
            _package_semantics(expected),
            _package_semantics(result),
            category="semantic",
        )
    )
    checks.extend(_preservation_checks(source, result))
    checks.append(
        _byte_observation(
            "raw-untouched-part-equality",
            before,
            after,
            exclude_names={table_part, bullet_part},
        )
    )
    return checks


def _chart_data(before: dict[str, bytes], after: dict[str, bytes]) -> list[Check]:
    try:
        source = PackageModel(before)
        result = PackageModel(after)
        source_links = _chart_links(source)
        output_links = _chart_links(result)
        expected = _chart_contract()
        checks = [
            _equal(
                "chart-topology-and-relations",
                source_links,
                output_links,
                category="preservation",
            ),
            _equal(
                "chart-slide-structure",
                _slide_shape_trees(source),
                _slide_shape_trees(result),
            ),
        ]
        expected_cells: dict[tuple[int, str, str], str] = {}
        for slide_number, expected_chart in expected.items():
            chart_key, _ = output_links[slide_number]
            root = result.root(result.name_of[chart_key])
            chart_name = str(expected_chart["name"])
            refs = _validate_chart(root, expected_chart)
            checks.append(
                _check(
                    f"{chart_name}-chart-exact-series-and-points",
                    refs is not None,
                    "chart type, series, or point mapping differs",
                )
            )
            if refs is not None:
                for formula, values in refs:
                    for sheet, coordinate, value in _formula_cells(formula, values):
                        key = (slide_number, sheet, coordinate)
                        if key in expected_cells and expected_cells[key] != value:
                            raise ValueError("chart series assign conflicting workbook values")
                        expected_cells[key] = value
        checks.append(
            _workbook_values_check(source, result, source_links, output_links, expected_cells)
        )
        checks.append(_workbook_structure_check(source, result, source_links, output_links))
        checks.append(
            _equal(
                "chart-formatting-preserved",
                _chart_formats(source, source_links),
                _chart_formats(result, output_links),
                category="preservation",
            )
        )
        checks.extend(_preservation_checks(source, result, exclude_kinds=_CHART_OWNED))
        checks.append(
            _byte_observation(
                "raw-untouched-part-equality",
                before,
                after,
                exclude_names=_owned_names(source, _CHART_OWNED)
                | _owned_names(result, _CHART_OWNED),
            )
        )
        return checks
    except (KeyError, TypeError, ValueError, ET.ParseError, InvalidOperation) as exc:
        return [_check("chart-package-semantics", False, str(exc))]


def _chart_contract() -> dict[int, dict[str, object]]:
    charts = FROZEN_RECIPE["charts"]
    if not isinstance(charts, dict):  # defensive guard against edited recipe data
        raise TypeError("invalid chart fixture recipe")
    category = charts["category"]
    xy = charts["xy"]
    bubble = charts["bubble"]
    if not isinstance(category, dict) or not isinstance(xy, dict) or not isinstance(bubble, dict):
        raise TypeError("invalid chart fixture recipe")
    return {
        1: {
            "name": "category",
            "kind": "barChart",
            "series": category["series"],
            "target": category["target"],
        },
        2: {
            "name": "xy",
            "kind": "scatterChart",
            "series": xy["series"],
            "target": xy["target"],
        },
        3: {
            "name": "bubble",
            "kind": "bubbleChart",
            "series": bubble["series"],
            "target": bubble["target"],
        },
    }


def _validate_chart(
    root: ET.Element, expected: dict[str, object]
) -> list[tuple[str, list[str]]] | None:
    kind = expected["kind"]
    series_name = expected["series"]
    target = expected["target"]
    if not isinstance(kind, str) or not isinstance(series_name, str):
        raise TypeError("invalid chart contract")
    charts = [element for element in root.iter() if _local(element.tag) == kind]
    if len(charts) != 1:
        return None
    series = [element for element in charts[0] if _local(element.tag) == "ser"]
    if len(series) != 1 or _series_name(series[0]) != series_name:
        return None
    if kind == "barChart":
        if not isinstance(target, dict):
            raise ValueError("invalid category chart contract")
        categories = _ref_values(series[0], "cat")
        values = _ref_values(series[0], "val")
        expected_categories = target.get("categories")
        expected_values = target.get("values")
        if not isinstance(expected_categories, list) or not isinstance(expected_values, list):
            raise ValueError("invalid category chart values")
        if categories is None or values is None:
            return None
        if categories[1] != [
            str(value) for value in expected_categories
        ] or not _numeric_sequence_equal(values[1], expected_values):
            return None
        return [(categories[0], categories[1]), (values[0], values[1])]
    if not isinstance(target, list):
        raise TypeError("invalid XY chart contract")
    keys = ("xVal", "yVal") if kind == "scatterChart" else ("xVal", "yVal", "bubbleSize")
    columns = list(zip(*target, strict=True))
    refs: list[tuple[str, list[str]]] = []
    for key, column in zip(keys, columns, strict=True):
        actual = _ref_values(series[0], key)
        if actual is None or not _numeric_sequence_equal(actual[1], list(column)):
            return None
        refs.append(actual)
    return refs


def _workbook_values_check(
    source: PackageModel,
    result: PackageModel,
    source_links: dict[int, tuple[str, str]],
    output_links: dict[int, tuple[str, str]],
    expected_cells: dict[tuple[int, str, str], str],
) -> Check:
    if {number: link[1] for number, link in source_links.items()} != {
        number: link[1] for number, link in output_links.items()
    }:
        return _check("related-workbook-identity", False, "chart workbook relations changed")
    expected_grids = {
        number: _workbook_cells(source.data(workbook))
        for number, (_, workbook) in source_links.items()
    }
    actual_grids = {
        number: _workbook_cells(result.data(workbook))
        for number, (_, workbook) in output_links.items()
    }
    for (number, sheet, coordinate), expected_value in expected_cells.items():
        try:
            expected_grids[number][sheet][coordinate] = _canonical_value(expected_value)
        except KeyError:
            return _check(
                "related-workbook-values",
                False,
                "chart formula points outside the source workbook",
            )
    return _equal("related-workbook-values", expected_grids, actual_grids)


def _workbook_structure_check(
    source: PackageModel,
    result: PackageModel,
    source_links: dict[int, tuple[str, str]],
    output_links: dict[int, tuple[str, str]],
) -> Check:
    return _equal(
        "related-workbook-structure-and-formatting",
        {
            number: _workbook_structure(source.data(workbook))
            for number, (_, workbook) in source_links.items()
        },
        {
            number: _workbook_structure(result.data(workbook))
            for number, (_, workbook) in output_links.items()
        },
        category="preservation",
    )


def _preservation_checks(
    source: PackageModel,
    result: PackageModel,
    *,
    exclude_kinds: frozenset[str] = frozenset(),
) -> list[Check]:
    source_excluded = source.descendants(set(exclude_kinds))
    result_excluded = result.descendants(set(exclude_kinds))

    def notes(model: PackageModel, excluded: set[str]) -> dict[str, object]:
        return {
            key: model.node(key) for key in model.keys_of_type("notesSlide") if key not in excluded
        }

    return [
        _equal(
            "notes-preserved",
            notes(source, source_excluded),
            notes(result, result_excluded),
            category="preservation",
        ),
        _equal(
            "media-preserved",
            source.media(exclude=source_excluded),
            result.media(exclude=result_excluded),
            category="preservation",
        ),
        _equal(
            "relationships-preserved",
            _relationship_graph(source),
            _relationship_graph(result),
            category="preservation",
        ),
        _equal(
            "opaque-parts-preserved",
            source.opaque_parts(exclude=source_excluded),
            _retained(
                source.opaque_parts(exclude=source_excluded),
                result.opaque_parts(exclude=result_excluded),
            ),
            category="preservation",
        ),
    ]


def _check(
    name: str,
    passed: bool,
    detail: str | None = None,
    *,
    category: str = "semantic",
    scored: bool = True,
) -> Check:
    return {
        "name": name,
        "outcome": "success" if passed else "failure",
        "detail": None if passed else detail,
        "category": category,
        "scored": scored,
    }


def _unscored(name: str, detail: str) -> Check:
    return {
        "name": name,
        "outcome": "unscored",
        "detail": detail,
        "category": "feature",
        "scored": False,
    }


def _equal(name: str, before: T, after: T, *, category: str = "semantic") -> Check:
    return _check(name, before == after, "source and result differ", category=category)


def _byte_observation(
    name: str,
    before: dict[str, bytes],
    after: dict[str, bytes],
    *,
    exclude_names: set[str] | None = None,
) -> Check:
    omitted = exclude_names or set()
    before_untouched = {part: value for part, value in before.items() if part not in omitted}
    after_untouched = {part: value for part, value in after.items() if part not in omitted}
    return _check(
        name,
        before_untouched == after_untouched,
        "raw untouched part bytes differ",
        category="byte-only",
        scored=False,
    )


def _owned_names(model: PackageModel, kinds: frozenset[str]) -> set[str]:
    return {model.name_of[key] for key in model.descendants(set(kinds))}


# Slides, in presentation order -----------------------------------------------------


def _slides(model: PackageModel) -> list[str]:
    """Slide keys in the order the presentation's slide list shows them."""
    presentation = model.main_part()
    root = model.root(model.name_of[presentation])
    if root.tag != _P + "presentation":
        raise ValueError("main document is not a presentation")
    slides: list[str] = []
    for slide_id in root.iter(_P + "sldId"):
        rel_type, name = model.target(presentation, slide_id.attrib.get(_R + "id", ""))
        if name is None or short_type(rel_type) != "slide":
            raise ValueError("slide list entry does not resolve to a slide part")
        slides.append(model.key_of[name])
    return slides


def _numbered_slides(model: PackageModel) -> list[tuple[int, str]]:
    return list(enumerate(_slides(model), start=1))


def _slide_name(model: PackageModel, slides: list[str], number: int) -> str:
    if number > len(slides):
        raise ValueError(f"missing expected slide {number}")
    return model.name_of[slides[number - 1]]


def _slide_roots(model: PackageModel) -> list[tuple[int, str, ET.Element]]:
    return [
        (number, key, model.root(model.name_of[key])) for number, key in _numbered_slides(model)
    ]


def _slide_element_signatures(
    model: PackageModel, select: Callable[[ET.Element], list[ET.Element]]
) -> dict[int, list[object]]:
    return {
        number: [model.element_signature(key, element) for element in select(root)]
        for number, key, root in _slide_roots(model)
    }


def _slide_identifiers(model: PackageModel) -> list[tuple[int, str]]:
    presentation = model.root(model.name_of[model.main_part()])
    return [
        (number, slide_id.attrib.get("id", ""))
        for number, slide_id in enumerate(presentation.iter(_P + "sldId"), start=1)
    ]


def _slide_shape_trees(model: PackageModel) -> dict[int, object]:
    return {number: model.signature(key) for number, key in _numbered_slides(model)}


def _slide_text(model: PackageModel) -> dict[int, list[str]]:
    return {
        number: [element.text or "" for element in root.iter(_A + "t")]
        for number, _, root in _slide_roots(model)
    }


def _text_formatting(model: PackageModel) -> dict[int, list[object]]:
    return _slide_element_signatures(
        model,
        lambda root: [
            element for element in root.iter() if element.tag in {_A + "rPr", _A + "endParaRPr"}
        ],
    )


def _table_grid(model: PackageModel) -> dict[int, list[list[list[str]]]]:
    return {
        number: [
            [
                [_single_text(cell) or "" for cell in row.findall(_A + "tc")]
                for row in table.findall(_A + "tr")
            ]
            for table in root.iter(_A + "tbl")
        ]
        for number, _, root in _slide_roots(model)
    }


def _table_formatting(model: PackageModel) -> dict[int, list[object]]:
    return _slide_element_signatures(model, lambda root: list(root.iter(_A + "tbl")))


def _paragraph_properties(model: PackageModel) -> dict[int, list[object]]:
    return _slide_element_signatures(model, lambda root: list(root.iter(_A + "p")))


def _shape_geometry(model: PackageModel) -> dict[int, list[object]]:
    return _slide_element_signatures(
        model,
        lambda root: [
            element
            for element in root.iter()
            if _local(element.tag) in {"xfrm", "off", "ext", "chOff", "chExt"}
        ],
    )


# Whole-package views ---------------------------------------------------------------


def _nodes(model: PackageModel, kind: str) -> dict[str, object]:
    return {key: model.node(key) for key in model.keys_of_type(kind)}


def _relationship_graph(model: PackageModel) -> dict[str, list[tuple[str, str]]]:
    graph = {key: model.outgoing(key) for key in model.edges}
    return {key: edges for key, edges in sorted(graph.items()) if edges}


def _package_semantics(model: PackageModel) -> dict[str, object]:
    """Every reachable part's content, type, and relationships; opaque parts aside."""
    result: dict[str, object] = {ROOT: model.outgoing(ROOT)}
    for key in model.part_keys():
        if not model.is_save_time(key) and not model.is_opaque(key):
            result[key] = model.node(key)
    return result


def _is_absent_feature(value: object) -> bool:
    return value in ({}, [], (), 0, None)


def _retained(expected: list[T], actual: list[T]) -> list[T]:
    """Expected entries still present in ``actual``, counted as a multiset.

    Unreachable leftovers a writer adds are invisible to every reader, so only
    losing a source part counts; added reachable parts change the relationship graph.
    """
    available = Counter(repr(entry) for entry in actual)
    kept: list[T] = []
    for entry in expected:
        if available[repr(entry)]:
            available[repr(entry)] -= 1
            kept.append(entry)
    return kept


# Template targets --------------------------------------------------------------------


def _shape_named(root: ET.Element, name: str) -> ET.Element:
    for shape in root.iter():
        if _local(shape.tag) not in {"sp", "graphicFrame"}:
            continue
        properties = shape.find(".//" + _P + "cNvPr")
        if properties is not None and properties.attrib.get("name") == name:
            return shape
    raise ValueError(f"required shape is missing: {name}")


def _table_cell(root: ET.Element, shape_name: str, row: int, column: int) -> ET.Element:
    shape = _shape_named(root, shape_name)
    table = shape.find(".//" + _A + "tbl")
    if table is None:
        raise ValueError(f"shape {shape_name!r} is not a table")
    rows = table.findall(_A + "tr")
    if row >= len(rows):
        raise ValueError("target table row is missing")
    cells = rows[row].findall(_A + "tc")
    if column >= len(cells):
        raise ValueError("target table column is missing")
    return cells[column]


def _shape_paragraph(root: ET.Element, shape_name: str, index: int) -> ET.Element:
    shape = _shape_named(root, shape_name)
    paragraphs = shape.findall(".//" + _A + "p")
    if index >= len(paragraphs):
        raise ValueError("target paragraph is missing")
    return paragraphs[index]


def _single_text(element: ET.Element) -> str | None:
    texts = [node.text or "" for node in element.iter(_A + "t")]
    return texts[0] if len(texts) == 1 else None


def _set_single_text(element: ET.Element, value: str) -> None:
    texts = list(element.iter(_A + "t"))
    if len(texts) != 1:
        raise ValueError("target must contain exactly one text node")
    texts[0].text = value


# Charts and embedded workbooks -------------------------------------------------------


def _chart_links(model: PackageModel) -> dict[int, tuple[str, str]]:
    """Chart slides mapped to (chart, embedded workbook) as a consumer resolves them."""
    links: dict[int, tuple[str, str]] = {}
    for number, key in _numbered_slides(model):
        root = model.root(model.name_of[key])
        chart_ids = [
            element.attrib[_R + "id"]
            for element in root.iter(_C + "chart")
            if _R + "id" in element.attrib
        ]
        if not chart_ids:
            continue
        if len(chart_ids) != 1:
            raise ValueError(f"slide {number} has an ambiguous chart relationship")
        rel_type, chart_name = model.target(key, chart_ids[0])
        if chart_name is None or short_type(rel_type) != "chart":
            raise ValueError(f"slide {number} chart relationship is missing")
        chart_key = model.key_of[chart_name]
        external = [
            element.attrib.get(_R + "id", "")
            for element in model.root(chart_name).iter(_C + "externalData")
        ]
        if len(external) != 1:
            raise ValueError(f"slide {number} chart must reference exactly one workbook")
        workbook_type, workbook_name = model.target(chart_key, external[0])
        if workbook_name is None or short_type(workbook_type) != "package":
            raise ValueError(f"slide {number} chart workbook is not an embedded package")
        links[number] = (chart_key, model.key_of[workbook_name])
    if set(links) != {1, 2, 3}:
        raise ValueError("chart fixture must contain exactly the declared three chart slides")
    return links


def _chart_formats(model: PackageModel, links: dict[int, tuple[str, str]]) -> dict[int, object]:
    return {
        number: model.signature(chart, transform=_chart_formatting)
        for number, (chart, _) in links.items()
    }


def _chart_formatting(root: ET.Element) -> ET.Element:
    """Chart formatting only: cached points change with every edit; other checks verify them.

    Axis ids are already renumbered by the package model, like every internal id family.
    """
    stripped = copy.deepcopy(root)
    for cache in stripped.iter():
        if _local(cache.tag) in {"strCache", "numCache", "strLit", "numLit"}:
            for value in cache.iter(_C + "v"):
                value.text = None
    return stripped


def _series_name(series: ET.Element) -> str | None:
    tx = series.find(_C + "tx")
    if tx is None:
        return None
    direct = tx.find(_C + "v")
    if direct is not None:
        return direct.text
    reference = tx.find(_C + "strRef")
    if reference is None:
        return None
    values = _cache_values(reference)
    return values[0] if len(values) == 1 else None


def _ref_values(series: ET.Element, local_name: str) -> tuple[str, list[str]] | None:
    container = series.find(_C + local_name)
    if container is None:
        return None
    references = [
        child
        for child in container
        if _local(child.tag) in {"strRef", "numRef", "strLit", "numLit"}
    ]
    if len(references) != 1:
        return None
    formula = references[0].findtext(_C + "f")
    if formula is None:
        return None
    return formula, _cache_values(references[0])


def _cache_values(reference: ET.Element) -> list[str]:
    cache = next(
        (
            child
            for child in reference
            if _local(child.tag) in {"strCache", "numCache", "strLit", "numLit"}
        ),
        None,
    )
    if cache is None:
        return []
    points = [child for child in cache if _local(child.tag) == "pt"]
    indexed: list[tuple[int, str]] = []
    for point in points:
        index = point.attrib.get("idx")
        value = point.findtext(_C + "v")
        if index is None or value is None:
            return []
        indexed.append((int(index), value))
    if [index for index, _ in indexed] != list(range(len(indexed))):
        return []
    return [value for _, value in indexed]


def _numeric_sequence_equal(actual: list[str], expected: list[object]) -> bool:
    if len(actual) != len(expected):
        return False
    try:
        return all(
            Decimal(value) == Decimal(str(target))
            for value, target in zip(actual, expected, strict=True)
        )
    except InvalidOperation:
        return False


def _formula_cells(formula: str, values: list[str]) -> list[tuple[str, str, str]]:
    sheet, coordinates = _formula_coordinates(formula)
    if len(coordinates) != len(values):
        raise ValueError("chart formula length does not match its cached points")
    return [
        (sheet, coordinate, value) for coordinate, value in zip(coordinates, values, strict=True)
    ]


def _formula_coordinates(formula: str) -> tuple[str, list[str]]:
    sheet_token, separator, range_token = formula.rpartition("!")
    if not separator:
        raise ValueError(f"unsupported chart formula: {formula}")
    sheet = sheet_token.strip("'").replace("''", "'")
    match = _CELL_RANGE.fullmatch(range_token)
    if not sheet or match is None:
        raise ValueError(f"unsupported chart formula: {formula}")
    start_column, start_row, end_column, end_row = match.groups()
    end_column = end_column or start_column
    end_row = end_row or start_row
    coordinates: list[str] = []
    for row in range(int(start_row), int(end_row) + 1):
        for column in range(_column_number(start_column), _column_number(end_column) + 1):
            coordinates.append(f"{_column_name(column)}{row}")
    return sheet, coordinates


def _column_number(column: str) -> int:
    value = 0
    for character in column:
        value = value * 26 + ord(character) - ord("A") + 1
    return value


def _column_name(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def _workbook(value: bytes) -> tuple[PackageModel, str]:
    model = PackageModel(package_bytes(value))
    return model, model.main_part()


def _workbook_cells(value: bytes) -> dict[str, dict[str, str]]:
    """Cell values by sheet name and coordinate, independent of storage form."""
    model, workbook = _workbook(value)
    shared = _shared_strings(model, workbook)
    sheets: dict[str, dict[str, str]] = {}
    for sheet in model.root(model.name_of[workbook]).iter():
        if _local(sheet.tag) != "sheet":
            continue
        name = sheet.attrib.get("name")
        relation_id = sheet.attrib.get(_R + "id")
        if not name or not relation_id:
            raise ValueError("malformed workbook sheet relationship")
        _, sheet_part = model.target(workbook, relation_id)
        if sheet_part is None:
            raise ValueError("workbook sheet part is missing")
        cells: dict[str, str] = {}
        for cell in model.root(sheet_part).iter():
            if _local(cell.tag) != "c":
                continue
            coordinate = cell.attrib.get("r")
            if not coordinate:
                raise ValueError("worksheet cell has no coordinate")
            cells[coordinate] = _cell_value(cell, shared)
        sheets[name] = cells
    return sheets


def _shared_strings(model: PackageModel, workbook: str) -> list[str]:
    tables = model.related(workbook, "sharedStrings")
    if not tables:
        return []
    if len(tables) != 1:
        raise ValueError("workbook has more than one shared string table")
    return [
        "".join(node.text or "" for node in item.iter() if _local(node.tag) == "t")
        for item in model.root(model.name_of[tables[0]]).iter()
        if _local(item.tag) == "si"
    ]


def _cell_value(cell: ET.Element, shared: list[str]) -> str:
    cell_type = cell.attrib.get("t", "n")
    if cell_type == "inlineStr":
        inline = next((element for element in cell.iter() if _local(element.tag) == "is"), None)
        if inline is None:
            raise ValueError("inline string cell has no value")
        return "".join(node.text or "" for node in inline.iter() if _local(node.tag) == "t")
    value = next((element.text for element in cell if _local(element.tag) == "v"), None)
    if value is None:
        raise ValueError("worksheet cell has no cached value")
    if cell_type == "s":
        index = int(value)
        if index < 0 or index >= len(shared):
            raise ValueError("shared string index is out of range")
        return shared[index]
    if cell_type == "n":
        return _canonical_value(value)
    return value


def _canonical_value(value: str) -> str:
    """One spelling per number: ``42``, ``42.0`` and ``4.2E1`` are the same cell value."""
    try:
        number = Decimal(value.strip())
    except InvalidOperation:
        return value
    if not number.is_finite():
        return value
    normalized = number.normalize()
    return format(normalized, "f") if normalized != 0 else "0"


def _workbook_structure(value: bytes) -> dict[str, object]:
    """Embedded workbook parts without cell values, which the values check verifies."""
    model, _ = _workbook(value)
    result: dict[str, object] = {ROOT: model.outgoing(ROOT)}
    for key in model.part_keys():
        if model.is_save_time(key):
            continue
        kinds = {short_type(rel_type) for _, rel_type in model.incoming.get(key, [])}
        if "sharedStrings" in kinds:
            continue
        if "worksheet" in kinds:
            result[key] = (
                model.content_type(model.name_of[key]),
                model.signature(key, transform=_without_cell_values),
                model.outgoing(key),
            )
        else:
            result[key] = model.node(key)
    return result


def _without_cell_values(root: ET.Element) -> ET.Element:
    """Keep cell placement and style; drop value storage (type, value, inline string)."""
    stripped = copy.deepcopy(root)
    for cell in stripped.iter():
        if _local(cell.tag) != "c":
            continue
        cell.attrib.pop("t", None)
        for child in [child for child in cell if _local(child.tag) in {"is", "v"}]:
            cell.remove(child)
    return stripped


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]
