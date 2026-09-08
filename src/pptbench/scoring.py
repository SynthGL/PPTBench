"""Exact Open XML semantic scorers with separate byte-preservation observations."""

from __future__ import annotations

import posixpath
import re
import zipfile
from collections.abc import Callable
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, TypeVar
from xml.etree import ElementTree as ET

from .fixtures import FROZEN_RECIPE, verify_fixture
from .util import package_bytes, package_parts, sha256_bytes, xml_root

Check = dict[str, Any]
T = TypeVar("T")

_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_SLIDE_NAME = re.compile(r"ppt/slides/slide(\d+)\.xml$")
_CELL_RANGE = re.compile(r"^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$")


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
        return _feature_matrix(before, after)
    if lane == "template-mutation":
        try:
            return _template_mutation(before, after)
        except (KeyError, ValueError, ET.ParseError) as exc:
            return [_check("template-package-semantics", False, str(exc))]
    if lane == "chart-data":
        return _chart_data(before, after)
    raise AssertionError("known lane dispatch is incomplete")


def _feature_matrix(before: dict[str, bytes], after: dict[str, bytes]) -> list[Check]:
    checks = [_check("package-readable", True)]
    contracts: list[tuple[str, Callable[[dict[str, bytes]], object]]] = [
        ("slide-count", _slide_count),
        ("slide-order-and-identifiers", _slide_identifiers),
        (
            "slide-relationship-graph",
            lambda parts: _relationship_graph(parts, "ppt/slides/_rels/"),
        ),
        ("shape-tree-semantics", _slide_shape_trees),
        ("text-content-and-order", _slide_text),
        ("text-run-formatting", _text_formatting),
        ("table-cell-grid-and-text", _table_grid),
        ("table-cell-formatting", _table_formatting),
        ("bullet-paragraph-properties", _paragraph_properties),
        ("layout-geometry", _shape_geometry),
        (
            "slide-layout-semantics",
            lambda parts: _xml_part_signatures(parts, "ppt/slideLayouts/"),
        ),
        ("notes-content", _notes_semantics),
        (
            "notes-relationships",
            lambda parts: _relationship_graph(parts, "ppt/notesSlides/_rels/"),
        ),
        ("media-part-presence", lambda parts: _names(parts, "ppt/media/")),
        ("media-exact-bytes", lambda parts: _hashes(parts, "ppt/media/")),
        ("theme-semantics", lambda parts: _xml_part_signatures(parts, "ppt/theme/")),
        (
            "master-semantics",
            lambda parts: _xml_part_signatures(parts, "ppt/slideMasters/"),
        ),
        ("relationship-semantics", lambda parts: _relationship_graph(parts, "")),
        ("opaque-package-part", lambda parts: _hashes(parts, "ppt/unknown/")),
    ]
    for name, extractor in contracts:
        expected = extractor(before)
        if _is_absent_feature(expected):
            checks.append(_unscored(name, "fixture does not contain this declared feature"))
        else:
            checks.append(_equal(name, expected, extractor(after), category="feature"))
    checks.append(_byte_observation("raw-untouched-part-equality", before, after))
    return checks


def _template_mutation(before: dict[str, bytes], after: dict[str, bytes]) -> list[Check]:
    expected, locations = _template_expected_parts(before)
    result: list[Check] = []
    for name, expected_value, actual_value in locations(after):
        result.append(_check(name, actual_value == expected_value, "wrong target value"))
    result.append(
        _equal(
            "only-declared-text-nodes-changed",
            _part_semantics(expected),
            _part_semantics(after),
            category="semantic",
        )
    )
    result.extend(_preservation_checks(before, after))
    result.append(
        _byte_observation(
            "raw-untouched-part-equality",
            before,
            after,
            exclude_names={"ppt/slides/slide5.xml", "ppt/slides/slide7.xml"},
        )
    )
    return result


def _template_expected_parts(
    before: dict[str, bytes],
) -> tuple[dict[str, bytes], Callable[[dict[str, bytes]], list[tuple[str, str, str | None]]]]:
    expected = dict(before)
    table_part = _slide_part(before, 5)
    bullet_part = _slide_part(before, 7)
    table = xml_root(before[table_part])
    bullet = xml_root(before[bullet_part])
    table_targets = [(1, 1, "UPDATED-TABLE-A"), (2, 2, "UPDATED-TABLE-B")]
    for row, column, value in table_targets:
        _set_single_text(_table_cell(table, "PPTBenchTable", row, column), value)
    _set_single_text(_shape_paragraph(bullet, "PPTBenchBullets", 1), "UPDATED-BULLET")
    expected[table_part] = ET.tostring(table, encoding="utf-8")
    expected[bullet_part] = ET.tostring(bullet, encoding="utf-8")

    def locations(parts: dict[str, bytes]) -> list[tuple[str, str, str | None]]:
        table_root = xml_root(parts[_slide_part(parts, 5)])
        bullet_root = xml_root(parts[_slide_part(parts, 7)])
        return [
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

    return expected, locations


def _chart_data(before: dict[str, bytes], after: dict[str, bytes]) -> list[Check]:
    try:
        source_links = _chart_links(before)
        output_links = _chart_links(after)
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
                _chart_slide_structure(before),
                _chart_slide_structure(after),
            ),
        ]
        expected_cells: dict[tuple[str, str, str], str] = {}
        for slide_number, expected_chart in expected.items():
            chart_part, workbook_part = output_links[slide_number]
            root = xml_root(after[chart_part])
            chart_name = str(expected_chart["name"])
            check_name = f"{chart_name}-chart-exact-series-and-points"
            refs = _validate_chart(root, expected_chart)
            checks.append(
                _check(
                    check_name,
                    refs is not None,
                    "chart type, series, or point mapping differs",
                )
            )
            if refs is not None:
                for formula, values in refs:
                    for sheet, coordinate, value in _formula_cells(
                        after[workbook_part], formula, values
                    ):
                        key = (workbook_part, sheet, coordinate)
                        if key in expected_cells and expected_cells[key] != value:
                            raise ValueError("chart series assign conflicting workbook values")
                        expected_cells[key] = value
        checks.append(_workbook_values_check(before, after, source_links, expected_cells))
        checks.append(_workbook_structure_check(before, after, source_links))
        checks.append(
            _equal(
                "chart-formatting-preserved",
                _chart_formats(before),
                _chart_formats(after),
                category="preservation",
            )
        )
        checks.extend(
            _preservation_checks(before, after, exclude_prefixes=("ppt/charts/", "ppt/embeddings/"))
        )
        checks.append(
            _byte_observation(
                "raw-untouched-part-equality",
                before,
                after,
                exclude_prefixes=("ppt/charts/", "ppt/embeddings/"),
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
    before: dict[str, bytes],
    after: dict[str, bytes],
    source_links: dict[int, tuple[str, str]],
    expected_cells: dict[tuple[str, str, str], str],
) -> Check:
    source_workbooks = {workbook for _, workbook in source_links.values()}
    output_workbooks = {part for part, _, _ in expected_cells}
    if source_workbooks != output_workbooks:
        return _check("related-workbook-identity", False, "chart workbook relations changed")
    expected_grids = {part: _workbook_cells(before[part]) for part in source_workbooks}
    actual_grids = {part: _workbook_cells(after[part]) for part in source_workbooks}
    for (part, sheet, coordinate), expected_value in expected_cells.items():
        try:
            expected_grids[part][sheet][coordinate] = expected_value
        except KeyError:
            return _check(
                "related-workbook-values",
                False,
                "chart formula points outside the source workbook",
            )
    return _equal("related-workbook-values", expected_grids, actual_grids)


def _workbook_structure_check(
    before: dict[str, bytes], after: dict[str, bytes], links: dict[int, tuple[str, str]]
) -> Check:
    workbooks = {workbook for _, workbook in links.values()}
    return _equal(
        "related-workbook-structure-and-formatting",
        {part: _workbook_structure(before[part]) for part in workbooks},
        {part: _workbook_structure(after[part]) for part in workbooks},
        category="preservation",
    )


def _preservation_checks(
    before: dict[str, bytes],
    after: dict[str, bytes],
    *,
    exclude_prefixes: tuple[str, ...] = (),
) -> list[Check]:
    def filtered(parts: dict[str, bytes], prefix: str) -> dict[str, bytes]:
        return {
            name: data
            for name, data in parts.items()
            if name.startswith(prefix) and not name.startswith(exclude_prefixes)
        }

    return [
        _equal(
            "notes-preserved",
            _part_semantics(filtered(before, "ppt/notesSlides/")),
            _part_semantics(filtered(after, "ppt/notesSlides/")),
            category="preservation",
        ),
        _equal(
            "media-preserved",
            _hashes(filtered(before, "ppt/media/"), ""),
            _hashes(filtered(after, "ppt/media/"), ""),
            category="preservation",
        ),
        _equal(
            "relationships-preserved",
            _relationship_graph(before, ""),
            _relationship_graph(after, ""),
            category="preservation",
        ),
        _equal(
            "opaque-parts-preserved",
            _hashes(filtered(before, "ppt/unknown/"), ""),
            _hashes(filtered(after, "ppt/unknown/"), ""),
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
    exclude_prefixes: tuple[str, ...] = (),
) -> Check:
    omitted = exclude_names or set()
    before_untouched = {
        part: value
        for part, value in before.items()
        if part not in omitted and not part.startswith(exclude_prefixes)
    }
    after_untouched = {
        part: value
        for part, value in after.items()
        if part not in omitted and not part.startswith(exclude_prefixes)
    }
    return _check(
        name,
        before_untouched == after_untouched,
        "raw untouched part bytes differ",
        category="byte-only",
        scored=False,
    )


def _names(parts: dict[str, bytes], prefix: str) -> list[str]:
    return sorted(name for name in parts if name.startswith(prefix))


def _hashes(parts: dict[str, bytes], prefix: str) -> dict[str, str]:
    return {name: sha256_bytes(data) for name, data in parts.items() if name.startswith(prefix)}


def _slide_count(parts: dict[str, bytes]) -> int:
    return len(_slide_parts(parts))


def _slide_parts(parts: dict[str, bytes]) -> list[tuple[int, str]]:
    result: list[tuple[int, str]] = []
    for name in parts:
        match = _SLIDE_NAME.fullmatch(name)
        if match:
            result.append((int(match.group(1)), name))
    return sorted(result)


def _slide_part(parts: dict[str, bytes], number: int) -> str:
    part = f"ppt/slides/slide{number}.xml"
    if part not in parts:
        raise ValueError(f"missing expected slide part: {part}")
    return part


def _slide_identifiers(parts: dict[str, bytes]) -> list[tuple[int, str]]:
    return _slide_parts(parts)


def _slide_shape_trees(parts: dict[str, bytes]) -> dict[str, object]:
    return {name: _xml_signature(parts[name]) for _, name in _slide_parts(parts)}


def _slide_text(parts: dict[str, bytes]) -> dict[str, list[str]]:
    return {
        name: [
            element.text or ""
            for element in xml_root(parts[name]).iter()
            if element.tag == _A + "t"
        ]
        for _, name in _slide_parts(parts)
    }


def _text_formatting(parts: dict[str, bytes]) -> dict[str, list[object]]:
    return {
        name: [
            _element_signature(element)
            for element in xml_root(parts[name]).iter()
            if element.tag in {_A + "rPr", _A + "endParaRPr"}
        ]
        for _, name in _slide_parts(parts)
    }


def _table_grid(parts: dict[str, bytes]) -> dict[str, list[list[list[str]]]]:
    output: dict[str, list[list[list[str]]]] = {}
    for _, name in _slide_parts(parts):
        grids: list[list[list[str]]] = []
        for table in xml_root(parts[name]).iter(_A + "tbl"):
            grids.append(
                [
                    [_single_text(cell) or "" for cell in row.findall(_A + "tc")]
                    for row in table.findall(_A + "tr")
                ]
            )
        output[name] = grids
    return output


def _table_formatting(parts: dict[str, bytes]) -> dict[str, list[object]]:
    return {
        name: [_element_signature(table) for table in xml_root(parts[name]).iter(_A + "tbl")]
        for _, name in _slide_parts(parts)
    }


def _paragraph_properties(parts: dict[str, bytes]) -> dict[str, list[object]]:
    return {
        name: [_element_signature(paragraph) for paragraph in xml_root(parts[name]).iter(_A + "p")]
        for _, name in _slide_parts(parts)
    }


def _shape_geometry(parts: dict[str, bytes]) -> dict[str, list[object]]:
    return {
        name: [
            _element_signature(element)
            for element in xml_root(parts[name]).iter()
            if _local(element.tag) in {"xfrm", "off", "ext", "chOff", "chExt"}
        ]
        for _, name in _slide_parts(parts)
    }


def _notes_semantics(parts: dict[str, bytes]) -> dict[str, object]:
    return _xml_part_signatures(parts, "ppt/notesSlides/")


def _xml_part_signatures(parts: dict[str, bytes], prefix: str) -> dict[str, object]:
    return {
        name: _xml_signature(data)
        for name, data in sorted(parts.items())
        if name.startswith(prefix) and name.endswith((".xml", ".rels"))
    }


def _relationship_graph(
    parts: dict[str, bytes], prefix: str
) -> dict[str, list[tuple[str, str, str]]]:
    graph: dict[str, list[tuple[str, str, str]]] = {}
    for name, data in parts.items():
        if name.endswith(".rels") and name.startswith(prefix):
            root = xml_root(data)
            graph[name] = [
                (
                    element.attrib.get("Id", ""),
                    element.attrib.get("Type", ""),
                    element.attrib.get("Target", ""),
                )
                for element in root
            ]
    return dict(sorted(graph.items()))


def _part_semantics(parts: dict[str, bytes]) -> dict[str, object]:
    return {
        name: _xml_signature(data) if name.endswith((".xml", ".rels")) else sha256_bytes(data)
        for name, data in sorted(parts.items())
    }


def _xml_signature(data: bytes) -> object:
    return _element_signature(xml_root(data))


def _element_signature(element: ET.Element, *, in_text: bool = False) -> object:
    text_sensitive = in_text or element.tag == _A + "t"
    text = (
        element.text if text_sensitive else (element.text if (element.text or "").strip() else None)
    )
    return (
        element.tag,
        tuple(sorted(element.attrib.items())),
        text,
        tuple(_element_signature(child, in_text=text_sensitive) for child in element),
    )


def _is_absent_feature(value: object) -> bool:
    return value in ({}, [], (), 0, None)


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


def _chart_links(parts: dict[str, bytes]) -> dict[int, tuple[str, str]]:
    links: dict[int, tuple[str, str]] = {}
    for number, slide_part in _slide_parts(parts):
        root = xml_root(parts[slide_part])
        chart_ids = [
            element.attrib[_R + "id"]
            for element in root.iter(_C + "chart")
            if _R + "id" in element.attrib
        ]
        if not chart_ids:
            continue
        if len(chart_ids) != 1:
            raise ValueError(f"slide {number} has an ambiguous chart relationship")
        relationships = _relationships(parts, slide_part)
        relation = relationships.get(chart_ids[0])
        if relation is None:
            raise ValueError(f"slide {number} chart relationship is missing")
        chart_part = _resolve_part(slide_part, relation[1])
        chart_relationships = _relationships(parts, chart_part)
        workbooks = [
            _resolve_part(chart_part, target)
            for _, target in chart_relationships.values()
            if target.lower().endswith(".xlsx")
        ]
        if len(workbooks) != 1:
            raise ValueError(f"chart {chart_part} must relate to exactly one embedded workbook")
        links[number] = (chart_part, workbooks[0])
    if set(links) != {1, 2, 3}:
        raise ValueError("chart fixture must contain exactly the declared three chart slides")
    return links


def _relationships(parts: dict[str, bytes], source_part: str) -> dict[str, tuple[str, str]]:
    relation_part = posixpath.join(
        posixpath.dirname(source_part),
        "_rels",
        posixpath.basename(source_part) + ".rels",
    )
    if relation_part not in parts:
        raise ValueError(f"missing relationships for {source_part}")
    root = xml_root(parts[relation_part])
    if root.tag != _REL + "Relationships":
        raise ValueError(f"invalid relationships root: {relation_part}")
    result: dict[str, tuple[str, str]] = {}
    for element in root:
        relation_id = element.attrib.get("Id")
        relation_type = element.attrib.get("Type")
        target = element.attrib.get("Target")
        if not relation_id or not relation_type or not target or relation_id in result:
            raise ValueError(f"malformed relationship in {relation_part}")
        result[relation_id] = (relation_type, target)
    return result


def _resolve_part(source_part: str, target: str) -> str:
    if target.startswith("/") or "\\" in target:
        raise ValueError("unsafe relationship target")
    resolved = posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))
    if resolved.startswith("../") or resolved == "..":
        raise ValueError("relationship escapes package")
    return resolved


def _chart_slide_structure(parts: dict[str, bytes]) -> dict[str, object]:
    return {name: _xml_signature(parts[name]) for _, name in _slide_parts(parts)}


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


def _formula_cells(workbook: bytes, formula: str, values: list[str]) -> list[tuple[str, str, str]]:
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


def _workbook_cells(value: bytes) -> dict[str, dict[str, str]]:
    parts = package_bytes(value)
    shared = _shared_strings(parts)
    workbook = xml_root(parts["xl/workbook.xml"])
    relations = _relationships(parts, "xl/workbook.xml")
    sheets: dict[str, dict[str, str]] = {}
    for sheet in workbook.iter():
        if _local(sheet.tag) != "sheet":
            continue
        name = sheet.attrib.get("name")
        relation_id = sheet.attrib.get(_R + "id")
        if not name or not relation_id or relation_id not in relations:
            raise ValueError("malformed workbook sheet relationship")
        sheet_part = _resolve_part("xl/workbook.xml", relations[relation_id][1])
        if sheet_part not in parts:
            raise ValueError("workbook sheet part is missing")
        cells: dict[str, str] = {}
        for cell in xml_root(parts[sheet_part]).iter():
            if _local(cell.tag) != "c":
                continue
            coordinate = cell.attrib.get("r")
            if not coordinate:
                raise ValueError("worksheet cell has no coordinate")
            cells[coordinate] = _cell_value(cell, shared)
        sheets[name] = cells
    return sheets


def _shared_strings(parts: dict[str, bytes]) -> list[str]:
    part = "xl/sharedStrings.xml"
    if part not in parts:
        return []
    return [
        "".join(node.text or "" for node in item.iter() if _local(node.tag) == "t")
        for item in xml_root(parts[part]).iter()
        if _local(item.tag) == "si"
    ]


def _cell_value(cell: ET.Element, shared: list[str]) -> str:
    cell_type = cell.attrib.get("t")
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
    return value


def _workbook_structure(value: bytes) -> dict[str, object]:
    parts = package_bytes(value)
    result: dict[str, object] = {}
    for name, data in parts.items():
        if name == "xl/sharedStrings.xml":
            continue
        if name == "docProps/core.xml":
            result[name] = _core_properties_signature(xml_root(data))
        elif name.startswith("xl/worksheets/") and name.endswith(".xml"):
            result[name] = _worksheet_structure(xml_root(data))
        elif name.endswith((".xml", ".rels")):
            result[name] = _xml_signature(data)
        else:
            result[name] = sha256_bytes(data)
    return result


def _core_properties_signature(root: ET.Element) -> object:
    def signature(element: ET.Element) -> object:
        text = (
            None
            if _local(element.tag) in {"created", "modified"}
            else (element.text if (element.text or "").strip() else None)
        )
        return (
            element.tag,
            tuple(sorted(element.attrib.items())),
            text,
            tuple(signature(child) for child in element),
        )

    return signature(root)


def _worksheet_structure(root: ET.Element) -> object:
    def signature(element: ET.Element) -> object:
        local_name = _local(element.tag)
        attributes = tuple(
            sorted(
                (key, value)
                for key, value in element.attrib.items()
                if not (local_name == "c" and key == "t")
            )
        )
        children = (
            tuple(signature(child) for child in element if _local(child.tag) not in {"is", "v"})
            if local_name == "c"
            else tuple(signature(child) for child in element)
        )
        return (
            element.tag,
            attributes,
            element.text if (element.text or "").strip() else None,
            children,
        )

    return signature(root)


def _chart_formats(parts: dict[str, bytes]) -> dict[str, object]:
    return {
        name: _chart_format_signature(xml_root(data))
        for name, data in parts.items()
        if name.startswith("ppt/charts/chart") and name.endswith(".xml")
    }


def _chart_format_signature(root: ET.Element) -> object:
    def signature(element: ET.Element, data_cache: bool = False) -> object:
        in_cache = data_cache or _local(element.tag) in {
            "strCache",
            "numCache",
            "strLit",
            "numLit",
        }
        text = (
            None
            if in_cache and _local(element.tag) == "v"
            else (element.text if (element.text or "").strip() else None)
        )
        return (
            element.tag,
            tuple(sorted(element.attrib.items())),
            text,
            tuple(signature(child, in_cache) for child in element),
        )

    return signature(root)


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]
