"""PPTBench Python-UNO script for LibreOffice Impress.

LibreOffice runs this file inside its own embedded Python as a user-profile
script (``vnd.sun.star.script:pptbench_uno.py$run?language=Python&location=user``);
``pptbench.worker`` installs it into a fresh, run-owned profile for every
invocation. Arguments arrive through environment variables because script URLs
carry none:

- ``PPTBENCH_UNO_LANE``: ``feature-matrix``, ``template-mutation``, or ``chart-data``
- ``PPTBENCH_UNO_INPUT`` / ``PPTBENCH_UNO_OUTPUT``: absolute ``.pptx`` paths
- ``PPTBENCH_UNO_STATUS``: absolute path of a JSON status file this script writes

Every lane loads the deck through Impress's PowerPoint import, edits it through
the UNO API, and stores it with the ``Impress Office Open XML`` export filter.
"""

from __future__ import annotations

import json
import os
import traceback

import uno
from com.sun.star.beans import PropertyValue

# Base of every UNO exception (RuntimeException, IllegalArgumentException,
# IOException, ...), resolved once so pyuno maps raised UNO errors onto it.
_UNO_EXCEPTION = uno.getClass("com.sun.star.uno.Exception")
_EXPECTED_FAILURES = (
    _UNO_EXCEPTION,
    AttributeError,
    LookupError,
    OSError,
    TypeError,
    ValueError,
)


def _property(name, value):
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item


def _shape_named(page, name):
    for index in range(page.Count):
        shape = page.getByIndex(index)
        if shape.Name == name:
            return shape
    raise LookupError(f"required shape is missing: {name}")


def _paragraph(text, index):
    paragraphs = text.createEnumeration()
    position = 0
    while paragraphs.hasMoreElements():
        paragraph = paragraphs.nextElement()
        if position == index:
            return paragraph
        position += 1
    raise LookupError(f"paragraph {index} is missing")


def _mutate_template(document):
    pages = document.DrawPages
    table = _shape_named(pages.getByIndex(4), "PPTBenchTable").Model
    # XTable.getCellByPosition takes (column, row).
    table.getCellByPosition(1, 1).setString("UPDATED-TABLE-A")
    table.getCellByPosition(2, 2).setString("UPDATED-TABLE-B")
    bullets = _shape_named(pages.getByIndex(6), "PPTBenchBullets")
    _paragraph(bullets.Text, 1).setString("UPDATED-BULLET")


def _single_series(chart):
    diagram = chart.getFirstDiagram()
    coordinate_systems = diagram.getCoordinateSystems()
    if len(coordinate_systems) != 1:
        raise ValueError("expected one coordinate system")
    chart_types = coordinate_systems[0].getChartTypes()
    if len(chart_types) != 1:
        raise ValueError("expected one chart type")
    series = chart_types[0].getDataSeries()
    if len(series) != 1:
        raise ValueError("expected one data series")
    return coordinate_systems[0], series[0]


def _set_role_values(chart, series, values_by_role):
    provider = chart.getDataProvider()
    seen = set()
    for labeled in series.getDataSequences():
        sequence = labeled.getValues()
        role = sequence.Role
        if role not in values_by_role:
            continue
        provider.setDataByRangeRepresentation(
            sequence.getSourceRangeRepresentation(), tuple(values_by_role[role])
        )
        seen.add(role)
    missing = set(values_by_role) - seen
    if missing:
        raise ValueError(f"series has no data sequence for roles {sorted(missing)}")


def _replace_chart_data(document):
    pages = document.DrawPages
    category = _shape_named(pages.getByIndex(0), "Chart 1").Model
    coordinate_system, series = _single_series(category)
    _set_role_values(category, series, {"values-y": (42.0, 84.0, 126.0)})
    categories = coordinate_system.getAxisByDimension(0, 0).getScaleData().Categories
    if categories is None:
        raise ValueError("category chart has no category sequence")
    category.getDataProvider().setDataByRangeRepresentation(
        categories.getValues().getSourceRangeRepresentation(), ("East", "42", "84")
    )

    xy = _shape_named(pages.getByIndex(1), "Chart 1").Model
    _, series = _single_series(xy)
    _set_role_values(xy, series, {"values-x": (5.0, 21.0, 55.0), "values-y": (13.0, 34.0, 89.0)})

    bubble = _shape_named(pages.getByIndex(2), "Chart 1").Model
    _, series = _single_series(bubble)
    _set_role_values(
        bubble,
        series,
        {
            "values-x": (5.0, 34.0, 8.0),
            "values-y": (13.0, 55.0, 13.0),
            "values-size": (21.0, 89.0, 21.0),
        },
    )


def _edit(desktop, lane, source, target):
    document = desktop.loadComponentFromURL(
        uno.systemPathToFileUrl(source),
        "_blank",
        0,
        (_property("Hidden", True), _property("ReadOnly", False)),
    )
    if document is None:
        raise RuntimeError("LibreOffice could not load the input deck")
    try:
        if lane == "template-mutation":
            _mutate_template(document)
        elif lane == "chart-data":
            _replace_chart_data(document)
        elif lane != "feature-matrix":
            raise ValueError(f"unsupported lane: {lane}")
        document.storeToURL(
            uno.systemPathToFileUrl(target),
            (_property("FilterName", "Impress Office Open XML"),),
        )
    finally:
        document.close(True)


def run(*_args):
    status = {"outcome": "failure"}
    context = uno.getComponentContext()
    desktop = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", context
    )
    try:
        _edit(
            desktop,
            os.environ["PPTBENCH_UNO_LANE"],
            os.environ["PPTBENCH_UNO_INPUT"],
            os.environ["PPTBENCH_UNO_OUTPUT"],
        )
        status = {"outcome": "success"}
    except _EXPECTED_FAILURES as exc:  # report every expected failure to the launcher
        status = {
            "outcome": "failure",
            "error": f"{type(exc).__name__}: {exc}",
            "traceback": traceback.format_exc(limit=8),
        }
    finally:
        with open(os.environ["PPTBENCH_UNO_STATUS"], "w", encoding="utf-8") as handle:
            json.dump(status, handle)
        desktop.terminate()


g_exportedScripts = (run,)
