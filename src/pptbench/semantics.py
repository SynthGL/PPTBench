"""Name-independent semantic model of an Open Packaging Convention package.

Scored checks must observe only what a consumer of the file can observe. A part is
therefore identified by how a consumer reaches it: the chain of relationship types
from the package root, with same-type siblings ordered by where the owning XML first
references them (document order is user-visible). Part names, relationship ids, ZIP
order, XML prefixes, and attribute order never contribute to a comparison. Media
targets are identified by content hash, and unrecognised parts by their content
digest plus the relationship types that reach them.
"""

from __future__ import annotations

import posixpath
from collections import Counter, deque
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from io import BytesIO
from urllib.parse import unquote
from xml.etree import ElementTree as ET

from .util import sha256_bytes, xml_root

ROOT = "package"

_CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_A_TEXT = "{http://schemas.openxmlformats.org/drawingml/2006/main}t"
_RELATIONSHIP_ATTRIBUTE_NAMESPACES = (
    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}",
    "{http://purl.oclc.org/ooxml/officeDocument/relationships}",
)
_VML_RELID = "{urn:schemas-microsoft-com:office:office}relid"
_MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
# Attributes whose values name XML prefixes; resolved to namespace URIs at parse time.
_PREFIX_VALUED_ATTRIBUTES = {
    _MC + "Ignorable",
    _MC + "ProcessContent",
    _MC + "PreserveElements",
    _MC + "PreserveAttributes",
    _MC + "MustUnderstand",
    "{http://www.w3.org/2001/XMLSchema-instance}type",
}
_MC_CHOICE = _MC + "Choice"
# xsd:boolean has two lexical forms per value.
_BOOLEAN_CANONICAL = {"true": "1", "false": "0"}


def _defaults(namespace: str, table: dict[str, dict[str, str]]) -> dict[tuple[str, str], str]:
    return {
        (namespace + element, attribute): value
        for element, attributes in table.items()
        for attribute, value in attributes.items()
    }


# Writer application/version stamps, rewritten on every save (save-time metadata).
_APP_VERSION_ELEMENTS = frozenset(
    {"{http://schemas.openxmlformats.org/spreadsheetml/2006/main}fileVersion"}
)
_APP_VERSION_ATTRIBUTES = frozenset(
    {("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}calcPr", "calcId")}
)

_CHART_BOOLEANS_DEFAULT_TRUE = (
    "applyToEnd",
    "applyToFront",
    "applyToSides",
    "auto",
    "autoTitleDeleted",
    "bubble3D",
    "date1904",
    "delete",
    "invertIfNegative",
    "marker",
    "noMultiLvlLbl",
    "overlay",
    "plotVisOnly",
    "roundedCorners",
    "showBubbleSize",
    "showCatName",
    "showDLblsOverMax",
    "showHorzBorder",
    "showKeys",
    "showLeaderLines",
    "showLegendKey",
    "showNegBubbles",
    "showOutline",
    "showPercent",
    "showSerName",
    "showVal",
    "showVertBorder",
    "smooth",
    "varyColors",
    "wireframe",
)
# Attribute defaults declared by the ECMA-376 Part 1 transitional schemas (sml.xsd,
# pml.xsd, dml-main.xsd, dml-chart.xsd), keyed by element and attribute, with
# booleans in canonical form. Writing one of these out explicitly changes no meaning.
_SCHEMA_DEFAULTS: dict[tuple[str, str], str] = {
    **_defaults(
        "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}",
        {
            "c": {"s": "0", "t": "n", "cm": "0", "vm": "0", "ph": "0"},
            "row": {
                "s": "0",
                "customFormat": "0",
                "hidden": "0",
                "customHeight": "0",
                "outlineLevel": "0",
                "collapsed": "0",
                "thickTop": "0",
                "thickBot": "0",
                "ph": "0",
            },
            "col": {
                "style": "0",
                "hidden": "0",
                "bestFit": "0",
                "customWidth": "0",
                "phonetic": "0",
                "outlineLevel": "0",
                "collapsed": "0",
            },
            "sheetView": {
                "windowProtection": "0",
                "showFormulas": "0",
                "showGridLines": "1",
                "showRowColHeaders": "1",
                "showZeros": "1",
                "rightToLeft": "0",
                "tabSelected": "0",
                "showRuler": "1",
                "showOutlineSymbols": "1",
                "defaultGridColor": "1",
                "showWhiteSpace": "1",
                "view": "normal",
                "colorId": "64",
                "zoomScale": "100",
                "zoomScaleNormal": "0",
                "zoomScaleSheetLayoutView": "0",
                "zoomScalePageLayoutView": "0",
            },
            "sheetFormatPr": {
                "baseColWidth": "8",
                "customHeight": "0",
                "zeroHeight": "0",
                "thickTop": "0",
                "thickBottom": "0",
                "outlineLevelRow": "0",
                "outlineLevelCol": "0",
            },
            "sheet": {"state": "visible"},
            "workbookView": {
                "visibility": "visible",
                "minimized": "0",
                "showHorizontalScroll": "1",
                "showVerticalScroll": "1",
                "showSheetTabs": "1",
                "tabRatio": "600",
                "firstSheet": "0",
                "activeTab": "0",
                "autoFilterDateGrouping": "1",
            },
            "xf": {"quotePrefix": "0", "pivotButton": "0"},
        },
    ),
    **_defaults(
        "{http://schemas.openxmlformats.org/presentationml/2006/main}",
        {
            "cNvPr": {"descr": "", "hidden": "0", "title": ""},
            "cNvSpPr": {"txBox": "0"},
            "xfrm": {"rot": "0", "flipH": "0", "flipV": "0"},
            "sp": {"useBgFill": "0"},
            "ph": {"type": "obj", "orient": "horz", "sz": "full", "idx": "0"},
            "sld": {"show": "1", "showMasterSp": "1", "showMasterPhAnim": "1"},
            "sldLayout": {
                "showMasterSp": "1",
                "showMasterPhAnim": "1",
                "matchingName": "",
                "type": "cust",
                "preserve": "0",
                "userDrawn": "0",
            },
            "sldMaster": {"preserve": "0"},
            "sldSz": {"type": "custom"},
            "presentation": {
                "firstSlideNum": "1",
                "showSpecialPlsOnTitleSld": "1",
                "rtl": "0",
                "removePersonalInfoOnSave": "0",
                "compatMode": "0",
                "strictFirstAndLastChars": "1",
                "embedTrueTypeFonts": "0",
                "saveSubsetFonts": "0",
                "autoCompressPictures": "1",
                "bookmarkIdSeed": "1",
            },
        },
    ),
    **_defaults(
        "{http://schemas.openxmlformats.org/drawingml/2006/main}",
        {
            "cNvPr": {"descr": "", "hidden": "0", "title": ""},
            "xfrm": {"rot": "0", "flipH": "0", "flipV": "0"},
            "tblPr": {
                "rtl": "0",
                "firstRow": "0",
                "firstCol": "0",
                "lastRow": "0",
                "lastCol": "0",
                "bandRow": "0",
                "bandCol": "0",
            },
            "tc": {"rowSpan": "1", "gridSpan": "1", "hMerge": "0", "vMerge": "0"},
            "tcPr": {
                "marL": "91440",
                "marR": "91440",
                "marT": "45720",
                "marB": "45720",
                "vert": "horz",
                "anchor": "t",
                "anchorCtr": "0",
                "horzOverflow": "clip",
            },
            "rPr": {"dirty": "1", "err": "0", "smtClean": "1", "smtId": "0"},
            "endParaRPr": {"dirty": "1", "err": "0", "smtClean": "1", "smtId": "0"},
            "defRPr": {"dirty": "1", "err": "0", "smtClean": "1", "smtId": "0"},
        },
    ),
    **_defaults(
        "{http://schemas.openxmlformats.org/drawingml/2006/chart}",
        {element: {"val": "1"} for element in _CHART_BOOLEANS_DEFAULT_TRUE},
    ),
}

# Relationship types (final path segment) whose targets are media: identified by
# content hash, never by part name or by how many parts carry identical bytes.
MEDIA_TYPES = frozenset({"image", "media", "video", "audio", "hdphoto"})
# Package-level save-time metadata: regenerated on every save and not document content.
THUMBNAIL_TYPE = "thumbnail"
_CORE_PROPERTIES_TYPE = "core-properties"
_EXTENDED_PROPERTIES_TYPE = "extended-properties"
_SAVE_TIME_CORE = frozenset({"created", "modified", "lastModifiedBy", "revision"})
_USER_EXTENDED = frozenset({"Template", "Manager", "Company", "HyperlinkBase"})
# Relationship types a PresentationML consumer understands. Anything else, and any
# part no relationship reaches, is an opaque part that must survive unchanged.
_KNOWN_TYPES = frozenset(
    {
        "officeDocument",
        _CORE_PROPERTIES_TYPE,
        _EXTENDED_PROPERTIES_TYPE,
        "custom-properties",
        THUMBNAIL_TYPE,
        "presentation",
        "slide",
        "slideLayout",
        "slideMaster",
        "notesSlide",
        "notesMaster",
        "handoutMaster",
        "theme",
        "themeOverride",
        "presProps",
        "viewProps",
        "tableStyles",
        "commentAuthors",
        "comments",
        "printerSettings",
        "hyperlink",
        "chart",
        "chartUserShapes",
        "chartStyle",
        "chartColorStyle",
        "package",
        "oleObject",
        "diagramData",
        "diagramLayout",
        "diagramQuickStyle",
        "diagramColors",
        "diagramDrawing",
        "tags",
        "font",
        "vbaProject",
        "customXml",
        "customXmlProps",
        "worksheet",
        "chartsheet",
        "sharedStrings",
        "styles",
        "calcChain",
        "drawing",
        "table",
        "pivotTable",
        "pivotCacheDefinition",
        "pivotCacheRecords",
        "externalLink",
        "vmlDrawing",
        "ctrlProp",
        "person",
        "threadedComment",
        *MEDIA_TYPES,
    }
)


@dataclass(frozen=True)
class Relationship:
    rel_id: str
    rel_type: str
    target: str
    external: bool


def short_type(rel_type: str) -> str:
    return rel_type.rstrip("/").rsplit("/", 1)[-1]


def semantic_root(value: bytes) -> ET.Element:
    """Parse XML with prefix-bearing markup-compatibility values resolved to URIs."""
    xml_root(value)  # shared DTD/entity rejection and well-formedness check
    scopes: list[dict[str, str]] = [{}]
    pending: dict[str, str] = {}
    root: ET.Element | None = None
    for event, item in ET.iterparse(BytesIO(value), events=("start-ns", "start", "end")):
        if event == "start-ns":
            prefix, uri = item
            pending[prefix] = uri
        elif event == "start":
            element = item
            scope = {**scopes[-1], **pending}
            pending = {}
            scopes.append(scope)
            if root is None:
                root = element
            for name in list(element.attrib):
                if name in _PREFIX_VALUED_ATTRIBUTES or (
                    name == "Requires" and element.tag == _MC_CHOICE
                ):
                    element.set(name, _resolve_prefixes(element.attrib[name], scope))
        else:
            scopes.pop()
    assert root is not None  # xml_root already proved the document has a root
    return root


def _resolve_prefixes(value: str, scope: dict[str, str]) -> str:
    tokens: list[str] = []
    for token in value.split():
        prefix, separator, local = token.partition(":")
        if separator:
            tokens.append("{" + scope.get(prefix, prefix) + "}" + local)
        else:
            tokens.append(scope.get(token, token))
    return " ".join(sorted(tokens))


def element_signature(
    element: ET.Element,
    resolve: Callable[[str], object] | None = None,
    *,
    in_text: bool = False,
) -> object:
    """Order-preserving element signature; relationship-id attributes are resolved.

    Lexical variants with one schema meaning collapse: ``true``/``1`` booleans,
    attributes written out at their schema default, and empty ``count="0"`` lists.
    """
    text_sensitive = in_text or element.tag == _A_TEXT
    text = (
        element.text if text_sensitive else (element.text if (element.text or "").strip() else None)
    )
    attributes: list[tuple[str, object]] = []
    for name, value in element.attrib.items():
        if resolve is not None and _is_relationship_attribute(name):
            attributes.append((name, resolve(value)))
            continue
        value = _BOOLEAN_CANONICAL.get(value, value)
        if (
            _SCHEMA_DEFAULTS.get((element.tag, name)) == value
            or (element.tag, name) in _APP_VERSION_ATTRIBUTES
        ):
            continue
        attributes.append((name, value))
    return (
        element.tag,
        tuple(sorted(attributes, key=lambda item: item[0])),
        text,
        tuple(
            element_signature(child, resolve, in_text=text_sensitive)
            for child in element
            if not _is_empty_list(child) and child.tag not in _APP_VERSION_ELEMENTS
        ),
    )


def _is_empty_list(element: ET.Element) -> bool:
    return (
        element.attrib == {"count": "0"} and len(element) == 0 and not (element.text or "").strip()
    )


def _is_relationship_attribute(name: str) -> bool:
    return name == _VML_RELID or name.startswith(_RELATIONSHIP_ATTRIBUTE_NAMESPACES)


def relationship_part(source: str) -> str:
    if not source:
        return "_rels/.rels"
    return posixpath.join(posixpath.dirname(source), "_rels", posixpath.basename(source) + ".rels")


def resolve_target(source: str, target: str) -> str | None:
    """Resolve an internal relationship target to a part name; None if it escapes."""
    target = unquote(target.split("#", 1)[0])
    if not target or "\\" in target:
        return None
    if target.startswith("/"):
        resolved = posixpath.normpath(target.lstrip("/"))
    else:
        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(source), target))
    if resolved in {".", ".."} or resolved.startswith("../"):
        return None
    return resolved


class PackageModel:
    """Canonical, name-free view of an OPC package's parts and relationship graph."""

    def __init__(self, parts: dict[str, bytes]) -> None:
        self.parts = parts
        self._lookup = {name.lower(): name for name in parts}
        self._content_types = _content_types(parts)
        self._relationship_cache: dict[str, list[Relationship]] = {}
        self._root_cache: dict[str, ET.Element] = {}
        self._shallow_cache: dict[str, str] = {}
        self.key_of: dict[str, str] = {}
        self.name_of: dict[str, str] = {}
        self.edges: dict[str, list[tuple[str, str, str | None]]] = {}
        self.incoming: dict[str, list[tuple[str, str]]] = {}
        self._build()

    # Graph construction -------------------------------------------------------

    def relationships(self, source: str) -> list[Relationship]:
        """Relationships of a part (``""`` is the package), with resolved targets."""
        cached = self._relationship_cache.get(source)
        if cached is not None:
            return cached
        result: list[Relationship] = []
        part = relationship_part(source)
        if part in self.parts:
            root = xml_root(self.parts[part])
            if root.tag != _PKG_REL + "Relationships":
                raise ValueError(f"invalid relationships root: {part}")
            for element in root:
                rel_id = element.attrib.get("Id", "")
                rel_type = element.attrib.get("Type", "")
                target = element.attrib.get("Target", "")
                if element.attrib.get("TargetMode") == "External":
                    result.append(Relationship(rel_id, rel_type, target, True))
                    continue
                resolved = resolve_target(source, target)
                name = self._lookup.get(resolved.lower()) if resolved is not None else None
                result.append(Relationship(rel_id, rel_type, name or "", False))
        self._relationship_cache[source] = result
        return result

    def _build(self) -> None:
        self.key_of[""] = ROOT
        self.name_of[ROOT] = ""
        queue: deque[tuple[str, str]] = deque([("", ROOT)])
        while queue:
            source, key = queue.popleft()
            counters: Counter[str] = Counter()
            edges: list[tuple[str, str, str | None]] = []
            for relationship in self._ordered(source):
                kind = short_type(relationship.rel_type)
                index = counters[kind]
                counters[kind] += 1
                target_name = relationship.target
                if relationship.external:
                    edges.append((relationship.rel_type, "external:" + target_name, None))
                    continue
                if not target_name:
                    edges.append((relationship.rel_type, "missing", None))
                    continue
                if target_name not in self.key_of:
                    child = f"{key}/{kind}{index}"
                    self.key_of[target_name] = child
                    self.name_of[child] = target_name
                    queue.append((target_name, child))
                target_key = self.key_of[target_name]
                edges.append((relationship.rel_type, target_key, target_name))
                self.incoming.setdefault(target_key, []).append((key, relationship.rel_type))
            self.edges[key] = edges

    def _ordered(self, source: str) -> list[Relationship]:
        relationships = self.relationships(source)
        by_id = {relationship.rel_id: relationship for relationship in relationships}
        explicit: list[Relationship] = []
        if source and self.is_xml(source):
            seen: set[str] = set()
            for element in self.root(source).iter():
                for name, value in element.attrib.items():
                    if _is_relationship_attribute(name) and value in by_id and value not in seen:
                        seen.add(value)
                        explicit.append(by_id[value])
        listed = {relationship.rel_id for relationship in explicit}
        implicit = sorted(
            (relationship for relationship in relationships if relationship.rel_id not in listed),
            key=lambda relationship: (
                relationship.rel_type,
                relationship.external,
                relationship.target
                if relationship.external
                else self._shallow_digest(relationship.target),
            ),
        )
        return explicit + implicit

    def _shallow_digest(self, name: str) -> str:
        """Digest independent of every other part's position, used only to order ties."""
        if not name:
            return ""
        cached = self._shallow_cache.get(name)
        if cached is None:
            if self.is_xml(name):
                types = {
                    relationship.rel_id: short_type(relationship.rel_type)
                    for relationship in self.relationships(name)
                }
                cached = sha256_bytes(
                    repr(
                        element_signature(
                            self.root(name),
                            lambda value: ("rel", types.get(value, "dangling")),
                        )
                    ).encode()
                )
            else:
                cached = sha256_bytes(self.parts[name])
            self._shallow_cache[name] = cached
        return cached

    # Part access ---------------------------------------------------------------

    def content_type(self, name: str) -> str:
        override = self._content_types[0].get("/" + name.lower())
        if override is not None:
            return override
        extension = posixpath.splitext(name)[1].lstrip(".").lower()
        return self._content_types[1].get(extension, "")

    def is_xml(self, name: str) -> bool:
        content_type = self.content_type(name).lower()
        if content_type:
            return content_type.endswith(("+xml", "/xml"))
        return name.lower().endswith((".xml", ".rels"))

    def root(self, name: str) -> ET.Element:
        cached = self._root_cache.get(name)
        if cached is None:
            cached = semantic_root(self.parts[name])
            self._root_cache[name] = cached
        return cached

    def data(self, key: str) -> bytes:
        return self.parts[self.name_of[key]]

    def part_keys(self) -> list[str]:
        return [key for key in self.name_of if key != ROOT]

    def keys_of_type(self, kind: str) -> list[str]:
        """Parts reached through a relationship of the given short type."""
        return [
            key
            for key in self.part_keys()
            if any(short_type(rel_type) == kind for _, rel_type in self.incoming.get(key, []))
        ]

    def target(self, key: str, rel_id: str) -> tuple[str, str | None]:
        """Resolve a relationship id of a part to (type, target part name)."""
        source = self.name_of[key]
        for relationship in self.relationships(source):
            if relationship.rel_id == rel_id:
                return relationship.rel_type, (
                    None
                    if relationship.external or not relationship.target
                    else relationship.target
                )
        raise ValueError(f"relationship {rel_id!r} is missing")

    def related(self, key: str, kind: str) -> list[str]:
        """Targets of a part's relationships of one short type, in canonical order."""
        return [
            target
            for rel_type, target, name in self.edges.get(key, [])
            if name is not None and short_type(rel_type) == kind
        ]

    def main_part(self) -> str:
        documents = self.related(ROOT, "officeDocument")
        if len(documents) != 1:
            raise ValueError("package must have exactly one main document relationship")
        return documents[0]

    # Semantic views --------------------------------------------------------------

    def target_token(self, rel_type: str, target_key: str, target_name: str | None) -> str:
        if target_name is None:
            return target_key
        kind = short_type(rel_type)
        if kind in MEDIA_TYPES:
            return "media:" + sha256_bytes(self.parts[target_name])
        if kind not in _KNOWN_TYPES:
            return "opaque:" + self.content_digest(target_name)
        return target_key

    def replace_root(self, name: str, root: ET.Element) -> None:
        """Substitute a part's parsed content, e.g. to model an expected edit."""
        self._root_cache[name] = root
        self._shallow_cache.clear()

    def resolver(self, key: str) -> Callable[[str], object]:
        """Map a part's relationship ids to what they reach (type and target identity)."""
        tokens: dict[str, object] = {}
        for relationship in self.relationships(self.name_of[key]):
            if relationship.external:
                target = "external:" + relationship.target
            elif relationship.target:
                target = self.target_token(
                    relationship.rel_type,
                    self.key_of[relationship.target],
                    relationship.target,
                )
            else:
                target = "missing"
            tokens[relationship.rel_id] = ("rel", relationship.rel_type, target)
        return lambda value: tokens.get(value, ("rel", "dangling"))

    def element_signature(self, key: str, element: ET.Element) -> object:
        """Signature of one element of a part, with relationship ids resolved."""
        return element_signature(element, self.resolver(key))

    def signature(
        self, key: str, transform: Callable[[ET.Element], ET.Element] | None = None
    ) -> object:
        """Part content with relationship-id attributes replaced by what they reach."""
        name = self.name_of[key]
        if not self.is_xml(name):
            return ("binary", sha256_bytes(self.parts[name]))
        root = self.root(name) if transform is None else transform(self.root(name))
        kinds = {short_type(rel_type) for _, rel_type in self.incoming.get(key, [])}
        if _CORE_PROPERTIES_TYPE in kinds:
            # An empty property and an absent property are the same to a reader.
            root = _filtered(
                root,
                lambda element: (
                    _local(element.tag) not in _SAVE_TIME_CORE
                    and bool((element.text or "").strip() or len(element))
                ),
            )
        elif _EXTENDED_PROPERTIES_TYPE in kinds:
            root = _filtered(
                root,
                lambda element: (
                    _local(element.tag) in _USER_EXTENDED and bool((element.text or "").strip())
                ),
                shallow=True,
            )
        return element_signature(root, self.resolver(key))

    def outgoing(self, key: str) -> list[tuple[str, str]]:
        """Relationships as a set of (type, what the target is); rIds never matter."""
        return sorted(
            {
                (rel_type, self.target_token(rel_type, target, name))
                for rel_type, target, name in self.edges.get(key, [])
                if short_type(rel_type) != THUMBNAIL_TYPE
            }
        )

    def node(self, key: str) -> tuple[str, object, list[tuple[str, str]]]:
        return (
            self.content_type(self.name_of[key]),
            self.signature(key),
            self.outgoing(key),
        )

    def content_digest(self, name: str) -> str:
        """Content hash that survives XML re-encoding but not content change."""
        if self.is_xml(name):
            return sha256_bytes(repr(element_signature(self.root(name))).encode())
        return sha256_bytes(self.parts[name])

    def is_save_time(self, key: str) -> bool:
        return any(
            short_type(rel_type) == THUMBNAIL_TYPE for _, rel_type in self.incoming.get(key, [])
        )

    def is_opaque(self, key: str) -> bool:
        return any(
            short_type(rel_type) not in _KNOWN_TYPES for _, rel_type in self.incoming.get(key, [])
        )

    def orphans(self) -> list[str]:
        """Parts no relationship reaches (excluding package plumbing)."""
        return sorted(
            name
            for name in self.parts
            if name not in self.key_of
            and name != "[Content_Types].xml"
            and not name.lower().endswith(".rels")
        )

    def opaque_parts(self, *, exclude: Iterable[str] = ()) -> list[tuple[object, str]]:
        """Unknown parts keyed by the relationship types that reach them and content."""
        excluded = set(exclude)
        result: list[tuple[object, str]] = [
            (
                tuple(sorted(self.incoming.get(key, []))),
                self.content_digest(self.name_of[key]),
            )
            for key in self.part_keys()
            if key not in excluded and self.is_opaque(key)
        ]
        result.extend(((), self.content_digest(name)) for name in self.orphans())
        return sorted(result, key=repr)

    def media(self, *, exclude: Iterable[str] = ()) -> list[tuple[str, str, str]]:
        """Media references as (owner, relationship type, content hash)."""
        excluded = set(exclude)
        return sorted(
            {
                (key, rel_type, sha256_bytes(self.parts[name]))
                for key, edges in self.edges.items()
                if key not in excluded
                for rel_type, _, name in edges
                if name is not None and short_type(rel_type) in MEDIA_TYPES
            }
        )

    def descendants(self, kinds: set[str]) -> set[str]:
        """Parts whose canonical path passes through a relationship of these types."""
        return {
            key
            for key in self.part_keys()
            if any(segment.rstrip("0123456789") in kinds for segment in key.split("/")[1:])
        }


def _content_types(parts: dict[str, bytes]) -> tuple[dict[str, str], dict[str, str]]:
    root = xml_root(parts["[Content_Types].xml"])
    overrides: dict[str, str] = {}
    defaults: dict[str, str] = {}
    for element in root:
        if element.tag == _CT + "Override":
            overrides[element.attrib.get("PartName", "").lower()] = element.attrib.get(
                "ContentType", ""
            )
        elif element.tag == _CT + "Default":
            defaults[element.attrib.get("Extension", "").lower()] = element.attrib.get(
                "ContentType", ""
            )
    return overrides, defaults


def _filtered(
    root: ET.Element,
    keep: Callable[[ET.Element], bool],
    *,
    shallow: bool = False,
) -> ET.Element:
    copy = ET.Element(root.tag, root.attrib)
    copy.text = root.text
    for child in root:
        if keep(child):
            copy.append(child if shallow else _filtered(child, keep))
    return copy


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]
