"""Small deterministic helpers. Persisted paths are always run-root relative."""

from __future__ import annotations

import hashlib
import json
import os
import platform
import posixpath
import stat
import sys
import zipfile
from collections.abc import Iterable
from io import BytesIO
from pathlib import Path, PurePosixPath
from typing import Any, BinaryIO
from urllib.parse import unquote
from xml.etree import ElementTree as ET

_MAX_ZIP_MEMBERS = 4_096
_MAX_MEMBER_BYTES = 64 * 1024 * 1024
_MAX_PACKAGE_BYTES = 256 * 1024 * 1024
_XML_SUFFIXES = (".xml", ".rels")
_CONTENT_TYPES = "{http://schemas.openxmlformats.org/package/2006/content-types}Types"
_RELATIONSHIPS = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationships"
_PRESENTATION = "{http://schemas.openxmlformats.org/presentationml/2006/main}presentation"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def relative(path: Path, root: Path) -> str:
    return path.resolve().relative_to(root.resolve()).as_posix()


def json_dump(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def package_parts(path: Path) -> dict[str, bytes]:
    """Read a bounded, well-formed PPTX package without ZIP path ambiguity."""
    with path.open("rb") as source:
        return _package_parts(source, require_presentation=True)


def package_bytes(value: bytes, *, require_presentation: bool = False) -> dict[str, bytes]:
    """Read an embedded Open Packaging Convention archive under the same limits."""
    return _package_parts(BytesIO(value), require_presentation=require_presentation)


def xml_root(value: bytes) -> ET.Element:
    """Parse XML after rejecting constructs unsupported by safe ElementTree parsing."""
    lowered = value.lower()
    if b"<!doctype" in lowered or b"<!entity" in lowered:
        raise ValueError("XML DTDs and entity declarations are not permitted")
    try:
        return ET.fromstring(value)
    except ET.ParseError as exc:
        raise ValueError(f"malformed XML: {exc}") from exc


def _package_parts(source: BinaryIO, *, require_presentation: bool) -> dict[str, bytes]:
    with zipfile.ZipFile(source) as archive:
        infos = archive.infolist()
        if len(infos) > _MAX_ZIP_MEMBERS:
            raise ValueError("ZIP member limit exceeded")
        names: set[str] = set()
        total = 0
        parts: dict[str, bytes] = {}
        for info in infos:
            _validate_zip_member(info, names)
            if info.is_dir():
                continue
            if info.file_size > _MAX_MEMBER_BYTES:
                raise ValueError(f"ZIP member is too large: {info.filename}")
            total += info.file_size
            if total > _MAX_PACKAGE_BYTES:
                raise ValueError("ZIP uncompressed-size limit exceeded")
            data = archive.read(info)
            if len(data) != info.file_size:
                raise ValueError(f"truncated ZIP member: {info.filename}")
            parts[info.filename] = data
    _validate_opc(parts, require_presentation=require_presentation)
    return parts


def _validate_zip_member(info: zipfile.ZipInfo, names: set[str]) -> None:
    name = info.filename
    path = PurePosixPath(name)
    if not name or "\\" in name or path.is_absolute() or ".." in path.parts:
        raise ValueError(f"unsafe ZIP member path: {name!r}")
    if name in names:
        raise ValueError(f"duplicate ZIP member: {name}")
    names.add(name)
    if stat.S_ISLNK(info.external_attr >> 16):
        raise ValueError(f"symbolic-link ZIP member: {name}")


def _validate_opc(parts: dict[str, bytes], *, require_presentation: bool) -> None:
    for required in ("[Content_Types].xml", "_rels/.rels"):
        if required not in parts:
            raise ValueError(f"missing required package part: {required}")
    content_types = xml_root(parts["[Content_Types].xml"])
    if content_types.tag != _CONTENT_TYPES:
        raise ValueError("invalid [Content_Types].xml root")
    root_relationships = xml_root(parts["_rels/.rels"])
    if root_relationships.tag != _RELATIONSHIPS:
        raise ValueError("invalid package relationships root")
    if require_presentation:
        main = _main_document(root_relationships, parts)
        if main is None:
            raise ValueError("missing required PPTX presentation part")
        if xml_root(parts[main]).tag != _PRESENTATION:
            raise ValueError("invalid PPTX presentation root")
    for name, data in parts.items():
        if not name.endswith(_XML_SUFFIXES):
            continue
        root = xml_root(data)
        if name.endswith(".rels"):
            _validate_relationships(name, root)


def _main_document(relationships: ET.Element, parts: dict[str, bytes]) -> str | None:
    """The part the package's officeDocument relationship names, whatever it is called."""
    for relationship in relationships:
        if relationship.attrib.get("Type", "").endswith("/officeDocument"):
            target = unquote(relationship.attrib.get("Target", "")).lstrip("/")
            name = posixpath.normpath(target) if target else ""
            return name if name in parts else None
    return None


def _validate_relationships(name: str, root: ET.Element) -> None:
    if root.tag != _RELATIONSHIPS:
        raise ValueError(f"invalid relationships root: {name}")
    identifiers: set[str] = set()
    for relationship in root:
        if relationship.tag != _RELATIONSHIPS.removesuffix("Relationships") + "Relationship":
            raise ValueError(f"invalid relationship element: {name}")
        identifier = relationship.attrib.get("Id")
        if (
            not identifier
            or identifier in identifiers
            or not relationship.attrib.get("Type")
            or not relationship.attrib.get("Target")
        ):
            raise ValueError(f"malformed relationship: {name}")
        identifiers.add(identifier)


def environment_identity() -> dict[str, str]:
    return {
        "python": sys.version.split()[0],
        "implementation": platform.python_implementation(),
        "platform": platform.platform(),
        "executable": Path(sys.executable).name,
        "pid": str(os.getpid()),
    }


def stable_subset(parts: dict[str, bytes], prefixes: Iterable[str]) -> dict[str, str]:
    selected = tuple(prefixes)
    return {
        name: sha256_bytes(data)
        for name, data in sorted(parts.items())
        if name.startswith(selected)
    }
