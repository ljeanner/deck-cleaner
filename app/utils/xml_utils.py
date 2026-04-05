"""
xml_utils.py
------------
Helper functions for parsing and manipulating Open XML content inside a
.pptx package.

A .pptx file is a ZIP archive whose parts are XML files.  Relationships
between parts are stored in companion *.rels files that live in a
`_rels/` sub-folder next to the part they describe.

Useful namespaces
-----------------
* Relationships  xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
* Presentation   xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
"""

from __future__ import annotations

import posixpath
from typing import Optional

from lxml import etree

# ---------------------------------------------------------------------------
# Namespace constants
# ---------------------------------------------------------------------------

RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
PRES_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
DRAW_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

# Relationship type fragments used throughout the Open XML spec
REL_TYPE_SLIDE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
)
REL_TYPE_SLIDE_LAYOUT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)
REL_TYPE_SLIDE_MASTER = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
)
REL_TYPE_SLIDE_LAYOUT_IN_MASTER = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)


def parse_xml_bytes(data: bytes) -> etree._Element:
    """Parse raw XML bytes and return the root element."""
    return etree.fromstring(data)


def serialize_xml(root: etree._Element) -> bytes:
    """Serialize an lxml element tree back to bytes with XML declaration."""
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def rels_path_for(part_path: str) -> str:
    """
    Given a part path such as ``ppt/slides/slide1.xml`` return the
    companion relationship file path:
    ``ppt/slides/_rels/slide1.xml.rels``.
    """
    directory = posixpath.dirname(part_path)
    filename = posixpath.basename(part_path)
    return posixpath.join(directory, "_rels", filename + ".rels")


def get_relationships(
    rels_root: etree._Element,
    rel_type: Optional[str] = None,
) -> list[etree._Element]:
    """
    Return all <Relationship> child elements from a .rels root element,
    optionally filtered by *rel_type* (an exact match on the ``Type``
    attribute).
    """
    relationships = rels_root.findall(f"{{{RELS_NS}}}Relationship")
    if rel_type is None:
        return relationships
    return [r for r in relationships if r.get("Type") == rel_type]


def resolve_target(base_path: str, target: str) -> str:
    """
    Resolve a relationship *target* (which may be relative) against the
    directory that contains the part at *base_path*.

    Example::

        resolve_target("ppt/presentation.xml", "../slides/slide1.xml")
        # → "ppt/slides/slide1.xml"
    """
    base_dir = posixpath.dirname(base_path)
    resolved = posixpath.normpath(posixpath.join(base_dir, target))
    # posixpath.normpath may add a leading "./" — strip it
    if resolved.startswith("./"):
        resolved = resolved[2:]
    return resolved


def remove_relationship_by_target(rels_root: etree._Element, target: str) -> bool:
    """
    Remove the <Relationship> element whose ``Target`` attribute equals
    *target* (after normalisation).  Returns ``True`` if an element was
    removed.
    """
    for rel in rels_root.findall(f"{{{RELS_NS}}}Relationship"):
        if posixpath.normpath(rel.get("Target", "")) == posixpath.normpath(target):
            rels_root.remove(rel)
            return True
    return False
