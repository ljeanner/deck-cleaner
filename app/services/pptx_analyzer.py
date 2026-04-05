"""
pptx_analyzer.py
----------------
Analyse the internal structure of an extracted .pptx package to
determine:

* which slide layout parts are actually *used* by at least one slide
* which slide master parts are actually *used* by at least one used layout

The analysis is purely read-only; it does not modify any files.

Open XML structure recap
------------------------
A .pptx ZIP archive has the following relevant structure::

    ppt/
        presentation.xml          ← lists slideMasterId / sldMasterId elements
        _rels/
            presentation.xml.rels ← Relationships for presentation.xml
        slides/
            slide1.xml
            slide2.xml
            ...
            _rels/
                slide1.xml.rels   ← each slide references exactly one layout
                slide2.xml.rels
                ...
        slideLayouts/
            slideLayout1.xml
            ...
            _rels/
                slideLayout1.xml.rels  ← each layout references its master
        slideMasters/
            slideMaster1.xml
            ...
            _rels/
                slideMaster1.xml.rels  ← each master references its layouts
"""

from __future__ import annotations

import logging
import os
import posixpath
from dataclasses import dataclass, field

from lxml import etree

from app.utils.xml_utils import (
    RELS_NS,
    REL_TYPE_SLIDE,
    REL_TYPE_SLIDE_LAYOUT,
    REL_TYPE_SLIDE_MASTER,
    get_relationships,
    parse_xml_bytes,
    rels_path_for,
    resolve_target,
)

logger = logging.getLogger(__name__)


@dataclass
class PptxStructure:
    """
    Holds the analysis results for a single .pptx package.

    All paths are relative to the root of the extracted package
    (e.g. ``"ppt/slides/slide1.xml"``).
    """

    # The canonical path of presentation.xml (almost always "ppt/presentation.xml")
    presentation_path: str = ""

    # All slide part paths in deck order
    slide_paths: list[str] = field(default_factory=list)

    # slide path → layout path it references
    slide_to_layout: dict[str, str] = field(default_factory=dict)

    # layout path → master path it references
    layout_to_master: dict[str, str] = field(default_factory=dict)

    # Set of layout paths that are referenced by at least one slide
    used_layouts: set[str] = field(default_factory=set)

    # Set of master paths that are referenced by at least one used layout
    used_masters: set[str] = field(default_factory=set)

    # All layout paths discovered in the package (used + unused)
    all_layouts: set[str] = field(default_factory=set)

    # All master paths discovered in the package (used + unused)
    all_masters: set[str] = field(default_factory=set)


def _read_file(base_dir: str, rel_path: str) -> bytes:
    """Read a file from the extracted package directory."""
    full_path = os.path.join(base_dir, rel_path.replace("/", os.sep))
    with open(full_path, "rb") as fh:
        return fh.read()


def _file_exists(base_dir: str, rel_path: str) -> bool:
    full_path = os.path.join(base_dir, rel_path.replace("/", os.sep))
    return os.path.isfile(full_path)


def _parse_rels(base_dir: str, part_path: str) -> etree._Element | None:
    """
    Parse the .rels file that accompanies *part_path*.
    Returns ``None`` when the rels file does not exist.
    """
    rels_path = rels_path_for(part_path)
    if not _file_exists(base_dir, rels_path):
        logger.debug("No rels file found for %s", part_path)
        return None
    data = _read_file(base_dir, rels_path)
    return parse_xml_bytes(data)


def _find_presentation_xml(base_dir: str) -> str:
    """
    Locate presentation.xml by reading the package-level _rels/.rels file.
    Falls back to ``ppt/presentation.xml`` if unavailable.
    """
    dot_rels = os.path.join(base_dir, "_rels", ".rels")
    if os.path.isfile(dot_rels):
        with open(dot_rels, "rb") as fh:
            root = parse_xml_bytes(fh.read())
        for rel in get_relationships(root):
            rel_type = rel.get("Type", "")
            if rel_type.endswith("/officeDocument"):
                target = rel.get("Target", "")
                # Target is relative to the package root
                if target.startswith("/"):
                    target = target[1:]
                logger.debug("Presentation part found: %s", target)
                return target
    # Fallback
    return "ppt/presentation.xml"


def _collect_slides(base_dir: str, presentation_path: str) -> list[str]:
    """
    Return slide part paths in deck order by reading the presentation
    relationships file.
    """
    rels_root = _parse_rels(base_dir, presentation_path)
    if rels_root is None:
        logger.warning("No presentation relationships file found.")
        return []

    slides: list[str] = []
    for rel in get_relationships(rels_root, REL_TYPE_SLIDE):
        target = rel.get("Target", "")
        resolved = resolve_target(presentation_path, target)
        slides.append(resolved)
        logger.debug("Slide found: %s", resolved)

    return slides


def _collect_masters_from_presentation(
    base_dir: str, presentation_path: str
) -> list[str]:
    """
    Return all slide master part paths listed in the presentation
    relationships file.
    """
    rels_root = _parse_rels(base_dir, presentation_path)
    if rels_root is None:
        return []

    masters: list[str] = []
    for rel in get_relationships(rels_root, REL_TYPE_SLIDE_MASTER):
        target = rel.get("Target", "")
        resolved = resolve_target(presentation_path, target)
        masters.append(resolved)
        logger.debug("Master found in presentation rels: %s", resolved)

    return masters


def _collect_layouts_from_master(base_dir: str, master_path: str) -> list[str]:
    """
    Return all slide layout part paths referenced by *master_path*.
    """
    rels_root = _parse_rels(base_dir, master_path)
    if rels_root is None:
        return []

    layouts: list[str] = []
    # Masters reference layouts with the same relationship type as slides
    # reference layouts (slideLayout type).
    for rel in get_relationships(rels_root, REL_TYPE_SLIDE_LAYOUT):
        target = rel.get("Target", "")
        resolved = resolve_target(master_path, target)
        layouts.append(resolved)
        logger.debug("Layout found in master %s: %s", master_path, resolved)

    return layouts


def _get_layout_for_slide(base_dir: str, slide_path: str) -> str | None:
    """
    Return the slide layout part path that *slide_path* references, or
    ``None`` if the relationship cannot be resolved.
    """
    rels_root = _parse_rels(base_dir, slide_path)
    if rels_root is None:
        return None

    for rel in get_relationships(rels_root, REL_TYPE_SLIDE_LAYOUT):
        target = rel.get("Target", "")
        resolved = resolve_target(slide_path, target)
        logger.debug("Slide %s → layout %s", slide_path, resolved)
        return resolved  # There is exactly one layout per slide

    logger.warning("No layout relationship found for slide %s", slide_path)
    return None


def _get_master_for_layout(base_dir: str, layout_path: str) -> str | None:
    """
    Return the slide master part path that *layout_path* references, or
    ``None`` if the relationship cannot be resolved.
    """
    rels_root = _parse_rels(base_dir, layout_path)
    if rels_root is None:
        return None

    for rel in get_relationships(rels_root, REL_TYPE_SLIDE_MASTER):
        target = rel.get("Target", "")
        resolved = resolve_target(layout_path, target)
        logger.debug("Layout %s → master %s", layout_path, resolved)
        return resolved  # There is exactly one master per layout

    logger.warning("No master relationship found for layout %s", layout_path)
    return None


def analyze_pptx(base_dir: str) -> PptxStructure:
    """
    Analyse the extracted .pptx package rooted at *base_dir* and return
    a :class:`PptxStructure` describing which masters and layouts are used.

    Parameters
    ----------
    base_dir:
        Path to the directory that was produced by extracting the .pptx ZIP.

    Returns
    -------
    PptxStructure
        Fully populated analysis result.
    """
    structure = PptxStructure()

    # ------------------------------------------------------------------ #
    # 1. Locate presentation.xml                                          #
    # ------------------------------------------------------------------ #
    structure.presentation_path = _find_presentation_xml(base_dir)
    logger.info("Presentation part: %s", structure.presentation_path)

    # ------------------------------------------------------------------ #
    # 2. Collect all slides (in deck order)                               #
    # ------------------------------------------------------------------ #
    structure.slide_paths = _collect_slides(base_dir, structure.presentation_path)
    logger.info("Slides found: %d", len(structure.slide_paths))

    # ------------------------------------------------------------------ #
    # 3. Collect *all* masters and the layouts they declare               #
    # ------------------------------------------------------------------ #
    all_master_paths = _collect_masters_from_presentation(
        base_dir, structure.presentation_path
    )
    for master_path in all_master_paths:
        structure.all_masters.add(master_path)
        for layout_path in _collect_layouts_from_master(base_dir, master_path):
            structure.all_layouts.add(layout_path)
            structure.layout_to_master[layout_path] = master_path

    logger.info(
        "Total masters: %d, total layouts: %d",
        len(structure.all_masters),
        len(structure.all_layouts),
    )

    # ------------------------------------------------------------------ #
    # 4. For each slide, determine which layout it uses                   #
    # ------------------------------------------------------------------ #
    for slide_path in structure.slide_paths:
        layout_path = _get_layout_for_slide(base_dir, slide_path)
        if layout_path is None:
            logger.warning(
                "Could not determine layout for slide %s – treating as safe",
                slide_path,
            )
            continue
        structure.slide_to_layout[slide_path] = layout_path
        structure.used_layouts.add(layout_path)

    logger.info("Used layouts: %d", len(structure.used_layouts))

    # ------------------------------------------------------------------ #
    # 5. For each used layout, determine which master it uses             #
    # ------------------------------------------------------------------ #
    for layout_path in structure.used_layouts:
        master_path = structure.layout_to_master.get(layout_path)
        if master_path is None:
            # Try to resolve directly from the layout's rels (belt-and-braces)
            master_path = _get_master_for_layout(base_dir, layout_path)
        if master_path is None:
            logger.warning(
                "Could not determine master for layout %s – treating as safe",
                layout_path,
            )
            continue
        structure.used_masters.add(master_path)

    logger.info("Used masters: %d", len(structure.used_masters))

    return structure
