"""
pptx_optimizer.py
-----------------
Removes unused slide masters and unused slide layouts from a .pptx file.

Strategy
--------
A .pptx is a ZIP archive.  We:

1. Extract everything to a temporary directory.
2. Use the analyser to determine used / unused parts.
3. Remove unused XML part files from the extracted tree.
4. Remove the corresponding <Relationship> entries from every relevant
   .rels file.
5. Remove references to unused masters from presentation.xml
   (the <p:sldMasterIdLst> element).
6. For each *kept* master, remove layout references that belong to
   *unused* layouts from that master's own .rels and from the
   <p:sldLayoutIdLst> element inside the master XML.
7. Repack the directory tree into a new ZIP / .pptx file.
8. Return structured result metadata.

Safety rules
------------
* Never remove a part that is still referenced by a kept part.
* When in doubt, keep the file.
* We work on a copy; the original is never modified.

Open XML namespaces used
------------------------
* p  = http://schemas.openxmlformats.org/presentationml/2006/main
* r  = http://schemas.openxmlformats.org/officeDocument/2006/relationships
* Relationship elements live in the RELS_NS namespace (package relationships).
"""

from __future__ import annotations

import logging
import os
import posixpath
import shutil
import zipfile
from dataclasses import dataclass

from lxml import etree

from app.services.pptx_analyzer import PptxStructure, analyze_pptx
from app.utils.file_utils import make_temp_dir, remove_temp_dir
from app.utils.xml_utils import (
    RELS_NS,
    REL_TYPE_SLIDE_LAYOUT,
    REL_TYPE_SLIDE_MASTER,
    get_relationships,
    parse_xml_bytes,
    rels_path_for,
    remove_relationship_by_target,
    resolve_target,
    serialize_xml,
)

logger = logging.getLogger(__name__)

# Presentation namespace
PRES_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


# ---------------------------------------------------------------------------
# Result dataclass
# ---------------------------------------------------------------------------


@dataclass
class OptimizationResult:
    """Metadata returned after a successful optimization run."""

    output_path: str
    original_size: int
    optimized_size: int
    removed_layouts: int
    removed_masters: int


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _read_file(base_dir: str, rel_path: str) -> bytes:
    full = os.path.join(base_dir, rel_path.replace("/", os.sep))
    with open(full, "rb") as fh:
        return fh.read()


def _write_file(base_dir: str, rel_path: str, data: bytes) -> None:
    full = os.path.join(base_dir, rel_path.replace("/", os.sep))
    with open(full, "wb") as fh:
        fh.write(data)


def _delete_file(base_dir: str, rel_path: str) -> None:
    full = os.path.join(base_dir, rel_path.replace("/", os.sep))
    if os.path.isfile(full):
        os.remove(full)
        logger.info("Deleted part: %s", rel_path)
    else:
        logger.warning("Tried to delete non-existent file: %s", rel_path)


def _file_exists(base_dir: str, rel_path: str) -> bool:
    return os.path.isfile(os.path.join(base_dir, rel_path.replace("/", os.sep)))


def _parse_rels(base_dir: str, part_path: str) -> etree._Element | None:
    rels_path = rels_path_for(part_path)
    if not _file_exists(base_dir, rels_path):
        return None
    data = _read_file(base_dir, rels_path)
    return parse_xml_bytes(data)


def _save_rels(base_dir: str, part_path: str, root: etree._Element) -> None:
    rels_path = rels_path_for(part_path)
    _write_file(base_dir, rels_path, serialize_xml(root))


def _relative_target(from_path: str, to_path: str) -> str:
    """
    Compute the relative path from the directory containing *from_path*
    to *to_path*.

    Used when we need to rebuild Target attributes in .rels files.
    """
    from_dir = posixpath.dirname(from_path)
    rel = posixpath.relpath(to_path, from_dir)
    return rel


# ---------------------------------------------------------------------------
# Step: remove unused parts from presentation.xml relationships
# ---------------------------------------------------------------------------


def _remove_masters_from_presentation_rels(
    base_dir: str,
    presentation_path: str,
    unused_masters: set[str],
) -> None:
    """
    Remove <Relationship> entries for *unused_masters* from the
    presentation's .rels file.
    """
    rels_root = _parse_rels(base_dir, presentation_path)
    if rels_root is None:
        logger.error("Presentation .rels not found – skipping master rels cleanup")
        return

    for master_path in unused_masters:
        target_attr = _relative_target(presentation_path, master_path)
        removed = remove_relationship_by_target(rels_root, target_attr)
        if not removed:
            # Try the raw path as stored in the file
            for rel in rels_root.findall(f"{{{RELS_NS}}}Relationship"):
                if rel.get("Type") == REL_TYPE_SLIDE_MASTER:
                    stored = rel.get("Target", "")
                    resolved = resolve_target(presentation_path, stored)
                    if resolved == master_path:
                        rels_root.remove(rel)
                        logger.info(
                            "Removed master rel (alt match) for %s", master_path
                        )
                        removed = True
                        break
            if not removed:
                logger.warning(
                    "Could not remove master rel for %s from presentation rels",
                    master_path,
                )

    _save_rels(base_dir, presentation_path, rels_root)


def _remove_masters_from_presentation_xml(
    base_dir: str,
    presentation_path: str,
    unused_masters: set[str],
    structure: PptxStructure,
) -> None:
    """
    Remove <p:sldMasterId> entries in presentation.xml that reference
    an unused master.

    The presentation.xml contains::

        <p:sldMasterIdLst>
          <p:sldMasterId id="..." r:id="rIdN"/>
          ...
        </p:sldMasterIdLst>

    We look up each ``r:id`` in the presentation rels to get the target
    path, then remove it if it is unused.
    """
    pres_data = _read_file(base_dir, presentation_path)
    pres_root = parse_xml_bytes(pres_data)

    # Build rId → resolved_path map from the (already-updated) rels file
    rels_root = _parse_rels(base_dir, presentation_path)
    rid_to_path: dict[str, str] = {}
    if rels_root is not None:
        for rel in get_relationships(rels_root, REL_TYPE_SLIDE_MASTER):
            rid = rel.get("Id", "")
            target = rel.get("Target", "")
            rid_to_path[rid] = resolve_target(presentation_path, target)

    # Also build from the *original* rels before we removed entries,
    # but at this point they've already been removed, so we re-read from
    # the data on disk (which includes original + removed).  We need to
    # handle masters that no longer have a rels entry.
    # Build a complete rId → path map from the original rels in memory.
    # Since _remove_masters_from_presentation_rels already ran, we must
    # use the structure's all_masters to cross-check.

    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    sld_master_id_lst = pres_root.find(f"{{{PRES_NS}}}sldMasterIdLst")
    if sld_master_id_lst is None:
        logger.warning("No sldMasterIdLst found in presentation.xml")
        return

    to_remove = []
    for sld_master_id_elem in sld_master_id_lst:
        rid = sld_master_id_elem.get(f"{{{r_ns}}}id")
        if rid is None:
            continue
        resolved = rid_to_path.get(rid)
        if resolved is None:
            # This rId was removed from rels – it must be an unused master
            to_remove.append(sld_master_id_elem)
            logger.info(
                "Removing sldMasterId element with rId=%s (no longer in rels)", rid
            )
        elif resolved in unused_masters:
            to_remove.append(sld_master_id_elem)
            logger.info(
                "Removing sldMasterId element for unused master %s", resolved
            )

    for elem in to_remove:
        sld_master_id_lst.remove(elem)

    _write_file(base_dir, presentation_path, serialize_xml(pres_root))


# ---------------------------------------------------------------------------
# Step: remove unused layouts from a kept master
# ---------------------------------------------------------------------------


def _remove_layouts_from_master_rels(
    base_dir: str,
    master_path: str,
    unused_layouts: set[str],
) -> None:
    """
    Remove <Relationship> entries for *unused_layouts* from a master's
    .rels file.
    """
    rels_root = _parse_rels(base_dir, master_path)
    if rels_root is None:
        return

    for rel in list(get_relationships(rels_root, REL_TYPE_SLIDE_LAYOUT)):
        target = rel.get("Target", "")
        resolved = resolve_target(master_path, target)
        if resolved in unused_layouts:
            rels_root.remove(rel)
            logger.info(
                "Removed layout rel %s from master %s", resolved, master_path
            )

    _save_rels(base_dir, master_path, rels_root)


def _remove_layouts_from_master_xml(
    base_dir: str,
    master_path: str,
    unused_layouts: set[str],
) -> None:
    """
    Remove <p:sldLayoutId> entries that reference *unused_layouts* from
    the master XML's ``<p:sldLayoutIdLst>`` element.
    """
    master_data = _read_file(base_dir, master_path)
    master_root = parse_xml_bytes(master_data)

    # Rebuild rId → layout path from the *updated* rels file
    rels_root = _parse_rels(base_dir, master_path)
    rid_to_layout: dict[str, str] = {}
    if rels_root is not None:
        for rel in get_relationships(rels_root, REL_TYPE_SLIDE_LAYOUT):
            rid = rel.get("Id", "")
            target = rel.get("Target", "")
            rid_to_layout[rid] = resolve_target(master_path, target)

    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    sld_layout_id_lst = master_root.find(f"{{{PRES_NS}}}sldLayoutIdLst")
    if sld_layout_id_lst is None:
        logger.debug("No sldLayoutIdLst in master %s – nothing to clean", master_path)
        return

    to_remove = []
    for sld_layout_id_elem in sld_layout_id_lst:
        rid = sld_layout_id_elem.get(f"{{{r_ns}}}id")
        if rid is None:
            continue
        layout_path = rid_to_layout.get(rid)
        if layout_path is None:
            # rId was removed → belongs to an unused layout
            to_remove.append(sld_layout_id_elem)
            logger.info(
                "Removing sldLayoutId element with rId=%s (no longer in rels)", rid
            )
        elif layout_path in unused_layouts:
            to_remove.append(sld_layout_id_elem)
            logger.info(
                "Removing sldLayoutId element for unused layout %s", layout_path
            )

    for elem in to_remove:
        sld_layout_id_lst.remove(elem)

    _write_file(base_dir, master_path, serialize_xml(master_root))


# ---------------------------------------------------------------------------
# Step: repack directory into .pptx (ZIP)
# ---------------------------------------------------------------------------


def _repack_pptx(base_dir: str, output_path: str) -> None:
    """
    Walk *base_dir* and pack every file into a new ZIP at *output_path*.

    The [Content_Types].xml and _rels/.rels files must be present; other
    parts are included recursively.  We use ``ZIP_DEFLATED`` for
    compression, which is compatible with all PowerPoint versions.
    """
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for dir_path, _dir_names, filenames in os.walk(base_dir):
            for filename in filenames:
                abs_path = os.path.join(dir_path, filename)
                arc_name = os.path.relpath(abs_path, base_dir).replace(os.sep, "/")
                zf.write(abs_path, arc_name)
    logger.info("Repacked %s → %s", base_dir, output_path)


# ---------------------------------------------------------------------------
# Step: remove unused parts from [Content_Types].xml
# ---------------------------------------------------------------------------


def _remove_content_types(base_dir: str, removed_parts: set[str]) -> None:
    """
    Remove <Override> entries from ``[Content_Types].xml`` for parts that
    have been physically deleted.

    The content-types file maps part names to MIME types.  Keeping stale
    entries for deleted parts would produce a corrupt package.
    """
    ct_path = os.path.join(base_dir, "[Content_Types].xml")
    if not os.path.isfile(ct_path):
        logger.warning("[Content_Types].xml not found – skipping content-type cleanup")
        return

    with open(ct_path, "rb") as fh:
        ct_root = parse_xml_bytes(fh.read())

    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    to_remove = []
    for override in ct_root.findall(f"{{{ct_ns}}}Override"):
        part_name = override.get("PartName", "")
        # PartName is absolute: "/ppt/slides/slide1.xml"
        if part_name.startswith("/"):
            part_name_rel = part_name[1:]
        else:
            part_name_rel = part_name
        if part_name_rel in removed_parts:
            to_remove.append(override)
            logger.info("Removing content-type entry for %s", part_name_rel)

    for elem in to_remove:
        ct_root.remove(elem)

    with open(ct_path, "wb") as fh:
        fh.write(serialize_xml(ct_root))


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def optimize_pptx(input_path: str, output_path: str) -> OptimizationResult:
    """
    Optimise the .pptx file at *input_path* and write the result to
    *output_path*.

    Parameters
    ----------
    input_path:
        Absolute path to the source .pptx file.
    output_path:
        Absolute path where the optimised .pptx should be written.

    Returns
    -------
    OptimizationResult
        Metadata describing what was removed.

    Raises
    ------
    ValueError
        If the file is not a valid ZIP / .pptx archive.
    RuntimeError
        If a critical structural error is encountered.
    """
    original_size = os.path.getsize(input_path)

    # ------------------------------------------------------------------ #
    # 1. Extract the .pptx to a temp directory                            #
    # ------------------------------------------------------------------ #
    work_dir = make_temp_dir()
    try:
        try:
            with zipfile.ZipFile(input_path, "r") as zf:
                zf.extractall(work_dir)
        except zipfile.BadZipFile as exc:
            raise ValueError(f"The uploaded file is not a valid .pptx archive: {exc}")

        # ------------------------------------------------------------------ #
        # 2. Analyse the package structure                                    #
        # ------------------------------------------------------------------ #
        structure: PptxStructure = analyze_pptx(work_dir)

        unused_layouts = structure.all_layouts - structure.used_layouts
        unused_masters = structure.all_masters - structure.used_masters

        logger.info(
            "Unused layouts: %d, unused masters: %d",
            len(unused_layouts),
            len(unused_masters),
        )

        removed_parts: set[str] = set()

        # ------------------------------------------------------------------ #
        # 3. Remove unused layout files + clean up rels in kept masters       #
        # ------------------------------------------------------------------ #
        for layout_path in unused_layouts:
            # Remove the layout XML part itself
            _delete_file(work_dir, layout_path)
            removed_parts.add(layout_path)

            # Also delete the layout's own .rels file if it exists
            layout_rels_path = rels_path_for(layout_path)
            if _file_exists(work_dir, layout_rels_path):
                _delete_file(work_dir, layout_rels_path)

        # For every *kept* master, remove rels + XML refs to unused layouts
        for master_path in structure.used_masters:
            _remove_layouts_from_master_rels(work_dir, master_path, unused_layouts)
            _remove_layouts_from_master_xml(work_dir, master_path, unused_layouts)

        # ------------------------------------------------------------------ #
        # 4. Remove unused master files                                       #
        # ------------------------------------------------------------------ #
        for master_path in unused_masters:
            _delete_file(work_dir, master_path)
            removed_parts.add(master_path)

            master_rels_path = rels_path_for(master_path)
            if _file_exists(work_dir, master_rels_path):
                _delete_file(work_dir, master_rels_path)

        # ------------------------------------------------------------------ #
        # 5. Remove stale entries from presentation.xml and its rels          #
        # ------------------------------------------------------------------ #
        if unused_masters:
            _remove_masters_from_presentation_rels(
                work_dir, structure.presentation_path, unused_masters
            )
            _remove_masters_from_presentation_xml(
                work_dir, structure.presentation_path, unused_masters, structure
            )

        # ------------------------------------------------------------------ #
        # 6. Clean [Content_Types].xml                                        #
        # ------------------------------------------------------------------ #
        _remove_content_types(work_dir, removed_parts)

        # ------------------------------------------------------------------ #
        # 7. Repack into a new .pptx                                          #
        # ------------------------------------------------------------------ #
        _repack_pptx(work_dir, output_path)

    finally:
        remove_temp_dir(work_dir)

    optimized_size = os.path.getsize(output_path)

    return OptimizationResult(
        output_path=output_path,
        original_size=original_size,
        optimized_size=optimized_size,
        removed_layouts=len(unused_layouts),
        removed_masters=len(unused_masters),
    )
