"""
optimize.py
-----------
FastAPI router that exposes the .pptx optimization endpoints:

    POST /optimize   – Upload a .pptx, optimise it, return JSON summary.
    GET  /download/{filename} – Download a previously optimised file.
"""

from __future__ import annotations

import logging
import os
from pathlib import Path

from fastapi import APIRouter, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse

from app.services.pptx_optimizer import optimize_pptx
from app.utils.file_utils import (
    ensure_dir,
    generate_unique_filename,
    safe_filename,
)

logger = logging.getLogger(__name__)

router = APIRouter()

# Directories are created by main.py on startup; we just reference them here.
UPLOADS_DIR = Path("uploads")
OUTPUTS_DIR = Path("outputs")

MAX_FILE_SIZE = 200 * 1024 * 1024  # 200 MB safety cap


# ---------------------------------------------------------------------------
# POST /optimize
# ---------------------------------------------------------------------------


@router.post("/optimize")
async def optimize_endpoint(file: UploadFile) -> JSONResponse:
    """
    Accept a .pptx file, remove unused slide masters and layouts, and
    return a JSON summary::

        {
            "output_filename": "deck_optimized_3f2a1b.pptx",
            "original_size":  1234567,
            "optimized_size":  987654,
            "removed_layouts": 12,
            "removed_masters":  3
        }

    HTTP error codes
    ----------------
    400 – missing file, wrong extension, empty file, or file too large.
    422 – file is not a valid .pptx / ZIP archive.
    500 – unexpected server-side failure.
    """
    # ------------------------------------------------------------------ #
    # Validation                                                           #
    # ------------------------------------------------------------------ #
    if file.filename is None or file.filename == "":
        raise HTTPException(status_code=400, detail="No file provided.")

    clean_name = safe_filename(file.filename)

    if not clean_name.lower().endswith(".pptx"):
        raise HTTPException(
            status_code=400,
            detail="Only .pptx files are accepted.",
        )

    # Read the upload into memory first so we can size-check it.
    content = await file.read()

    if len(content) == 0:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")

    if len(content) > MAX_FILE_SIZE:
        raise HTTPException(
            status_code=400,
            detail=f"File exceeds the maximum allowed size of {MAX_FILE_SIZE // (1024*1024)} MB.",
        )

    # ------------------------------------------------------------------ #
    # Persist upload                                                       #
    # ------------------------------------------------------------------ #
    ensure_dir(UPLOADS_DIR)
    ensure_dir(OUTPUTS_DIR)

    upload_filename = generate_unique_filename(clean_name, suffix="_upload")
    upload_path = UPLOADS_DIR / upload_filename

    try:
        with open(upload_path, "wb") as fh:
            fh.write(content)
    except OSError as exc:
        logger.error("Failed to save upload: %s", exc)
        raise HTTPException(status_code=500, detail="Failed to save uploaded file.")

    # ------------------------------------------------------------------ #
    # Optimise                                                             #
    # ------------------------------------------------------------------ #
    output_filename = generate_unique_filename(clean_name)
    output_path = OUTPUTS_DIR / output_filename

    try:
        result = optimize_pptx(str(upload_path), str(output_path))
    except ValueError as exc:
        logger.warning("Invalid .pptx file: %s", exc)
        raise HTTPException(status_code=422, detail=str(exc))
    except Exception as exc:
        logger.exception("Unexpected error during optimization: %s", exc)
        raise HTTPException(
            status_code=500,
            detail="An unexpected error occurred while optimizing the file.",
        )
    finally:
        # Always clean up the upload regardless of outcome
        try:
            os.remove(upload_path)
        except OSError:
            pass

    return JSONResponse(
        content={
            "output_filename": output_filename,
            "original_size": result.original_size,
            "optimized_size": result.optimized_size,
            "removed_layouts": result.removed_layouts,
            "removed_masters": result.removed_masters,
        }
    )


# ---------------------------------------------------------------------------
# GET /download/{filename}
# ---------------------------------------------------------------------------


@router.get("/download/{filename}")
async def download_endpoint(filename: str) -> FileResponse:
    """
    Stream a previously optimised .pptx file back to the client.

    The *filename* must be a bare filename (no slashes), preventing
    path-traversal attacks.
    """
    # Guard against path traversal: sanitise the filename, then canonicalise
    # the resolved path and confirm it sits inside OUTPUTS_DIR.
    clean = safe_filename(filename)
    if clean != filename or "/" in filename or ".." in filename:
        raise HTTPException(status_code=400, detail="Invalid filename.")

    outputs_resolved = OUTPUTS_DIR.resolve()
    file_path = (outputs_resolved / clean).resolve()

    # Reject anything that would escape the outputs directory.
    try:
        file_path.relative_to(outputs_resolved)
    except ValueError:
        raise HTTPException(status_code=400, detail="Invalid filename.")

    if not file_path.is_file():
        raise HTTPException(status_code=404, detail="File not found.")

    return FileResponse(
        path=str(file_path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=clean,
    )
