"""
main.py
-------
FastAPI application entry point for Deck Cleaner.

Serves:
    GET  /            → HTML frontend (Jinja2 template)
    POST /optimize    → Upload + optimise a .pptx file
    GET  /download/{filename} → Download an optimised .pptx file
    Static files under /static/
"""

from __future__ import annotations

import logging
import os
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app.routes.optimize import router as optimize_router
from app.utils.file_utils import ensure_dir

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s – %(message)s",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).parent
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"
UPLOADS_DIR = BASE_DIR.parent / "uploads"
OUTPUTS_DIR = BASE_DIR.parent / "outputs"

# ---------------------------------------------------------------------------
# Application factory
# ---------------------------------------------------------------------------

app = FastAPI(
    title="Deck Cleaner",
    description="Optimise .pptx files by removing unused slide masters and layouts.",
    version="1.0.0",
)

# Serve static assets (CSS, JS)
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# Register route blueprints
app.include_router(optimize_router)

# Jinja2 for the single-page frontend
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


# ---------------------------------------------------------------------------
# Startup / shutdown hooks
# ---------------------------------------------------------------------------


@app.on_event("startup")
async def on_startup() -> None:
    """Create required directories on first run."""
    ensure_dir(UPLOADS_DIR)
    ensure_dir(OUTPUTS_DIR)
    logger.info("Deck Cleaner started – uploads: %s, outputs: %s", UPLOADS_DIR, OUTPUTS_DIR)


@app.on_event("shutdown")
async def on_shutdown() -> None:
    logger.info("Deck Cleaner shutting down.")


# ---------------------------------------------------------------------------
# Frontend route
# ---------------------------------------------------------------------------


@app.get("/", response_class=HTMLResponse)
async def index(request: Request) -> HTMLResponse:
    """Serve the single-page frontend."""
    return templates.TemplateResponse("index.html", {"request": request})
