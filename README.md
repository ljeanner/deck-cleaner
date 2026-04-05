# Deck Cleaner 🧹

A web application that optimises `.pptx` PowerPoint files by removing unused
slide masters and unused slide layouts, producing a smaller, cleaner file that
remains fully editable.

---

## Features

- Upload a `.pptx` file through a clean browser UI.
- The FastAPI backend analyses the Open XML structure to find every master and
  layout that is not referenced by any slide.
- Unused parts are removed from the ZIP package, along with their relationship
  entries and `[Content_Types].xml` registrations.
- The optimised file is made available for immediate download.
- A summary shows the original/optimised file size, number of layouts removed,
  and number of masters removed.

---

## Project structure

```
deck-cleaner/
├── app/
│   ├── main.py                 # FastAPI entry point
│   ├── routes/
│   │   └── optimize.py         # POST /optimize, GET /download/{filename}
│   ├── services/
│   │   ├── pptx_analyzer.py    # Read-only structural analysis
│   │   └── pptx_optimizer.py   # ZIP/XML manipulation + repacking
│   ├── utils/
│   │   ├── file_utils.py       # Unique names, temp dirs, size helpers
│   │   └── xml_utils.py        # lxml helpers, namespace constants
│   ├── static/
│   │   ├── style.css
│   │   └── app.js
│   └── templates/
│       └── index.html
├── uploads/                    # Transient upload storage (auto-created)
├── outputs/                    # Optimised file storage (auto-created)
├── requirements.txt
└── README.md
```

---

## Prerequisites

- Python 3.11 or later
- `pip`

---

## Quick start

### 1 – Clone the repository

```bash
git clone https://github.com/ljeanner/deck-cleaner.git
cd deck-cleaner
```

### 2 – Create and activate a virtual environment

```bash
# macOS / Linux
python3.11 -m venv .venv
source .venv/bin/activate

# Windows (PowerShell)
py -3.11 -m venv .venv
.venv\Scripts\Activate.ps1
```

### 3 – Install dependencies

```bash
pip install -r requirements.txt
```

### 4 – Run the FastAPI server

```bash
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

### 5 – Open the app in your browser

Navigate to **http://localhost:8000**

---

## API endpoints

| Method | Path                       | Description                                          |
|--------|----------------------------|------------------------------------------------------|
| `GET`  | `/`                        | Serves the HTML frontend                             |
| `POST` | `/optimize`                | Accepts a `.pptx`, returns JSON optimisation summary |
| `GET`  | `/download/{filename}`     | Downloads an optimised file                          |

### `POST /optimize` response body

```json
{
  "output_filename": "deck_optimized_3f2a1b.pptx",
  "original_size": 1234567,
  "optimized_size": 987654,
  "removed_layouts": 12,
  "removed_masters": 3
}
```

---

## How slide masters and layouts are detected and removed

A `.pptx` file is a ZIP archive of XML parts.  The relationships between parts
are declared in companion `.rels` files.

### Detection

1. **`ppt/presentation.xml`** lists every slide master via
   `<p:sldMasterIdLst>`.  Each master is resolved through the presentation's
   `.rels` file.

2. **Each slide master** (`ppt/slideMasters/slideMasterN.xml`) lists the slide
   layouts it owns via `<p:sldLayoutIdLst>`.  Each layout is resolved through
   the master's `.rels` file.

3. **Each slide** (`ppt/slides/slideN.xml`) carries a relationship of type
   `slideLayout` in its `.rels` file that points to exactly one layout.

4. A layout is **used** if at least one slide references it.
   A master is **used** if at least one *used* layout belongs to it.

### Removal

- Unused layout XML files are deleted from the package.
- The corresponding `<Relationship>` entries are removed from each kept
  master's `.rels` file.
- The `<p:sldLayoutId>` elements inside each kept master's XML are cleaned.
- Unused master XML files are deleted.
- Their `<Relationship>` entries are removed from `presentation.xml.rels`.
- The `<p:sldMasterId>` elements are removed from `presentation.xml`.
- Stale `<Override>` entries are removed from `[Content_Types].xml`.
- The cleaned directory tree is repacked into a valid `.pptx` ZIP.

The implementation never removes a part that is still referenced and errs on
the side of keeping files when any ambiguity is detected.
