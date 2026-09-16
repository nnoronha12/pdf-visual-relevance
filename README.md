# PDF Visual Relevance API

A deterministic, precision-oriented FastAPI microservice that shortlists PDF pages whose visual content is likely to add meaning beyond normal text extraction. It is designed as the inexpensive first stage of a RAG ingestion pipeline; it does not call a vision model.

## API

### Analyze PDF

```bash
curl -X POST \
  -F "file=@document.pdf" \
  https://DOMAIN/analyze-pdf
```

With a custom relevance threshold and compact debug metadata:

```bash
curl -X POST \
  -F "file=@document.pdf" \
  -F "min_score=0.60" \
  -F "include_debug=true" \
  https://DOMAIN/analyze-pdf
```

The JSON response exposes `relevant_pages` directly. Page numbers are one-based.

### Render pages

```bash
curl -X POST \
  -F "file=@document.pdf" \
  -F "page_numbers=3,6,9" \
  https://DOMAIN/render-pages \
  --output rendered_pages.zip
```

The ZIP contains `page_3.png`, `page_6.png`, and `page_9.png`, rendered at 144 DPI.

### Extract page text and tables

```bash
curl -X POST \
  -F "file=@document.pdf" \
  https://DOMAIN/extract-content
```

The JSON response contains one object per page. Each page includes its one-based page number, native extracted text, and any detected tables as normalized bounding boxes and arrays of rows/cells. This endpoint does not perform OCR, so scanned text remains the responsibility of the later vision/OCR stage.

### Convert Office files to PDF

```bash
curl -X POST \
  -F "file=@presentation.pptx" \
  https://DOMAIN/convert-to-pdf \
  --output presentation.pdf
```

The endpoint accepts `.docx`, `.ppt`, and `.pptx` files and returns an `application/pdf` response produced by headless LibreOffice. Inputs are stored only in a per-request temporary directory and removed immediately after conversion.

### Health

```bash
curl https://DOMAIN/health
```

## Scoring

The analyzer works in two passes. First it fingerprints every embedded raster image and records its normalized position and dimensions. A small asset occurring at approximately the same location on at least 60% of pages (and at least three pages) is treated as repeated branding. Second, each page is analyzed for meaningful raster regions and vector clusters.

Raster images below 2.5% of page area, small header/footer images, repeated small assets, low-resolution stretched backgrounds, and likely decorative full-page backgrounds behind abundant extractable text are ignored. Vector paths from `page.get_drawings()` are normalized, stripped of tiny marks, thin separators, and simple borders, then joined into nearby connected clusters. Clusters require both meaningful area and structural complexity. Regular orthogonal grids containing abundant extractable text are classified as ordinary tables and penalized.

Overlapping visual rectangles are combined on a 100 x 100 occupancy grid so overlap is not double-counted. The deterministic score is capped to `[0, 1]` and combines:

- 35% meaningful visual union area
- 30% strongest raster region
- 25% strongest vector cluster
- 10% multiple substantial regions
- 8% vector-path complexity
- 5% useful text/graphics balance
- positive gates for a raster over 15% or a complex vector cluster over 16%
- penalties for ordinary text-table grids and ignored noise

The exact formula and gates are documented beside `score_page()` in `app/visual_scoring.py`. The default threshold is `0.55`; raise it for still higher precision.

## Limits and error handling

- Maximum upload: 50 MB
- Maximum analysis length: 150 pages
- Maximum render request: 30 pages
- Office conversion timeout: 120 seconds
- Password-protected, empty, malformed, non-PDF, and out-of-range requests return clean JSON errors
- Uploaded data is held only for the request and is never persisted

## Local development and tests

```bash
python -m venv .venv
. .venv/bin/activate
pip install -r requirements-dev.txt
pytest -q
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

The generated fixture covers text-only content, a tiny logo, a large screenshot, a vector chart, a vector architecture diagram, a normal text table, and repeated header branding.

## n8n

Use an **HTTP Request** node with method `POST`, URL `https://DOMAIN/analyze-pdf`, and **Send Body** enabled. Choose **Form-Data**. Add a parameter of type **n8n Binary File**, name it `file`, and set **Input Data Field Name** (called **Binary Property** in older n8n versions) to `data`. Optional text parameters are `min_score=0.55` and `include_debug=false`. Keep response format as JSON.

Downstream expressions:

```text
{{$json.relevant_pages}}
{{$json.relevant_pages.join(",")}}
```
