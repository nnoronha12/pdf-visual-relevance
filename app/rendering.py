from __future__ import annotations

import io
import zipfile

import fitz

from .pdf_analyzer import PDFAnalysisError


def parse_page_numbers(value: str, page_count: int, max_render_pages: int = 30) -> list[int]:
    try:
        pages = list(dict.fromkeys(int(x.strip()) for x in value.split(",") if x.strip()))
    except ValueError as exc:
        raise PDFAnalysisError("page_numbers must be comma-separated integers") from exc
    if not pages:
        raise PDFAnalysisError("At least one page number is required")
    if len(pages) > max_render_pages:
        raise PDFAnalysisError(f"No more than {max_render_pages} pages may be rendered per request")
    if any(p < 1 or p > page_count for p in pages):
        raise PDFAnalysisError(f"Page numbers must be between 1 and {page_count}")
    return pages


def render_pages_zip(data: bytes, page_numbers: str, dpi: int = 144) -> bytes:
    try:
        doc = fitz.open(stream=data, filetype="pdf")
    except Exception as exc:
        raise PDFAnalysisError("The uploaded file is not a valid or readable PDF") from exc
    try:
        if doc.needs_pass:
            raise PDFAnalysisError("Password-protected PDFs are not supported")
        pages = parse_page_numbers(page_numbers, doc.page_count)
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for number in pages:
                try:
                    pixmap = doc[number - 1].get_pixmap(dpi=dpi, alpha=False)
                    archive.writestr(f"page_{number}.png", pixmap.tobytes("png"))
                except Exception as exc:
                    raise PDFAnalysisError(f"Failed to render page {number}") from exc
        return output.getvalue()
    finally:
        doc.close()
