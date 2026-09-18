from __future__ import annotations

from typing import Any

import fitz

from .models import ContentExtractionResponse, ExtractedPage, ExtractedTable
from .pdf_analyzer import PDFAnalysisError
from .visual_scoring import normalize_rect


def _clean_cell(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).replace("\x00", "").strip()
    return text or None


def extract_content(data: bytes) -> ContentExtractionResponse:
    """Extract native page text and tables without OCR or an LLM."""
    try:
        doc = fitz.open(stream=data, filetype="pdf")
    except Exception as exc:
        raise PDFAnalysisError("The uploaded file is not a valid or readable PDF") from exc

    try:
        if doc.needs_pass:
            raise PDFAnalysisError("Password-protected PDFs are not supported")
        if doc.page_count == 0:
            raise PDFAnalysisError("The PDF contains no pages")
        pages: list[ExtractedPage] = []
        for page_number, page in enumerate(doc, start=1):
            text = page.get_text("text", sort=True).replace("\x00", "").strip()
            tables: list[ExtractedTable] = []
            try:
                finder = page.find_tables()
                for table_index, table in enumerate(finder.tables, start=1):
                    rows = [[_clean_cell(cell) for cell in row] for row in table.extract()]
                    column_count = max((len(row) for row in rows), default=0)
                    tables.append(
                        ExtractedTable(
                            table_index=table_index,
                            bbox=[round(value, 4) for value in normalize_rect(fitz.Rect(table.bbox), page.rect)],
                            row_count=len(rows),
                            column_count=column_count,
                            rows=rows,
                        )
                    )
            except Exception:
                # Some malformed drawing streams can defeat table recognition;
                # native text is still useful and should be returned.
                tables = []
            pages.append(ExtractedPage(page=page_number, text=text, tables=tables))

        return ContentExtractionResponse(page_count=doc.page_count, pages=pages)
    except PDFAnalysisError:
        raise
    except Exception as exc:
        raise PDFAnalysisError("The PDF is malformed or could not be extracted") from exc
    finally:
        doc.close()
