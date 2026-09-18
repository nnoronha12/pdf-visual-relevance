from __future__ import annotations

import hashlib
from collections import Counter
from typing import Any

import fitz

from .models import AnalysisResponse, PageResult
from .visual_scoring import Region, cluster_drawings, normalize_rect, rect_area, score_page


class PDFAnalysisError(ValueError):
    pass


def _digest(info: dict[str, Any]) -> str:
    digest = info.get("digest")
    if isinstance(digest, bytes):
        return digest.hex()
    if digest:
        return str(digest)
    return hashlib.sha256(f'{info.get("xref")}:{info.get("width")}:{info.get("height")}'.encode()).hexdigest()


def _asset_key(info: dict[str, Any], box: tuple[float, float, float, float]) -> tuple[Any, ...]:
    # Position is deliberately coarse so minor producer rounding still matches.
    return (_digest(info), *(round(v / 0.025) for v in box))


def analyze_pdf(data: bytes, min_score: float = 0.55, include_debug: bool = False) -> AnalysisResponse:
    try:
        doc = fitz.open(stream=data, filetype="pdf")
    except Exception as exc:
        raise PDFAnalysisError("The uploaded file is not a valid or readable PDF") from exc
    try:
        if doc.needs_pass:
            raise PDFAnalysisError("Password-protected PDFs are not supported")
        if doc.page_count == 0:
            raise PDFAnalysisError("The PDF contains no pages")
        page_images: list[list[dict[str, Any]]] = []
        asset_pages: dict[tuple[Any, ...], set[int]] = {}
        for page_index, page in enumerate(doc):
            infos = page.get_image_info(hashes=True, xrefs=True)
            page_images.append(infos)
            for info in infos:
                box = normalize_rect(fitz.Rect(info["bbox"]), page.rect)
                asset_pages.setdefault(_asset_key(info, box), set()).add(page_index)

        repeat_threshold = max(3, int(doc.page_count * 0.6 + 0.999))
        repeated = {key for key, pages in asset_pages.items() if len(pages) >= repeat_threshold}
        results: list[PageResult] = []

        for page_index, page in enumerate(doc):
            text_dict = page.get_text("dict")
            text_blocks = [b for b in text_dict.get("blocks", []) if b.get("type") == 0]
            text = page.get_text("text")
            text_chars = len("".join(text.split()))
            raster_regions: list[Region] = []
            ignored: list[dict[str, Any]] = []
            for info in page_images[page_index]:
                box = normalize_rect(fitz.Rect(info["bbox"]), page.rect)
                area = rect_area(box)
                key = _asset_key(info, box)
                in_margin = box[3] <= 0.11 or box[1] >= 0.89
                reason = None
                pixel_count = int(info.get("width") or 0) * int(info.get("height") or 0)
                if area > 0.5 and pixel_count < 5000:
                    reason = "stretched_low_resolution_background"
                elif area > 0.85 and text_chars >= 700:
                    reason = "likely_decorative_page_background"
                elif key in repeated and area < 0.15:
                    reason = "repeated_branding"
                elif area < 0.025:
                    reason = "tiny"
                elif in_margin and area < 0.12:
                    reason = "header_or_footer"
                if reason:
                    ignored.append({"bbox": [round(v, 4) for v in box], "kind": "raster", "reason": reason, "area_ratio": round(area, 4)})
                else:
                    raster_regions.append(Region(box, "raster", area, meta={"width": info.get("width"), "height": info.get("height")}))

            try:
                drawings = page.get_drawings()
            except Exception:
                drawings = []
            vector_regions, vector_ignored = cluster_drawings(drawings, page.rect)
            score, visual_area, large_count, reason, score_debug = score_page(
                raster_regions, vector_regions, len(ignored) + sum(vector_ignored.values()), len(page_images[page_index]), text_chars, len(text_blocks)
            )
            debug = None
            if include_debug:
                debug = {
                    "raster_regions": [{"bbox": [round(v, 4) for v in r.bbox], "area_ratio": round(r.area, 4), **r.meta} for r in raster_regions],
                    "vector_regions": [{"bbox": [round(v, 4) for v in r.bbox], "area_ratio": round(r.area, 4), **r.meta} for r in vector_regions],
                    "ignored_regions": ignored,
                    "ignored_vector_counts": vector_ignored,
                    **score_debug,
                }
            results.append(PageResult(page=page_index + 1, relevant=score >= min_score, score=score, reason=reason, visual_area_ratio=visual_area, image_count=len(page_images[page_index]), large_visual_count=large_count, debug=debug))

        relevant_pages = [p.page for p in results if p.relevant]
        return AnalysisResponse(page_count=doc.page_count, relevant_pages=relevant_pages, pages=results)
    except PDFAnalysisError:
        raise
    except Exception as exc:
        raise PDFAnalysisError("The PDF is malformed or could not be analyzed") from exc
    finally:
        doc.close()
