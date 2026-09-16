from __future__ import annotations

from dataclasses import dataclass, field
from math import hypot
from typing import Any

import fitz


@dataclass
class Region:
    bbox: tuple[float, float, float, float]
    kind: str
    area: float
    weight: float = 1.0
    meta: dict[str, Any] = field(default_factory=dict)


def normalize_rect(rect: fitz.Rect, page_rect: fitz.Rect) -> tuple[float, float, float, float]:
    w, h = max(page_rect.width, 1), max(page_rect.height, 1)
    return (
        max(0.0, min(1.0, (rect.x0 - page_rect.x0) / w)),
        max(0.0, min(1.0, (rect.y0 - page_rect.y0) / h)),
        max(0.0, min(1.0, (rect.x1 - page_rect.x0) / w)),
        max(0.0, min(1.0, (rect.y1 - page_rect.y0) / h)),
    )


def rect_area(box: tuple[float, float, float, float]) -> float:
    return max(0.0, box[2] - box[0]) * max(0.0, box[3] - box[1])


def union_ratio(regions: list[Region], grid: int = 100) -> float:
    """Approximate rectangle union on a normalized grid, avoiding overlap inflation."""
    if not regions:
        return 0.0
    occupied = bytearray(grid * grid)
    for region in regions:
        x0, y0, x1, y1 = region.bbox
        ix0, iy0 = max(0, int(x0 * grid)), max(0, int(y0 * grid))
        ix1, iy1 = min(grid, max(ix0 + 1, int(x1 * grid + 0.999))), min(grid, max(iy0 + 1, int(y1 * grid + 0.999)))
        for y in range(iy0, iy1):
            start = y * grid + ix0
            occupied[start : y * grid + ix1] = b"\x01" * (ix1 - ix0)
    return sum(occupied) / float(grid * grid)


def _near(a: tuple[float, float, float, float], b: tuple[float, float, float, float], gap: float = 0.025) -> bool:
    return not (a[2] + gap < b[0] or b[2] + gap < a[0] or a[3] + gap < b[1] or b[3] + gap < a[1])


def _merge_boxes(boxes: list[tuple[float, float, float, float]]) -> tuple[float, float, float, float]:
    return (min(x[0] for x in boxes), min(x[1] for x in boxes), max(x[2] for x in boxes), max(x[3] for x in boxes))


def cluster_drawings(drawings: list[dict[str, Any]], page_rect: fitz.Rect) -> tuple[list[Region], dict[str, int]]:
    """Group nearby vector paths; preserve complexity rather than merely path area."""
    items: list[dict[str, Any]] = []
    ignored = {"tiny": 0, "separator": 0, "border": 0}
    for drawing in drawings:
        rect = drawing.get("rect")
        # A pure horizontal/vertical connector has a zero-area ("empty") Rect
        # in PyMuPDF, but is essential when joining diagram nodes.
        if not rect or rect.is_infinite:
            continue
        box = normalize_rect(rect, page_rect)
        width, height, area = box[2] - box[0], box[3] - box[1], rect_area(box)
        primitives = drawing.get("items", [])
        if area > 0.88 and len(primitives) <= 5:
            ignored["border"] += 1
            continue
        if (width < 0.003 or height < 0.003) and max(width, height) > 0.12:
            ignored["separator"] += 1
            continue
        if area < 0.00008 and max(width, height) < 0.018:
            ignored["tiny"] += 1
            continue
        curve_count = sum(1 for p in primitives if p and p[0] == "c")
        line_count = sum(1 for p in primitives if p and p[0] == "l")
        rect_count = sum(1 for p in primitives if p and p[0] in {"re", "qu"})
        items.append({"box": box, "paths": 1, "primitives": len(primitives), "curves": curve_count, "lines": line_count, "rects": rect_count})

    parent = list(range(len(items)))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def join(i: int, j: int) -> None:
        ri, rj = find(i), find(j)
        if ri != rj:
            parent[rj] = ri

    # Spatial bins avoid quadratic comparison on CAD / slide-deck pages with
    # thousands of paths while preserving the same 2.5%-page proximity rule.
    bins: dict[tuple[int, int], list[int]] = {}
    cell_size = 0.05
    for i, item in enumerate(items):
        x0, y0, x1, y1 = item["box"]
        cells = {
            (gx, gy)
            for gx in range(max(0, int((x0 - 0.025) / cell_size)), min(19, int((x1 + 0.025) / cell_size)) + 1)
            for gy in range(max(0, int((y0 - 0.025) / cell_size)), min(19, int((y1 + 0.025) / cell_size)) + 1)
        }
        candidates = {j for cell in cells for j in bins.get(cell, [])}
        for j in candidates:
            if _near(item["box"], items[j]["box"]):
                join(i, j)
        for cell in cells:
            bins.setdefault(cell, []).append(i)

    groups: dict[int, list[dict[str, Any]]] = {}
    for i, item in enumerate(items):
        groups.setdefault(find(i), []).append(item)

    regions: list[Region] = []
    for group in groups.values():
        box = _merge_boxes([x["box"] for x in group])
        area = rect_area(box)
        paths = sum(x["paths"] for x in group)
        primitives = sum(x["primitives"] for x in group)
        curves = sum(x["curves"] for x in group)
        lines = sum(x["lines"] for x in group)
        rects = sum(x["rects"] for x in group)
        # A vector cluster is meaningful only when it has both spatial extent and
        # sufficient structure. This rejects isolated rules, bullets, and borders.
        complexity = paths + primitives * 0.55 + curves * 0.8
        if area >= 0.055 and complexity >= 10 or area >= 0.13 and complexity >= 7:
            regions.append(Region(box, "vector", area, meta={"paths": paths, "primitives": primitives, "curves": curves, "lines": lines, "rectangles": rects, "complexity": round(complexity, 2)}))
    return regions, ignored


def is_table_like(region: Region, text_chars: int, raster_regions: list[Region]) -> bool:
    """Conservatively identify ordinary extractable text-table grids."""
    if raster_regions or region.kind != "vector" or text_chars < 220:
        return False
    m = region.meta
    orthogonal = m.get("lines", 0) + m.get("rectangles", 0)
    primitives = max(1, m.get("primitives", 0))
    return region.area > 0.12 and orthogonal / primitives > 0.78 and m.get("curves", 0) == 0 and text_chars / max(region.area, 0.01) > 900


def score_page(
    raster_regions: list[Region],
    vector_regions: list[Region],
    ignored_count: int,
    image_count: int,
    text_chars: int,
    text_blocks: int,
) -> tuple[float, float, int, str, dict[str, Any]]:
    table_regions = [v for v in vector_regions if is_table_like(v, text_chars, raster_regions)]
    meaningful_vectors = [v for v in vector_regions if v not in table_regions]
    regions = raster_regions + meaningful_vectors
    area = min(1.0, union_ratio(regions))
    max_raster = max((r.area for r in raster_regions), default=0.0)
    max_vector = max((r.area for r in meaningful_vectors), default=0.0)
    vector_complexity = max((r.meta.get("complexity", 0) for r in meaningful_vectors), default=0)
    substantial = sum(r.area >= 0.12 for r in regions)

    # Deterministic precision-biased formula (all terms capped):
    #   35% union visual area, 30% strongest raster, 25% strongest vector cluster,
    #   6% multiple substantial visuals, 8% vector complexity, 5% favorable
    #   text/graphics balance. Ordinary table grids and noise receive penalties.
    score = 0.0
    score += min(0.35, area * 0.85)
    score += min(0.30, max(0.0, max_raster - 0.04) * 1.55)
    score += min(0.25, max(0.0, max_vector - 0.035) * 1.18)
    score += min(0.10, max(0, substantial - 1) * 0.05)
    score += min(0.08, max(0.0, vector_complexity - 12) / 200)
    if area >= 0.15 and 30 <= text_chars <= 1800:
        score += 0.05
    if max_raster >= 0.15:
        score += 0.12
    if max_vector >= 0.16 and vector_complexity >= 18:
        score += 0.14
    if table_regions:
        score -= min(0.24, 0.12 + 0.04 * len(table_regions))
    if not regions:
        score -= min(0.08, ignored_count * 0.01)
    score = round(max(0.0, min(1.0, score)), 3)

    if max_raster >= max_vector and max_raster >= 0.15:
        reason = f"Large raster visual occupying approximately {max_raster:.0%} of the page"
    elif max_vector >= 0.18 and vector_complexity >= 10:
        reason = f"Substantial vector chart or diagram occupying approximately {max_vector:.0%} of the page"
    elif len(meaningful_vectors) >= 2 and area >= 0.25:
        reason = f"Multiple connected vector diagram regions occupy approximately {area:.0%} of the page"
    elif len(raster_regions) >= 2 and area >= 0.15:
        reason = f"Multiple substantial raster visuals occupy approximately {area:.0%} of the page"
    elif table_regions and not regions:
        reason = "Extractable text table with a regular vector grid; no substantial non-table visual"
    elif not regions and image_count:
        reason = "Only small, repeated, or margin graphics detected"
    elif not regions:
        reason = "Text-only page or no meaningful visual region detected"
    else:
        reason = f"Limited graphical content occupying approximately {area:.0%} of the page"

    debug = {
        "text_chars": text_chars,
        "text_blocks": text_blocks,
        "max_raster_ratio": round(max_raster, 4),
        "max_vector_ratio": round(max_vector, 4),
        "vector_complexity": round(vector_complexity, 2),
        "table_like_regions": len(table_regions),
        "scoring_components": {"visual_union": round(area, 4), "substantial_regions": substantial},
    }
    return score, round(area, 4), substantial, reason, debug
