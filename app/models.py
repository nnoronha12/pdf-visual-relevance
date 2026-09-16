from __future__ import annotations

from typing import Any

from pydantic import BaseModel, Field


class PageResult(BaseModel):
    page: int
    relevant: bool
    score: float = Field(ge=0, le=1)
    reason: str
    visual_area_ratio: float = Field(ge=0, le=1)
    image_count: int = Field(ge=0)
    large_visual_count: int = Field(ge=0)
    debug: dict[str, Any] | None = None


class AnalysisResponse(BaseModel):
    success: bool = True
    page_count: int
    relevant_pages: list[int]
    pages: list[PageResult]
