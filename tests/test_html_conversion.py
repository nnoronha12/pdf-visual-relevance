from __future__ import annotations

import io

import fitz
from fastapi.testclient import TestClient
from pptx import Presentation

from app.main import app


HTML = """<!doctype html>
<html>
<head>
  <style>
    @page { size: 13.333in 7.5in; margin: 0; }
    * { box-sizing: border-box; }
    body { margin: 0; }
    .slide {
      width: 13.333in;
      height: 7.5in;
      page-break-after: always;
      background: #10233f;
      color: white;
      font: 48px Arial, sans-serif;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    .slide:last-child { page-break-after: auto; }
  </style>
</head>
<body>
  <section class="slide">Example Slide One</section>
  <section class="slide">Example Slide Two</section>
</body>
</html>"""


def payload(output_format: str) -> dict[str, object]:
    return {
        "html": HTML,
        "format": output_format,
        "filename": f"example-test.{output_format}",
        "deck_id": "97000000-0000-4000-8000-000000000001",
        "revision": 3,
    }


def test_convert_pdf_and_pptx() -> None:
    with TestClient(app) as client:
        pdf_response = client.post("/convert", json=payload("pdf"))
        assert pdf_response.status_code == 200
        assert pdf_response.headers["content-type"] == "application/pdf"
        assert 'filename="example-test.pdf"' in pdf_response.headers["content-disposition"]
        assert pdf_response.headers["x-deck-revision"] == "3"
        assert pdf_response.content.startswith(b"%PDF-")
        pdf = fitz.open(stream=pdf_response.content, filetype="pdf")
        assert pdf.page_count == 2
        pdf.close()

        pptx_response = client.post("/convert", json=payload("pptx"))
        assert pptx_response.status_code == 200
        assert pptx_response.headers["content-type"] == "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        assert pptx_response.content.startswith(b"PK")
        presentation = Presentation(io.BytesIO(pptx_response.content))
        assert len(presentation.slides) == 2
        assert round(presentation.slide_width / 914400, 2) == 13.33
        assert round(presentation.slide_height / 914400, 2) == 7.5


def test_convert_rejects_unsupported_format() -> None:
    response = TestClient(app).post("/convert", json=payload("docx"))
    assert response.status_code == 422
    assert response.json()["success"] is False
