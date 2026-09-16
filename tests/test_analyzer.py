from __future__ import annotations

import io
import zipfile

import fitz
from fastapi.testclient import TestClient
from PIL import Image, ImageDraw

from app.main import app
from app.pdf_analyzer import analyze_pdf


def png_bytes(kind: str = "photo", size: tuple[int, int] = (900, 600)) -> bytes:
    image = Image.new("RGB", size, "#e9eef5")
    draw = ImageDraw.Draw(image)
    if kind == "logo":
        draw.rounded_rectangle((4, 4, size[0] - 4, size[1] - 4), radius=10, fill="#155eef")
        draw.text((18, size[1] // 3), "ACME", fill="white")
    else:
        for x in range(0, size[0], 60):
            color = (30 + x % 180, 80 + x % 140, 130 + x % 100)
            draw.rectangle((x, 50, min(x + 48, size[0]), size[1] - 50), fill=color)
        draw.rectangle((80, 120, size[0] - 80, size[1] - 120), outline="white", width=8)
        draw.text((110, 150), "INFORMATION-RICH SCREENSHOT", fill="white")
    out = io.BytesIO()
    image.save(out, format="PNG")
    return out.getvalue()


def make_fixture_pdf() -> bytes:
    doc = fitz.open()
    page_size = (595, 842)

    # 1: text only.
    page = doc.new_page(width=page_size[0], height=page_size[1])
    for i in range(18):
        page.insert_text((60, 80 + i * 28), f"Normal paragraph line {i + 1}: searchable text for the RAG pipeline.", fontsize=11)

    # 2: one tiny logo plus paragraphs.
    page = doc.new_page(width=page_size[0], height=page_size[1])
    page.insert_image(fitz.Rect(505, 25, 565, 50), stream=png_bytes("logo", (240, 100)))
    for i in range(15):
        page.insert_text((60, 100 + i * 30), f"Narrative content {i + 1} without meaningful graphics.", fontsize=11)

    # 3: large screenshot / photo.
    page = doc.new_page(width=page_size[0], height=page_size[1])
    page.insert_text((60, 55), "Application dashboard", fontsize=16)
    page.insert_image(fitz.Rect(55, 95, 540, 620), stream=png_bytes("photo"))

    # 4: chart made from vectors.
    page = doc.new_page(width=page_size[0], height=page_size[1])
    page.insert_text((60, 55), "Quarterly performance", fontsize=16)
    shape = page.new_shape()
    shape.draw_line((85, 650), (520, 650))
    shape.draw_line((85, 160), (85, 650))
    for i in range(12):
        x = 105 + i * 32
        height = 80 + (i * 37) % 340
        shape.draw_rect(fitz.Rect(x, 650 - height, x + 20, 650))
    for i in range(6):
        shape.draw_line((85, 650 - i * 90), (520, 650 - i * 90))
    shape.finish(color=(0.1, 0.3, 0.7), fill=None, width=1.2)
    shape.commit()

    # 5: architecture/flow diagram made entirely from vector paths and labels.
    page = doc.new_page(width=page_size[0], height=page_size[1])
    page.insert_text((60, 55), "Platform architecture", fontsize=16)
    boxes = []
    for row in range(4):
        for col in range(3):
            x, y = 65 + col * 170, 120 + row * 140
            boxes.append(fitz.Rect(x, y, x + 125, y + 72))
            shape = page.new_shape()
            shape.draw_rect(boxes[-1])
            shape.finish(color=(0.1, 0.25, 0.5), fill=(0.88, 0.93, 1), width=1.5)
            shape.commit()
            page.insert_text((x + 14, y + 40), f"Service {row + 1}.{col + 1}", fontsize=10)
    for idx in range(9):
        a, b = boxes[idx], boxes[idx + 3]
        shape = page.new_shape()
        shape.draw_line((a.x0 + 62, a.y1), (b.x0 + 62, b.y0))
        shape.finish(color=(0.15, 0.15, 0.15), width=2)
        shape.commit()

    # 6: ordinary text table; line grid is extractable and should be suppressed.
    page = doc.new_page(width=page_size[0], height=page_size[1])
    page.insert_text((60, 55), "Employee directory", fontsize=16)
    shape = page.new_shape()
    left, top, width, row_h = 55, 100, 485, 30
    for r in range(16):
        shape.draw_line((left, top + r * row_h), (left + width, top + r * row_h))
    for x in (left, left + 150, left + 300, left + width):
        shape.draw_line((x, top), (x, top + 15 * row_h))
    shape.finish(color=(0.4, 0.4, 0.4), width=0.7)
    shape.commit()
    for r in range(15):
        page.insert_text((left + 8, top + 20 + r * row_h), f"Employee {r + 1}", fontsize=9)
        page.insert_text((left + 158, top + 20 + r * row_h), f"Department {r % 4 + 1}", fontsize=9)
        page.insert_text((left + 308, top + 20 + r * row_h), f"employee{r + 1}@example.com", fontsize=9)

    # 7-11: repeated header logo across most pages with normal text.
    logo = png_bytes("logo", (240, 100))
    for p in range(5):
        page = doc.new_page(width=page_size[0], height=page_size[1])
        page.insert_image(fitz.Rect(480, 24, 565, 59), stream=logo)
        for i in range(16):
            page.insert_text((60, 100 + i * 29), f"Policy section {p + 1}, paragraph {i + 1}: text content only.", fontsize=11)

    result = doc.tobytes(deflate=True)
    doc.close()
    return result


def make_minimal_docx() -> bytes:
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            "</Types>",
        )
        archive.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            "</Relationships>",
        )
        archive.writestr(
            "word/document.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p><w:r><w:t>Converted document test</w:t></w:r></w:p>'
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>',
        )
    return output.getvalue()


def test_precision_fixture() -> None:
    result = analyze_pdf(make_fixture_pdf(), include_debug=True)
    assert result.pages[0].relevant is False
    assert result.pages[1].relevant is False
    assert result.pages[2].relevant is True
    assert result.pages[3].relevant is True
    assert result.pages[4].relevant is True
    assert result.pages[5].relevant is False
    assert all(page.relevant is False for page in result.pages[6:])
    assert result.relevant_pages == [3, 4, 5]


def test_api_and_rendering() -> None:
    client = TestClient(app)
    data = make_fixture_pdf()
    health = client.get("/health")
    assert health.json() == {"status": "ok"}
    response = client.post("/analyze-pdf", files={"file": ("fixture.pdf", data, "application/pdf")}, data={"min_score": "0.55"})
    assert response.status_code == 200
    assert response.json()["relevant_pages"] == [3, 4, 5]
    rendered = client.post("/render-pages", files={"file": ("fixture.pdf", data, "application/pdf")}, data={"page_numbers": "3,4,5"})
    assert rendered.status_code == 200
    assert rendered.headers["content-type"] == "application/zip"

    extracted = client.post("/extract-content", files={"file": ("fixture.pdf", data, "application/pdf")})
    assert extracted.status_code == 200
    content = extracted.json()
    assert content["page_count"] == 11
    assert "Normal paragraph line 1" in content["pages"][0]["text"]
    assert content["pages"][5]["tables"]
    assert content["pages"][5]["tables"][0]["row_count"] >= 10


def test_invalid_upload() -> None:
    response = TestClient(app).post("/analyze-pdf", files={"file": ("bad.pdf", b"not a pdf", "application/pdf")})
    assert response.status_code == 415


def test_repeated_asset_is_identified_as_branding() -> None:
    doc = fitz.open()
    logo = png_bytes("logo", (400, 180))
    for index in range(5):
        page = doc.new_page(width=595, height=842)
        # About 4% of page area and not in the header/footer: size alone would
        # retain it, so this specifically exercises cross-page suppression.
        page.insert_image(fitz.Rect(420, 120, 560, 260), stream=logo)
        page.insert_text((60, 100), f"Narrative page {index + 1}", fontsize=12)
    data = doc.tobytes()
    doc.close()
    result = analyze_pdf(data, include_debug=True)
    assert result.relevant_pages == []
    assert all(any(x["reason"] == "repeated_branding" for x in page.debug["ignored_regions"]) for page in result.pages)


def test_docx_conversion_endpoint() -> None:
    response = TestClient(app).post(
        "/convert-to-pdf",
        files={"file": ("sample.docx", make_minimal_docx(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
    )
    assert response.status_code == 200
    assert response.headers["content-type"] == "application/pdf"
    assert response.content.startswith(b"%PDF-")
