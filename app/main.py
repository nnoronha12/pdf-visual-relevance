from __future__ import annotations

import io
from contextlib import asynccontextmanager
from typing import Annotated
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, Request, Response, UploadFile
from fastapi.encoders import jsonable_encoder
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse, StreamingResponse
from playwright.async_api import async_playwright
from pydantic import BaseModel, ConfigDict, Field, field_validator
from starlette.concurrency import run_in_threadpool

from .content_extraction import extract_content
from .conversion import DocumentConversionError, convert_office_to_pdf
from .html_conversion import HTMLConversionError, pdf_to_pptx, render_html_to_pdf
from .models import AnalysisResponse, ContentExtractionResponse
from .pdf_analyzer import PDFAnalysisError, analyze_pdf
from .rendering import render_pages_zip

MAX_UPLOAD_BYTES = 50 * 1024 * 1024
MAX_HTML_BYTES = 20 * 1024 * 1024


@asynccontextmanager
async def lifespan(application: FastAPI):
    async with async_playwright() as playwright:
        browser = await playwright.chromium.launch(args=["--no-sandbox"])
        application.state.chromium = browser
        yield
        await browser.close()


app = FastAPI(title="Document Services API", version="1.1.0", lifespan=lifespan)


class ConvertRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    html: str = Field(min_length=1)
    format: str
    filename: str = Field(min_length=1, max_length=255)
    deck_id: str = Field(min_length=1, max_length=128)
    revision: int = Field(ge=0)

    @field_validator("format")
    @classmethod
    def supported_format(cls, value: str) -> str:
        normalized = value.lower()
        if normalized not in {"pdf", "pptx"}:
            raise ValueError("format must be either 'pdf' or 'pptx'")
        return normalized

    @field_validator("html")
    @classmethod
    def html_size_limit(cls, value: str) -> str:
        if len(value.encode("utf-8")) > MAX_HTML_BYTES:
            raise ValueError("html exceeds the 20 MB request limit")
        return value


def output_filename(requested_name: str, output_format: str) -> str:
    # Strip path components/control characters and make the extension agree
    # with the actual response type.
    name = requested_name.replace("\\", "/").rsplit("/", 1)[-1]
    name = "".join(character for character in name if character >= " " and character != "\x7f").strip()
    stem = name.rsplit(".", 1)[0] if "." in name else name
    return f"{stem or 'deck'}.{output_format}"


def attachment_header(filename: str) -> str:
    ascii_name = filename.encode("ascii", "ignore").decode() or f"deck.{filename.rsplit('.', 1)[-1]}"
    ascii_name = ascii_name.replace('"', "")
    return f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{quote(filename)}"


async def read_pdf(upload: UploadFile) -> bytes:
    if upload.content_type not in {"application/pdf", "application/octet-stream", None} and not (upload.filename or "").lower().endswith(".pdf"):
        raise HTTPException(status_code=415, detail="The file must be a PDF")
    chunks = bytearray()
    while chunk := await upload.read(1024 * 1024):
        chunks.extend(chunk)
        if len(chunks) > MAX_UPLOAD_BYTES:
            raise HTTPException(status_code=413, detail="PDF exceeds the 50 MB upload limit")
    await upload.close()
    if not chunks:
        raise HTTPException(status_code=400, detail="The uploaded file is empty")
    if not bytes(chunks[:5]).startswith(b"%PDF-"):
        raise HTTPException(status_code=415, detail="The uploaded file does not appear to be a PDF")
    return bytes(chunks)


@app.exception_handler(PDFAnalysisError)
async def analysis_error_handler(_request, exc: PDFAnalysisError):
    return JSONResponse(status_code=422, content={"success": False, "error": str(exc)})


@app.exception_handler(DocumentConversionError)
async def conversion_error_handler(_request, exc: DocumentConversionError):
    return JSONResponse(status_code=422, content={"success": False, "error": str(exc)})


@app.exception_handler(HTMLConversionError)
async def html_conversion_error_handler(_request, exc: HTMLConversionError):
    return JSONResponse(status_code=422, content={"success": False, "error": str(exc)})


@app.exception_handler(HTTPException)
async def http_error_handler(_request: Request, exc: HTTPException):
    return JSONResponse(status_code=exc.status_code, content={"success": False, "error": str(exc.detail)})


@app.exception_handler(RequestValidationError)
async def validation_error_handler(_request: Request, exc: RequestValidationError):
    details = jsonable_encoder(exc.errors(), custom_encoder={Exception: str})
    return JSONResponse(status_code=422, content={"success": False, "error": "Invalid request fields", "details": details})


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/convert")
async def convert_html_endpoint(payload: ConvertRequest, request: Request) -> Response:
    pdf = await render_html_to_pdf(request.app.state.chromium, payload.html)
    filename = output_filename(payload.filename, payload.format)
    headers = {
        "Content-Disposition": attachment_header(filename),
        "X-Deck-Id": payload.deck_id,
        "X-Deck-Revision": str(payload.revision),
    }
    if payload.format == "pdf":
        return Response(content=pdf, media_type="application/pdf", headers=headers)

    pptx = await run_in_threadpool(pdf_to_pptx, pdf)
    return Response(
        content=pptx,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers=headers,
    )


@app.post("/analyze-pdf", response_model=AnalysisResponse, response_model_exclude_none=True)
async def analyze_endpoint(
    file: Annotated[UploadFile, File(...)],
    min_score: Annotated[float, Form()] = 0.55,
    include_debug: Annotated[bool, Form()] = False,
) -> AnalysisResponse:
    if not 0 <= min_score <= 1:
        raise HTTPException(status_code=422, detail="min_score must be between 0 and 1")
    data = await read_pdf(file)
    return await run_in_threadpool(analyze_pdf, data, min_score, include_debug)


@app.post("/extract-content", response_model=ContentExtractionResponse)
async def extract_content_endpoint(file: Annotated[UploadFile, File(...)]) -> ContentExtractionResponse:
    data = await read_pdf(file)
    return await run_in_threadpool(extract_content, data)


@app.post("/render-pages")
async def render_endpoint(
    file: Annotated[UploadFile, File(...)],
    page_numbers: Annotated[str, Form(...)],
) -> StreamingResponse:
    data = await read_pdf(file)
    archive = await run_in_threadpool(render_pages_zip, data, page_numbers)
    return StreamingResponse(io.BytesIO(archive), media_type="application/zip", headers={"Content-Disposition": 'attachment; filename="rendered_pages.zip"'})


@app.post("/convert-to-pdf")
async def convert_to_pdf_endpoint(file: Annotated[UploadFile, File(...)]) -> StreamingResponse:
    filename = file.filename
    clean_name = filename or "document"
    extension = clean_name.rsplit(".", 1)[-1].lower() if "." in clean_name else ""
    if extension not in {"docx", "ppt", "pptx"}:
        raise HTTPException(status_code=415, detail="Supported input formats are .docx, .ppt, and .pptx")
    chunks = bytearray()
    while chunk := await file.read(1024 * 1024):
        chunks.extend(chunk)
        if len(chunks) > MAX_UPLOAD_BYTES:
            raise HTTPException(status_code=413, detail="Document exceeds the 50 MB upload limit")
    await file.close()
    if not chunks:
        raise HTTPException(status_code=400, detail="The uploaded file is empty")
    pdf, output_name = await run_in_threadpool(convert_office_to_pdf, bytes(chunks), filename)
    return StreamingResponse(
        io.BytesIO(pdf),
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{output_name}"'},
    )
