from __future__ import annotations

import io
from typing import Annotated

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse, StreamingResponse
from starlette.concurrency import run_in_threadpool

from .content_extraction import extract_content
from .conversion import DocumentConversionError, convert_office_to_pdf
from .models import AnalysisResponse, ContentExtractionResponse
from .pdf_analyzer import PDFAnalysisError, analyze_pdf
from .rendering import render_pages_zip

MAX_UPLOAD_BYTES = 50 * 1024 * 1024

app = FastAPI(title="PDF Visual Relevance API", version="1.0.0")


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


@app.exception_handler(HTTPException)
async def http_error_handler(_request: Request, exc: HTTPException):
    return JSONResponse(status_code=exc.status_code, content={"success": False, "error": str(exc.detail)})


@app.exception_handler(RequestValidationError)
async def validation_error_handler(_request: Request, exc: RequestValidationError):
    return JSONResponse(status_code=422, content={"success": False, "error": "Invalid request fields", "details": exc.errors()})


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


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
