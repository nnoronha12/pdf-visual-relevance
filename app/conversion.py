from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path


class DocumentConversionError(ValueError):
    pass


SUPPORTED_EXTENSIONS = {".docx", ".ppt", ".pptx"}


def safe_filename(filename: str | None) -> str:
    original = Path(filename or "document").name
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(original).stem).strip("._") or "document"
    suffix = Path(original).suffix.lower()
    if suffix not in SUPPORTED_EXTENSIONS:
        raise DocumentConversionError("Supported input formats are .docx, .ppt, and .pptx")
    return f"{stem}{suffix}"


def convert_office_to_pdf(data: bytes, filename: str | None, timeout_seconds: int = 120) -> tuple[bytes, str]:
    clean_name = safe_filename(filename)
    executable = shutil.which("libreoffice") or shutil.which("soffice")
    if not executable:
        raise DocumentConversionError("Document conversion is unavailable on this server")

    with tempfile.TemporaryDirectory(prefix="pdf-convert-") as temp_dir:
        root = Path(temp_dir)
        source = root / clean_name
        output_dir = root / "output"
        profile_dir = root / "libreoffice-profile"
        output_dir.mkdir()
        profile_dir.mkdir()
        source.write_bytes(data)
        command = [
            executable,
            "--headless",
            "--nologo",
            "--nodefault",
            "--nolockcheck",
            f"-env:UserInstallation={profile_dir.as_uri()}",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_dir),
            str(source),
        ]
        try:
            completed = subprocess.run(command, capture_output=True, text=True, timeout=timeout_seconds, env=os.environ.copy())
        except subprocess.TimeoutExpired as exc:
            raise DocumentConversionError(f"Document conversion exceeded {timeout_seconds} seconds") from exc
        except OSError as exc:
            raise DocumentConversionError("Could not start the document converter") from exc

        output = output_dir / f"{source.stem}.pdf"
        if completed.returncode != 0 or not output.exists() or output.stat().st_size == 0:
            detail = (completed.stderr or completed.stdout or "unknown conversion error").strip()[-500:]
            raise DocumentConversionError(f"LibreOffice could not convert the document: {detail}")
        return output.read_bytes(), f"{source.stem}.pdf"
