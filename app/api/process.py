from pathlib import Path
import shutil
import tempfile
import os

from fastapi import APIRouter, Form, HTTPException, UploadFile, File, Query
from fastapi.responses import FileResponse

from app.utils.chamber_l500_converter import (
    chamber_l500_convert,
    validate_document as validate_chamber_l500,
)
from app.utils.l500_chamber_converter import (
    l500_chamber_convert,
    validate_document as validate_l500_chamber,
)
from app.utils.docx_to_pdf import convert_docx_to_pdf  # <-- You must implement this

router = APIRouter()

@router.post("/convert")
async def convert_document_endpoint(
    file: UploadFile = File(...),
    mode: str = Form(...),
    preview: bool = Query(False),
    download: bool = Query(False),
):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_in:
        shutil.copyfileobj(file.file, tmp_in)
        tmp_in_path = tmp_in.name

    try:
        if mode == "l500_chamber":
            is_valid = await validate_l500_chamber(tmp_in_path)
            if not is_valid:
                raise HTTPException(status_code=400, detail="Invalid L500 document")
            template_path = Path(__file__).resolve().parent.parent / "utils" / "templateDestination.docx"
            output_path = await l500_chamber_convert(tmp_in_path, template_path)

        elif mode == "chamber_l500":
            is_valid = await validate_chamber_l500(tmp_in_path)
            if not is_valid:
                raise HTTPException(status_code=400, detail="Invalid Chamber document")
            template_path = Path(__file__).resolve().parent.parent / "utils" / "legal 500.doc"
            output_path = await chamber_l500_convert(tmp_in_path, template_path)

        else:
            raise HTTPException(status_code=400, detail="Invalid conversion mode")

        # PREVIEW: Convert DOCX to PDF and return that
        if preview:
            pdf_path = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
            convert_docx_to_pdf(output_path, pdf_path)
            return FileResponse(
                pdf_path,
                media_type="application/pdf",
                filename="preview.pdf",
            )

        # DOWNLOAD: Return the actual DOCX
        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename="converted.docx",
        )

    finally:
        if os.path.exists(tmp_in_path):
            os.remove(tmp_in_path)
