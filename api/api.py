import base64
import binascii
import os
import re
from typing import Optional

from fastapi import FastAPI, File, Header, HTTPException, UploadFile
from fastapi.responses import Response
from pydantic import BaseModel, Field

from processor import process_excel_file

app = FastAPI(title="RM Assist Excel Processor API", version="1.0.0")

API_KEY = os.getenv("PROCESSOR_API_KEY", "")
MAX_FILE_BYTES = int(os.getenv("MAX_FILE_BYTES", str(16 * 1024 * 1024)))
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv"}


class ProcessBase64Request(BaseModel):
    file_name: str = Field(min_length=1, max_length=255)
    file_content_base64: str = Field(min_length=1)


def check_auth(x_api_key: Optional[str]):
    # Configure PROCESSOR_API_KEY in deployment. Refuse unauthenticated calls if configured.
    if API_KEY and x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="Invalid or missing API key.")


def validate_file(file_name: str, content: bytes):
    ext = os.path.splitext(file_name)[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(status_code=400, detail="Supported file types: .xlsx, .xls, .csv")
    if not content:
        raise HTTPException(status_code=400, detail="The uploaded file is empty.")
    if len(content) > MAX_FILE_BYTES:
        raise HTTPException(status_code=413, detail=f"File exceeds the {MAX_FILE_BYTES // (1024 * 1024)} MiB limit.")


def make_response(output_bytes: bytes, source_rows: int, output_rows: int):
    headers = {
        "Content-Disposition": 'attachment; filename="RM_Assist_Processed_Output.xlsx"',
        "X-Source-Rows": str(source_rows),
        "X-Output-Rows": str(output_rows),
    }
    return Response(
        content=output_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.get("/health")
def health():
    return {"status": "ok", "service": "rm-assist-excel-processor"}


@app.post("/api/excel/process-base64")
def process_base64(payload: ProcessBase64Request, x_api_key: Optional[str] = Header(default=None)):
    check_auth(x_api_key)
    try:
        content = base64.b64decode(payload.file_content_base64, validate=True)
    except (binascii.Error, ValueError):
        raise HTTPException(status_code=400, detail="file_content_base64 is not valid Base64.")
    validate_file(payload.file_name, content)
    try:
        output, source_rows, output_rows = process_excel_file(content, payload.file_name)
    except Exception as exc:
        raise HTTPException(status_code=422, detail=f"Could not process workbook: {str(exc)}")
    return make_response(output, source_rows, output_rows)


@app.post("/api/excel/process")
async def process_multipart(file: UploadFile = File(...), x_api_key: Optional[str] = Header(default=None)):
    check_auth(x_api_key)
    file_name = file.filename or "upload.xlsx"
    content = await file.read()
    validate_file(file_name, content)
    try:
        output, source_rows, output_rows = process_excel_file(content, file_name)
    except Exception as exc:
        raise HTTPException(status_code=422, detail=f"Could not process workbook: {str(exc)}")
    return make_response(output, source_rows, output_rows)
