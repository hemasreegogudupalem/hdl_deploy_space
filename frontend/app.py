"""FastAPI application for the Hdl Automation frontend and generation APIs."""

import os
from pathlib import Path
import shutil
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from frontend import service

app = FastAPI(
    title="Hdl Automation",
    description="Web Interface for Oracle HCM Data Loader automation",
    version="2.0.0",
)

# CORS configuration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
TEMPLATES_DIR = BASE_DIR / "templates"

STATIC_DIR.mkdir(parents=True, exist_ok=True)
TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)

# Mount static files
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


class GenerateRequest(BaseModel):
    file_path: str
    template_name: str = "Worker"
    legislation_code: Optional[str] = None
    legal_employer_name: Optional[str] = None
    source_system_owner: Optional[str] = None


class BatchTemplateSelection(BaseModel):
    template_name: str
    column_mappings: Dict[str, Optional[str]] = Field(default_factory=dict)


class BatchGenerateRequest(BaseModel):
    file_path: str
    templates: List[BatchTemplateSelection]
    legislation_code: Optional[str] = None
    legal_employer_name: Optional[str] = None
    source_system_owner: Optional[str] = None


@app.get("/", response_class=HTMLResponse)
async def serve_index():
    """Serve the single-page application interface."""
    index_file = TEMPLATES_DIR / "index.html"
    if not index_file.exists():
        raise HTTPException(status_code=404, detail="Index template not found")
    return HTMLResponse(content=index_file.read_text(encoding="utf-8"))


@app.get("/api/status")
async def get_system_status():
    """Return database health, connection details, and summary stats."""
    db_status = service.test_database_connection()
    inputs = service.list_input_files()
    files = service.list_reports_and_outputs()
    return {
        "database": db_status,
        "input_files_count": len(inputs),
        "outputs_count": len(files["outputs"]),
        "reports_count": len(files["reports"]),
    }


@app.post("/api/test-db")
async def test_db_connection():
    """Trigger manual database connectivity test."""
    return service.test_database_connection()


@app.get("/api/templates")
async def get_templates():
    """List all available HDL templates from MySQL."""
    return service.get_available_templates()


@app.get("/api/templates/{name}")
async def get_template_details(name: str):
    """Get deep object and attribute structure for a specific template."""
    try:
        return service.get_template_details(name)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.get("/api/generation/templates/{name}")
async def get_batch_template_schema(name: str):
    """Return user-mappable source fields for the batch generation UI."""
    try:
        return service.get_batch_template_schema(name)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.get("/api/inputs")
async def get_input_files():
    """List all Excel files available in the input directory."""
    return service.list_input_files()


@app.delete("/api/inputs")
async def delete_input_file(filename: str = Query(..., min_length=1)):
    """Delete an Excel workbook from the input directory by filename."""
    try:
        return service.delete_input_file(filename)
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="Workbook not found in the input directory.")
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/upload")
async def upload_excel_file(file: UploadFile = File(...)):
    """Upload a new Excel spreadsheet into the input directory."""
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(
            status_code=400,
            detail="Only Excel workbooks (.xlsx, .xls) are supported.",
        )

    target_dir = Path("input")
    target_dir.mkdir(parents=True, exist_ok=True)
    target_path = target_dir / file.filename

    with target_path.open("wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    return {
        "success": True,
        "filename": file.filename,
        "path": str(target_path.as_posix()),
        "size_kb": round(target_path.stat().st_size / 1024, 1),
    }


@app.get("/api/excel/preview")
async def preview_excel(
    file_path: str = Query(..., description="Relative path to Excel file"),
    page: int = Query(1, ge=1),
    page_size: int = Query(15, ge=1, le=100),
    search: str = Query("", description="Optional search filter"),
):
    """Retrieve paginated rows and columns from an Excel file."""
    try:
        return service.inspect_excel_data(file_path, page, page_size, search)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Failed to read Excel: {exc}")


@app.get("/api/defaults")
async def get_defaults(
    file_path: str = Query(..., description="Excel file path"),
    template_name: str = Query("Worker", description="Template name"),
):
    """Auto-detect generation defaults from the specified Excel file."""
    try:
        return service.detect_generation_defaults(file_path, template_name)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Failed to detect defaults: {exc}")


@app.get("/api/mappings/preview")
async def preview_mappings(
    file_path: str = Query(..., description="Excel file path"),
    template_name: str = Query("Worker", description="Template name"),
):
    """Analyze and preview field mappings between Excel and template."""
    try:
        return service.analyze_field_mappings(file_path, template_name)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Failed to analyze mappings: {exc}")


@app.post("/api/generate")
async def generate_hdl(req: GenerateRequest):
    """Execute the HDL generation process."""
    try:
        overrides = {}
        if req.legislation_code:
            overrides["legislation_code"] = req.legislation_code
        if req.legal_employer_name:
            overrides["legal_employer_name"] = req.legal_employer_name
        if req.source_system_owner:
            overrides["source_system_owner"] = req.source_system_owner

        result = service.run_hdl_generation_pipeline(
            file_path=req.file_path,
            template_name=req.template_name,
            generation_overrides=overrides,
        )
        return result
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/generate/batch/validate")
async def validate_batch_hdl(req: BatchGenerateRequest):
    """Dry-run selected templates and return validation by template."""
    try:
        overrides = {
            key: value
            for key, value in {
                "legislation_code": req.legislation_code,
                "legal_employer_name": req.legal_employer_name,
                "source_system_owner": req.source_system_owner,
            }.items()
            if value
        }
        return service.validate_batch_generation(
            file_path=req.file_path,
            templates=[
                selection.dict()
                for selection in req.templates
            ],
            generation_overrides=overrides,
        )
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/generate/batch")
async def generate_batch_hdl(req: BatchGenerateRequest):
    """Generate selected templates in HDL load order and create one ZIP."""
    try:
        overrides = {
            key: value
            for key, value in {
                "legislation_code": req.legislation_code,
                "legal_employer_name": req.legal_employer_name,
                "source_system_owner": req.source_system_owner,
            }.items()
            if value
        }
        return service.generate_batch_hdl(
            file_path=req.file_path,
            templates=[
                selection.dict()
                for selection in req.templates
            ],
            generation_overrides=overrides,
        )
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.get("/api/reports")
async def list_reports():
    """Retrieve all generated outputs and reports."""
    return service.list_reports_and_outputs()


@app.get("/api/reports/download/{category}/{filename}")
async def download_file(category: str, filename: str):
    """Download a generated DAT file or report."""
    if category == "output":
        file_path = Path("output") / filename
    elif category == "reports":
        file_path = Path("reports") / filename
    else:
        raise HTTPException(status_code=400, detail="Invalid file category")

    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Requested file not found")

    return FileResponse(
        path=str(file_path),
        filename=filename,
        media_type="application/octet-stream",
    )


@app.get("/api/reports/preview-dat/{filename}")
async def preview_dat_file(
    filename: str,
    max_lines: int = Query(300, ge=10, le=2000),
):
    """Get contents of an HDL DAT file for preview."""
    try:
        return service.read_full_dat_file(filename, max_lines=max_lines)
    except Exception as exc:
        raise HTTPException(status_code=404, detail=str(exc))


@app.get("/api/reports/preview-csv/{filename}")
async def preview_csv_report(
    filename: str,
    max_rows: int = Query(150, ge=10, le=1000),
):
    """Get preview rows of a CSV report (e.g. validation_errors.csv)."""
    try:
        return service.read_report_csv(filename, max_rows=max_rows)
    except Exception as exc:
        raise HTTPException(status_code=404, detail=str(exc))


@app.get("/api/reports/preview-excel/{filename}")
async def preview_excel_report(
    filename: str,
    max_rows: int = Query(250, ge=10, le=2000),
):
    """Get preview rows of an Excel report (e.g. mapping_report.xlsx)."""
    try:
        return service.read_report_excel(filename, max_rows=max_rows)
    except Exception as exc:
        raise HTTPException(status_code=404, detail=str(exc))
