#!/usr/bin/env python3
"""HTTP API for the PDF URL capture web app."""
from __future__ import annotations

import multiprocessing as mp
import os
import re
import shutil
import tempfile
import threading
import time
import uuid
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed
from contextlib import asynccontextmanager
from pathlib import Path

import pymupdf as fitz
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from starlette.background import BackgroundTask

from pdf_url_extractor import (
    extract_pdf,
    extract_pdf_pages_worker,
    extract_pdf_worker,
    footer_for_pdf,
    write_excel,
)
from url_status import attach_url_status

ROOT = Path(__file__).resolve().parent
TEMP_ROOT = Path(tempfile.gettempdir()) / "url-extractor-tmp"

MAX_UPLOAD_BYTES = 50 * 1024 * 1024
MAX_FILES = 25
ALLOWED_SUFFIXES = {".pdf"}
_FILE_POOL: ProcessPoolExecutor | None = None
_FILE_POOL_LOCK = threading.Lock()
_POOL_WORKERS = 2


def _pool_worker_init() -> None:
    try:
        from vision_scan import warmup
        warmup()
    except Exception:
        pass


def _warmup_models() -> None:
    try:
        from vision_scan import warmup
        warmup()
    except Exception:
        pass


def _start_file_pool() -> None:
    """PDFs are read in this process. Extra processes made each file wait on startup."""
    return


def _stop_file_pool() -> None:
    global _FILE_POOL
    with _FILE_POOL_LOCK:
        if _FILE_POOL is None:
            return
        _FILE_POOL.shutdown(wait=False, cancel_futures=True)
        _FILE_POOL = None


def _job_dir(job_id: str) -> Path:
    path = TEMP_ROOT / job_id
    path.mkdir(parents=True, exist_ok=True)
    return path


def _cleanup_job_dir(job_dir: str) -> None:
    shutil.rmtree(job_dir, ignore_errors=True)


def _remove_project_jobs_folder() -> None:
    shutil.rmtree(ROOT / "jobs", ignore_errors=True)
    for pattern in ("*qr*.png", "*qr*.jpg", "*_qr_clip.png"):
        for path in ROOT.glob(pattern):
            try:
                path.unlink()
            except Exception:
                pass


@asynccontextmanager
async def lifespan(_app: FastAPI):
    _remove_project_jobs_folder()
    _start_file_pool()
    threading.Thread(target=_warmup_models, daemon=True).start()
    yield
    _stop_file_pool()
    shutil.rmtree(TEMP_ROOT, ignore_errors=True)
    _remove_project_jobs_folder()


app = FastAPI(title="PDF URL Capture", lifespan=lifespan)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
)

_jobs: dict[str, dict] = {}
_lock = threading.Lock()


class ProcessBody(BaseModel):
    job_id: str


def _snapshot(job: dict) -> dict:
    payload = {
        "job_id": job["job_id"],
        "status": job["status"],
        "progress": job["progress"],
        "message": job["message"],
        "stats": job["stats"],
        "recentUrls": job["recent_urls"],
    }
    if job.get("results"):
        payload["results"] = job["results"]
    if job.get("error"):
        payload["error"] = job["error"]
    return payload


def _update(job_id: str, **fields) -> None:
    with _lock:
        job = _jobs.get(job_id)
        if not job:
            return
        job.update(fields)


def _safe_name(filename: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", filename) or "document.pdf"


def _page_count(path: Path) -> int:
    document = fitz.open(path)
    try:
        return int(document.page_count)
    finally:
        document.close()


def _empty_stats(files_total: int = 0) -> dict:
    return {
        "filesTotal": files_total,
        "filesDone": 0,
        "pagesScanned": 0,
        "urlsFound": 0,
        "headerUrls": 0,
        "footerUrls": 0,
        "qrUrls": 0,
        "imageUrls": 0,
        "workingUrls": 0,
        "failedUrls": 0,
        "urlsChecked": 0,
        "phase": "extract",
    }


def _count_stats(rows: list[dict], files_total: int, files_done: int, pages_done: int, phase: str) -> dict:
    return {
        "filesTotal": files_total,
        "filesDone": files_done,
        "pagesScanned": pages_done,
        "urlsFound": len(rows),
        "headerUrls": sum(1 for row in rows if row.get("Location") == "header"),
        "footerUrls": sum(1 for row in rows if row.get("Location") == "footer"),
        "qrUrls": sum(1 for row in rows if row.get("Source") == "qr"),
        "imageUrls": sum(1 for row in rows if row.get("Source") in {"image", "logo"}),
        "workingUrls": sum(1 for row in rows if str(row.get("Status") or "") == "Working"),
        "failedUrls": sum(1 for row in rows if "not working" in str(row.get("Status") or "").lower()),
        "urlsChecked": sum(1 for row in rows if row.get("Status")),
        "phase": phase,
    }


def _preview_row(row: dict) -> dict:
    return {
        "fileName": row.get("Filename") or "",
        "fileDate": row.get("Filename Date") or "",
        "last4": row.get("Last 4 Digits") or "",
        "clientId": row.get("Client ID") or "",
        "formType": row.get("Form Type") or "",
        "pageNumber": row.get("Page Number") or "",
        "url": row.get("URL") or "",
        "location": row.get("Location") or "body",
        "source": row.get("Source") or "text",
        "status": row.get("Status") or "",
    }


def _run_job(job_id: str) -> None:
    with _lock:
        job = _jobs[job_id]
        input_files: list[dict] = job["input_files"]

    started = time.perf_counter()
    _update(job_id, status="running", progress=4, message="Scanning PDFs...", started_at=started, stats=_empty_stats(len(input_files)))

    all_urls: list[dict] = []
    all_text: list[dict] = []
    pages_done = 0
    files_done = 0
    recent: list[dict] = []
    file_errors: list[str] = []
    progress_lock = threading.Lock()
    extracted: dict[int, tuple[list[dict], list[dict]]] = {}

    def _record_result(index: int, url_rows: list[dict], text_rows: list[dict], error: str | None) -> None:
        nonlocal files_done, pages_done, recent
        with progress_lock:
            files_done += 1
            if error:
                file_errors.append(error)
            else:
                extracted[index] = (url_rows, text_rows)
                pages_done += max(len(text_rows), 1)
                for row in reversed(url_rows[-8:]):
                    recent = [
                        {
                            "id": f"{job_id}-{len(recent)}-{row['URL']}",
                            "url": row["URL"],
                            "file": row.get("Filename") or "",
                            "page": row["Page Number"],
                            "location": row.get("Location") or "body",
                            "source": row.get("Source") or "text",
                            "status": row.get("Status") or "",
                        },
                        *recent,
                    ][:10]
            merged = [row for item in extracted.values() for row in item[0]]
            _update(
                job_id,
                progress=min(8 + int((files_done / max(len(input_files), 1)) * 62), 68),
                message=f"Extracting {len(input_files)} PDF(s) · {files_done} of {len(input_files)} complete",
                recent_urls=recent,
                stats=_count_stats(merged, len(input_files), files_done, pages_done, "extract"),
            )

    file_count = len(input_files)
    running_message = "Reading the PDF..." if file_count == 1 else f"Reading {file_count} PDFs..."
    _update(
        job_id,
        progress=8,
        message=running_message,
        stats=_empty_stats(file_count),
    )

    def _consume(index: int, url_rows: list[dict], text_rows: list[dict], error: str | None) -> None:
        _record_result(index, url_rows, text_rows, error)

    def _page_groups(page_count: int, workers: int) -> list[list[int]]:
        groups: list[list[int]] = [[] for _ in range(max(1, workers))]
        for index in range(page_count):
            groups[index % len(groups)].append(index)
        return [group for group in groups if group]

    for index, item in enumerate(input_files):
        try:
            url_rows, text_rows = extract_pdf(
                item["path"],
                file_name=item["name"],
                check_urls=False,
                page_workers=1,
            )
            _consume(index, url_rows, text_rows, None)
        except Exception as exc:
            _consume(index, [], [], f"{item['name']}: {exc}")

    for index in range(len(input_files)):
        if index not in extracted:
            continue
        url_rows, text_rows = extracted[index]
        all_urls.extend(url_rows)
        all_text.extend(text_rows)

    if not input_files:
        _update(job_id, status="failed", error="No PDF files were uploaded.", message="Unable to process files.")
        return
    if files_done == 0:
        _update(job_id, status="failed", error=file_errors[0] if file_errors else "The PDF files could not be read.", message="Unable to process files.")
        return

    if all_urls:
        _update(job_id, progress=72, message="Checking URLs...", stats=_count_stats(all_urls, len(input_files), files_done, pages_done, "verify"))

        def on_check_progress(done: int, total: int) -> None:
            progress = 72 + int((done / max(total, 1)) * 22)
            _update(
                job_id,
                progress=min(progress, 94),
                message=f"Checking URLs · {done} of {total}",
                stats={**_count_stats(all_urls, len(input_files), files_done, pages_done, "verify"), "urlsChecked": done},
            )

        attach_url_status(all_urls, on_progress=on_check_progress)
        recent = [
            {
                "id": f"{job_id}-final-{index}-{row['URL']}",
                "url": row["URL"],
                "file": row.get("Filename") or "",
                "page": row["Page Number"],
                "location": row.get("Location") or "body",
                "source": row.get("Source") or "text",
                "status": row.get("Status") or "",
            }
            for index, row in enumerate(reversed(all_urls[-10:]))
        ]

    _update(job_id, progress=96, message="Preparing Excel file...")
    if len(input_files) == 1:
        output_name = f"{Path(input_files[0]['name']).stem}_urls.xlsx"
    else:
        output_name = "extracted_urls.xlsx"
    job_dir = Path(input_files[0]["path"]).parent if input_files else _job_dir(job_id)
    output_path = job_dir / output_name
    write_excel(all_urls, all_text, output_path)
    for item in input_files:
        try:
            Path(item["path"]).unlink(missing_ok=True)
        except Exception:
            pass

    elapsed = round(time.perf_counter() - started, 1)
    preview = [_preview_row(row) for row in all_urls[:120]]
    error_note = f" {len(file_errors)} file(s) could not be read." if file_errors else ""
    stats = _count_stats(all_urls, len(input_files), files_done, pages_done, "done")
    _update(
        job_id,
        status="completed",
        progress=100,
        message="Capture complete" + error_note,
        output_path=str(output_path),
        output_name=output_name,
        recent_urls=recent,
        stats=stats,
        results={
            "totalUrls": len(all_urls),
            "filesProcessed": files_done,
            "pagesScanned": pages_done,
            "headerUrls": stats["headerUrls"],
            "footerUrls": stats["footerUrls"],
            "qrUrls": stats["qrUrls"],
            "imageUrls": stats["imageUrls"],
            "workingUrls": stats["workingUrls"],
            "failedUrls": stats["failedUrls"],
            "processingTimeSeconds": elapsed,
            "downloadFilename": output_name,
            "preview": preview,
            "warnings": file_errors,
        },
    )


@app.get("/health")
def health():
    return {"ok": True}


@app.post("/upload")
async def upload(files: list[UploadFile] = File(...)):
    if not files:
        raise HTTPException(status_code=400, detail="Please upload at least one PDF file.")
    if len(files) > MAX_FILES:
        raise HTTPException(status_code=400, detail=f"Please upload at most {MAX_FILES} PDF files at a time.")

    saved: list[dict] = []
    job_id = str(uuid.uuid4())
    job_dir = _job_dir(job_id)

    for index, upload_file in enumerate(files):
        filename = upload_file.filename or f"document-{index + 1}.pdf"
        suffix = Path(filename).suffix.lower()
        if suffix not in ALLOWED_SUFFIXES:
            raise HTTPException(status_code=400, detail="Please upload PDF files only.")
        data = await upload_file.read()
        if not data:
            raise HTTPException(status_code=400, detail=f"{filename} is empty.")
        if len(data) > MAX_UPLOAD_BYTES:
            raise HTTPException(status_code=400, detail=f"{filename} is larger than 50 MB.")
        safe = _safe_name(filename)
        input_path = job_dir / f"{index:03d}_{safe}"
        input_path.write_bytes(data)
        saved.append({"name": filename, "path": input_path, "size": len(data)})

    with _lock:
        _jobs[job_id] = {
            "job_id": job_id,
            "input_files": saved,
            "status": "queued",
            "progress": 0,
            "message": "Files ready",
            "stats": _empty_stats(len(saved)),
            "recent_urls": [],
            "results": None,
            "error": None,
            "output_path": None,
            "output_name": None,
            "job_dir": str(job_dir),
        }

    return {"job_id": job_id, "files": [{"name": item["name"], "size": item["size"]} for item in saved]}


@app.post("/process")
def process(body: ProcessBody):
    with _lock:
        job = _jobs.get(body.job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Job not found")
        if job["status"] == "running":
            return {"job_id": body.job_id}
        job["status"] = "running"
        job["message"] = "Opening files..."
    thread = threading.Thread(target=_run_job, args=(body.job_id,), daemon=True)
    thread.start()
    return {"job_id": body.job_id}


@app.get("/status/{job_id}")
def status(job_id: str):
    with _lock:
        job = _jobs.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Job not found")
        return _snapshot(job)


@app.get("/download/{job_id}")
def download(job_id: str):
    with _lock:
        job = _jobs.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Job not found")
        if job["status"] != "completed" or not job.get("output_path"):
            raise HTTPException(status_code=409, detail="The Excel file is not ready yet.")
        output_path = Path(job["output_path"])
        output_name = job.get("output_name") or output_path.name
        job_dir = job.get("job_dir")
    if not output_path.exists():
        raise HTTPException(status_code=404, detail="The Excel file could not be found.")
    return FileResponse(
        path=output_path,
        filename=output_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        background=BackgroundTask(_cleanup_job_dir, job_dir) if job_dir else None,
    )


def _attach_web_ui() -> None:
    dist = ROOT / "static"
    if not dist.is_dir():
        @app.get("/")
        def ui_missing():
            return {"message": "Web UI files are missing from the static/ folder.", "api": "The PDF URL capture API is running."}
        return

    assets = dist / "assets"
    if assets.is_dir():
        app.mount("/assets", StaticFiles(directory=str(assets)), name="assets")

    @app.get("/")
    def ui_index():
        return FileResponse(dist / "index.html")

    @app.get("/{full_path:path}")
    def ui_static(full_path: str):
        target = (dist / full_path).resolve()
        if dist.resolve() not in target.parents and target != dist.resolve():
            raise HTTPException(status_code=404, detail="Not found")
        if target.is_file():
            return FileResponse(target)
        return FileResponse(dist / "index.html")


_attach_web_ui()


def _workbench_http_options() -> tuple[str, int, str, bool]:
    import os
    from argparse import ArgumentParser

    parser = ArgumentParser(add_help=False)
    parser.add_argument("--anaconda-project-host", action="append", default=[])
    parser.add_argument("--anaconda-project-port", type=int, default=None)
    parser.add_argument("--anaconda-project-iframe-hosts", action="append")
    parser.add_argument("--anaconda-project-no-browser", action="store_true")
    parser.add_argument("--anaconda-project-use-xheaders", action="store_true")
    parser.add_argument("--anaconda-project-url-prefix", default="")
    parser.add_argument("--anaconda-project-address", default="0.0.0.0")
    args, _unknown = parser.parse_known_args()
    address = args.anaconda_project_address or "0.0.0.0"
    port = args.anaconda_project_port
    if port is None:
        port = int(os.environ.get("PORT", "4747"))
    prefix = (args.anaconda_project_url_prefix or "").rstrip("/")
    return address, port, prefix, bool(args.anaconda_project_use_xheaders)


if __name__ == "__main__":
    import uvicorn

    address, port, url_prefix, _use_xheaders = _workbench_http_options()
    serve_app = app
    if url_prefix:
        wrapper = FastAPI()
        wrapper.add_middleware(
            CORSMiddleware,
            allow_origins=["*"],
            allow_methods=["*"],
            allow_headers=["*"],
            expose_headers=["Content-Disposition"],
        )
        wrapper.mount(url_prefix, app)
        serve_app = wrapper
    uvicorn.run(serve_app, host=address, port=port, proxy_headers=True, forwarded_allow_ips="*")
