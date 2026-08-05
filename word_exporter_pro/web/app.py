"""
Word Page Exporter Pro - Web Application (Flask)
Browser interface for the high-fidelity page range export engine.

Run with:
    python run_web.py
"""

import io
import multiprocessing
import os
import queue
import tempfile
import threading
import uuid
import zipfile

from flask import Flask, jsonify, make_response, render_template, request, send_file

from word_exporter_pro.core.com_engine import DocumentInspector, aw
from word_exporter_pro.core.pdf_engine import PdfInspector
from word_exporter_pro.core.preview import ensure_preview_async, render_page_preview
from word_exporter_pro.core.batch_processor import ExportJobConfig
from word_exporter_pro.core.naming_formatter import NamingFormatter
from word_exporter_pro.web.job_manager import JobManager
from word_exporter_pro.utils.logger import get_logger

logger = get_logger()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WEB_DATA_DIR = os.getenv(
    "WORD_EXPORTER_WEB_DATA_DIR",
    os.path.join(tempfile.gettempdir(), "word_exporter_pro_web"),
)
UPLOAD_DIR = os.path.join(WEB_DATA_DIR, "uploads")
OUTPUT_DIR = os.path.join(WEB_DATA_DIR, "outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

ALLOWED_EXTENSIONS = {".docx", ".doc", ".docm", ".dotx", ".dotm", ".rtf", ".pdf"}
ALLOWED_FORMATS = ["docx", "pdf", "doc", "rtf", "docm"]

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 512 * 1024 * 1024  # 512 MB upload cap
job_manager = JobManager()
application = app

# Large Word documents can take longer to paginate than a reverse proxy allows
# for one HTTP request. Keep inspection work out of that request path and let
# the browser poll a lightweight status endpoint instead.
_inspection_jobs: dict[str, dict] = {}
_inspection_jobs_lock = threading.Lock()

if aw is None:
    logger.warning(
        "Aspose.Words is not available. Word-document exports will be rejected; "
        "check the Render build log for the aspose-words installation."
    )
else:
    logger.info("Aspose.Words is available for Linux Word-document pagination.")


def _inspect_document_process(path: str, result_queue) -> None:
    """Run CPU-intensive page layout outside the Gunicorn worker process."""
    try:
        if os.path.splitext(path)[1].lower() == ".pdf":
            info = PdfInspector.get_info(path)
        else:
            info = DocumentInspector.get_info(path)
        result_queue.put({"status": "done", "info": info})
    except Exception as e:
        result_queue.put({"status": "error", "error": str(e)})


def _is_allowed(filename: str) -> bool:
    return os.path.splitext(filename)[1].lower() in ALLOWED_EXTENSIONS


def _job_file(job, filename: str) -> str:
    """Resolve an output path while keeping it inside the job folder."""
    filepath = os.path.abspath(os.path.join(job.output_dir, filename))
    if os.path.commonpath((job.output_dir, filepath)) != job.output_dir:
        raise FileNotFoundError(filename)
    return filepath


def _store_upload(file_storage) -> dict:
    """Stores an uploaded file under its original name in a unique subfolder."""
    safe_name = os.path.basename(file_storage.filename)
    subdir = os.path.join(UPLOAD_DIR, uuid.uuid4().hex[:8])
    os.makedirs(subdir, exist_ok=True)
    dest = os.path.join(subdir, safe_name)
    file_storage.save(dest)
    return {"name": safe_name, "size": os.path.getsize(dest)}


def _resolve_upload(name: str) -> str:
    """Resolves an uploaded filename or direct server file path to its stored path."""
    if not name:
        raise FileNotFoundError()
    
    # 1. Direct absolute or relative file path on server filesystem
    if os.path.exists(name) and os.path.isfile(name):
        return os.path.abspath(name)

    # 2. Otherwise, treat it as an uploaded file basename.
    # Normalize separators to strip Windows backslashes even if running on a Linux server.
    base = name.replace("\\", "/").split("/")[-1]
    
    direct = os.path.join(UPLOAD_DIR, base)
    if os.path.isfile(direct):
        return direct
    best, best_mtime = None, 0.0
    for entry in os.listdir(UPLOAD_DIR):
        sub = os.path.join(UPLOAD_DIR, entry)
        if os.path.isdir(sub):
            candidate = os.path.join(sub, base)
            if os.path.isfile(candidate):
                mtime = os.path.getmtime(candidate)
                if mtime >= best_mtime:
                    best, best_mtime = candidate, mtime
    if best:
        return best
    raise FileNotFoundError(name)


@app.route("/")
def index():
    try:
        import pythoncom
        import win32com.client
        is_windows_com = (pythoncom is not None and win32com is not None)
    except ImportError:
        is_windows_com = False
    return render_template(
        "index.html",
        is_windows_com=is_windows_com
    )


@app.route("/api/upload", methods=["POST"])
def api_upload():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files provided."}), 400

    saved = []
    for f in files:
        if not f or not f.filename:
            continue
        if not _is_allowed(f.filename):
            return jsonify({
                "error": f"Unsupported file type: {f.filename}. "
                         f"Allowed: {sorted(e.lstrip('.') for e in ALLOWED_EXTENSIONS)}"
            }), 400
        saved.append(_store_upload(f))

    return jsonify({"files": saved})


@app.route("/api/clear-storage", methods=["POST"])
def api_clear_storage():
    """Deletes all uploaded documents stored in the server's temporary upload path."""
    import shutil
    reclaimed_bytes = 0
    file_count = 0

    if os.path.exists(UPLOAD_DIR):
        for root, dirs, files in os.walk(UPLOAD_DIR):
            for f in files:
                fp = os.path.join(root, f)
                try:
                    reclaimed_bytes += os.path.getsize(fp)
                    os.remove(fp)
                    file_count += 1
                except Exception as e:
                    logger.warning(f"Could not remove file '{fp}': {e}")
        for item in os.listdir(UPLOAD_DIR):
            ip = os.path.join(UPLOAD_DIR, item)
            if os.path.isdir(ip):
                try:
                    shutil.rmtree(ip, ignore_errors=True)
                except Exception:
                    pass

    os.makedirs(UPLOAD_DIR, exist_ok=True)
    reclaimed_mb = round(reclaimed_bytes / (1024 * 1024), 2)
    logger.info(f"Server document storage cleared: {file_count} file(s), {reclaimed_mb} MB freed from '{UPLOAD_DIR}'")
    return jsonify({
        "message": "Server document storage cleared successfully.",
        "file_count": file_count,
        "reclaimed_bytes": reclaimed_bytes,
        "reclaimed_mb": reclaimed_mb
    })


@app.route("/api/inspect", methods=["POST"])
def api_inspect():
    data = request.get_json(silent=True) or {}
    name = str(data.get("name", ""))
    try:
        path = _resolve_upload(name)
    except FileNotFoundError:
        return jsonify({"error": f"Uploaded file not found: {name}"}), 404
    key = os.path.abspath(path)
    with _inspection_jobs_lock:
        existing = _inspection_jobs.get(key)
        if existing and existing["status"] == "done":
            return jsonify({"status": "done", "info": existing["info"]})
        if existing and existing["status"] == "running":
            return jsonify({"status": "running"}), 202
        _inspection_jobs[key] = {"status": "running"}

    result_queue = queue.Queue(maxsize=1)
    thread = threading.Thread(
        target=_inspect_document_process,
        args=(path, result_queue),
        daemon=True,
    )
    thread.start()

    def collect_result():
        try:
            result = result_queue.get()
            with _inspection_jobs_lock:
                _inspection_jobs[key] = result
        except Exception as e:
            logger.error(f"Background inspect result failed via web: {e}")
            with _inspection_jobs_lock:
                _inspection_jobs[key] = {"status": "error", "error": str(e)}

    threading.Thread(target=collect_result, daemon=True).start()
    return jsonify({"status": "running"}), 202


@app.route("/api/inspect/<path:name>")
def api_inspect_status(name):
    """Return the state of a background document-inspection request."""
    try:
        path = _resolve_upload(name)
    except FileNotFoundError:
        return jsonify({"error": "Uploaded file not found."}), 404
    with _inspection_jobs_lock:
        result = _inspection_jobs.get(os.path.abspath(path))
        if result is None:
            return jsonify({"error": "Inspection has not been started."}), 404
        if result["status"] == "done":
            return jsonify({"status": "done", "info": result["info"]})
        if result["status"] == "error":
            return jsonify({"status": "error", "error": result["error"]}), 500
    return jsonify({"status": "running"}), 202


@app.route("/api/naming-preview", methods=["POST"])
def api_naming_preview():
    data = request.get_json(silent=True) or {}
    pattern = str(data.get("pattern", "")).strip() or NamingFormatter.DEFAULT_PATTERN
    fmt = str(data.get("format", "docx")).strip() or "docx"
    sample = str(data.get("sample", "SampleReport.docx")) or "SampleReport.docx"
    try:
        name = NamingFormatter.generate_filename(
            pattern=pattern,
            original_filepath=sample,
            page_range=(1, 3),
            total_pages=10,
            output_ext=fmt,
            batch_index=1,
        )
        return jsonify({"preview": name})
    except Exception as e:
        return jsonify({"error": str(e)}), 400


def _preview_response(path: str):
    """Renders one page of a document and returns an image/png response."""
    try:
        page = max(1, int(request.args.get("page", 1)))
    except ValueError:
        page = 1
    try:
        width = max(300, min(2400, int(request.args.get("w", 1000))))
    except ValueError:
        width = 1000
    try:
        data, total, mime_type = render_page_preview(path, page=page, max_width=width)
    except Exception as e:
        logger.error(f"Preview failed via web: {e}")
        return jsonify({"error": str(e)}), 500
    resp = make_response(data)
    resp.headers["Content-Type"] = mime_type
    resp.headers["X-Total-Pages"] = str(total)
    return resp


@app.route("/api/preview/<path:name>")
def api_preview(name):
    try:
        path = _resolve_upload(name)
    except FileNotFoundError:
        return jsonify({"error": "Uploaded file not found."}), 404

    # Word files need a preview PDF first; generate it in the background
    # and let the client poll until it is ready.
    if not path.lower().endswith(".pdf") and not ensure_preview_async(path):
        return jsonify({"status": "generating"}), 202
    return _preview_response(path)


@app.route("/api/output-preview/<job_id>/<path:filename>")
def api_output_preview(job_id, filename):
    job = job_manager.get(job_id)
    if not job:
        return jsonify({"error": "Job not found."}), 404
    try:
        filepath = _job_file(job, filename)
    except FileNotFoundError:
        return jsonify({"error": "Output file not found."}), 404
    if not os.path.isfile(filepath):
        return jsonify({"error": f"Output file not found: {filename}"}), 404
    return _preview_response(filepath)


@app.route("/api/export", methods=["POST"])
def api_export():
    data = request.get_json(silent=True) or {}
    names = data.get("files", [])
    if not names:
        return jsonify({"error": "No files selected for export."}), 400

    paths = []
    for n in names:
        try:
            paths.append(_resolve_upload(str(n)))
        except FileNotFoundError:
            continue
    if not paths:
        return jsonify({"error": "None of the selected files exist on the server."}), 400

    export_format = str(data.get("format", "docx")).lower()
    if export_format not in ALLOWED_FORMATS:
        return jsonify({"error": f"Unsupported export format '{export_format}'."}), 400

    engine_mode = str(data.get("engine_mode", "trimming"))
    if engine_mode not in ("trimming", "aspose", "selection"):
        engine_mode = "trimming"

    # A remote browser cannot write to a local user folder.  Keep exports in
    # a server-owned temporary folder and expose them through download routes.
    output_dir = os.path.join(OUTPUT_DIR, uuid.uuid4().hex)

    naming_pattern = str(data.get("naming_pattern", "")).strip() or NamingFormatter.DEFAULT_PATTERN

    config = ExportJobConfig(
        source_files=paths,
        range_expression=str(data.get("range", "1-end")),
        output_dir=output_dir,
        export_format=export_format,
        naming_pattern=naming_pattern,
        overwrite=bool(data.get("overwrite", False)),
        engine_mode=engine_mode,
        visible=bool(data.get("visible", False)),
        clear_storage_after_export=bool(data.get("clear_storage_after_export", False)),
    )

    job = job_manager.create(config)
    job_manager.start(job)
    logger.info(f"Web export job started: {job.job_id} ({len(paths)} file(s))")
    return jsonify({"job_id": job.job_id}), 202


@app.route("/api/job/<job_id>")
def api_job(job_id):
    job = job_manager.get(job_id)
    if not job:
        return jsonify({"error": "Job not found."}), 404
    return jsonify(job.snapshot())


@app.route("/api/job/<job_id>/cancel", methods=["POST"])
def api_job_cancel(job_id):
    job = job_manager.get(job_id)
    if not job:
        return jsonify({"error": "Job not found."}), 404
    job_manager.cancel(job)
    return jsonify({"job_id": job_id, "status": "cancelling"})


@app.route("/api/download/<job_id>/<path:filename>")
def api_download(job_id, filename):
    job = job_manager.get(job_id)
    if not job:
        return jsonify({"error": "Job not found."}), 404
    try:
        filepath = _job_file(job, filename)
    except FileNotFoundError:
        return jsonify({"error": "Output file not found."}), 404
    if not os.path.isfile(filepath):
        return jsonify({"error": f"Output file not found: {filename}"}), 404
    return send_file(filepath, as_attachment=True, download_name=os.path.basename(filepath))


@app.route("/api/download/<job_id>/zip")
def api_download_zip(job_id):
    job = job_manager.get(job_id)
    if not job:
        return jsonify({"error": "Job not found."}), 404
    with job._lock:
        outputs = list(job.outputs)
    if not outputs:
        return jsonify({"error": "No output files for this job."}), 404

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in outputs:
            try:
                filepath = _job_file(job, name)
            except FileNotFoundError:
                continue
            if os.path.isfile(filepath):
                zf.write(filepath, arcname=name)
    buf.seek(0)

    zip_name = f"word_pdf_exports_{job.job_id}.zip"
    return send_file(
        buf,
        as_attachment=True,
        download_name=zip_name,
        mimetype="application/zip",
    )


def main(host: str | None = None, port: int | None = None, debug: bool | None = None):
    host = host or os.getenv("HOST", "0.0.0.0")
    port = int(port or os.getenv("PORT", "8000"))
    debug = debug if debug is not None else os.getenv("DEBUG", "False").lower() in {"1", "true", "yes", "on"}

    print("=" * 60)
    print(" Microsoft Word & PDF Page Exporter Pro - Web Interface")
    print("=" * 60)
    print(f" Local:   http://127.0.0.1:{port}")
    try:
        import socket
        for ip in socket.gethostbyname_ex(socket.gethostname())[2]:
            if ip and not ip.startswith("127."):
                print(f" Network: http://{ip}:{port}  (open from other devices)")
    except Exception:
        pass
    print(" Press Ctrl+C to stop the server.")
    print("=" * 60)
    app.run(host=host, port=port, debug=debug, threaded=True, use_reloader=False)


if __name__ == "__main__":
    main()
