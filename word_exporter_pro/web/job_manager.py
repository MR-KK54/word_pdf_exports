"""
Word Page Exporter Pro - Web Job Manager
Tracks background batch export jobs with disk-backed state persistence for multi-worker support.
"""

import json
import os
import tempfile
import threading
import uuid
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from word_exporter_pro.core.batch_processor import BatchProcessor, ExportJobConfig
from word_exporter_pro.utils.logger import get_logger

logger = get_logger()

MAX_LOG_LINES = 500
STATE_DIR = os.path.join(tempfile.gettempdir(), "word_exporter_pro_jobs")
os.makedirs(STATE_DIR, exist_ok=True)


def _save_job_disk(snap: dict):
    try:
        j_path = os.path.join(STATE_DIR, f"{snap['job_id']}.json")
        with open(j_path, "w", encoding="utf-8") as f:
            json.dump(snap, f)
    except Exception as e:
        logger.warning(f"Could not save job disk state for '{snap.get('job_id')}': {e}")


def _load_job_disk(job_id: str) -> Optional[dict]:
    try:
        j_path = os.path.join(STATE_DIR, f"{job_id}.json")
        if os.path.exists(j_path):
            with open(j_path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception as e:
        logger.warning(f"Could not read job disk state for '{job_id}': {e}")
    return None


@dataclass
class WebJob:
    """Runtime state for a single web-triggered export job."""

    job_id: str
    config: ExportJobConfig
    processor: Optional[BatchProcessor] = None
    status: str = "queued"
    completed: int = 0
    total: int = 0
    current_status: str = "Waiting to start..."
    success_count: int = 0
    fail_count: int = 0
    errors: List[str] = field(default_factory=list)
    logs: List[dict] = field(default_factory=list)
    outputs: List[str] = field(default_factory=list)
    output_dir: str = ""
    _lock: threading.Lock = field(default_factory=threading.Lock)

    def snapshot(self) -> dict:
        with self._lock:
            return {
                "job_id": self.job_id,
                "status": self.status,
                "completed": self.completed,
                "total": self.total,
                "current_status": self.current_status,
                "success_count": self.success_count,
                "fail_count": self.fail_count,
                "errors": list(self.errors),
                "logs": list(self.logs),
                "outputs": list(self.outputs),
                "output_dir": self.output_dir,
            }


class JobManager:
    """Registry of running/completed web jobs with multi-worker disk persistence."""

    def __init__(self):
        self._jobs: Dict[str, WebJob] = {}
        self._lock = threading.Lock()

    def create(self, config: ExportJobConfig) -> WebJob:
        job = WebJob(
            job_id=uuid.uuid4().hex[:12],
            config=config,
            output_dir=os.path.abspath(config.output_dir),
        )
        with self._lock:
            self._jobs[job.job_id] = job
        _save_job_disk(job.snapshot())
        return job

    def get(self, job_id: str) -> Optional[WebJob]:
        with self._lock:
            job = self._jobs.get(job_id)
            if job:
                return job

        snap = _load_job_disk(job_id)
        if snap:
            fallback_job = WebJob(
                job_id=snap["job_id"],
                config=ExportJobConfig(source_files=[], range_expression="1-end", output_dir=snap.get("output_dir", "")),
                status=snap.get("status", "done"),
                completed=snap.get("completed", 0),
                total=snap.get("total", 0),
                current_status=snap.get("current_status", ""),
                success_count=snap.get("success_count", 0),
                fail_count=snap.get("fail_count", 0),
                errors=snap.get("errors", []),
                logs=snap.get("logs", []),
                outputs=snap.get("outputs", []),
                output_dir=snap.get("output_dir", ""),
            )
            with self._lock:
                self._jobs[job_id] = fallback_job
            return fallback_job
        return None

    def start(self, job: WebJob) -> None:
        def _run_starter():
            def on_progress(completed: int, total: int, filename: str, status: str):
                with job._lock:
                    job.status = "running"
                    job.completed = completed
                    job.total = total
                    job.current_status = status
                    _save_job_disk(job.snapshot())

            def on_finished(success: int, fail: int, errors: List[str]):
                with job._lock:
                    job.success_count = success
                    job.fail_count = fail
                    job.errors.extend(errors)
                    job.completed = job.total
                    cancelled = job.processor is not None and job.processor.cancel_event.is_set()
                    job.status = "cancelled" if cancelled else "done"
                    _save_job_disk(job.snapshot())
                get_logger().remove_listener(_listener)

            def _listener(timestamp: str, level: str, message: str):
                with job._lock:
                    job.logs.append({"time": timestamp, "level": level, "message": message})
                    if len(job.logs) > MAX_LOG_LINES:
                        job.logs = job.logs[-MAX_LOG_LINES:]
                    _save_job_disk(job.snapshot())

            def on_file_created(path: str):
                with job._lock:
                    name = os.path.basename(path)
                    if name not in job.outputs:
                        job.outputs.append(name)
                        _save_job_disk(job.snapshot())

            get_logger().add_listener(_listener)
            processor = BatchProcessor(job.config)
            with job._lock:
                job.processor = processor
            processor.start_async(on_progress, on_finished, on_file_created)

        threading.Thread(target=_run_starter, daemon=True).start()

    def cancel(self, job: WebJob) -> None:
        with job._lock:
            if job.status in ("queued", "running"):
                job.status = "cancelling"
                _save_job_disk(job.snapshot())
        if job.processor:
            job.processor.cancel()
