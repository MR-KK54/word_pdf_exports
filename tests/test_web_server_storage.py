import os

from word_exporter_pro.web import app as web_app
from word_exporter_pro.web.job_manager import WebJob
from word_exporter_pro.core.batch_processor import ExportJobConfig


def test_web_exports_use_an_isolated_server_temp_folder(tmp_path, monkeypatch):
    client = web_app.app.test_client()
    source = tmp_path / "source.pdf"
    source.write_bytes(b"%PDF-1.4\n%%EOF\n")

    upload_dir = tmp_path / "uploads"
    upload_dir.mkdir()
    uploaded = upload_dir / "source.pdf"
    uploaded.write_bytes(source.read_bytes())
    monkeypatch.setattr(web_app, "_resolve_upload", lambda _name: str(uploaded))
    monkeypatch.setattr(web_app.job_manager, "start", lambda _job: None)

    response = client.post(
        "/api/export",
        json={"files": ["source.pdf"], "output_dir": str(tmp_path / "client-folder")},
    )
    assert response.status_code == 202
    job = web_app.job_manager.get(response.get_json()["job_id"])

    assert job is not None
    assert os.path.basename(web_app.OUTPUT_DIR) == "outputs"
    assert os.path.commonpath((web_app.OUTPUT_DIR, job.output_dir)) == web_app.OUTPUT_DIR
    assert str(tmp_path / "client-folder") != job.output_dir


def test_job_file_rejects_path_traversal(tmp_path):
    output_dir = tmp_path / "job"
    output_dir.mkdir()
    config = ExportJobConfig([], "1", str(output_dir))
    job = WebJob("test", config, output_dir=str(output_dir))

    assert web_app._job_file(job, "result.pdf") == str(output_dir / "result.pdf")
    try:
        web_app._job_file(job, "../outside.pdf")
    except FileNotFoundError:
        pass
    else:
        raise AssertionError("Path traversal must be rejected")
