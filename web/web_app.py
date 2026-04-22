"""
SBA File Renamer — Flask Web Backend
Run:  python web_app.py
Then open http://localhost:5000 in a browser on the school network.
"""

import os
import tempfile
import zipfile
from pathlib import Path
from flask import (Flask, render_template, request, send_file,
                   jsonify, redirect, url_for)
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.engine import (
    CSEC_SUBJECTS, CSEC_SUBJECT_NAMES,
    extract_title_page_text, extract_candidate_info,
    process_job, RenameJob, validate_candidate_number, build_filename,
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024   # 64 MB max upload
UPLOAD_FOLDER = Path(tempfile.gettempdir()) / "sba_renamer_uploads"
OUTPUT_FOLDER = Path(tempfile.gettempdir()) / "sba_renamer_outputs"
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html",
                           subjects=CSEC_SUBJECT_NAMES,
                           subject_codes=CSEC_SUBJECTS)


@app.route("/api/scan", methods=["POST"])
def api_scan():
    """Accept a PDF upload, return detected candidate number & name."""
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400
    f = request.files["file"]
    tmp = UPLOAD_FOLDER / f.filename
    f.save(tmp)
    try:
        text = extract_title_page_text(tmp)
        info = extract_candidate_info(text)
        return jsonify(info)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        tmp.unlink(missing_ok=True)


@app.route("/api/rename", methods=["POST"])
def api_rename():
    """
    Accept one or more PDFs + form fields.
    Returns a ZIP of renamed files (or a single file if only one).
    """
    files = request.files.getlist("files")
    candidate_number = request.form.get("candidate_number", "").strip()
    subject          = request.form.get("subject", "").strip()
    file_type        = request.form.get("file_type", "SBA").strip()
    doc_number       = int(request.form.get("doc_number", 1))

    if not files or not candidate_number or not subject:
        return jsonify({"error": "Missing required fields"}), 400

    ok, err = validate_candidate_number(candidate_number)
    if not ok:
        return jsonify({"error": err}), 400

    moderation_code = CSEC_SUBJECTS.get(subject)
    if not moderation_code:
        return jsonify({"error": f"Unknown subject: {subject}"}), 400

    saved = []
    for i, f in enumerate(files):
        src = UPLOAD_FOLDER / f.filename
        f.save(src)
        doc_num = doc_number + i if file_type in ("SBA", "Mark Scheme") else doc_number
        job = RenameJob(
            source_path=src,
            candidate_number=candidate_number,
            moderation_code=moderation_code,
            file_type=file_type,
            doc_number=doc_num,
            output_dir=OUTPUT_FOLDER,
        )
        result = process_job(job)
        if result.ok:
            saved.append(result.dest)
        src.unlink(missing_ok=True)

    if not saved:
        return jsonify({"error": "No files were successfully renamed"}), 500

    if len(saved) == 1:
        response = send_file(saved[0], as_attachment=True, download_name=saved[0].name)
        saved[0].unlink(missing_ok=True)
        return response

    # Multiple files → ZIP
    zip_path = OUTPUT_FOLDER / f"{candidate_number}_SBA_files.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        for p in saved:
            zf.write(p, p.name)
            p.unlink(missing_ok=True)
    response = send_file(zip_path, as_attachment=True, download_name=zip_path.name)
    zip_path.unlink(missing_ok=True)
    return response


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
