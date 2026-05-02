from __future__ import annotations

import os
from pathlib import Path

from flask import Flask, render_template, request

from membership_store import MembershipStore


def _database_path() -> Path:
    explicit = os.environ.get("MEMBERSHIP_DB_PATH")
    if explicit:
        return Path(explicit)
    data_dir = Path(os.environ.get("DATA_DIR", Path.cwd() / "data"))
    return data_dir / "membership.sqlite3"


app = Flask(__name__)
store = MembershipStore(_database_path())


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", db_count=store.count_records())


@app.route("/upload", methods=["POST"])
def upload_excel():
    file = request.files.get("excel_file")
    if not file or not file.filename:
        return render_template(
            "index.html",
            db_count=store.count_records(),
            upload_error="Select an Excel file to import.",
        )
    try:
        file.stream.seek(0)
        summary = store.sync_from_excel(file.stream)
    except Exception as exc:  # noqa: BLE001 - surface parsing errors to UI
        return render_template(
            "index.html",
            db_count=store.count_records(),
            upload_error=f"Could not load workbook: {exc}",
        )
    return render_template(
        "index.html",
        db_count=store.count_records(),
        sync_summary=summary,
    )


@app.route("/scan", methods=["POST"])
def scan_membership():
    scan_text = (request.form.get("scan_text") or "").strip()
    if not scan_text:
        return render_template(
            "index.html",
            db_count=store.count_records(),
            scan_error="Enter a membership number, email, or name to scan.",
        )
    results = store.lookup(scan_text)
    return render_template(
        "index.html",
        db_count=store.count_records(),
        scan_text=scan_text,
        results=results,
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")))
