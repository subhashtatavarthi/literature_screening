"""Flask web UI for the Literature Screening Pipeline."""

import os
import json
import shutil
import tempfile
from pathlib import Path

import pandas as pd
from flask import Flask, render_template, request, jsonify, send_from_directory

# ---------------------------------------------------------------------------
# Make the pipeline importable — set env vars before importing
# ---------------------------------------------------------------------------
os.environ.setdefault("UNPAYWALL_EMAIL", "your.email@example.com")

import literature_pipeline as pipeline

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB upload limit


@app.route("/")
def index():
    """Serve the main UI page."""
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    """Accept file upload, run the pipeline, and return JSON results."""
    # --- Validate inputs ---
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    uploaded = request.files["file"]
    if not uploaded.filename:
        return jsonify({"error": "No file selected"}), 400

    email = request.form.get("email", "").strip()
    if not email:
        return jsonify({"error": "Unpaywall email is required"}), 400

    sheet = request.form.get("sheet", "").strip() or None

    # --- Configure pipeline ---
    os.environ["UNPAYWALL_EMAIL"] = email
    pipeline.EMAIL = email
    pipeline.USER_AGENT = f"LiteratureScreeningTool/1.0 (mailto:{email})"
    pipeline.session.headers.update({"User-Agent": pipeline.USER_AGENT})

    if sheet:
        os.environ["LIT_INPUT_SHEET"] = sheet
        pipeline.INPUT_SHEET = sheet
    else:
        os.environ.pop("LIT_INPUT_SHEET", None)
        pipeline.INPUT_SHEET = None

    # --- Save uploaded file ---
    ext = uploaded.filename.rsplit(".", 1)[-1].lower()
    input_path = Path("input/literature.xlsx")
    input_path.parent.mkdir(parents=True, exist_ok=True)

    if ext == "csv":
        # Convert CSV → XLSX so the pipeline can read it
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as tmp:
            uploaded.save(tmp.name)
            df_csv = pd.read_csv(tmp.name)
            df_csv.to_excel(str(input_path), index=False)
            os.unlink(tmp.name)
    else:
        uploaded.save(str(input_path))

    pipeline.INPUT_FILE = str(input_path)

    # --- Run the pipeline ---
    try:
        pipeline.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        pipeline.PDF_DIR.mkdir(parents=True, exist_ok=True)

        df = pipeline.detect_duplicates(pipeline.read_input_excel())
        pipeline.create_screening_workbook(df)

        results = []
        for _, row in df.iterrows():
            try:
                result = pipeline.process_record(row)
            except Exception as exc:
                result = {
                    "UI": pipeline.normalize_ui(row[pipeline.UI_COLUMN]),
                    "Title": pipeline.clean_value(row[pipeline.TITLE_COLUMN]),
                    "DOI": pipeline.normalize_doi(row[pipeline.DOI_COLUMN]),
                    "Duplicate": bool(row["Duplicate"]),
                    "Status": "ERROR",
                    "OA Source": "",
                    "PDF URL": "",
                    "Publisher": "",
                    "Publisher URL": "",
                    "Purchase Cost": "",
                    "Notes": str(exc),
                }
            results.append(result)
            import time
            time.sleep(pipeline.REQUEST_DELAY)

        results_df = pd.DataFrame(results)
        results_df.to_excel(str(pipeline.LOG_OUTPUT), index=False)
        pipeline.create_purchase_list(results_df)
        pipeline.create_zip()

    except Exception as exc:
        return jsonify({"error": str(exc)}), 500

    # --- Build JSON response ---
    screening_df = df[
        [pipeline.UI_COLUMN, pipeline.TITLE_COLUMN, pipeline.DOI_COLUMN, "Duplicate", "Duplicate Reason"]
    ].copy()
    screening_data = _df_to_records(screening_df)

    log_data = _df_to_records(results_df)

    purchase_df = results_df[results_df["Status"] == "PURCHASE_REQUIRED"][
        ["UI", "Title", "DOI", "Publisher", "Publisher URL", "Purchase Cost"]
    ].copy()
    purchase_df.rename(columns={"Publisher URL": "Website/site link for purchase"}, inplace=True)
    purchase_data = _df_to_records(purchase_df)

    return jsonify({
        "screening": screening_data,
        "log": log_data,
        "purchase": purchase_data,
    })


def _df_to_records(df):
    """Convert a DataFrame to a list of dicts, replacing NaN with empty strings."""
    records = df.where(df.notna(), "").to_dict(orient="records")
    # Ensure booleans are preserved properly
    for rec in records:
        for key, val in rec.items():
            if isinstance(val, (bool,)):
                continue
            if val is True or val is False:
                rec[key] = bool(val)
    return records


@app.route("/download/<filename>")
def download(filename):
    """Download an output file."""
    allowed = {
        "Screening_Output.xlsx",
        "Download_Log.xlsx",
        "Purchase_List.xlsx",
        "Literature_PDFs.zip",
    }
    if filename not in allowed:
        return jsonify({"error": "File not found"}), 404

    output_dir = Path("output").resolve()
    filepath = output_dir / filename
    if not filepath.exists():
        return jsonify({"error": f"{filename} has not been generated yet"}), 404

    return send_from_directory(str(output_dir), filename, as_attachment=True)


if __name__ == "__main__":
    print("\n  📚 Literature Screening Pipeline — Web UI")
    print("  Open http://127.0.0.1:5000 in your browser\n")
    app.run(debug=True, host="127.0.0.1", port=5000)
