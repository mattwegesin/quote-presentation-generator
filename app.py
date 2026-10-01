import os
import io
import logging
from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify
from werkzeug.utils import secure_filename
from generate_quote_deck import generate_presentation

# Configure structured logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
logger = logging.getLogger("quote_deck_web")

app = Flask(__name__)
# Secret key for session/flashed messages
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "prod-quote-deck-secret-key-change-in-env")

# Upload limits and allowed extensions
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB
ALLOWED_PPTX_EXTENSIONS = {".pptx"}
ALLOWED_EXCEL_EXTENSIONS = {".xlsx", ".xls"}

def has_allowed_extension(filename, allowed_extensions):
    if not filename or "." not in filename:
        return False
    _, ext = os.path.splitext(filename)
    return ext.lower() in allowed_extensions

@app.route("/", methods=["GET"])
def index():
    """Serves the main presentation upload interface."""
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    """
    Handles presentation generation:
    1. Validates base presentation (.pptx) and quote data (.xlsx) uploads.
    2. Ingests uploaded files into memory streams (io.BytesIO) avoiding disk writes.
    3. Synthesizes slide deck with financial calculations, summary cards, and disclaimers.
    4. Directly streams compiled PPTX back to the client as an attachment.
    """
    # Verify both file parts exist in multipart request
    if "base_pptx" not in request.files or "quote_excel" not in request.files:
        msg = "Both 'Base Presentation (.pptx)' and 'Quote Data (.xlsx)' files are required."
        logger.warning(msg)
        if request.headers.get("X-Requested-With") == "XMLHttpRequest" or request.accept_mimetypes.accept_json:
            return jsonify({"success": False, "error": msg}), 400
        flash(msg, "error")
        return redirect(url_for("index"))

    base_pptx_file = request.files["base_pptx"]
    quote_excel_file = request.files["quote_excel"]

    # Verify filenames
    if not base_pptx_file.filename or not quote_excel_file.filename:
        msg = "Please select both files before clicking Generate Presentation."
        logger.warning(msg)
        if request.headers.get("X-Requested-With") == "XMLHttpRequest" or request.accept_mimetypes.accept_json:
            return jsonify({"success": False, "error": msg}), 400
        flash(msg, "error")
        return redirect(url_for("index"))

    # Validate file extensions
    if not has_allowed_extension(base_pptx_file.filename, ALLOWED_PPTX_EXTENSIONS):
        msg = f"Invalid base presentation format '{base_pptx_file.filename}'. Only .pptx files are supported."
        logger.warning(msg)
        if request.headers.get("X-Requested-With") == "XMLHttpRequest" or request.accept_mimetypes.accept_json:
            return jsonify({"success": False, "error": msg}), 400
        flash(msg, "error")
        return redirect(url_for("index"))

    if not has_allowed_extension(quote_excel_file.filename, ALLOWED_EXCEL_EXTENSIONS):
        msg = f"Invalid quote data format '{quote_excel_file.filename}'. Only .xlsx or .xls files are supported."
        logger.warning(msg)
        if request.headers.get("X-Requested-With") == "XMLHttpRequest" or request.accept_mimetypes.accept_json:
            return jsonify({"success": False, "error": msg}), 400
        flash(msg, "error")
        return redirect(url_for("index"))

    try:
        # Read uploaded files directly into in-memory BytesIO buffers
        logger.info(
            "Buffering files in memory: base='%s', quote='%s'",
            secure_filename(base_pptx_file.filename),
            secure_filename(quote_excel_file.filename)
        )
        pptx_stream = io.BytesIO(base_pptx_file.read())
        excel_stream = io.BytesIO(quote_excel_file.read())

        # Generate the compiled presentation in-memory
        output_stream = generate_presentation(pptx_stream, excel_stream)

        logger.info("Successfully compiled proposal deck. Streaming 'Generated_Proposal.pptx' to client.")

        return send_file(
            output_stream,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name="Generated_Proposal.pptx"
        )

    except Exception as exc:
        logger.exception("Failed to generate proposal presentation: %s", exc)
        err_msg = f"Failed to generate proposal deck: {str(exc)}"
        if request.headers.get("X-Requested-With") == "XMLHttpRequest" or request.accept_mimetypes.accept_json:
            return jsonify({"success": False, "error": err_msg}), 500
        flash(err_msg, "error")
        return redirect(url_for("index"))

@app.route("/health", methods=["GET"])
def health():
    """Health check endpoint for Render zero-downtime deploys and monitoring."""
    return jsonify({"status": "healthy", "service": "quote-presentation-generator"}), 200

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug)
