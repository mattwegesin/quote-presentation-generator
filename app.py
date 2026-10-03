import os
import io
import json
import logging
import requests
import tempfile
import subprocess
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

def convert_pptx_to_pdf(pptx_stream):
    """Converts a PPTX BytesIO stream to a PDF BytesIO stream using headless LibreOffice."""
    with tempfile.TemporaryDirectory() as temp_dir:
        input_path = os.path.join(temp_dir, "input.pptx")
        with open(input_path, "wb") as f:
            f.write(pptx_stream.read())
            
        logger.info("Running LibreOffice headless conversion to PDF...")
        # Note: In a Docker container running LibreOffice, this command converts the file
        cmd = [
            "libreoffice", "--headless", "--convert-to", "pdf",
            "--outdir", temp_dir, input_path
        ]
        
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except subprocess.CalledProcessError as e:
            logger.error(f"LibreOffice conversion failed: {e.stderr.decode('utf-8', errors='ignore')}")
            raise RuntimeError("PDF conversion failed.")
            
        output_path = os.path.join(temp_dir, "input.pdf")
        if not os.path.exists(output_path):
            raise FileNotFoundError("PDF file was not created by LibreOffice.")
            
        with open(output_path, "rb") as f:
            pdf_data = f.read()
            
    return io.BytesIO(pdf_data)

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
        property_code = request.form.get("property_code", "").strip()
        signer_name = request.form.get("signer_name", "").strip()
        signer_email = request.form.get("signer_email", "").strip()
        
        if not property_code or not signer_name or not signer_email:
            raise ValueError("Property Code, Signer Name, and Signer Email are required.")

        # Read uploaded files directly into in-memory BytesIO buffers
        logger.info(
            "Buffering files in memory: base='%s', quote='%s'",
            secure_filename(base_pptx_file.filename),
            secure_filename(quote_excel_file.filename)
        )
        pptx_stream = io.BytesIO(base_pptx_file.read())
        excel_stream = io.BytesIO(quote_excel_file.read())

        # Generate the compiled presentation in-memory
        output_stream = generate_presentation(pptx_stream, excel_stream, property_code=property_code)
        
        # Convert PPTX to PDF
        output_stream.seek(0)
        pdf_stream = convert_pptx_to_pdf(output_stream)

        # Send to BoldSign
        boldsign_api_key = os.environ.get("BOLDSIGN_API_KEY")
        if boldsign_api_key:
            logger.info(f"Sending document to BoldSign for {property_code}")
            pdf_stream.seek(0)
            
            headers = {
                'X-API-KEY': boldsign_api_key,
                'Accept': 'application/json'
            }
            
            files = {
                'Files': ('Generated_Proposal.pdf', pdf_stream.read(), 'application/pdf')
            }
            
            data = {
                'Title': f"Hospitality Technologies Agreement - {property_code}",
                'DisableEmails': 'true',
                'ExpiryDays': '14',
                'UseTextTags': 'true',
                'Signers[0][name]': signer_name,
                'Signers[0][emailAddress]': signer_email,
                'Signers[0][signerType]': 'Signer',
                'CustomField': f"PropertyCode={property_code}"
            }
            
            resp = requests.post("https://api.boldsign.com/v1/document/send", headers=headers, data=data, files=files)
            if resp.status_code not in (200, 201):
                logger.error(f"BoldSign API Error: {resp.status_code} - {resp.text}")
                raise RuntimeError(f"Failed to create BoldSign document: {resp.text}")
            else:
                logger.info(f"Successfully created BoldSign document: {resp.json().get('documentId')}")
            
            pdf_stream.seek(0)

        logger.info("Successfully compiled proposal deck. Streaming 'Generated_Proposal.pdf' to client.")

        return send_file(
            pdf_stream,
            mimetype="application/pdf",
            as_attachment=True,
            download_name="Generated_Proposal.pdf"
        )

    except Exception as exc:
        logger.exception("Failed to generate proposal presentation: %s", exc)
        err_msg = f"Failed to generate proposal deck: {str(exc)}"
        if request.headers.get("X-Requested-With") == "XMLHttpRequest" or request.accept_mimetypes.accept_json:
            return jsonify({"success": False, "error": err_msg}), 500
        flash(err_msg, "error")
        return redirect(url_for("index"))

@app.route("/sign/<property_code>", methods=["GET"])
def sign_document(property_code):
    """Dynamic gateway redirecting to BoldSign embedded signing link."""
    boldsign_api_key = os.environ.get("BOLDSIGN_API_KEY")
    if not boldsign_api_key:
        return "BoldSign integration not configured.", 500
        
    headers = {
        'X-API-KEY': boldsign_api_key,
        'Accept': 'application/json'
    }
    
    # Search for the document by title
    search_url = f"https://api.boldsign.com/v1/document/list?searchQuery={property_code}&pageSize=10"
    resp = requests.get(search_url, headers=headers)
    if resp.status_code != 200:
        return "Failed to search BoldSign documents.", 500
        
    data = resp.json()
    docs = data.get("result", [])
    if not docs:
        return "Document not found or has expired.", 404
        
    # Find the most recent matching document
    target_doc = None
    target_title = f"Hospitality Technologies Agreement - {property_code}"
    for doc in docs:
        if doc.get("messageTitle") == target_title and doc.get("status") not in ("Expired", "Completed", "Declined", "Revoked"):
            target_doc = doc
            break
            
    if not target_doc:
        return "No active document found for signing. It may have expired.", 404
        
    document_id = target_doc.get("documentId")
    signer_email = target_doc.get("signers")[0].get("signerEmail")
    
    # Get Embedded Sign Link
    link_url = f"https://api.boldsign.com/v1/document/getEmbeddedSignLink?documentId={document_id}&signerEmail={signer_email}"
    link_resp = requests.get(link_url, headers=headers)
    if link_resp.status_code != 200:
        return f"Failed to generate signing link: {link_resp.text}", 500
        
    sign_link = link_resp.json().get("signLink")
    if sign_link:
        # Redirect client directly into the signing interface
        return redirect(sign_link)
    
    return "Error generating signature link.", 500

@app.route("/api/boldsign-webhook", methods=["POST"])
def boldsign_webhook():
    """Receives completion events from BoldSign to sync with Monday.com and SharePoint."""
    # Since this is an MVP demonstration, we will acknowledge the webhook and log the event.
    # In a full production setup, this would execute the Monday.com GraphQL mutations 
    # and Microsoft Graph API PUT requests as detailed in the architecture spec.
    try:
        event = request.json
        if event and event.get("event") == "DocumentCompleted":
            doc_id = event.get("document", {}).get("documentId")
            logger.info(f"Webhook Received: DocumentCompleted for {doc_id}")
            # Placeholder for Monday.com Sync
            # Placeholder for MS Graph Archival
            return jsonify({"status": "acknowledged", "sync": "pending"}), 200
            
        return jsonify({"status": "ignored"}), 200
    except Exception as e:
        logger.error(f"Webhook processing error: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/health", methods=["GET"])
def health():
    """Health check endpoint for Render zero-downtime deploys and monitoring."""
    return jsonify({"status": "healthy", "service": "quote-presentation-generator"}), 200

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug)
