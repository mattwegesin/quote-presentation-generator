import os
import io
import json
import logging
import requests
import tempfile
import subprocess
from functools import wraps
from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify, Response
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
from generate_quote_deck import generate_presentation
from search_quotes import get_quotes_by_inn_code, mark_quote_won, get_quote_notes
from download_quote import download_ihg_quote_xlsx

# Load environment variables from .env if present
load_dotenv()

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

def check_auth(username, password):
    """This function is called to check if a username / password combination is valid."""
    # Attempt to load a dictionary of users from the APP_USERS environment variable
    app_users_str = os.environ.get("APP_USERS")
    
    if app_users_str:
        try:
            # Expecting format: {"matt": "pass123", "john": "pass456"}
            users_dict = json.loads(app_users_str)
            if username in users_dict:
                return users_dict[username] == password
            return False
        except json.JSONDecodeError:
            logger.error("APP_USERS environment variable is not valid JSON.")
            return False
            
    # Fallback to single-user mode if APP_USERS is not set
    expected_user = os.environ.get("APP_USERNAME", "admin")
    expected_pass = os.environ.get("APP_PASSWORD", "password")
    return username == expected_user and password == expected_pass

def authenticate():
    """Sends a 401 response that enables basic auth"""
    return Response(
        'Could not verify your access level for that URL.\n'
        'You have to login with proper credentials', 401,
        {'WWW-Authenticate': 'Basic realm="Login Required"'})

def requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()
        return f(*args, **kwargs)
    return decorated

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
            "libreoffice", 
            f"-env:UserInstallation=file://{temp_dir}/profile",
            "--headless", 
            "--convert-to", "pdf",
            "--outdir", temp_dir, 
            input_path
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
@requires_auth
def index():
    """Serves the main presentation upload interface."""
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
@requires_auth
def generate():
    """
    Handles presentation generation:
    1. Validates base presentation (.pptx) and quote data (.xlsx) uploads.
    2. Ingests uploaded files into memory streams (io.BytesIO) avoiding disk writes.
    3. Synthesizes slide deck with financial calculations, summary cards, and disclaimers.
    4. Directly streams compiled PPTX back to the client as an attachment.
    """
    # Verify file parts exist
    base_pptx_choice = request.form.get("base_pptx_choice", "custom")
    if base_pptx_choice == "custom" and "base_pptx" not in request.files:
        msg = "The 'Base Presentation (.pptx)' file is required for a custom upload."
        logger.warning(msg)
        if request.headers.get("X-Requested-With") == "XMLHttpRequest" or request.accept_mimetypes.accept_json:
            return jsonify({"success": False, "error": msg}), 400
        flash(msg, "error")
        return redirect(url_for("index"))
        
    internal_quote_id = request.form.get("internal_quote_id")
    
    quote_excel_file = None
    excel_filename = None
    
    if internal_quote_id:
        logger.info(f"Downloading internal quote ID {internal_quote_id}")
        filepath = download_ihg_quote_xlsx(internal_quote_id, output_dir="/tmp")
        if not filepath:
            msg = f"Failed to download internal quote {internal_quote_id}."
            logger.warning(msg)
            if request.headers.get("X-Requested-With") == "XMLHttpRequest" or request.accept_mimetypes.accept_json:
                return jsonify({"success": False, "error": msg}), 400
            flash(msg, "error")
            return redirect(url_for("index"))
            
        excel_filename = os.path.basename(filepath)
        quote_excel_file = open(filepath, 'rb')
    else:
        if "quote_excel" not in request.files or not request.files["quote_excel"].filename:
            msg = "The 'Quote Data (.xlsx)' file or an internal quote selection is required."
            logger.warning(msg)
            if request.headers.get("X-Requested-With") == "XMLHttpRequest" or request.accept_mimetypes.accept_json:
                return jsonify({"success": False, "error": msg}), 400
            flash(msg, "error")
            return redirect(url_for("index"))

        quote_excel_file = request.files["quote_excel"]
        excel_filename = quote_excel_file.filename

        # Validate excel extension
        if not has_allowed_extension(excel_filename, ALLOWED_EXCEL_EXTENSIONS):
            msg = f"Invalid quote data format '{excel_filename}'. Only .xlsx or .xls files are supported."
            logger.warning(msg)
            if request.headers.get("X-Requested-With") == "XMLHttpRequest" or request.accept_mimetypes.accept_json:
                return jsonify({"success": False, "error": msg}), 400
            flash(msg, "error")
            return redirect(url_for("index"))

    try:
        property_code = request.form.get("property_code", "").strip().upper()
        monday_item_id = request.form.get("monday_item_id", "").strip()
        monday_item_name = request.form.get("monday_item_name", "").strip()
        signer_name = request.form.get("signer_name", "").strip()
        signer_email = request.form.get("signer_email", "").strip()
        
        if not property_code or not signer_name or not signer_email:
            raise ValueError("Property Code, Signer Name, and Signer Email are required.")

        # Build dynamic filename
        import re
        from datetime import datetime
        safe_excel_filename = secure_filename(excel_filename)
        quote_match = re.search(r'quote-(\d+)', safe_excel_filename, re.IGNORECASE)
        quote_num = quote_match.group(1) if quote_match else "UNKNOWN"

        sales_owner = ""
        quote_date = datetime.now().strftime("%Y-%m-%d")
        
        if monday_item_name:
            parts = [p.strip() for p in monday_item_name.split('|')]
            if len(parts) >= 3:
                sales_owner = parts[2]
        
        filename_parts = [property_code, f"Proposal #{quote_num}"]
        if sales_owner and sales_owner != "N/A":
            filename_parts.append(sales_owner)
        if quote_date and quote_date != "N/A":
            filename_parts.append(quote_date)
            
        generated_filename = " - ".join(filename_parts) + ".pdf"

        # Check if GSA generation is requested
        generate_gsa = request.form.get("generate_gsa", "false").lower() == "true"
        gsa_notes = request.form.get("gsa_notes", "").strip()

        # Check if Cabling Proposal is requested
        generate_cabling = request.form.get("generate_cabling", "false").lower() == "true"
        cabling_monday_item_id = request.form.get("cabling_monday_item_id", "").strip()
        cabling_internal_quote_id = request.form.get("cabling_internal_quote_id", "").strip()

        # Load PPTX into stream
        pptx_stream = io.BytesIO()
        if base_pptx_choice == "custom":
            base_pptx_file = request.files.get("base_pptx")
            if not has_allowed_extension(base_pptx_file.filename, ALLOWED_PPTX_EXTENSIONS):
                msg = f"Invalid base presentation format '{base_pptx_file.filename}'. Only .pptx files are supported."
                logger.warning(msg)
                if request.headers.get("X-Requested-With") == "XMLHttpRequest" or request.accept_mimetypes.accept_json:
                    return jsonify({"success": False, "error": msg}), 400
                flash(msg, "error")
                return redirect(url_for("index"))
            logger.info("Buffering custom PPTX file in memory: '%s'", secure_filename(base_pptx_file.filename))
            pptx_stream.write(base_pptx_file.read())
        else:
            local_path = os.path.join(os.path.dirname(__file__), base_pptx_choice)
            if not os.path.exists(local_path):
                raise FileNotFoundError(f"Standard template {base_pptx_choice} not found on server.")
            logger.info("Loading standard PPTX template from disk: '%s'", base_pptx_choice)
            with open(local_path, "rb") as f:
                pptx_stream.write(f.read())
        
        pptx_stream.seek(0)
        
        logger.info("Buffering Quote Excel file in memory: '%s'", secure_filename(excel_filename))
        excel_stream = io.BytesIO(quote_excel_file.read())

        custom_notes = None
        if internal_quote_id:
            fetched_notes = get_quote_notes(internal_quote_id)
            if fetched_notes:
                custom_notes = [note.strip() for note in fetched_notes.split('\n') if note.strip()]

        import concurrent.futures
        executor = concurrent.futures.ThreadPoolExecutor(max_workers=6)
        
        # ---------------------------------------------------------
        # GENERATE FOR BOLDSIGN (No CTA Button, Has Signature Lines)
        # ---------------------------------------------------------
        output_stream_bs, target_page, grand_total = generate_presentation(pptx_stream, excel_stream, property_code=property_code, for_boldsign=True, signer_name=signer_name, custom_notes=custom_notes)
        output_stream_bs.seek(0)
        future_main_bs = executor.submit(convert_pptx_to_pdf, output_stream_bs)

        # ---------------------------------------------------------
        # GENERATE FOR CLIENT (Has CTA Button, No Signature Lines)
        # ---------------------------------------------------------
        pptx_stream.seek(0)
        excel_stream.seek(0)
        output_stream_client, _, _ = generate_presentation(pptx_stream, excel_stream, property_code=property_code, for_boldsign=False, signer_name=signer_name, custom_notes=custom_notes)
        output_stream_client.seek(0)
        future_main_client = executor.submit(convert_pptx_to_pdf, output_stream_client)

        if monday_item_id and quote_num != "UNKNOWN":
            logger.info(f"Updating Monday.com item {monday_item_id} with quote number and total.")
            monday_token = os.environ.get("MONDAY_API_TOKEN")
            if monday_token:
                mutation = """
                mutation ($itemId: ID!, $boardId: ID!, $colValues: JSON!) {
                  change_multiple_column_values(item_id: $itemId, board_id: $boardId, column_values: $colValues) {
                    id
                  }
                }
                """
                col_vals = {
                    "text_mm8v74nz": quote_num,
                    "numeric_mm8vcbpj": grand_total
                }
                try:
                    import json
                    headers_mon = {
                        "Authorization": monday_token,
                        "API-Version": "2024-01",
                        "Content-Type": "application/json"
                    }
                    resp_mon = requests.post("https://api.monday.com/v2", headers=headers_mon, json={
                        "query": mutation,
                        "variables": {
                            "itemId": monday_item_id,
                            "boardId": "18424851511",
                            "colValues": json.dumps(col_vals)
                        }
                    })
                    if "errors" in resp_mon.json():
                        logger.error(f"Error updating Monday quote columns: {resp_mon.json()['errors']}")
                except Exception as e:
                    logger.error(f"Failed to update Monday quote columns: {e}")

        # ---------------------------------------------------------
        # GENERATE CABLING PROPOSAL IF REQUESTED
        # ---------------------------------------------------------
        cabling_client_stream = None
        cabling_bs_stream = None
        cabling_target_page = None
        if generate_cabling and cabling_internal_quote_id:
            logger.info(f"Cabling Proposal requested for {property_code}. Downloading quote {cabling_internal_quote_id}.")
            cabling_excel_path = download_ihg_quote_xlsx(cabling_internal_quote_id, output_dir="/tmp")
            if cabling_excel_path:
                with open(cabling_excel_path, 'rb') as f:
                    cabling_excel_stream = io.BytesIO(f.read())
                
                cabling_pptx_path = os.path.join(os.path.dirname(__file__), "Cabling_Proposal_Template.pptx")
                if os.path.exists(cabling_pptx_path):
                    with open(cabling_pptx_path, "rb") as f:
                        cabling_pptx_stream = io.BytesIO(f.read())
                    
                    cabling_custom_notes = None
                    cabling_fetched_notes = get_quote_notes(cabling_internal_quote_id)
                    if cabling_fetched_notes:
                        cabling_custom_notes = [note.strip() for note in cabling_fetched_notes.split('\n') if note.strip()]

                    # Generate for BoldSign
                    cabling_output_bs, cabling_target_page, _ = generate_presentation(
                        cabling_pptx_stream, cabling_excel_stream, property_code=property_code, 
                        for_boldsign=True, signer_name=signer_name, custom_notes=cabling_custom_notes
                    )
                    cabling_output_bs.seek(0)
                    future_cabling_bs = executor.submit(convert_pptx_to_pdf, cabling_output_bs)

                    # Generate for Client
                    cabling_pptx_stream.seek(0)
                    cabling_excel_stream.seek(0)
                    cabling_output_client, _, _ = generate_presentation(
                        cabling_pptx_stream, cabling_excel_stream, property_code=property_code, 
                        for_boldsign=False, signer_name=signer_name, custom_notes=cabling_custom_notes
                    )
                    cabling_output_client.seek(0)
                    future_cabling_client = executor.submit(convert_pptx_to_pdf, cabling_output_client)
                else:
                    logger.error("Cabling_Proposal_Template.pptx not found on server.")
            else:
                logger.error(f"Failed to download Cabling Quote {cabling_internal_quote_id}.")

        # ---------------------------------------------------------
        # GENERATE GSA LOCALLY IF REQUESTED
        # ---------------------------------------------------------
        gsa_client_stream = None
        gsa_bs_stream = None
        if generate_gsa:
            logger.info(f"GSA generation requested for {property_code}. Extracting details.")
            hotel_name = property_code
            room_count = 0
            address = ""
            
            # Fetch Monday details for Address
            if monday_item_id:
                monday_token = os.environ.get("MONDAY_API_TOKEN")
                if monday_token:
                    monday_headers = {
                        "Authorization": monday_token, 
                        "API-Version": "2024-01", 
                        "Content-Type": "application/json"
                    }
                    query = """
                    query ($itemId: [ID!]) {
                      items(ids: $itemId) {
                        column_values {
                          column { title }
                          text
                        }
                      }
                    }
                    """
                    try:
                        item_resp = requests.post("https://api.monday.com/v2", headers=monday_headers, json={
                            "query": query, "variables": {"itemId": [monday_item_id]}
                        })
                        if item_resp.status_code == 200:
                            cols = item_resp.json().get("data", {}).get("items", [{}])[0].get("column_values", [])
                            for col in cols:
                                title = col.get("column", {}).get("title")
                                if title == "Property Address" and col.get("text"):
                                    address = col.get("text")
                    except Exception as e:
                        logger.error(f"Failed to fetch Monday details for GSA: {e}")
            
            # Extract from Excel
            try:
                import openpyxl
                excel_stream.seek(0)
                wb = openpyxl.load_workbook(excel_stream, data_only=True)
                ws = wb.active
                
                extracted_hotel_name = ws['C4'].value
                if extracted_hotel_name:
                    hotel_name = str(extracted_hotel_name).strip()
                    
                extracted_room_count = ws['C11'].value
                if extracted_room_count:
                    try:
                        room_count = int(extracted_room_count)
                    except (ValueError, TypeError):
                        pass
            except Exception as e:
                logger.error(f"Failed to extract GSA details from Excel: {e}")
                
            gsa_term = request.form.get("gsa_term", "")
            gsa_studio = request.form.get("gsa_studio", "NO")
            try:
                gsa_meeting = float(request.form.get("gsa_meeting", 0))
            except (ValueError, TypeError):
                gsa_meeting = 0.0

            from generate_gsa_docx import generate_gsa_docx_to_pdf
            template_path = os.path.join(os.path.dirname(__file__), "GSA_TAGS.docx")
            
            # Generate GSA for BoldSign (No CTA Button)
            future_gsa_bs = executor.submit(
                generate_gsa_docx_to_pdf,
                template_path, property_code, signer_name, signer_email, hotel_name, address, room_count, gsa_term, gsa_studio, gsa_meeting, gsa_notes, True
            )

            # Generate GSA for Client (Has CTA Button)
            future_gsa_client = executor.submit(
                generate_gsa_docx_to_pdf,
                template_path, property_code, signer_name, signer_email, hotel_name, address, room_count, gsa_term, gsa_studio, gsa_meeting, gsa_notes, False
            )

            # Check for existing GSA quote and create if missing
            new_gsa_monday_id = None
            try:
                from search_quotes import get_quotes_by_inn_code, create_and_approve_gsa_quote
                existing_quotes = get_quotes_by_inn_code(property_code)
                gsa_quote_id = None
                visible_gsa_quote_num = None
                original_rep = None
                
                for q in existing_quotes:
                    if "GSA ONLY" in q.get("package", "").upper():
                        if q.get("links") and len(q["links"]) > 0:
                            match = [l for l in q["links"] if "customerView.php?id=" in l]
                            if match:
                                gsa_quote_id = match[0].split("id=")[1].split("&")[0]
                                visible_gsa_quote_num = q.get("job", "").replace("#", "").strip()
                    elif not original_rep:
                        # Grab the assigned rep from the first non-GSA quote we find
                        original_rep = q.get("assigned_rep")

                if not gsa_quote_id:
                    logger.info(f"No existing GSA quote found for {property_code}. Creating one now. Assigning to: {original_rep}")
                    gsa_quote_id, visible_gsa_quote_num = create_and_approve_gsa_quote(property_code, hotel_name, total_monthly, room_count, gsa_meeting, gsa_notes, sales_rep_name=original_rep)
                    if gsa_quote_id:
                        logger.info(f"Successfully created and approved GSA quote ID: {gsa_quote_id}")
                else:
                    logger.info(f"GSA quote already exists: ID {gsa_quote_id}. Skipping creation.")
                    
                # Duplicate the Monday item for GSA synchronously to grab ID, then finish async
                if monday_item_id and gsa_quote_id:
                    monday_token = os.environ.get("MONDAY_API_TOKEN")
                    if monday_token:
                        monday_headers = {
                            "Authorization": monday_token,
                            "API-Version": "2024-01",
                            "Content-Type": "application/json"
                        }
                        dup_query = """
                        mutation ($itemId: ID!) {
                          duplicate_item (board_id: 18424851511, item_id: $itemId, with_updates: false) {
                            id
                          }
                        }
                        """
                        dup_resp = requests.post("https://api.monday.com/v2", headers=monday_headers, json={"query": dup_query, "variables": {"itemId": monday_item_id}})
                        dup_data = dup_resp.json()
                        if "errors" not in dup_data:
                            new_gsa_monday_id = dup_data.get("data", {}).get("duplicate_item", {}).get("id")
                            logger.info(f"Successfully duplicated item to new GSA item {new_gsa_monday_id}")
                            
                            from monday_helper import process_monday_gsa_duplicate_async
                            import threading
                            threading.Thread(target=process_monday_gsa_duplicate_async, args=(new_gsa_monday_id, property_code, monday_token, visible_gsa_quote_num, total_monthly)).start()
            except Exception as e:
                logger.error(f"Failed during GSA quote automation: {e}")

        logger.info("Waiting for all PDF conversions to complete...")
        pdf_stream_bs = future_main_bs.result()
        pdf_stream_client = future_main_client.result()
        
        if generate_cabling and cabling_internal_quote_id and 'future_cabling_bs' in locals():
            cabling_bs_stream = future_cabling_bs.result()
            cabling_client_stream = future_cabling_client.result()
            
        if generate_gsa and 'future_gsa_bs' in locals():
            gsa_bs_stream, total_monthly = future_gsa_bs.result()
            gsa_client_stream, _ = future_gsa_client.result()
            
        executor.shutdown()

        # ---------------------------------------------------------
        # SEND PROPOSAL TO BOLDSIGN
        # ---------------------------------------------------------
        boldsign_api_key = os.environ.get("BOLDSIGN_API_KEY")
        if not boldsign_api_key:
            raise ValueError("BOLDSIGN_API_KEY is missing from Render Environment Variables. Please add it to generate the document.")

        logger.info(f"Sending document to BoldSign for {property_code}")
        pdf_stream_bs.seek(0)

        headers = {
            'X-API-KEY': boldsign_api_key,
            'Accept': 'application/json'
        }

        # Initialize the file list with the Proposal
        proposal_files = [
            ('Files', ('Generated_Proposal.pdf', pdf_stream_bs.read(), 'application/pdf'))
        ]
        pdf_stream_bs.seek(0)

        title_str = f"Hospitality Technologies Agreement - {property_code}"
        if monday_item_id:
            title_str += f" | ID:{monday_item_id}"
        if internal_quote_id:
            title_str += f" | QID:{internal_quote_id}"

        proposal_data = {
            'Title': title_str,
            'ExpiryDays': '14',
            'Signers[0][name]': signer_name,
            'Signers[0][emailAddress]': signer_email,
            'Signers[0][signerType]': 'Signer',
            'Signers[0][recipientNotificationSettings][signatureRequest]': 'false'
        }

        # Inject exact coordinates for the Proposal Signature Block
        if target_page:
            proposal_data.update({
                f'Signers[0][formFields][0][id]': 'proposal_signature',
                f'Signers[0][formFields][0][fieldType]': 'Signature',
                f'Signers[0][formFields][0][pageNumber]': str(target_page),
                f'Signers[0][formFields][0][bounds][x]': '344',
                f'Signers[0][formFields][0][bounds][y]': '584',
                f'Signers[0][formFields][0][bounds][width]': '214',
                f'Signers[0][formFields][0][bounds][height]': '30',
                f'Signers[0][formFields][0][isRequired]': 'true'
            })

        resp_prop = requests.post("https://api.boldsign.com/v1/document/send", headers=headers, data=proposal_data, files=proposal_files)
        if resp_prop.status_code not in (200, 201):
            logger.error(f"BoldSign Proposal API Error: {resp_prop.status_code} - {resp_prop.text}")
            raise RuntimeError(f"Failed to create BoldSign Proposal document: {resp_prop.text}")
        else:
            logger.info(f"Successfully created BoldSign Proposal document: {resp_prop.json().get('documentId')}")

        # ---------------------------------------------------------
        # SEND GSA TO BOLDSIGN (IF REQUESTED)
        # ---------------------------------------------------------
        if generate_gsa and gsa_bs_stream:
            gsa_title_str = f"Guest Service Agreement - {property_code}"
            if new_gsa_monday_id:
                gsa_title_str += f" | ID:{new_gsa_monday_id}"
            elif monday_item_id:
                gsa_title_str += f" | ID:{monday_item_id}"
            if gsa_quote_id:
                gsa_title_str += f" | QID:{gsa_quote_id}"

            gsa_data = {
                'Title': gsa_title_str,
                'ExpiryDays': '180',
                'UseTextTags': 'true',
                'EnableSigningOrder': 'true',
                'Signers[0][name]': signer_name,
                'Signers[0][emailAddress]': signer_email,
                'Signers[0][signerType]': 'Signer',
                'Signers[0][signerOrder]': '1',
                'Signers[0][recipientNotificationSettings][signatureRequest]': 'false',
                'Signers[1][name]': 'HT Contracts',
                'Signers[1][emailAddress]': 'contracts@hospitalitytechnologies.com',
                'Signers[1][signerType]': 'Signer',
                'Signers[1][signerOrder]': '2'
            }

            gsa_files = [
                ('Files', ('Guest_Service_Agreement.pdf', gsa_bs_stream.read(), 'application/pdf'))
            ]

            resp_gsa = requests.post("https://api.boldsign.com/v1/document/send", headers=headers, data=gsa_data, files=gsa_files)
            if resp_gsa.status_code not in (200, 201):
                logger.error(f"BoldSign GSA API Error: {resp_gsa.status_code} - {resp_gsa.text}")
                raise RuntimeError(f"Failed to create BoldSign GSA document: {resp_gsa.text}")
            else:
                logger.info(f"Successfully created BoldSign GSA document: {resp_gsa.json().get('documentId')}")

        # ---------------------------------------------------------
        # SEND CABLING PROPOSAL TO BOLDSIGN (IF REQUESTED)
        # ---------------------------------------------------------
        if generate_cabling and cabling_bs_stream:
            cabling_title_str = f"Cabling Proposal - {property_code}"
            if cabling_monday_item_id:
                cabling_title_str += f" | ID:{cabling_monday_item_id}"
            if cabling_internal_quote_id:
                cabling_title_str += f" | QID:{cabling_internal_quote_id}"

            cabling_data = {
                'Title': cabling_title_str,
                'ExpiryDays': '14',
                'Signers[0][name]': signer_name,
                'Signers[0][emailAddress]': signer_email,
                'Signers[0][signerType]': 'Signer',
                'Signers[0][recipientNotificationSettings][signatureRequest]': 'false'
            }

            if cabling_target_page:
                cabling_data.update({
                    f'Signers[0][formFields][0][id]': 'cabling_signature',
                    f'Signers[0][formFields][0][fieldType]': 'Signature',
                    f'Signers[0][formFields][0][pageNumber]': str(cabling_target_page),
                    f'Signers[0][formFields][0][bounds][x]': '344',
                    f'Signers[0][formFields][0][bounds][y]': '584',
                    f'Signers[0][formFields][0][bounds][width]': '214',
                    f'Signers[0][formFields][0][bounds][height]': '30',
                    f'Signers[0][formFields][0][isRequired]': 'true'
                })

            cabling_files = [
                ('Files', ('Cabling_Proposal.pdf', cabling_bs_stream.read(), 'application/pdf'))
            ]

            resp_cabling = requests.post("https://api.boldsign.com/v1/document/send", headers=headers, data=cabling_data, files=cabling_files)
            if resp_cabling.status_code not in (200, 201):
                logger.error(f"BoldSign Cabling API Error: {resp_cabling.status_code} - {resp_cabling.text}")
                raise RuntimeError(f"Failed to create BoldSign Cabling document: {resp_cabling.text}")
            else:
                logger.info(f"Successfully created BoldSign Cabling document: {resp_cabling.json().get('documentId')}")


        logger.info(f"Successfully compiled documents. Preparing download stream.")
        pdf_stream_client.seek(0)

        if generate_gsa or generate_cabling:
            import zipfile
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                zf.writestr(generated_filename, pdf_stream_client.read())

                if generate_gsa and gsa_client_stream:
                    zf.writestr(f"GSA_{property_code}.pdf", gsa_client_stream.read())

                if generate_cabling and cabling_client_stream:
                    zf.writestr(f"Cabling_Proposal_{property_code}.pdf", cabling_client_stream.read())

            zip_buffer.seek(0)

            zip_name_parts = ["Proposal"]
            if generate_gsa: zip_name_parts.append("GSA")
            if generate_cabling: zip_name_parts.append("Cabling")

            zip_filename = f"{'_'.join(zip_name_parts)}_{property_code}.zip"
            return send_file(
                zip_buffer,
                mimetype="application/zip",
                as_attachment=True,
                download_name=zip_filename
            )

        return send_file(
            pdf_stream_client,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=generated_filename
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
    property_code = property_code.upper()
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
    target_title_prefix = f"Hospitality Technologies Agreement - {property_code}"
    for doc in docs:
        if doc.get("messageTitle", "").startswith(target_title_prefix) and doc.get("status") not in ("Expired", "Completed", "Declined", "Revoked"):
            target_doc = doc
            break

    if not target_doc:
        return "No active document found for signing. It may have expired.", 404
        
    document_id = target_doc.get("documentId")
    try:
        signer_email = target_doc.get("signerDetails")[0].get("signerEmail")
    except (TypeError, IndexError, KeyError):
        return "Signer details could not be parsed from BoldSign.", 500
    
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

@app.route("/sign_gsa/<property_code>", methods=["GET"])
def sign_gsa_document(property_code):
    """Dynamic gateway redirecting to BoldSign embedded signing link specifically for the GSA."""
    property_code = property_code.upper()
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
        
    # Find the most recent matching document for the GSA
    target_doc = None
    target_title_prefix = f"Guest Service Agreement - {property_code}"
    for doc in docs:
        if doc.get("messageTitle", "").startswith(target_title_prefix) and doc.get("status") not in ("Expired", "Completed", "Declined", "Revoked"):
            target_doc = doc
            break

    if not target_doc:
        return "No active GSA document found for signing. It may have expired.", 404
        
    document_id = target_doc.get("documentId")
    try:
        signer_email = target_doc.get("signerDetails")[0].get("signerEmail")
    except (TypeError, IndexError, KeyError):
        return "Signer details could not be parsed from BoldSign.", 500
    
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

@app.route("/api/debug/columns", methods=["GET"])
@requires_auth
def debug_columns():
    monday_token = os.environ.get("MONDAY_API_TOKEN")
    board_id = "18424851511"
    query = "query { boards(ids: 18424851511) { columns { id title type settings_str } } }"
    headers = {"Authorization": monday_token, "API-Version": "2024-01", "Content-Type": "application/json"}
    resp = requests.post("https://api.monday.com/v2", headers=headers, json={"query": query})
    return jsonify(resp.json())

@app.route("/api/get_monday_quotes", methods=["POST"])
@requires_auth
def get_monday_quotes():
    property_code = request.form.get("property_code", "").strip().upper()
    if not property_code:
        return jsonify({"error": "Property code required"}), 400
        
    monday_token = os.environ.get("MONDAY_API_TOKEN")
    if not monday_token:
        return jsonify({"error": "Monday API token not configured on server"}), 500
        
    board_id = "18424851511"
    
    query = """
    query ($boardId: [ID!], $propCode: String!) {
      boards(ids: $boardId) {
        items_page(limit: 50, query_params: {rules: [{column_id: "name", compare_value: [$propCode], operator: contains_text}]}) {
          items {
            id
            name
            column_values {
              column { title }
              text
            }
          }
        }
      }
    }
    """
    
    headers = {
        "Authorization": monday_token,
        "API-Version": "2024-01",
        "Content-Type": "application/json"
    }
    
    try:
        resp = requests.post("https://api.monday.com/v2", headers=headers, json={
            "query": query,
            "variables": {
                "boardId": [board_id],
                "propCode": property_code
            }
        })
        resp.raise_for_status()
        data = resp.json()
        
        if "errors" in data:
            logger.error(f"Monday API Error: {data['errors']}")
            return jsonify({"error": "Error querying Monday.com. Check server logs."}), 500
            
        raw_items = data.get("data", {}).get("boards", [{}])[0].get("items_page", {}).get("items", [])
        
        formatted_items = []
        for item in raw_items:
            item_name = item.get("name", "")
            quote_type = "N/A"
            sales_owner = "N/A"
            item_date = "N/A"
            contact_name = ""
            contact_email = ""
            
            for col in item.get("column_values", []):
                title = col.get("column", {}).get("title", "")
                text = col.get("text") or ""
                
                title_lower = title.lower()
                if "quote type" in title_lower or title_lower == "type":
                    if text: quote_type = text
                elif "sales owner" in title_lower or title_lower == "owner":
                    if text: sales_owner = text
                elif "date" in title_lower and "update" not in title_lower:
                    if text: item_date = text
                elif "contact name" in title_lower:
                    if text: contact_name = text
                elif "contact e-mail" in title_lower or "contact email" in title_lower:
                    if text: contact_email = text
            
            display_name = f"{item_name} | {quote_type} | {sales_owner} | {item_date}"
            formatted_items.append({
                "id": item["id"], 
                "name": display_name,
                "contact_name": contact_name,
                "contact_email": contact_email
            })

        return jsonify(formatted_items)
        
    except Exception as e:
        logger.error(f"Failed to query Monday.com: {e}")
        return jsonify({"error": "Failed to connect to Monday.com"}), 500

@app.route("/api/internal_quotes", methods=["POST"])
def get_internal_quotes():
    property_code = request.form.get("property_code", "").strip().upper()
    if not property_code:
        return jsonify({"error": "Property code required"}), 400
    
    try:
        quotes = get_quotes_by_inn_code(property_code)
        return jsonify(quotes)
    except Exception as e:
        logger.error(f"Failed to fetch internal quotes: {e}")
        return jsonify({"error": "Failed to fetch quotes from internal system"}), 500

@app.route("/api/boldsign/status/<inn_code>", methods=["GET"])
@requires_auth
def get_boldsign_status(inn_code):
    boldsign_api_key = os.environ.get("BOLDSIGN_API_KEY")
    if not boldsign_api_key:
        return jsonify({"error": "BoldSign integration not configured."}), 500

    headers = {
        'X-API-KEY': boldsign_api_key,
        'Accept': 'application/json'
    }
    
    search_url = f"https://api.boldsign.com/v1/document/list?searchQuery={inn_code}&pageSize=10"
    try:
        resp = requests.get(search_url, headers=headers)
        if resp.status_code != 200:
            return jsonify({"error": "Failed to search BoldSign documents."}), 500
        
        data = resp.json()
        docs = data.get("result", [])
        
        # Format the response
        results = []
        for doc in docs:
            # Strictly filter by the Inn Code provided since BoldSign's search is broad
            if inn_code.upper() not in doc.get("messageTitle", "").upper():
                continue
                
            # Check signers if in progress to give better details
            signer_statuses = []
            if doc.get("status") == "InProgress":
                try:
                    prop_resp = requests.get(f"https://api.boldsign.com/v1/document/properties?documentId={doc['documentId']}", headers=headers)
                    if prop_resp.status_code == 200:
                        prop_data = prop_resp.json()
                        for s in prop_data.get("signerDetails", []):
                            signer_statuses.append(f"{s.get('signerName')} ({s.get('status', 'Pending')})")
                except Exception as e:
                    logger.error(f"Failed to fetch document properties for {doc['documentId']}: {e}")
            
            results.append({
                "documentId": doc.get("documentId"),
                "title": doc.get("messageTitle"),
                "status": doc.get("status"),
                "createdDate": doc.get("createdDate"),
                "signers": ", ".join(signer_statuses) if signer_statuses else ""
            })
            
        return jsonify(results)
    except Exception as e:
        logger.error(f"Error querying BoldSign status for {inn_code}: {e}")
        return jsonify({"error": "Internal server error"}), 500

@app.route("/api/boldsign/revoke/<doc_id>", methods=["POST"])
@requires_auth
def revoke_boldsign_document(doc_id):
    boldsign_api_key = os.environ.get("BOLDSIGN_API_KEY")
    if not boldsign_api_key:
        return jsonify({"error": "BoldSign integration not configured."}), 500

    headers = {
        'X-API-KEY': boldsign_api_key,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
    }
    
    payload = request.get_json() or {}
    message = payload.get("message", "Document revoked by admin.")
    
    revoke_url = f"https://api.boldsign.com/v1/document/revoke?documentId={doc_id}"
    try:
        resp = requests.post(revoke_url, headers=headers, json={"message": message})
        if resp.status_code in (200, 204):
            return jsonify({"success": True, "message": "Document successfully revoked."})
        else:
            logger.error(f"Failed to revoke BoldSign document {doc_id}: {resp.status_code} - {resp.text}")
            return jsonify({"error": f"Failed to revoke document: {resp.text}"}), 500
    except Exception as e:
        logger.error(f"Error revoking BoldSign document {doc_id}: {e}")
        return jsonify({"error": "Internal server error"}), 500

# Track processed document IDs to prevent duplicate webhook processing
processed_doc_ids = set()

@app.route("/api/boldsign-webhook", methods=["POST"])
def boldsign_webhook():
    """Receives completion events from BoldSign to sync with Monday.com and SharePoint."""
    try:
        event = request.json
        if not event:
            return jsonify({"status": "ignored"}), 200

        event_type = event.get("event", {}).get("eventType") if "eventType" in event.get("event", {}) else event.get("event")

        # Handle BoldSign Webhook Verification Ping
        if event_type == "Verification":
            logger.info("BoldSign Webhook Verification Received.")
            return jsonify({"status": "verified"}), 200
            
        if event_type == "Completed":
            doc_id = event.get("data", {}).get("documentId")
            if not doc_id:
                logger.error("Webhook received 'Completed' event but no documentId was found.")
                return jsonify({"error": "Missing documentId"}), 400
                
            # Deduplication check
            if doc_id in processed_doc_ids:
                logger.info(f"Webhook Received: Document {doc_id} already processed. Skipping duplicate webhook.")
                return jsonify({"status": "accepted"}), 200
            processed_doc_ids.add(doc_id)
            
            # Run the heavy processing in a background thread to prevent BoldSign timeout (which causes retries)
            import threading
            threading.Thread(target=process_boldsign_webhook_async, args=(event,)).start()
            return jsonify({"status": "accepted"}), 200
            
        return jsonify({"status": "ignored"}), 200

    except Exception as e:
        logger.error(f"Webhook initialization error: {e}")
        return jsonify({"error": str(e)}), 500

def process_boldsign_webhook_async(event):
    """Background task to handle Monday.com updates and SharePoint archival."""
    import base64
    import time
    try:
        doc_id = event.get("data", {}).get("documentId")
        if not doc_id:
            logger.error("Webhook received 'Completed' event but no documentId was found.")
            return
            
        logger.info(f"Background Thread: Processing Completed for {doc_id}")

        # 1. Fetch BoldSign Document Properties to get MondayItemId from the Title string
        bs_api_key = os.environ.get("BOLDSIGN_API_KEY")
        bs_headers = {'X-API-KEY': bs_api_key, 'Accept': 'application/json'}
        prop_resp = requests.get(f"https://api.boldsign.com/v1/document/properties?documentId={doc_id}", headers=bs_headers)
        prop_resp.raise_for_status()

        message_title = prop_resp.json().get("messageTitle", "")
        monday_item_id = None
        internal_quote_id = None
        
        if " | QID:" in message_title:
            internal_quote_id = message_title.split(" | QID:")[1].strip()
            # Clean up message_title so monday_item_id splitting works if QID is appended
            message_title_for_id = message_title.split(" | QID:")[0]
        else:
            message_title_for_id = message_title
            
        if " | ID:" in message_title_for_id:
            monday_item_id = message_title_for_id.split(" | ID:")[1].strip()

        if not monday_item_id:
            logger.error(f"No MondayItemId found in BoldSign Title string: {message_title}")
            return
            
        inn_code = "Unknown"
        if " - " in message_title and " | ID:" in message_title:
            inn_code = message_title.split(" - ")[1].split(" | ID:")[0].strip().upper()

        is_gsa = "Guest Service Agreement" in message_title
        
        if internal_quote_id:
            logger.info(f"Background Thread: Marking internal Quote {internal_quote_id} as WON.")
            success = mark_quote_won(internal_quote_id)
            if success:
                logger.info(f"Successfully updated internal Quote {internal_quote_id} to WON.")
            else:
                logger.error(f"Failed to update internal Quote {internal_quote_id} to WON.")

        # 2. Update Monday.com Status and Extract SharePoint Link
        monday_token = os.environ.get("MONDAY_API_TOKEN")
        board_id = "18424851511"
        monday_headers = {
            "Authorization": monday_token, 
            "API-Version": "2024-01", 
            "Content-Type": "application/json"
        }

        if not is_gsa:
            # Update Status to WON for Proposal
            mutation = """
            mutation ($itemId: ID!, $boardId: ID!) {
              change_column_value(item_id: $itemId, board_id: $boardId, column_id: "color_mm6sy6mm", value: "{\\"label\\":\\"WON\\"}") {
                id
              }
            }
            """
            mut_resp = requests.post("https://api.monday.com/v2", headers=monday_headers, json={
                "query": mutation,
                "variables": {"itemId": monday_item_id, "boardId": board_id}
            })
            
            try:
                mut_data = mut_resp.json()
                if "errors" in mut_data:
                    logger.error(f"Monday Mutation Error: {mut_data['errors']}")
                else:
                    logger.info(f"Successfully updated Monday item {monday_item_id} to WON.")
            except Exception as e:
                logger.error(f"Failed to parse Monday mutation response: {e}")
        else:
            try:
                    # 1. Find and mark "Contract Signed by Customer" as Done
                    sub_query = """
                    query ($itemId: [ID!]) {
                      items (ids: $itemId) {
                        subitems { id name }
                      }
                    }
                    """
                    sub_resp = requests.post("https://api.monday.com/v2", headers=monday_headers, json={
                        "query": sub_query,
                        "variables": {"itemId": [monday_item_id]}
                    })
                    
                    items_list = sub_resp.json().get('data', {}).get('items', [])
                    current_subitems = items_list[0].get('subitems', []) if items_list else []
                    
                    signed_subitem = next((s for s in current_subitems if s.get('name') == "Contract Signed by Customer"), None)
                    
                    update_sub_query = """
                    mutation ($itemId: ID!, $boardId: ID!, $colVals: JSON!) {
                      change_multiple_column_values (item_id: $itemId, board_id: $boardId, column_values: $colVals) { id }
                    }
                    """
                    if signed_subitem:
                        logger.info("Marking 'Contract Signed by Customer' as Done...")
                        requests.post("https://api.monday.com/v2", headers=monday_headers, json={
                            "query": update_sub_query,
                            "variables": {
                                "itemId": signed_subitem['id'],
                                "boardId": "18424851639",
                                "colVals": json.dumps({"status": {"label": "Done"}})
                            }
                        })
                    else:
                        logger.warning("'Contract Signed by Customer' subitem not found.")
                    
                    # 2. Wait for post-signature subitems to be created by the new automation, then mark them Done
                    import time
                    max_auto_retries = 15
                    logger.info("Polling for post-signature subitems...")
                    sp_subitem = None
                    won_subitem = None
                    for attempt in range(max_auto_retries):
                        sub_resp = requests.post("https://api.monday.com/v2", headers=monday_headers, json={
                            "query": sub_query,
                            "variables": {"itemId": [monday_item_id]}
                        })
                        try:
                            items_list2 = sub_resp.json().get('data', {}).get('items', [])
                            current_subitems2 = items_list2[0].get('subitems', []) if items_list2 else []
                            sp_subitem = next((s for s in current_subitems2 if s.get('name') == "Upload Signed GSA to Sharepoint"), None)
                            won_subitem = next((s for s in current_subitems2 if s.get('name') == "Mark Status in Quote Tool as WON"), None)
                            if sp_subitem or won_subitem:
                                break
                        except Exception:
                            pass
                        time.sleep(2.5)
                        
                    for sub_name, sub_obj in [("Upload Signed GSA to Sharepoint", sp_subitem), ("Mark Status in Quote Tool as WON", won_subitem)]:
                        if sub_obj:
                            logger.info(f"Marking '{sub_name}' as Done...")
                            requests.post("https://api.monday.com/v2", headers=monday_headers, json={
                                "query": update_sub_query,
                                "variables": {
                                    "itemId": sub_obj['id'],
                                    "boardId": "18424851639",
                                    "colVals": json.dumps({"status": {"label": "Done"}})
                                }
                            })
                        else:
                            logger.warning(f"Timed out waiting for '{sub_name}' subitem to appear.")
                        
                    # 3. Finally Update Status to Contract Signed for GSA
                    mutation = """
                    mutation ($itemId: ID!, $boardId: ID!) {
                      change_column_value(item_id: $itemId, board_id: $boardId, column_id: "color_mm6sy6mm", value: "{\\"label\\":\\"Contract Signed\\"}") {
                        id
                      }
                    }
                    """
                    mut_resp = requests.post("https://api.monday.com/v2", headers=monday_headers, json={
                        "query": mutation,
                        "variables": {"itemId": monday_item_id, "boardId": board_id}
                    })
                    try:
                        mut_data = mut_resp.json()
                        if "errors" in mut_data:
                            logger.error(f"Monday Mutation Error: {mut_data['errors']}")
                        else:
                            logger.info(f"Successfully updated GSA Monday item {monday_item_id} to Contract Signed.")
                    except Exception as e:
                        logger.error(f"Failed to parse Monday mutation response: {e}")
                        
            except Exception as e:
                logger.error(f"Failed in webhook subitem execution: {e}")

        # Fetch the Files (SharePoint) link
        query = """
        query ($itemId: [ID!]) {
          items(ids: $itemId) {
            column_values {
              column { title }
              text
              value
            }
          }
        }
        """
        item_resp = requests.post("https://api.monday.com/v2", headers=monday_headers, json={
            "query": query,
            "variables": {"itemId": [monday_item_id]}
        })
        item_data = item_resp.json()

        sp_link = None
        for col in item_data.get("data", {}).get("items", [])[0].get("column_values", []):
            if col.get("column", {}).get("title") == "Files":
                # Parse link from text or raw JSON value
                import json as sys_json
                sp_link = col.get("text")
                if not sp_link and col.get("value"):
                    sp_link = sys_json.loads(col.get("value")).get("url")
                break

        if not sp_link:
            logger.warning(f"Could not extract SharePoint link from Monday.com item {monday_item_id}")
            return

        # 3. Download signed PDF from BoldSign
        dl_resp = requests.get(f"https://api.boldsign.com/v1/document/download?documentId={doc_id}", headers=bs_headers)
        dl_resp.raise_for_status()
        pdf_bytes = dl_resp.content

        # 4. Upload to SharePoint via MS Graph API
        tenant_id = os.environ.get("MS_GRAPH_TENANT_ID")
        client_id = os.environ.get("MS_GRAPH_CLIENT_ID")
        client_secret = os.environ.get("MS_GRAPH_CLIENT_SECRET")

        if not all([tenant_id, client_id, client_secret]):
            logger.error("Missing MS Graph credentials in environment variables")
            return

        token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        token_resp = requests.post(token_url, data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default"
        })
        token_resp.raise_for_status()
        ms_token = token_resp.json().get("access_token")

        # Convert SharePoint web URL to Graph API Sharing URL
        encoded_url = base64.b64encode(sp_link.encode("utf-8")).decode("utf-8")
        encoded_url = "u!" + encoded_url.rstrip("=").replace("/", "_").replace("+", "-")

        folder_url = f"https://graph.microsoft.com/v1.0/shares/{encoded_url}/driveItem"
        ms_headers = {"Authorization": f"Bearer {ms_token}"}
        folder_resp = requests.get(folder_url, headers=ms_headers)
        folder_resp.raise_for_status()

        folder_info = folder_resp.json()
        drive_id = folder_info.get("parentReference", {}).get("driveId")
        item_id = folder_info.get("id")

        # Upload the file
        file_name = f"Signed_Agreement_{doc_id[:8]}.pdf"
        upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}:/{file_name}:/content"
        upload_resp = requests.put(upload_url, headers={"Authorization": f"Bearer {ms_token}", "Content-Type": "application/pdf"}, data=pdf_bytes)
        upload_resp.raise_for_status()

        logger.info("Successfully updated Monday.com status and uploaded PDF to SharePoint.")
        
    except Exception as e:
        logger.error(f"Webhook processing error: {e}")

@app.route("/health", methods=["GET"])
def health():
    """Health check endpoint for Render zero-downtime deploys and monitoring."""
    return jsonify({"status": "healthy", "service": "quote-presentation-generator"}), 200

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug)
