import io
import os
import tempfile
import subprocess
from datetime import datetime
import fitz  # PyMuPDF
from docx import Document

def generate_gsa_docx_to_pdf(template_path, property_code, signer_name, signer_email, hotel_name, room_count, term, studio, meeting_fee, notes, for_boldsign=False):
    # Perform math calculations
    connect_support = room_count * 1.50
    studio_support = 1.00 if studio == "YES" else 0.00
    total_studio = room_count * studio_support
    total_monthly = connect_support + total_studio + meeting_fee
    
    # Format as currency
    def format_currency(val):
        return "${:,.2f}".format(val)
        
    replacements = {
        "{{SIGNER_NAME}}": signer_name or "",
        "{{SIGNER_EMAIL}}": signer_email or "",
        "{{HOTEL_NAME}}": hotel_name or "",
        "{{INN_CODE}}": property_code.upper() if property_code else "",
        "{{TERM}}": term or "",
        "{{DATE}}": datetime.now().strftime("%B %d, %Y"),
        "{{ROOM_COUNT}}": str(room_count),
        "{{CONNECT_SUPPORT}}": format_currency(connect_support),
        "{{STUDIO_SUPPORT}}": format_currency(studio_support),
        "{{TOTAL_STUDIO}}": format_currency(total_studio),
        "{{MEETING}}": format_currency(meeting_fee),
        "{{TOTAL MONTHLY}}": format_currency(total_monthly),
        "{{NOTES}}": notes or "",
        "{{EIN}}": "" # User left blank
    }
    
    doc = Document(template_path)
    
    # Replace in paragraphs
    for para in doc.paragraphs:
        for key, value in replacements.items():
            if key in para.text:
                for run in para.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, value)
                        
    # Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for key, value in replacements.items():
                        if key in para.text:
                            for run in para.runs:
                                if key in run.text:
                                    run.text = run.text.replace(key, value)
                                    
    # Convert docx to PDF using LibreOffice headless
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_docx = os.path.join(temp_dir, "temp_gsa.docx")
        doc.save(temp_docx)
        
        libreoffice_exec = "libreoffice"
        import shutil
        if not shutil.which("libreoffice"):
            if os.path.exists("/Applications/LibreOffice.app/Contents/MacOS/soffice"):
                libreoffice_exec = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
                
        cmd = [
            libreoffice_exec, "--headless", "--convert-to", "pdf",
            "--outdir", temp_dir, temp_docx
        ]
        
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        
        temp_pdf = os.path.join(temp_dir, "temp_gsa.pdf")
        
        # Open PDF with PyMuPDF to add CTA Button if needed
        pdf_doc = fitz.open(temp_pdf)
        
        if not for_boldsign:
            last_page = pdf_doc[-1]
            page_width = last_page.rect.width
            page_height = last_page.rect.height
            
            btn_width = 300
            btn_height = 40
            btn_x = (page_width - btn_width) / 2
            btn_y = page_height - 80
            
            btn_rect = fitz.Rect(btn_x, btn_y, btn_x + btn_width, btn_y + btn_height)
            
            # Draw Blue Button
            shape = last_page.new_shape()
            shape.draw_rect(btn_rect)
            shape.finish(color=(0, 0.447, 0.807), fill=(0, 0.447, 0.807)) # Primary Blue
            shape.commit()
            
            # Add CTA Text
            last_page.insert_textbox(btn_rect, "Ready to Move Forward? Click to E-Sign", 
                                     fontsize=12, fontname="hebo", color=(1, 1, 1), 
                                     align=fitz.TEXT_ALIGN_CENTER)
                                     
            link_url = f"https://quote-presentation-generator.onrender.com/sign_gsa/{property_code}"
            last_page.insert_link({"kind": fitz.LINK_URI, "from": btn_rect, "uri": link_url})
            
        output_stream = io.BytesIO()
        pdf_doc.save(output_stream)
        pdf_doc.close()
        
    output_stream.seek(0)
    return output_stream