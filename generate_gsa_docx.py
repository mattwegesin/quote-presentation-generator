import io
import os
from datetime import datetime
from docx import Document
from python_docx_replace.paragraph import Paragraph

def generate_gsa_docx(template_path, property_code, signer_name, signer_email, hotel_name, address, room_count, term, studio, meeting_fee, notes):
    # Perform math calculations
    try:
        room_count = int(room_count)
    except (ValueError, TypeError):
        room_count = 0
        
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
        "{{ADDRESS}}": address or "",
        "{{INN_CODE}}": property_code.upper() if property_code else "",
        "{{TERM}}": term or "",
        "{{DATE}}": datetime.now().strftime("%B %d, %Y"),
        "{{ROOM_COUNT}}": str(room_count),
        "{{CONNECT_SUPPORT}}": format_currency(connect_support),
        "{{STUDIO_SUPPORT}}": format_currency(studio_support),
        "{{TOTAL_STUDIO}}": format_currency(total_studio),
        "{{MEETING}}": format_currency(meeting_fee),
        "{{TOTAL_MONTHLY}}": format_currency(total_monthly),
        "{{TOTAL MONTHLY}}": format_currency(total_monthly),
        "{{NOTES}}": notes or "",
        "{{EIN}}": "" # User left blank
    }
    
    doc = Document(template_path)
    
    # Use python_docx_replace to cleanly replace text across runs without destroying styling
    for p in Paragraph.get_all(doc):
        paragraph = Paragraph(p)
        for key, value in replacements.items():
            paragraph.replace_key(key, str(value))
    
    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)
    return output_stream