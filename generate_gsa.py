import io
import os
import fitz  # PyMuPDF

def generate_gsa_pdf(template_path, hotel_name, address, notes, property_code, for_boldsign=False):
    """
    Overlays dynamic text onto the blank GSA PDF and adds an interactive CTA button.
    If for_boldsign is True, the button is omitted.
    """
    doc = fitz.open(template_path)
    
    text_color = (0, 0, 0) # Black
    font_size = 11
    font_name = "helv"
    
    # Page 1 - Cover Page
    page1 = doc[0]
    # Injecting slightly to the right of the found labels
    page1.insert_text((180, 373), hotel_name or "", fontsize=font_size, fontname=font_name, color=text_color)
    page1.insert_text((180, 398), address or "", fontsize=font_size, fontname=font_name, color=text_color)
    
    # Page 12 - Addendum
    if len(doc) >= 12:
        page12 = doc[11]
        page12.insert_text((180, 184), hotel_name or "", fontsize=font_size, fontname=font_name, color=text_color)
        page12.insert_text((180, 214), address or "", fontsize=font_size, fontname=font_name, color=text_color)
        
        if notes:
            # Wrap notes in a text box
            notes_rect = fitz.Rect(93, 600, 500, 720)
            page12.insert_textbox(notes_rect, notes, fontsize=font_size, fontname=font_name, color=text_color)
            
    # Add Interactive Button if it is the client copy
    if not for_boldsign:
        last_page = doc[-1]
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
                                 fontsize=12, fontname="helv-b", color=(1, 1, 1), 
                                 align=fitz.TEXT_ALIGN_CENTER)
                                 
        # Add Hyperlink pointing to the exact same Render gateway used by the quote deck
        # Note: If running locally, this points to production. It routes via property_code.
        link_url = f"https://quote-presentation-generator.onrender.com/sign/{property_code}"
        last_page.insert_link({"kind": fitz.LINK_URI, "from": btn_rect, "uri": link_url})
        
    output_stream = io.BytesIO()
    doc.save(output_stream)
    doc.close()
    
    output_stream.seek(0)
    return output_stream