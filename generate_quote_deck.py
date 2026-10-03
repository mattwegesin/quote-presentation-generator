import io
import os
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image, ImageDraw, ImageFont

def create_cta_button_image():
    """Generates a high-quality modern CTA button as an in-memory PNG to bypass PDF hyperlink styling issues."""
    # 1200x168 matches the 5.0" x 0.7" aspect ratio perfectly to prevent PowerPoint clipping
    width, height = 1200, 168
    img = Image.new('RGBA', (width, height), color=(0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    
    # Draw button with shadow-like aesthetic (flat modern)
    # Background border/shadow
    d.rounded_rectangle([(0, 0), (width, height)], radius=20, fill=(2, 132, 199))
    # Main button face
    d.rounded_rectangle([(2, 2), (width-2, height-6)], radius=20, fill=(14, 165, 233))
    
    # Try to load a nice font, fallback to default if not available
    try:
        # Absolute path for Debian/Render Docker environment
        font = ImageFont.truetype("/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf", 48)
    except:
        try:
            # Works on macOS
            font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 48)
        except:
            try:
                font = ImageFont.truetype("LiberationSans-Bold.ttf", 48)
            except:
                font = ImageFont.load_default()
            
    text = "READY TO MOVE FORWARD? CLICK TO E-SIGN"
    # Center text
    d.text((width/2, height/2 - 2), text, fill="white", anchor="mm", font=font)
    
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')
    img_byte_arr.seek(0)
    return img_byte_arr

def move_slide(prs, old_index, new_index):
    """Moves the closing slide to the absolute end of the deck."""
    xml_slides = prs.slides._sldIdLst  
    slides = list(xml_slides)
    xml_slides.remove(slides[old_index])
    xml_slides.insert(new_index, slides[old_index])

def clean_float(val):
    """Safely handles parsed currency string conversions."""
    if isinstance(val, (int, float)): return float(val)
    if isinstance(val, str):
        val = val.replace(',', '').replace('$', '').strip()
        try: return float(val)
        except ValueError: return 0.0
    return 0.0

def add_logo_to_slide(slide, logo_data):
    """Pastes the extracted top-right logo image onto the new slide."""
    if logo_data:
        image_stream = io.BytesIO(logo_data['blob'])
        slide.shapes.add_picture(
            image_stream, 
            logo_data['left'], 
            logo_data['top'], 
            logo_data['width'], 
            logo_data['height']
        )

def create_pdf_style_slide(prs, title, chunks, logo_data, is_monthly=False):
    """Generates the main financial tables matching the PDF design guidelines."""
    blue_bg = RGBColor.from_string('0072CE')
    slate_text = RGBColor.from_string('0F172A')
    light_grey_bg = RGBColor.from_string('F1F5F9')
    red_text = RGBColor.from_string('BE123C')
    header_grey = RGBColor.from_string('E2E8F0')
    
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_logo_to_slide(slide, logo_data)
    
    # Red Kicker
    tb1 = slide.shapes.add_textbox(Inches(0.83), Inches(0.12), Inches(10), Inches(0.5))
    p1 = tb1.text_frame.paragraphs[0]
    p1.text = "FINANCIAL INVESTMENT OVERVIEW"
    p1.font.name = "Montserrat"
    p1.font.size = Pt(12)
    p1.font.bold = True
    p1.font.color.rgb = red_text
    
    # Main Page Title
    tb2 = slide.shapes.add_textbox(Inches(0.83), Inches(0.29), Inches(10), Inches(0.5))
    p2 = tb2.text_frame.paragraphs[0]
    p2.text = title
    p2.font.name = "Montserrat"
    p2.font.size = Pt(24)
    p2.font.bold = True
    p2.font.color.rgb = slate_text
    
    if is_monthly:
        tb3 = slide.shapes.add_textbox(Inches(0.83), Inches(0.8), Inches(11.5), Inches(0.5))
        p3 = tb3.text_frame.paragraphs[0]
        p3.text = "24x7x365 support by dedicated Hospitality Technologies agents for guests and hotel staff"
        p3.font.name = 'Montserrat'
        p3.font.size = Pt(12)
        p3.font.color.rgb = slate_text
        
    data = []
    for s in chunks:
        data.append({"type": "section_title", "text": s["section"]})
        data.append({"type": "header"})
        subtotal = 0
        for it in s["items"]:
            unit_cost = it.get("unit_cost", 0.0)
            data.append({"type": "item", "desc": it["desc"], "part": it["part"], "qty": it["qty"], "unit_cost": unit_cost, "price": it["total"]})
            subtotal += it["total"]
        if is_monthly:
            data.append({"type": "monthly", "val": subtotal})
        else:
            data.append({"type": "subtotal", "val": subtotal})
        data.append({"type": "spacer"})

    if data and data[-1]["type"] == "spacer":
        data = data[:-1]
        
    rows = len(data)
    if rows == 0: return slide
    
    start_y = 1.1 if not is_monthly else 1.3
    table_shape = slide.shapes.add_table(rows, 5, Inches(0.83), Inches(start_y), Inches(11.5), Inches(0.3 * rows))
    table = table_shape.table
    
    table.columns[0].width = Inches(4.5)
    table.columns[1].width = Inches(2.5)
    table.columns[2].width = Inches(1.1)
    table.columns[3].width = Inches(1.7)
    table.columns[4].width = Inches(1.7)
    
    for i, row in enumerate(data):
        cells = [table.cell(i, j) for j in range(5)]
        
        if row["type"] == "section_title":
            cells[0].merge(cells[4])
            cells[0].text = "   " + row["text"]
            cells[0].fill.solid()
            cells[0].fill.fore_color.rgb = blue_bg
            p = cells[0].text_frame.paragraphs[0]
            p.font.name = 'Montserrat'
            p.font.size = Pt(14)
            p.font.bold = True
            p.font.color.rgb = RGBColor.from_string('FFFFFF')
            p.alignment = PP_ALIGN.LEFT
            
        elif row["type"] == "header":
            headers = ["Description", "Item", "Quantity", "Unit Cost", "Price"]
            for j, h in enumerate(headers):
                cells[j].text = h
                cells[j].fill.solid()
                cells[j].fill.fore_color.rgb = header_grey
                p = cells[j].text_frame.paragraphs[0]
                p.font.name = 'Montserrat'
                p.font.size = Pt(11)
                p.font.bold = True
                p.font.color.rgb = slate_text
                if j == 2:
                    p.alignment = PP_ALIGN.CENTER
                elif j >= 3:
                    p.alignment = PP_ALIGN.RIGHT
                
        elif row["type"] == "item":
            cells[0].text = row["desc"]
            cells[1].text = row["part"]
            cells[2].text = str(row["qty"])
            cells[3].text = f"${row['unit_cost']:,.2f}"
            cells[4].text = f"${row['price']:,.2f}"
            for j in range(5):
                cells[j].fill.solid()
                cells[j].fill.fore_color.rgb = RGBColor.from_string('FFFFFF') if i % 2 == 0 else light_grey_bg
                p = cells[j].text_frame.paragraphs[0]
                p.font.name = 'Montserrat'
                p.font.size = Pt(10)
                p.font.color.rgb = slate_text
                if j == 2:
                    p.alignment = PP_ALIGN.CENTER
                elif j >= 3:
                    p.alignment = PP_ALIGN.RIGHT
                
        elif row["type"] in ("subtotal", "monthly"):
            cells[0].merge(cells[3])
            cells[0].text = "SUBTOTAL" if row["type"] == "subtotal" else "MONTHLY"
            cells[4].text = f"${row['val']:,.2f}"
            for j in (0, 4):
                cells[j].fill.solid()
                cells[j].fill.fore_color.rgb = RGBColor.from_string('FFFFFF')
                p = cells[j].text_frame.paragraphs[0]
                p.font.name = 'Montserrat'
                p.font.size = Pt(12)
                p.font.bold = True
                p.font.color.rgb = slate_text
                p.alignment = PP_ALIGN.RIGHT
                
        elif row["type"] == "spacer":
            cells[0].merge(cells[4])
            cells[0].text = ""

def create_summary_cards_slide(prs, total_inv, total_mo, logo_data):
    """Generates the Investment Summary metric cards."""
    blue_bg = RGBColor.from_string('0072CE')
    slate_text = RGBColor.from_string('0F172A')
    red_text = RGBColor.from_string('BE123C')
    
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_logo_to_slide(slide, logo_data)
    
    tb1 = slide.shapes.add_textbox(Inches(0.83), Inches(0.12), Inches(10), Inches(0.5))
    p1 = tb1.text_frame.paragraphs[0]
    p1.text = "FINANCIAL INVESTMENT OVERVIEW"
    p1.font.name = "Montserrat"
    p1.font.size = Pt(12)
    p1.font.bold = True
    p1.font.color.rgb = red_text
    
    tb2 = slide.shapes.add_textbox(Inches(0.83), Inches(0.29), Inches(10), Inches(0.5))
    p2 = tb2.text_frame.paragraphs[0]
    p2.text = "Investment Summary"
    p2.font.name = "Montserrat"
    p2.font.size = Pt(24)
    p2.font.bold = True
    p2.font.color.rgb = slate_text

    card_y = Inches(1.8)
    card_w = Inches(4.5)
    card_h = Inches(3.5)
    
    # Left Card (Total Investment)
    left_x = Inches(1.66)
    c1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_x, card_y, card_w, card_h)
    c1.fill.solid()
    c1.fill.fore_color.rgb = RGBColor.from_string('FFFFFF')
    c1.line.color.rgb = RGBColor.from_string('E2E8F0')
    c1.line.width = Pt(1)
    
    c1_top = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_x, card_y, card_w, Inches(0.1))
    c1_top.fill.solid()
    c1_top.fill.fore_color.rgb = blue_bg
    c1_top.line.fill.background()
    
    tb_c1 = slide.shapes.add_textbox(left_x, card_y + Inches(0.5), card_w, Inches(1.0))
    tb_c1.text_frame.word_wrap = True
    p = tb_c1.text_frame.paragraphs[0]
    p.text = "Total Project\nInvestment"
    p.font.name = 'Montserrat'
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = slate_text
    p.alignment = PP_ALIGN.CENTER
    
    tb_v1 = slide.shapes.add_textbox(left_x, card_y + Inches(1.8), card_w, Inches(1.0))
    p = tb_v1.text_frame.paragraphs[0]
    p.text = f"${total_inv:,.2f}"
    p.font.name = 'Montserrat'
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = blue_bg
    p.alignment = PP_ALIGN.CENTER

    # Right Card (Monthly Services)
    right_x = left_x + card_w + Inches(1.0)
    c2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, right_x, card_y, card_w, card_h)
    c2.fill.solid()
    c2.fill.fore_color.rgb = RGBColor.from_string('FFFFFF')
    c2.line.color.rgb = RGBColor.from_string('E2E8F0')
    c2.line.width = Pt(1)
    
    c2_top = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, right_x, card_y, card_w, Inches(0.1))
    c2_top.fill.solid()
    c2_top.fill.fore_color.rgb = red_text
    c2_top.line.fill.background()
    
    tb_c2 = slide.shapes.add_textbox(right_x, card_y + Inches(0.5), card_w, Inches(1.0))
    tb_c2.text_frame.word_wrap = True
    p = tb_c2.text_frame.paragraphs[0]
    p.text = "Monthly\nSupport Services"
    p.font.name = 'Montserrat'
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = slate_text
    p.alignment = PP_ALIGN.CENTER
    
    tb_v2 = slide.shapes.add_textbox(right_x, card_y + Inches(1.8), card_w, Inches(1.0))
    p = tb_v2.text_frame.paragraphs[0]
    p.text = f"${total_mo:,.2f}"
    p.font.name = 'Montserrat'
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = red_text
    p.alignment = PP_ALIGN.CENTER

def create_acceptance_slides(prs, logo_data, property_code=None, for_boldsign=False):
    """Generates two full-width legal disclaimer slides."""
    slate_text = RGBColor.from_string('0F172A')
    red_text = RGBColor.from_string('BE123C')
    
    def setup_slide(title_text):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_logo_to_slide(slide, logo_data)
        
        tb1 = slide.shapes.add_textbox(Inches(0.83), Inches(0.12), Inches(10), Inches(0.5))
        p1 = tb1.text_frame.paragraphs[0]
        p1.text = "FINANCIAL INVESTMENT OVERVIEW"
        p1.font.name = "Montserrat"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = red_text
        
        tb2 = slide.shapes.add_textbox(Inches(0.83), Inches(0.29), Inches(10), Inches(0.5))
        p2 = tb2.text_frame.paragraphs[0]
        p2.text = title_text
        p2.font.name = "Montserrat"
        p2.font.size = Pt(24)
        p2.font.bold = True
        p2.font.color.rgb = slate_text
        return slide
    
    def add_p(tf, text, bold=False, bullet=False):
        p = tf.add_paragraph() if tf.paragraphs[0].text else tf.paragraphs[0]
        p.text = text
        p.font.name = 'Montserrat'
        p.font.size = Pt(11)
        p.font.bold = bold
        p.font.color.rgb = slate_text
        if bullet:
            p.level = 1
        else:
            p.space_after = Pt(10)
        return p

    # --- SLIDE 1: Disclaimer & Caveats ---
    slide1 = setup_slide("Acceptance: Terms & Caveats")
    tb = slide1.shapes.add_textbox(Inches(0.83), Inches(1.0), Inches(11.5), Inches(5.5))
    tf = tb.text_frame
    tf.word_wrap = True
        
    add_p(tf, "Supply Chain & Pricing Disclaimer", bold=True)
    add_p(tf, "Due to ongoing global supply constraints affecting RAM and related electronic components, product availability, pricing, and delivery timelines may change without notice. All quotes and pricing are based on current supplier costs and component availability at the time issued.")
    add_p(tf, "Hospitality Technologies reserves the right to adjust pricing, revise delivery schedules, or substitute equivalent components if manufacturer pricing, allocations, or supply conditions change prior to order fulfillment. Hospitality Technologies shall not be held liable for shortages, backorders, delays, or price increases resulting from RAM supply constraints or other upstream component availability issues.")
    add_p(tf, "Acceptance of a quote, purchase order, or invoice acknowledges these potential supply chain conditions.")
    
    add_p(tf, "Caveats", bold=True)
    add_p(tf, "Any applicable License fees will be prorated according to each individual hotel's billing agreement and payment history with Insight.", bullet=True)
    add_p(tf, "Any access holes needed, in order to run Cat6 Ethernet wiring, are to be approved in advance by the customer.", bullet=True)
    add_p(tf, "It is the responsibility of the hotel to obtain approval from your particular Brand or any Owning Partners for any wire molding, access panels or access point enclosures that may be provided by Hospitality Technologies and its installers. If the hotel decides to forego access point enclosures so that the access points are in public view, it is their responsibility to obtain approval from their Brand or any Owning Partners.", bullet=True)
    add_p(tf, "Drywall and paint repairs are to be performed by the customer as needed unless otherwise discussed in writing.", bullet=True)
    add_p(tf, "Existing Cat5e Ethernet may be re-used throughout as needed provided they meet with our end to end cable testing.", bullet=True)
    add_p(tf, "Any electrical drops needed for a power source in the MDF and IDFs are to be completed by the customer before the install is scheduled.", bullet=True)
    add_p(tf, "Upon completion of installation, if quoted, Hospitality Technologies will provide a final heatmap to ensure -65dBm signal strength throughout as required by the latest Industry Standards.", bullet=True)

    # --- SLIDE 2: Disclosures ---
    slide2 = setup_slide("Acceptance: Disclosures")
    tb2 = slide2.shapes.add_textbox(Inches(0.83), Inches(1.0), Inches(11.5), Inches(5.5))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    
    add_p(tf2, "Disclosures", bold=True)
    add_p(tf2, "This proposal is hereby accepted and Hospitality Technologies is hereby authorized to proceed with work, contingent upon credit approval by Hospitality Technologies and receiving 80% down payment (not necessary with Lease Option). The final 20% is due no later than 10 days after completion. All orders received which are $5000.00 or less will be paid in full. If this quote is an equipment or license-only purchase, 100% of the total quoted is due upon acceptance. All credit card transactions are subject to a 5% processing fee.", bullet=True)
    add_p(tf2, "In the event that there are unforeseen circumstances which require extra equipment to be installed, Hospitality Technologies reserves the right to change the final quoted amount based on the extra equipment and/or labor required to install said equipment.", bullet=True)
    add_p(tf2, "License fees shown are estimated. Actual fees will be prorated from date of installation.", bullet=True)
    add_p(tf2, "Any required lodging shall be provided by the customer unless otherwise stated in advance.", bullet=True)
    add_p(tf2, "This installation is based on a single visit. If multiple visits are required, Hospitality Technologies will charge additional fees for travel on the final invoice.", bullet=True)
    add_p(tf2, "IPTV requires additional equipment. If customer chooses to add IPTV after the quote is completed a change order will be required for additional equipment and labor costs.", bullet=True)
    add_p(tf2, "It is further understood that acceptance of this proposal includes acceptance of all the attached terms, conditions and monthly charges. The total price listed includes all shipping and handling charges, but does not include any applicable state, local or national sales taxes. These additional state, local or national sales taxes will be included, as required, upon final billing. Any required lodging shall be provided by the customer unless otherwise stated in advance. Please sign and date all pages of the quote.", bullet=True)
    add_p(tf2, "* Monthly Lease Price is contingent on approved financing.", bold=True)

    if property_code:
        button_width = Inches(5.0)
        button_height = Inches(0.7)
        # Center horizontally: (13.333 - 5.0) / 2 = 4.166
        button_left = Inches(4.166)
        button_top = Inches(6.0)

        # 1. Add the modern CTA Button Image
        btn_image_stream = create_cta_button_image()
        cta_pic = slide2.shapes.add_picture(btn_image_stream, button_left, button_top, button_width, button_height)
        
        # 2. Apply hyperlink directly to the image object
        link_url = f"https://quote-presentation-generator.onrender.com/sign/{property_code}"
        cta_pic.click_action.hyperlink.address = link_url

        # 3. Inject BoldSign text tag BELOW the button so it is not hidden or stripped by LibreOffice
        tag_top = button_top + button_height + Inches(0.1)
        tag_tb = slide2.shapes.add_textbox(button_left, tag_top, button_width, Inches(0.4))
        tag_p = tag_tb.text_frame.paragraphs[0]
        tag_r = tag_p.add_run()
        # Valid BoldSign text tag syntax padded to establish field width
        tag_r.text = r"{{             sign|1|*             }}"
        # Use an off-white color (#FEFEFE) so the human eye can't see it easily, but the PDF renderer doesn't delete it as invisible
        tag_r.font.size = Pt(8)
        tag_r.font.color.rgb = RGBColor.from_string('FEFEFE')
        tag_p.alignment = PP_ALIGN.CENTER

def generate_presentation(pptx_source, excel_source, property_code=None, output_target=None):
    """
    Generates the presentation deck by extracting quote figures from an Excel workbook
    and injecting formatted slides into a base PowerPoint presentation.

    :param pptx_source: File path (str/Path) or file-like binary stream (e.g., io.BytesIO).
    :param excel_source: File path (str/Path) or file-like binary stream (e.g., io.BytesIO).
    :param output_target: Optional file path or stream to save to. If None, returns an in-memory io.BytesIO.
    :return: io.BytesIO stream or path where the presentation was saved.
    """
    if hasattr(pptx_source, 'seek'):
        pptx_source.seek(0)
    if hasattr(excel_source, 'seek'):
        excel_source.seek(0)

    prs = Presentation(pptx_source)
    orig_last_idx = len(prs.slides) - 1

    # Target and extract top-right logo blob
    logo_data = None
    if len(prs.slides) >= 6:
        slide_6 = prs.slides[5]
        for shape in slide_6.shapes:
            if shape.shape_type == 13 and shape.left > 8000000 and shape.top < 1000000:
                logo_data = {
                    'blob': shape.image.blob,
                    'left': shape.left,
                    'top': shape.top,
                    'width': shape.width,
                    'height': shape.height
                }
                break

    if not logo_data:
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.shape_type == 13 and shape.left > 8000000 and shape.top < 1000000:
                    logo_data = {
                        'blob': shape.image.blob,
                        'left': shape.left,
                        'top': shape.top,
                        'width': shape.width,
                        'height': shape.height
                    }
                    break
            if logo_data:
                break

    df = pd.read_excel(excel_source, sheet_name=0)

    sections = []
    current_section = None
    for i in range(len(df)):
        row = df.iloc[i]
        if len(row) < 6:
            continue
        col1 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        col2 = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        col3 = row.iloc[3]
        col4 = row.iloc[4]
        col5 = row.iloc[5]

        # Identify section headers
        if pd.isna(row.iloc[2]) and pd.isna(row.iloc[3]) and pd.notna(row.iloc[1]):
            if col1 and col1 != "Description" and not col1.startswith("SUBTOTAL") and not col1.startswith("Total") and "Acceptance" not in col1:
                current_section = col1
                sections.append({"section": current_section, "items": []})
                continue

        # Parse items
        if current_section:
            if col1.startswith("SUBTOTAL") or col1.startswith("Total"):
                current_section = None
                continue
            if col1 and col1 != "Description" and col1 != "Spares":
                qty = clean_float(col3)
                total = clean_float(col5)
                unit_cost = clean_float(col4)
                if unit_cost == 0.0 and qty > 0 and total > 0:
                    unit_cost = round(total / qty, 2)
                if qty > 0:
                    sections[-1]["items"].append({
                        "desc": col1,
                        "part": col2,
                        "qty": int(qty),
                        "unit_cost": unit_cost,
                        "total": total
                    })

    sections = [s for s in sections if s["items"]]

    # Pagination Mapping safely handling variable section counts
    hw = [sections[0]] if len(sections) > 0 else []
    sw_misc = sections[1:3] if len(sections) > 1 else []
    cabling_inst = sections[3:5] if len(sections) > 3 else []
    pm = [sections[5]] if len(sections) > 5 else []
    shipping = [sections[6]] if len(sections) > 6 else []
    monthly = sections[7:] if len(sections) > 7 else []

    grand_total = sum(it["total"] for s in sections for it in s["items"] if "Monthly" not in s["section"])
    monthly_total = sum(it["total"] for s in monthly for it in s["items"])

    # Slide Generation Sequence
    if hw:
        create_pdf_style_slide(prs, "IHG Connect Solution Investment", hw, logo_data=logo_data)
    if sw_misc:
        create_pdf_style_slide(prs, "IHG Connect Solution Investment (Cont.)", sw_misc, logo_data=logo_data)
    if cabling_inst:
        create_pdf_style_slide(prs, "Labor & Project Management Costs", cabling_inst, logo_data=logo_data)
    if pm:
        create_pdf_style_slide(prs, "Labor & Project Management Costs (Cont.)", pm, logo_data=logo_data)
    if shipping:
        create_pdf_style_slide(prs, "Shipping and Travel", shipping, logo_data=logo_data)

    if monthly:
        create_pdf_style_slide(prs, "Support After Installation", monthly, logo_data=logo_data, is_monthly=True)

    create_summary_cards_slide(prs, grand_total, monthly_total, logo_data=logo_data)
    create_acceptance_slides(prs, logo_data=logo_data, property_code=property_code)

    # Shift standard closer slide back to the end
    if orig_last_idx >= 0 and len(prs.slides) > 1:
        move_slide(prs, orig_last_idx, len(prs.slides) - 1)

    if output_target is None:
        out_stream = io.BytesIO()
        prs.save(out_stream)
        out_stream.seek(0)
        return out_stream
    elif isinstance(output_target, (str, bytes, os.PathLike)):
        prs.save(output_target)
        return output_target
    else:
        prs.save(output_target)
        if hasattr(output_target, 'seek'):
            output_target.seek(0)
        return output_target

def main(pptx_path=None, excel_path=None, output_filename=None):
    if pptx_path is None:
        pptx_path = "Design Overview - PDXCM.pptx"
    if excel_path is None:
        excel_path = "PDXCM-quote-27093.xlsx"
    if output_filename is None:
        output_filename = "Design Overview - Automated.pptx"

    result = generate_presentation(pptx_path, excel_path, output_filename)
    print(f"Success! Finished slide generation -> {output_filename}")
    return result

if __name__ == "__main__":
    
    # ------------------
    # FILE SETTINGS
    # ------------------
    BASE_PRESENTATION = "Design Overview - PDXCM.pptx"
    EXCEL_QUOTE_FILE  = "PDXCM-quote-27093.xlsx" 
    OUTPUT_FILE_NAME  = "Design Overview - Automated.pptx"
    
    main(BASE_PRESENTATION, EXCEL_QUOTE_FILE, OUTPUT_FILE_NAME)
