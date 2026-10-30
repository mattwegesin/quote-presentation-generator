import os
import sys

# Add current directory to path so we can import generate_gsa_docx
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from generate_gsa_docx import generate_gsa_docx_to_pdf

def test_generation():
    template_path = "GSA_TAGS.docx"
    
    if not os.path.exists(template_path):
        print(f"Error: Could not find {template_path}")
        return
        
    print("Generating DOCX to PDF test document...")
    
    # Generate client copy (with CTA button)
    output_stream = generate_gsa_docx_to_pdf(
        template_path=template_path,
        property_code="TEST1234",
        signer_name="John Doe",
        signer_email="john@example.com",
        hotel_name="Test Hotel & Suites",
        address="1234 Test Ave\nCity, ST 12345",
        room_count=200,
        term="36 Months",
        studio="YES",
        meeting_fee=150.00,
        notes="These are test notes added to the end of the document.",
        for_boldsign=False
    )
    
    output_filename = "TEST_OUTPUT_GSA_DOCX.pdf"
    with open(output_filename, "wb") as f:
        f.write(output_stream.getvalue())
        
    print(f"Success! Test document saved as {output_filename}")

if __name__ == "__main__":
    test_generation()
