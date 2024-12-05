import os
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.errors import PdfReadError
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from io import BytesIO

def create_overlay(text_header, text_footer, page_size=A4):
    """
    Create a single-page PDF for header and footer overlay.
    """
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=page_size)
    
    # Set header
    can.setFont("Helvetica-Bold", 12)
    can.drawString(40, page_size[1] - 30, text_header)  # Header position

    # Set footer
    can.setFont("Helvetica", 10)
    can.drawString(40, 30, text_footer)  # Footer position

    can.save()
    packet.seek(0)
    return packet

def add_header_footer_to_pdf(input_pdf_path, output_pdf_path, header_text, footer_text):
    """
    Add header and footer to PDF files.
    """
    try:
        reader = PdfReader(input_pdf_path)
        writer = PdfWriter()
        
        # Validate PDF file integrity
        if len(reader.pages) == 0:
            print(f"Warning: {input_pdf_path} appears to be empty or corrupted")
            return False
            
        # Main processing logic remains unchanged
        for page in reader.pages:
            # Create header and footer overlay content
            overlay_pdf = create_overlay(header_text, footer_text)
            overlay_reader = PdfReader(overlay_pdf)
            overlay_page = overlay_reader.pages[0]

            # Merge header and footer to current page
            page.merge_page(overlay_page)
            writer.add_page(page)

        # Save new PDF file
        with open(output_pdf_path, "wb") as output_file:
            writer.write(output_file)
        
        return True
        
    except PdfReadError as e:
        print(f"Error: Problem processing {input_pdf_path}: {str(e)}")
        return False
    except Exception as e:
        print(f"Unexpected error: Problem processing {input_pdf_path}: {str(e)}")
        return False

def process_folder(folder_path, output_folder, header_text, footer_text):
    """
    Recursively scan folder and its subfolders for all PDFs and add header and footer.
    
    Args:
        folder_path: Source folder path
        output_folder: Output folder path
        header_text: Header text
        footer_text: Footer text
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    success_count = 0
    error_count = 0
    
    for root, dirs, files in os.walk(folder_path):
        # Calculate relative path to create same directory structure in output
        rel_path = os.path.relpath(root, folder_path)
        current_output_dir = os.path.join(output_folder, rel_path)

        # Ensure output subfolder exists
        if rel_path != '.':
            os.makedirs(current_output_dir, exist_ok=True)

        # Process all PDF files in current folder
        for file_name in files:
            if file_name.lower().endswith('.pdf'):
                input_path = os.path.join(root, file_name)
                output_path = os.path.join(current_output_dir, file_name)
                
                if add_header_footer_to_pdf(input_path, output_path, header_text, footer_text):
                    success_count += 1
                    print(f"Successfully processed: {os.path.relpath(input_path, folder_path)}")
                else:
                    error_count += 1
    
    print(f"\nProcessing complete! Success: {success_count} files, Failed: {error_count} files")

# Define paths and text
input_folder = "data"  # Replace with your PDF folder path
output_folder = "ProcessedData"  # Replace with output folder path
footer = "New Zealand educational sources"
header = "only use for EduGPT Callaghan Innovation"

# Process all PDF files in the folder
process_folder(input_folder, output_folder, header, footer)
