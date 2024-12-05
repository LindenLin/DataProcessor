import os
import re
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.errors import PdfReadError
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from io import BytesIO

def create_overlay(text_header, text_footer, skill_type, year_level, page_size=A4):
    """
    Create a single-page PDF for header, footer and labels overlay.
    """
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=page_size)
    
    # Set header
    can.setFont("Helvetica-Bold", 12)
    can.drawString(40, page_size[1] - 30, text_header)  # Header position

    # Set skill type (Reading/Writing)
    can.setFont("Helvetica-Bold", 11)
    can.drawString(40, page_size[1] - 50, f"Skill: {skill_type}")

    # Set year level
    can.drawString(40, page_size[1] - 70, f"Year Level: {year_level}")

    # Set footer
    can.setFont("Helvetica", 10)
    can.drawString(40, 30, text_footer)  # Footer position

    can.save()
    packet.seek(0)
    return packet

def get_year_level(filename):
    """Extract year level from filename's six digit number."""
    match = re.search(r'(\d{6})', filename)
    if match:
        year_num = int(match.group(1)[:2])
        return f"Year {year_num}" if 1 <= year_num <= 7 else "Unknown Year"
    return "Unknown Year"

def get_skill_type(folder_path):
    """Determine if the folder path contains Reading or Writing."""
    folder_path = folder_path.lower()
    if 'reading' in folder_path:
        return 'Reading'
    elif 'writing' in folder_path:
        return 'Writing'
    return 'Unknown Skill'

def add_header_footer_to_pdf(input_pdf_path, output_pdf_path, header_text, footer_text, skill_type, year_level):
    """Add header, footer and labels to PDF files."""
    try:
        reader = PdfReader(input_pdf_path)
        writer = PdfWriter()
        
        if len(reader.pages) == 0:
            print(f"Warning: {input_pdf_path} appears to be empty or corrupted")
            return False
            
        for page in reader.pages:
            overlay_pdf = create_overlay(header_text, footer_text, skill_type, year_level)
            overlay_reader = PdfReader(overlay_pdf)
            overlay_page = overlay_reader.pages[0]
            page.merge_page(overlay_page)
            writer.add_page(page)

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
    """Process PDFs with additional labels based on folder name and filename."""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    success_count = 0
    error_count = 0
    
    for root, dirs, files in os.walk(folder_path):
        rel_path = os.path.relpath(root, folder_path)
        current_output_dir = os.path.join(output_folder, rel_path)
        
        if rel_path != '.':
            os.makedirs(current_output_dir, exist_ok=True)

        # Get skill type from current folder path
        skill_type = get_skill_type(root)

        for file_name in files:
            if file_name.lower().endswith('.pdf'):
                # Get year level from filename
                year_level = get_year_level(file_name)
                
                input_path = os.path.join(root, file_name)
                output_path = os.path.join(current_output_dir, file_name)
                
                if add_header_footer_to_pdf(input_path, output_path, header_text, footer_text, skill_type, year_level):
                    success_count += 1
                    print(f"Successfully processed: {os.path.relpath(input_path, folder_path)}")
                    print(f"Added labels - Skill: {skill_type}, {year_level}")
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
