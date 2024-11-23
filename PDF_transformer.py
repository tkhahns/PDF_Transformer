import os
import sys
import io
from PyPDF2 import PdfReader, PdfWriter, Transformation, PageObject
from reportlab.pdfgen import canvas
import win32com.client

wdFormatPDF = 17  # Word file format for PDF export

def convert_word_to_pdf(input_docx_path, output_pdf_path):
    """
    Convert a Word document to a PDF.
    """
    word = win32com.client.Dispatch('Word.Application')
    try:
        doc = word.Documents.Open(os.path.abspath(input_docx_path))
        doc.SaveAs(os.path.abspath(output_pdf_path), FileFormat=wdFormatPDF)
        print(f"Converted Word to PDF: {output_pdf_path}")
    finally:
        doc.Close()
        word.Quit()

def create_blank_page(width, height):
    """
    Create a blank PDF page with specified width and height.
    """
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(width, height))
    can.showPage()
    can.save()

    packet.seek(0)
    return PdfReader(packet).pages[0]

def merge_pages_horizontally(left_page, blank_page):
    """
    Merge two pages horizontally. Place the original page on the left and the blank page on the right.
    """
    new_page = PageObject.create_blank_page(
        width=left_page.mediabox.width + blank_page.mediabox.width,
        height=left_page.mediabox.height
    )
    new_page.merge_page(left_page)
    blank_page.add_transformation(Transformation().translate(left_page.mediabox.width, 0))
    new_page.merge_page(blank_page)
    return new_page

def merge_pages_vertically(top_page, blank_page):
    """
    Merge two pages vertically. Place the original page at the top and the blank page at the bottom.
    """
    new_page = PageObject.create_blank_page(
        width=top_page.mediabox.width,
        height=top_page.mediabox.height + blank_page.mediabox.height
    )
    new_page.merge_page(top_page)
    blank_page.add_transformation(Transformation().translate(0, -top_page.mediabox.height))
    new_page.merge_page(blank_page)
    return new_page

def transform_pdf(input_pdf_path, output_pdf_path):
    """
    Transform a PDF by adding blank pages horizontally or vertically based on page orientation.
    """
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()

    for page_num in range(len(reader.pages)):
        original_page = reader.pages[page_num]
        width = original_page.mediabox.width
        height = original_page.mediabox.height

        # Create a blank page with the same dimensions
        blank_page = create_blank_page(width, height)

        # Merge based on orientation
        if height > width:  # Portrait mode
            new_page = merge_pages_horizontally(original_page, blank_page)
        else:  # Landscape mode
            new_page = merge_pages_vertically(original_page, blank_page)

        writer.add_page(new_page)

    # Write the output to the new PDF file
    with open(output_pdf_path, "wb") as output_file:
        writer.write(output_file)
    print(f"Transformed PDF saved to: {output_pdf_path}")

def main():
    """
    Main function to handle Word to PDF conversion and PDF transformation.
    """
    if len(sys.argv) < 2:
        print("Usage: python script.py <input_file.docx>")
        sys.exit(1)

    input_file = sys.argv[1]
    if not input_file.lower().endswith(".docx"):
        print("Input file must be a .docx Word document.")
        sys.exit(1)

    # Generate output file name
    input_name = os.path.splitext(os.path.basename(input_file))[0]
    output_pdf_path = f"my_{input_name}.pdf"

    # Convert Word to PDF
    temp_pdf_path = input_file.replace(".docx", ".pdf")
    convert_word_to_pdf(input_file, temp_pdf_path)

    # Transform the PDF
    transform_pdf(temp_pdf_path, output_pdf_path)

    # Clean up temporary PDF
    if os.path.exists(temp_pdf_path):
        os.remove(temp_pdf_path)
        print(f"Temporary PDF file removed: {temp_pdf_path}")

if __name__ == "__main__":
    main()
