from pdf2docx import Converter
import fitz
from docx import Document
from PIL import Image
from io import BytesIO

# Provide the paths for your PDF and Word files
pdf_file_path = f"C:\\Users\\PabloTinajas\\Downloads\\N1 INCYE300 L=18.21m.pdf"
docx_file_path = f"C:\\Users\\PabloTinajas\\Downloads\\N1 INCYE300 L=18.21m.docx"

# Open the PDF file
pdf_document = fitz.open(pdf_file_path)

# Create a new Word document
doc = Document()

# Iterate through each page of the PDF
for page_number in range(pdf_document.page_count):
    page = pdf_document.load_page(page_number)
    image_list = page.get_images(full=True)
    for img in image_list:
        xref = img[0]
        base_image = pdf_document.extract_image(xref)
        image_data = base_image["image"]
        image = Image.open(BytesIO(image_data))
        image.save(f"C:\\Users\\PabloTinajas\\Downloadsimage{page_number}_{img[1]}.png")  # save images as PNG
        doc.add_picture(f"C:\\Users\\PabloTinajas\\Downloadsimage{page_number}_{img[1]}.png")

# Save the Word document
doc.save(docx_file_path)
