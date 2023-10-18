from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from tkinter import Tk, filedialog
import io
import getpass
username = getpass.getuser()


import PyPDF2

def insert_image(pdf_in, img, pdf_out):

  pdf_reader = PdfReader(pdf_in)
  pdf_writer = PdfWriter()

  for page_num in range(len(pdf_reader.pages)):

    pdf_page = pdf_reader.pages[page_num]
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)

    # Draw the existing page content first 
    can.saveState() 
    can.translate(0,0)
    can.drawImage(pdf_page, 0, 300)
    can.restoreState()

    # Draw the new image on top
    can.drawImage(img, 200, 200)  

    can.save()

    packet.seek(0)  
    new_pdf = PdfReader(packet)  
    pdf_writer.add_page(new_pdf.pages[0])

  with open(pdf_out, "wb") as f:
    pdf_writer.write(f)

def main():
    root = Tk()
    root.withdraw()
    pdf_path = filedialog.askopenfilename()
    image1 = f"C:\\Users\\{username}\\Incye\\Ingenieria - Documentos\\12_Aplicaciones\\BeamMod\\incyelogo.jpg"
    image2 = f"C:\\Users\\{username}\\Incye\\Ingenieria - Documentos\\12_Aplicaciones\\BeamMod\\imagen (2).png"
      # Absolute paths of the images

    output_path = pdf_path.split(".pdf")[0] + "_edited.pdf"

    insert_image(pdf_path, image1, output_path)
    print(f"New PDF generated at {output_path}")
    
if __name__ == "__main__":
    main()
