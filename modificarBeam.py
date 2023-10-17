from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from tkinter import Tk, filedialog
import io
import getpass
username = getpass.getuser()


def insert_images(pdf_path, images, output_path):
    existing_pdf = PdfReader(pdf_path)
    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        page = existing_pdf.pages[page_num]

        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)

        if page_num % 2 == 0:
            can.drawImage(ImageReader(images[0]), 100, 700, width=200, height=100)  # Adjust the position as needed
        else:
            can.drawImage(ImageReader(images[1]), 100, 100, width=200, height=100)  # Adjust the position as needed

        can.save()

        # Move to the beginning of the buffer
        packet.seek(0)
        new_pdf = PdfReader(packet)
        for new_page in new_pdf.pages:
            output.add_page(new_page)

    with open(output_path, "wb") as out:
        output.write(out)


def main():
    root = Tk()
    root.withdraw()
    pdf_path = filedialog.askopenfilename()
    images = [
        f"C:\\Users\\{username}\\Incye\\Ingenieria - 12_Aplicaciones\\BeamMod\\incyelogo.jpg",
        f"C:\\Users\\{username}\\Incye\\Ingenieria - 12_Aplicaciones\\BeamMod\\imagen (2).png",
    ]  # Absolute paths of the images

    output_path = pdf_path.split(".pdf")[0] + "_edited.pdf"

    insert_images(pdf_path, images, output_path)
    print(f"New PDF generated at {output_path}")


if __name__ == "__main__":
    main()
