# aplicación de cartas de interlocutores 
# necesitamos una API potente entre Excel, Word, python y una interfaz visual facilona

import win32com.client
import datetime
import tkinter as tk
from mailmerge import MailMerge
import openpyxl
import locale
import re
import os
from tkinter import *
from tkinter import messagebox, ttk
from ttkthemes import ThemedStyle
from docx import Document
from docx.shared import Inches
from docx2pdf import convert
from tkinter import filedialog
import PyPDF2

locale.setlocale(locale.LC_ALL, '')

class DocumentEditor:
    def __init__(self, document_path):
        self.document_path = document_path
        self.document = Document(self.document_path)

    def get_username(self):
        return os.getlogin()
    username = get_username()  # You can write a function to get the current username

# en esta clase se escribirá todo lo necesario para poder generar y editar el documento word





work_code = user_input.get()  # Get the work code entered by the user
username = self.get_username()  # You can write a function to get the current username
city_code = work_code[-1]  # Extract the city code from the work code
excel_path = f'C:\\Users\\{username}\\Incye\\Proyectos - Documentos\\{city_code}\\{work_code}\\01 Info\\{work_code}.xlsm'



###################### idea para sacar todos los datos necesarios del excel ###################################################################
workbook = openpyxl.load_workbook(f'C:\\Users\\{username}\\Incye\\Proyectos - Documentos\\{city_code}\\{work_code}\\01 Info\\{work_code}.xlsm')

# Select the "DATOS" sheet
worksheet = workbook['DATOS']

# Extract data from specific cells and store them in variables
delegado = worksheet['C7'].value
nombre_cliente = worksheet['C15'].value
contacto_cliente = worksheet['C17'].value
nombre_obra = worksheet['C19'].value
tlf_cliente = worksheet['C21'].value
email_cliente = worksheet['C23'].value
direccion_obra = worksheet['C25'].value


# Close the workbook when done
workbook.close()













































import datetime
import tkinter as tk
from mailmerge import MailMerge
import locale
import re
import os
from tkinter import *
from tkinter import messagebox, ttk
from ttkthemes import ThemedStyle
from docx import Document
from docx.shared import Inches
from docx2pdf import convert
from tkinter import filedialog
import PyPDF2
import getpass
username = getpass.getuser()



locale.setlocale(locale.LC_ALL, '')

class DocumentEditor:
    def __init__(self, document_path):
        self.document_path = document_path
        self.document = Document(self.document_path)

    def remove_empty_paragraphs_between_range(self, start_index, end_index):
        in_range = False

        paragraphs_to_remove = []

        for i, paragraph in enumerate(self.document.paragraphs):
            if i == start_index:
                in_range = True
            elif i == end_index + 1:
                in_range = False
            
            if in_range and not paragraph.text.strip():
                paragraphs_to_remove.append(paragraph)
        
        for paragraph in paragraphs_to_remove:
            p = paragraph._element
            p.getparent().remove(p)

    def get_username(self):
        return os.getlogin()
    username = get_username()  # You can write a function to get the current username


    def convert_to_pdf(word_file_path, pdf_file_path):
        convert(word_file_path, pdf_file_path)

    ################################################### Guardado del documento ######################################################################

    def save_document(self, output_path):
        self.document.save(output_path)
        pdf_output_path = os.path.splitext(output_path)[0] + ".pdf"
        convert(output_path, pdf_output_path)        


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("Carta de interlocutores")
        self.master.geometry("880x930")
        self.master.configure(background="#F5F5F5")
        self.pack(fill=tk.BOTH, expand=True)
        self.create_widgets()
                # Button styling 

    def create_widgets(self):

        # Codigo de obra
        self.codigo_frame = tk.Frame(self, bg="#F5F5F5")
        self.codigo_frame.pack(pady=15)
        self.codigo_label = tk.Label(self.codigo_frame, text="Codigo de la obra:", font=("Helvetica", 14), bg="#F5F5F5", fg="#333333")
        self.codigo_label.pack(side=tk.LEFT, padx=15)
        self.codigo_entry = tk.Entry(self.codigo_frame, font=("Helvetica", 14))
        self.codigo_entry.pack(side=tk.RIGHT, padx=15, expand=True, fill=tk.X)


        # jefe_frame de equipo
        self.jefe_frame = tk.Frame(self, bg="#F5F5F5")
        self.jefe_frame.pack(pady=15)
        self.jefe_label = tk.Label(self.jefe_frame, text="Jefe de equipo:", font=("Helvetica", 14), bg="#F5F5F5", fg="#333333")
        self.jefe_label.pack(side=tk.LEFT, padx=15)
        self.jefe_entry = tk.Entry(self.jefe_frame, font=("Helvetica", 14))
        self.jefe_entry.pack(side=tk.RIGHT, padx=15, expand=True, fill=tk.X)


        # Modificar button bg="#3986F3"
        self.fill_button = tk.Button(text="Generar", command=self.fill_template, font=("Helvetica", 16), bg="#FF6E40", fg="white",
                               padx=70,
                               pady=20)
        self.fill_button.pack()

        # Modificar button bg="#3986F3"
        self.fill_button = tk.Button(text="Lanzar correo", command=self.lanzar_correo, font=("Helvetica", 16), bg="#FF6E40", fg="white",
                               padx=70,
                               pady=20)
        self.fill_button.pack()

    def fill_template(self):
###################### idea para sacar todos los datos necesarios del excel ###################################################################

        codigo = self.codigo_entry.get()
        jefe = self.jefe_entry.get() 
        ciudad = codigo[-1]  # Extract the city code from the work code

        workbook = openpyxl.load_workbook(f'C:\\Users\\{username}\\Incye\\Proyectos - Documentos\\{ciudad}\\{codigo}\\01 Info\\{codigo}.xlsm')

        # Select the "DATOS" sheet
        worksheet = workbook['DATOS']

        # Extract data from specific cells and store them in variables
        delegado = worksheet['C7'].value
        nombre_cliente = worksheet['C15'].value
        contacto_cliente = worksheet['C17'].value
        nombre_obra = worksheet['C19'].value
        tlf_cliente = worksheet['C21'].value
        email_cliente = worksheet['C23'].value
        direccion_obra = worksheet['C25'].value


        # Close the workbook when done
        workbook.close()
        
        additional_info = {
        "JoseManuelMaldonadoM": "José M. Maldonado",
        "DavidLara": "David Lara",
        "": "Ezequiel Sánchez",
        "": "Andrés Rodríguez Pérez",
        "AlbertoAldamaMartine": "Alberto Aldama Martínez",
        "": "Adelaida Sáez Castejón",
        "AlejandroBuiles": "Alejandro Ángel Builes",
        "": "Juan José Morón Blanco",
        "": "Manuel González-Arquiso Madrigal",
        "RafaelMansilla.": "Rafael Mansilla Correa",
        "EstebanLopezFernande": "Esteban López Fernández"
        }

        additional_info2 = {
        "JoseManuelMaldonadoM": "624 563 882",
        "DavidLara": "624 296 118",
        "Ezequiel Sánchez.": "627 177 912",
        "Andrés Rodríguez.": "627 172 908",
        "AlbertoAldamaMartine": "627 172 717",
        "Adelaida Sáez.": "624 454 082",
        "AlejandroBuiles": "+573126286401",
        "Juan José Morón.": "627 197 867",
        "Manuel González.": "611 069 632",
        "RafaelMansilla.": "613 105 723", 
        "EstebanLopezFernande": ""
        }

        additional_info3 = {
        "José M. Maldonado": "Málaga",
        "David Lara.": "Madrid",
        "Ezequiel Sánchez.": "Madrid",
        "Andrés Rodríguez.": "Madrid",
        "Jorge Nebreda.": "Madrid",
        "Alberto Aldama.": "Bilbao",
        "Adelaida Sáez.": "Valencia",
        "Alejandro Ángel Builes.": "Madrid",
        "Juan José Morón.": "Sevilla",
        "Manuel González.": "Valladolid",
        "Rafael Mansilla.": "Madrid"
        }

        tecnico = additional_info.get(username, "")
        tel_tecnico = additional_info2.get(username, "")
        

        revisor_nota = additional_info.get(selected_option2, "")
        result2 = re.sub(' +', ' ', revisor_nota)
        revisor_nota = result2   

        siglas_autor = additional_info2.get(selected_option, "")
        siglas_rev = additional_info2.get(selected_option2, "")
        ciudad = additional_info3.get(selected_option, "")

        # Fechas
        current_date = datetime.datetime.now()
        formatted_date = current_date.strftime("%d/%m/%Y")
        dia = current_date.strftime("%d")
        mes = current_date.strftime("%B")  # la B es el mes en formato palabra
        anyo = current_date.strftime("%Y") # mis primeras 

        # cargamos la plantilla
        if checkbox es motaje :
            template = f"C:\Users\{username}\Incye\Ingenieria - 12_Aplicaciones\interlocutores\\ci_montaje.docx"
        if checkbox es sin montaje :
            template = f"C:\Users\{username}\Incye\Ingenieria - 12_Aplicaciones\interlocutores\\ci_sinmontaje.docx"

        document = MailMerge(template)

        # Sustituimos valores
        document.merge(Nombre_Cliente=nombre_cliente, Obra=obra, Direccion_Obra=Direccion_obra, Codigo_Obra=codigo, Fecha=formatted_date, Dia=dia, Mes=mes, Anyo=anyo, Autor_NotaC=autor_nota, Revisor_NotaC=revisor_nota, Inic_AutorNC = siglas_autor, Inic_RevNC = siglas_rev, Ciudad=ciudad)

        # Obtener los valores de las checkboxes
        checkbox_values = list(self.checkbar.state())
        checkbox_values2 = list(self.checkbar2.state())

        # Guardar los valores en el documento
        document.merge(Pipeshor6=checkbox_values[0], Pipeshor4L=checkbox_values[1], Pipeshor4S=checkbox_values[2], Megaprop=checkbox_values[3])
        document.merge(INCYE300=checkbox_values2[0], INCYE450=checkbox_values2[1], INCYE600=checkbox_values2[2], SuperSlim=checkbox_values2[3])
        
        if checkbox_values[0]:   # Alshor Plus
            document.merge(AL="El sistema Alshor Plus es un sistema de cimbra de aluminio con una capacidad de carga de hasta 120 kN por pie, compuesto por gatos ajustables, verticales, bastidores y cazoletas. Junto con el sistema Alshor Plus se utilizarán vigas Albeam como vigas de reparto.")
        if checkbox_values[1]:  # Shoring 75
            document.merge(SH="El sistema Kwikstage Shoring 75 es un sistema de cimbra de acero con una capacidad de carga de hasta 75 kN, compuesto por bases ajustables, verticales, espigas, horizontales y horquillas ajustables. Junto con el sistema Kwikstage Shoring 75 se utilizarán vigas Superslim como vigas de reparto.")
        if checkbox_values[2]:   # SuperSlim
            document.merge(SS="El sistema Superslim (área neta de la sección 19,64 cm2) es un sistema de perfilería constituido por vigas compuestas formadas por dos perfiles en C, unidos mediante presillas situadas en distintas posiciones que configuran un sistema de vigas modulares de gran versatilidad.")
        if checkbox_values[3]:   # Megaprop
            document.merge(MP="El sistema Megaprop (área sección 58,45 cm2) es un sistema de perfilería constituido por vigas compuestas formadas por dos perfiles en C, unidos mediante presillas situadas en distintas posiciones que configuran un sistema de vigas modulares de gran versatilidad. El acero utilizado para su fabricación es de la calidad S355.")
        if checkbox_values2[0]:   # Pipeshor 4L
            document.merge(Pip4L="El sistema Pipeshor 4L, con área de sección 100.13 cm2, es un sistema de puntales formados por módulos de tubos de 406 mm de diámetro y sus elementos asociados. Fabricado con acero S355 de 8 milímetros de espesor.")
        if checkbox_values2[1]:   # Pipeshor 4S
            document.merge(Pip4S="El sistema Pipeshor 4S, con área sección 196,24 cm2, es un sistema de puntales formados por módulos de tubos de 406 mm de diámetro y sus elementos asociados. Fabricado con acero S355 de 16 milímetros de espesor.")
        if checkbox_values2[2]:  # Pipeshor 6
            document.merge(Pip6="El sistema Pipeshor 6 (área sección 234,4 cm2) está formado por tubos de 610 mm de diámetro y sus elementos asociados. Fabricado con acero de calidad S355 y un espesor de 12,5 milímetros.")
        if checkbox_values2[3]:  # Granshor
            document.merge(GS="El sistema Granshor (área sección 72,59 cm2/cordón x 2 cordones) es un sistema de de celosías modular y sus elementos asociados. Fabricado con acero S355.")
        
        
        
        ## textos de la metodología de cálculo sabes?

        # 1000
        if checkbox_values[0] and not checkbox_values[1] and not checkbox_values[2] and not checkbox_values[3]:      
            document.merge(SST="Nota: En el Technical Data Sheet del sistema Alshor los valores incluidos en las gráficas son valores en Estado Límite de Servicio, es decir, son valores ya minorados por un coeficiente de 1,50. Para comparar frente a cargas mayoradas es necesario multiplicar los valores admisibles de las gráficas por 1,50 para no tener en cuenta un factor de mayoración duplicado.")
        # 1001
        if checkbox_values[0] and not checkbox_values[1] and not checkbox_values[2] and checkbox_values[3]:    
            document.merge(SST="Nota: En el Technical Data Sheet del sistema Alshor, así como los valores del Megaprop, los valores incluidos en las gráficas son valores en Estado Límite de Servicio, es decir, son valores ya minorados por un coeficiente de 1,50. Para comparar frente a cargas mayoradas es necesario multiplicar los valores admisibles de las gráficas por 1,50 para no tener en cuenta un factor de mayoración duplicado.")
        # 1010
        if checkbox_values[0] and not checkbox_values[1] and checkbox_values[2] and not checkbox_values[3]:
            document.merge(SST="Nota: En el Technical Data Sheet del sistema Alshor, así como en el del Superslim, los valores incluidos en las gráficas son valores en Estado Límite de Servicio, es decir, son valores ya minorados por un coeficiente de 1,50. Para comparar frente a cargas mayoradas es necesario multiplicar los valores admisibles de las gráficas por 1,50 para no tener en cuenta un factor de mayoración duplicado.")
        # 1011
        if checkbox_values[0] and not checkbox_values[1] and checkbox_values[2] and checkbox_values[3]:
            document.merge(SST="Nota: En el Technical Data Sheet del sistema Alshor, así como en el del Superslim y el Megaprop, los valores incluidos en las gráficas son valores en Estado Límite de Servicio, es decir, son valores ya minorados por un coeficiente de 1,50. Para comparar frente a cargas mayoradas es necesario multiplicar los valores admisibles de las gráficas por 1,50 para no tener en cuenta un factor de mayoración duplicado.")
        # 1100
        if checkbox_values[0] and checkbox_values[1] and not checkbox_values[2] and not checkbox_values[3]:
            document.merge(SST="Nota: En el Technical Data Sheet del sistema Alshor, así como en el del KS, los valores incluidos en las gráficas son valores en Estado Límite de Servicio, es decir, son valores ya minorados por un coeficiente de 1,50. Para comparar frente a cargas mayoradas es necesario multiplicar los valores admisibles de las gráficas por 1,50 para no tener en cuenta un factor de mayoración duplicado.")
        # 1101
        if checkbox_values[0] and checkbox_values[1] and not checkbox_values[2] and checkbox_values[3]:
            document.merge(SST="Nota: En el Technical Data Sheet del sistema Alshor, así como en el del KS y el Megaprop, los valores incluidos en las gráficas son valores en Estado Límite de Servicio, es decir, son valores ya minorados por un coeficiente de 1,50. Para comparar frente a cargas mayoradas es necesario multiplicar los valores admisibles de las gráficas por 1,50 para no tener en cuenta un factor de mayoración duplicado.")
        # 1110
        if checkbox_values[0] and checkbox_values[1] and checkbox_values[2] and not checkbox_values[3]:
            document.merge(SST="Nota: En el Technical Data Sheet del sistema Alshor, así como en el del KS y el Superslim, los valores incluidos en las gráficas son valores en Estado Límite de Servicio, es decir, son valores ya minorados por un coeficiente de 1,50. Para comparar frente a cargas mayoradas es necesario multiplicar los valores admisibles de las gráficas por 1,50 para no tener en cuenta un factor de mayoración duplicado.")
        # 1111
        if checkbox_values[0] and checkbox_values[1] and checkbox_values[2] and checkbox_values[3]:
            document.merge(SST="Nota: En el Technical Data Sheet del sistema Alshor, así como en el del Superslim, el KS y el Megaprop, los valores incluidos en las gráficas son valores en Estado Límite de Servicio, es decir, son valores ya minorados por un coeficiente de 1,50. Para comparar frente a cargas mayoradas es necesario multiplicar los valores admisibles de las gráficas por 1,50 para no tener en cuenta un factor de mayoración duplicado.")
        # 0001
        if not checkbox_values[0] and not checkbox_values[1] and not checkbox_values[2] and checkbox_values[3]:
            document.merge(SST="Nota: En el Technical Data Sheet del sistema Megaprop los valores incluidos en las gráficas son valores en Estado Límite de Servicio, es decir, son valores ya minorados por un coeficiente de 1,50. Para comparar frente a cargas mayoradas es necesario multiplicar los valores admisibles de las gráficas por 1,50 para no tener en cuenta un factor de mayoración duplicado.")
        # 0010
        if not checkbox_values[0] and not checkbox_values[1] and checkbox_values[2] and not checkbox_values[3]:
            document.merge(SST="Nota: En el Technical Data Sheet del sistema Superslim los valores incluidos en las gráficas son valores en Estado Límite de Servicio, es decir, son valores ya minorados por un coeficiente de 1,50. Para comparar frente a cargas mayoradas es necesario multiplicar los valores admisibles de las gráficas por 1,50 para no tener en cuenta un factor de mayoración duplicado.")
        # 0011
        if not checkbox_values[0] and not checkbox_values[1] and checkbox_values[2] and checkbox_values[3]:
            document.merge(SST="Nota: En el Technical Data Sheet del sistema Superslim, así como en el del Megaprop, los valores incluidos en las gráficas son valores en Estado Límite de Servicio, es decir, son valores ya minorados por un coeficiente de 1,50. Para comparar frente a cargas mayoradas es necesario multiplicar los valores admisibles de las gráficas por 1,50 para no tener en cuenta un factor de mayoración duplicado.")
        # 0100
        if not checkbox_values[0] and checkbox_values[1] and not checkbox_values[2] and not checkbox_values[3]:
            document.merge(SST="Nota: En el Technical Data Sheet del sistema KS los valores incluidos en las gráficas son valores en Estado Límite de Servicio, es decir, son valores ya minorados por un coeficiente de 1,50. Para comparar frente a cargas mayoradas es necesario multiplicar los valores admisibles de las gráficas por 1,50 para no tener en cuenta un factor de mayoración duplicado.")
        # 0101
        if not checkbox_values[0] and checkbox_values[1] and not checkbox_values[2] and checkbox_values[3]:
            document.merge(SST="Nota: En el Technical Data Sheet del sistema KS, así como en el del Megaprop, los valores incluidos en las gráficas son valores en Estado Límite de Servicio, es decir, son valores ya minorados por un coeficiente de 1,50. Para comparar frente a cargas mayoradas es necesario multiplicar los valores admisibles de las gráficas por 1,50 para no tener en cuenta un factor de mayoración duplicado.")
        # 0110
        if not checkbox_values[0] and checkbox_values[1] and checkbox_values[2] and not checkbox_values[3]:
            document.merge(SST="Nota: En el Technical Data Sheet del sistema KS, así como en el del Superslim, los valores incluidos en las gráficas son valores en Estado Límite de Servicio, es decir, son valores ya minorados por un coeficiente de 1,50. Para comparar frente a cargas mayoradas es necesario multiplicar los valores admisibles de las gráficas por 1,50 para no tener en cuenta un factor de mayoración duplicado.")
        # 0111
        if not checkbox_values[0] and checkbox_values[1] and checkbox_values[2] and checkbox_values[3]:
            document.merge(SST="Nota: En el Technical Data Sheet del sistema KS, así como en el del Superslim y el Megaprop, los valores incluidos en las gráficas son valores en Estado Límite de Servicio, es decir, son valores ya minorados por un coeficiente de 1,50. Para comparar frente a cargas mayoradas es necesario multiplicar los valores admisibles de las gráficas por 1,50 para no tener en cuenta un factor de mayoración duplicado.")

        # Ruta de guardado
        doc_modified = "C:/Memorias y servidor/Apeos/Plantilla_apeos.docx"
        document.write(doc_modified)
        
        if __name__ == '__main__':
            document_path = 'C:/Memorias y servidor/Apeos/Plantilla_apeos.docx'

            # texto para meter pdfs
            texto_apendice = "Y MECÁNICAS MATERIALES INCYE"
            
            # texto e imágenes del Superslim
            texto_SS = "El sistema Superslim (área neta de la sección 19,64 cm2) es un sistema de perfilería constituido por vigas compuestas formadas por dos perfiles en C, unidos mediante presillas situadas en distintas posiciones que configuran un sistema de vigas modulares de gran versatilidad."
           
            start_paragraph_index = 40
            end_paragraph_index = 49 

            if added_imagen_AL or added_imagen_SH or added_imagen_SS or added_imagen_MP or added_imagen_GS or added_imagen_PS4 or added_imagen_PS2 or added_imagen_PS6 or added_image_TDS_SH or added_imagen_TDS_SS or added_image_TDS_P or added_imagen_TDS_GS or added_imagen_TDS_MP:
                if self.output_path:
                    document_editor.remove_empty_paragraphs_between_range(start_paragraph_index, end_paragraph_index)
                    
                    document_editor.save_document(self.output_path)
                    pdf_path = os.path.splitext(self.output_path)[0] + ".pdf"
                    document_editor.in_planos(self.planos_paths, pdf_path) 
                    document_editor.append_pdf(self.apendice_paths, pdf_path) 
                    #pdf_output_path = self.output_path.replace(".docx", ".pdf") 
                    #convert(self.output_path, pdf_output_path)
            else:
                print("The target paragraph was not found in the document. Image not added.")

        messagebox.showinfo("Completado!", "La nota de cálculo ha sido creada y guardada con éxito.")
        

class Checkbar(Frame):
    def __init__(self, parent=None, picks=[], side=LEFT, anchor=W, checkbox_font=None):
        Frame.__init__(self, parent)
        self.vars = []
        self.checkbox_font = checkbox_font if checkbox_font else ("Helvetica", 12)

        for pick in picks:
            var = IntVar()
            chk = Checkbutton(self, text=pick, variable=var, font=self.checkbox_font)
            chk.pack(side=side, anchor=anchor, expand=YES)
            self.vars.append(var)

    def state(self):
        return map((lambda var: var.get()), self.vars)


# Launch the UI
root = tk.Tk()
app = Application(master=root)
app.mainloop()


