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
import docx2pdf
from docx2pdf import convert
from tkinter import filedialog
import PyPDF2
import getpass
import openpyxl
import win32com.client
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
        self.master.geometry("430x350")
        self.master.configure(background="#F5F5F5")
        self.pack(fill=tk.BOTH, expand=True)
        self.create_widgets()
                # Button styling 

    def get_username(self):
        return os.getlogin()

    def create_widgets(self):

        # Codigo de obra
        self.codigo_frame = tk.Frame(self, bg="#F5F5F5")
        self.codigo_frame.pack(pady=15)
        self.codigo_label = tk.Label(self.codigo_frame, text="Codigo de la obra:", font=("Helvetica", 14), bg="#F5F5F5", fg="#333333")
        self.codigo_label.pack(side=tk.LEFT, padx=15)
        self.codigo_entry = tk.Entry(self.codigo_frame, font=("Helvetica", 14))
        self.codigo_entry.pack(side=tk.RIGHT, padx=15, expand=True, fill=tk.X)

        # Jefe de Equipo 
        self.combobox_frame = tk.Frame(self, bg="#F5F5F5")
        self.combobox_frame.pack(pady=15)

        self.label1 = ttk.Label(self.combobox_frame, text="Jefe de Equipo", font=("Arial", 14))
        self.label1.grid(column=0, row=0, padx=11, pady=11)
        self.opcion_autor = tk.StringVar()
        opciones = ("Alejandro Elías", "Emilio Fernández", "Fernando García", "José A. Sineiro", "Julián García", "Khalil Ghanfari", "Luis Ismael Rodríguez", "Luis Martínez", "Tomás González", "Aomar Buzeid", "Oumar Cisse", "Adrián Rodríguez", "Mayade Diagne", "Santiago Núñez")
        self.combobox_autor = ttk.Combobox(self.combobox_frame, width=30, textvariable=self.opcion_autor, values=opciones, font=("Arial", 12), style='Custom.TCombobox')
        self.combobox_autor.current(0)
        self.combobox_autor.grid(column=0, row=1, padx=11, pady=11)

        # Checkboxes
        self.checkboxes_frame = tk.Frame(self, bg="#F5F5F5")
        self.checkboxes_frame.pack(pady=15)
        self.checkbar = Checkbar(self.checkboxes_frame, ['Con montaje', 'Sin montaje'], checkbox_font=("Helvetica", 14))
        self.checkbar.pack(side=tk.TOP, fill=tk.X, padx=15)
        self.checkbar.config(relief=tk.GROOVE, bd=4)


        # Modificar button bg="#3986F3"
        self.fill_button = tk.Button(text="Generar", command=self.fill_template, font=("Helvetica", 16), bg="#990000", fg="white",
                               padx=50,
                               pady=13)
        self.fill_button.pack()

    def fill_template(self):
###################### idea para sacar todos los datos necesarios del excel ###################################################################
        username = self.get_username() 
        codigo = self.codigo_entry.get()
        lciudad = codigo[-1]  # Extract the city code from the work code
        jefe = self.combobox_autor.get()

        if lciudad == "M":
            ciudad = "MAD"
        if lciudad == "T":
            ciudad = "BCN"
        if lciudad == "N":
            ciudad = "BLB"
        if lciudad == "F":
            ciudad = "CRN"
        if lciudad == "E":
            ciudad = "EXP"
        if lciudad == "V":
            ciudad = "LEV"
        if lciudad == "P":
            ciudad = "PT"
        if lciudad == "X":
            ciudad = "SEV"
        

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
        "EzequielSanchezdelaG": "Ezequiel Sánchez",
        "": "Andrés Rodríguez Pérez",
        "AlbertoAldamaMartine": "Alberto Aldama Martínez",
        "": "Adelaida Sáez Castejón",
        "AlejandroBuiles": "Alejandro Ángel Builes",
        "JuanJoseMoron": "Juan José Morón Blanco",
        "": "Manuel González-Arquiso Madrigal",
        "RafaelMansilla.": "Rafael Mansilla Correa",
        "EstebanLopezFernande": "Esteban López Fernández", 
        "PabloTinajas": "David Lara"
        }

        additional_info2 = {
        "JoseManuelMaldonadoM": "624 563 882",
        "DavidLara": "624 296 118",
        "EzequielSanchezdelaG": "627 177 912",
        "": "627 172 908",
        "AlbertoAldamaMartine": "627 172 717",
        "Adelaida Sáez.": "624 454 082",
        "AlejandroBuiles": "+573126286401",
        "JuanJoseMoron": "627 197 867",
        "": "611 069 632",
        "RafaelMansilla.": "613 105 723", 
        "EstebanLopezFernande": "",
        "PabloTinajas": "624 296 118"
        }

        additional_info3 = {
        "MAD": "Alejandro/Luis Ismael",
        "BCN": "Julián",
        "BLB": "Luis Ismael",
        "CRN": "Alejandro/Luis Ismael",
        "EXP": "Luis Ismael",
        "FRA": "Leyla",
        "LEV": "Julián",
        "PT": "Julián",
        "SEV": "Alejandro"
        }

        additional_info4 = {
        "MAD": "676 554 345/624 402 367",
        "BCN": "636 972 592",
        "BLB": "624 402 367",
        "CRN": "676 554 345/624 402 367",
        "EXP": "624 402 367",
        "FRA": "Leyla",
        "LEV": "636 972 592",
        "PT": "636 972 592",
        "SEV": "676 554 345"
        }

        additional_info5 = {
        "AVS": "Antonio Vázquez Sánchez",
        "CENT": "Ignacio Merlín López",
        "BCN": "David Suárez",
        "BLB": "Ibai Marlasca",
        "EXP": "Antonio Vázquez Sánchez",
        "FRA": "Leyla",
        "LEV": "Ismael Pérez",
        "PT": "",
        "AND": "Javier Álvarez"
        }    

        additional_info6 = {
        "AVS": "629 075 233",
        "CENT": "615 201 952",
        "BCN": "627 197 582",
        "BLB": "623 255 811",
        "EXP": "",
        "FRA": "",
        "LEV": "611 069 601",
        "PT": "",
        "AND": "623 491 208"
        }    

        additional_info8 = {
        "AVS": "antonio.vazquez@incye.com",
        "CENT": "madrid@incye.com",
        "BCN": "barcelona@incye.com",
        "BLB": "bilbao@incye.com",
        "EXP": "antonio.vazquez@incye.com",
        "FRA": "leyla.bentahar@incye.com",
        "LEV": "valencia@incye.com",
        "PT": "",
        "AND": "malaga@incye.com"
        }  


        additional_info7 = {
        "Alejandro Elías": "676 554 345",
        "Emilio Fernández": "678 045 791",
        "Fernando García": "622 249 861",
        "José A. Sineiro": "628 467 283",
        "Julián García": "636 972 592",
        "Khalil Ghanfari": "627 191 627",
        "Luis Ismael Rodríguez": "624 402 367",
        "Luis Martínez": "626 142 081",
        "Tomás González": "627 882 046",
        "Aomar Buzeid": "627 178 935",
        "Oumar Cisse": "647 673 181",
        "Adrián Rodríguez": "647 673 189",
        "Mayade Diagne": "647 673 175",
        "Santiago Núñez": "647 673 207",
        }         

        tecnico = additional_info.get(username, "")
        tel_tecnico = additional_info2.get(username, "")
        encargado_prod = additional_info3.get(ciudad, "")
        tel_prod = additional_info4.get(ciudad, "")
        delegado1 = additional_info5.get(delegado, "")
        tel_delegado = additional_info6.get(delegado, "")
        tel_jefe = additional_info7.get(jefe, "")
        email_delegado = additional_info8.get(delegado, "")

        # Fechas
        current_date = datetime.datetime.now()
        formatted_date = current_date.strftime("%d/%m/%Y")

        # Obtener los valores de las checkboxes
        checkbox_values = list(self.checkbar.state())
        if checkbox_values[0]:
            template = f"C:\\Users\\{username}\\Incye\\Ingenieria - 12_Aplicaciones\\interlocutores\\ci_montaje.docx"
        if checkbox_values[1]:
            template = f"C:\\Users\\{username}\\Incye\\Ingenieria - 12_Aplicaciones\\interlocutores\\ci_sinmontaje.docx"

        document = MailMerge(template)

        # Sustituimos valores
        document.merge(Nombre_Contacto = contacto_cliente, Nombre_Cliente=nombre_cliente, Nombre_Obra=nombre_obra, Fecha=formatted_date, Delegado=delegado1, Tel_Delegado=tel_delegado, Tecnico=tecnico, Tel_Tecnico=tel_tecnico, Jefe=jefe, Tel_Jefe=tel_jefe, Encargado_Prod=encargado_prod, Tel_Prod=tel_prod)
        
        output_path = f'C:\\Users\\{username}\\Incye\\Proyectos - Documentos\\{ciudad}\\{codigo}\\09 Comunicados\\3_Docs\\{codigo}_CI.docx'
        document.write(output_path)
        pdf_path = output_path.replace(".docx", ".pdf")
        docx2pdf.convert(output_path, pdf_path)

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.To = email_cliente
        mail.Subject = f"{codigo} {nombre_obra} CARTA DE INTERLOCUTORES" 
        mail.CC = email_delegado
        mail.Body = "Estimado cliente, \nadjunto Carta de interlocutores con los datos de contacto del personal de INCYE relacionado con la obra. \nUn cordial saludo."
        mail.Attachments.Add(pdf_path) 

        mail.Display()


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


