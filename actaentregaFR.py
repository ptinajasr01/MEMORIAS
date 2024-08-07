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
from docx.shared import Pt
username = getpass.getuser()



locale.setlocale(locale.LC_ALL, '')

class DocumentEditor:
    def __init__(self, document_path):
        self.document_path = document_path
        self.document = Document(self.document_path)

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
        self.master.title("Carta de interlocutores FRANCIA")
        self.master.geometry("730x798")
        self.master.configure(background="#F5F5F5")
        self.pack(fill=tk.BOTH, expand=True)
        self.create_widgets()
                # Button styling 

    def get_username(self):
        return os.getlogin()

    def create_widgets(self):


        # Seleccionar familia de materiales
        self.familia_frame = tk.Frame(self, bg="#F5F5F5")
        self.familia_frame.pack(pady=15)
        self.familia_label = tk.Label(self.familia_frame, text="Seleccionar tipo de obra:", font=("Helvetica", 14), bg="#F5F5F5", fg="#333333")
        self.familia_label.pack(side=tk.LEFT, padx=15)

        # Checkboxes
        self.checkboxes_frame = tk.Frame(self, bg="#F5F5F5")
        self.checkboxes_frame.pack(pady=15)
        self.checkbar = Checkbar(self.checkboxes_frame, ['ACODALAMIENTO', 'APEO/APUNTALAMIENTO', 'ESTABILIZADOR'], checkbox_font=("Helvetica", 14))
        self.checkbar.pack(side=tk.TOP, fill=tk.X, padx=15)
        self.checkbar.config(relief=tk.GROOVE, bd=4)

        # Codigo de obra
        self.codigo_frame = tk.Frame(self, bg="#F5F5F5")
        self.codigo_frame.pack(pady=15)
        self.codigo_label = tk.Label(self.codigo_frame, text="Codigo de la obra:", font=("Helvetica", 14), bg="#F5F5F5", fg="#333333")
        self.codigo_label.pack(side=tk.LEFT, padx=15)
        self.codigo_entry = tk.Entry(self.codigo_frame, font=("Helvetica", 14))
        self.codigo_entry.pack(side=tk.RIGHT, padx=15, expand=True, fill=tk.X)

        # Estructura
        self.estr_frame = tk.Frame(self, bg="#F5F5F5")
        self.estr_frame.pack(pady=15)
        self.estr_label = tk.Label(self.estr_frame, text="Estructura:", font=("Helvetica", 14), bg="#F5F5F5", fg="#333333")
        self.estr_label.pack(side=tk.LEFT, padx=15)
        self.estr_entry = tk.Entry(self.estr_frame, font=("Helvetica", 14))
        self.estr_entry.pack(side=tk.RIGHT, padx=15, expand=True, fill=tk.X)

        # Jefe de Equipo 
        self.combobox_frame = tk.Frame(self, bg="#F5F5F5")
        self.combobox_frame.pack(pady=15)

        self.label1 = ttk.Label(self.combobox_frame, text="Jefe de Equipo", font=("Arial", 14))
        self.label1.grid(column=0, row=0, padx=11, pady=11)
        self.opcion_autor = tk.StringVar()
        opciones = ("Alejandro Elías", "Emilio Fernández", "Fernando García", "José A. Sineiro", "Julián García", "Khalil Ghanfari", "Luis Martínez", "Tomás González", "Aomar Buzeid", "Adrián Rodríguez", "Mayade Diagne", "Santiago Núñez", "Beñat Larinzgoitia", "Albert Petre", "Kevin Huacanca", "Johan Antezana")
        self.combobox_autor = ttk.Combobox(self.combobox_frame, width=30, textvariable=self.opcion_autor, values=opciones, font=("Arial", 12), style='Custom.TCombobox')
        self.combobox_autor.current(0)
        self.combobox_autor.grid(column=0, row=1, padx=11, pady=11)

        self.label_revisor = ttk.Label(self.combobox_frame, text="¿Nota de cálculo?", font=("Arial", 14))
        self.label_revisor.grid(column=0, row=2, padx=11, pady=11)
        self.opcion_revisor = tk.StringVar()
        opciones2= ("Sí", "No")
        self.combobox_revisor = ttk.Combobox(self.combobox_frame, width=30, textvariable=self.opcion_revisor, values=opciones2, font=("Arial", 12), style='Custom.TCombobox')
        self.combobox_revisor.current(0)
        self.combobox_revisor.grid(column=0, row=3, padx=11, pady=11)

        # Segundo Revisor
        self.combobox2_frame = tk.Frame(self, bg="#F5F5F5")
        self.combobox2_frame.pack(pady=15)

        self.label2 = ttk.Label(self.combobox2_frame, text="Segundo Revisor", font=("Arial", 14))
        self.label2.grid(column=0, row=0, padx=11, pady=11)
        self.opcion2_autor = tk.StringVar()
        opciones3 = ("Álvaro Abad", "Santiago Venencia", "Antonio Vázquez", "Ignacio Merlín", "Álvaro Milla", "Javier Álvarez", "David Suárez", "Ana Seoane", "Luis José Iglesias", "Ismael Pérez", "Iván Valero", "Ibai Marlasca")
        self.combobox2_autor = ttk.Combobox(self.combobox2_frame, width=30, textvariable=self.opcion2_autor, values=opciones3, font=("Arial", 12), style='Custom.TCombobox')
        self.combobox2_autor.current(0)
        self.combobox2_autor.grid(column=0, row=1, padx=11, pady=11)

        # Nota adicional
        self.not_frame = tk.Frame(self, bg="#F5F5F5")
        self.not_frame.pack(pady=15)
        self.not_label = tk.Label(self.not_frame, text="Nota adicional:", font=("Helvetica", 14), bg="#F5F5F5", fg="#333333")
        self.not_label.pack(side=tk.LEFT, padx=15)
        self.not_entry = tk.Entry(self.not_frame, font=("Helvetica", 14))
        self.not_entry.pack(side=tk.RIGHT, padx=15, expand=True, fill=tk.X)

        # Fecha inspección final:
        self.fech_frame = tk.Frame(self, bg="#F5F5F5")
        self.fech_frame.pack(pady=15)
        self.fech_label = tk.Label(self.fech_frame, text="Fecha de inspección final:", font=("Helvetica", 14), bg="#F5F5F5", fg="#333333")
        self.fech_label.pack(side=tk.LEFT, padx=15)
        self.fech_entry = tk.Entry(self.fech_frame, font=("Helvetica", 14))
        self.fech_entry.pack(side=tk.RIGHT, padx=15, expand=True, fill=tk.X)

        # Modificar button bg="#3986F3"
        self.fill_button = tk.Button(text="Generar", command=self.fill_template, font=("Helvetica", 16), bg="#2F4F4F", fg="white",
                               padx=50,
                               pady=13)
        self.fill_button.pack()

    def fill_template(self):
###################### idea para sacar todos los datos necesarios del excel ###################################################################
        username = self.get_username() 
        codigo = self.codigo_entry.get()
        estructura = self.estr_entry.get()
        lciudad = codigo[-1]  # Extract the city code from the work code
        jefe = self.combobox_autor.get()
        nota = self.combobox_revisor.get()
        fecha2 = self.fech_entry.get()
        segundorevisor = self.combobox2_autor.get()
        
        if self.not_entry.get() != "":
            nota_ad = self.not_entry.get()
            m = "Note:"
        if self.not_entry.get() == "":
            nota_ad = ""
            m = ""
    
        if nota == "Sí":
            nc = "Note de calcul,"
        if nota == "No":
            nc= ""

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
        

        workbook = openpyxl.load_workbook(f'C:\\Users\\{username}\\Incye\\France - Projets\\{codigo}\\01 Info\\{codigo}.xlsm')

        # Select the "DATOS" sheet
        worksheet = workbook['DONÉES']

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
        "EzequielSanchezdelaG": "Ezequiel Sánchez De La Guía",
        "AndresRodriguezPerez": "Andrés Rodríguez Pérez",
        "AlbertoAldamaMartine": "Alberto Aldama Martínez",
        "AdelaidaSaez": "Adelaida Sáez Castejón",
        "AlejandroBuiles": "Alejandro Ángel Builes",
        "JuanJoseMoron": "Juan José Morón Blanco",
        "": "Manuel González-Arquiso Madrigal",
        "RafaelMansilla": "Rafael Mansilla",
        "EstebanLopezFernande": "Esteban López Fernández", 
        "PabloTinajas": "David Lara",
        "FilippoBrusca": "Filippo Brusca"
        } 

        additional_info2 = {
        "JoseManuelMaldonadoM": "Máster Ingeniero de Caminos, CC. y PP.",
        "DavidLara": "Máster Ingeniero de Caminos, CC. y PP.",
        "EzequielSanchezdelaG": "Ingeniero Industrial.",
        "AndresRodriguezPerez": "Ingeniero Téc. Industrial",
        "AlbertoAldamaMartine": "Ingeniero Industrial.",
        "AdelaidaSáez": "Ing Téc. Industrial",
        "AlejandroBuiles": "Ingeniero Civil.",
        "JuanJoseMoron": "Delineante.",
        "": "Ing. Téc. Agrícola",
        "RafaelMansilla": "Máster Ingeniero de Caminos, CC. y PP.", 
        "EstebanLopezFernande": "Máster Ingeniero de Caminos, CC. y PP.",
        "PabloTinajas": "Máster Ingeniero de Caminos, CC. y PP.",
        "FilippoBrusca": "Máster Ingeniero de Caminos, CC. y PP."
        }

        additional_info8 = {
        "AVS": "antonio.vazquez@incye.com",
        "CENT": "madrid@incye.com",
        "BCN": "barcelona@incye.com",
        "BLB": "bilbao@incye.com",
        "ASP": "ana.seoane@incye.com", 
        "CRN": "ana.seoane@incye.com",
        "SPG": "galicia@incye.com",  
        "EXP": "antonio.vazquez@incye.com",
        "FR": "xavier.marty@incye.com",
        "LEV": "valencia@incye.com",
        "PT": "",
        "AND": "malaga@incye.com"
        }

        additional_info12 = {
        "Antonio Vázquez": ".",
        "Ignacio Merlín": ".",
        "David Suárez": ".",
        "Ibai Marlasca": ".",
        "Luis José Iglesias": ".", 
        "Ana Seoane": ".",  
        "Iván Valero": ".",
        "Álvaro Milla": ".",
        "Ismael Pérez": ".",
        "Javier Álvarez": ".",
        "Santiago Venencia": ", en qualité de Chef de Production de INCYE.", 
        "Álvaro Abad": ", en qualité de Chef de Production de INCYE."
        }  

        email_delegado = additional_info8.get(delegado, "")
        tecnico = additional_info.get(username, "")
        titulacion = additional_info2.get(username, "")

        # Fechas
        current_date = datetime.datetime.now()
        formatted_date = current_date.strftime("%d/%m/%Y")
        custom_date_format = current_date.strftime("%y%m%d")
        coletillaRV2 = additional_info12.get(segundorevisor, "")

        # Obtener los valores de las checkboxes
        checkbox_values = list(self.checkbar.state())
        if checkbox_values[0]:
            template = f"C:\\Users\\{username}\\Incye\\Ingenieria - Documentos\\12_Aplicaciones\\acta entrega\\ACTA ENTREGA ACODALAMIENTOSfr.docx"
        if checkbox_values[1]:
            template = f"C:\\Users\\{username}\\Incye\\Ingenieria - Documentos\\12_Aplicaciones\\acta entrega\\ACTA ENTREGA APEO-APUNTALAMIENTOfr.docx"
        if checkbox_values[2]:
            template = f"C:\\Users\\{username}\\Incye\\Ingenieria - Documentos\\12_Aplicaciones\\acta entrega\\ACTA ENTREGA ESTABILIZADORESfr.docx"

        document = MailMerge(template)

        # Sustituimos valores
        document.merge(Nombre_Obra=nombre_obra, Estructura = estructura, Codigo = codigo, Fecha=formatted_date, Nombre_Cliente=nombre_cliente, NC=nc, Fecha2=fecha2, Tecnico=tecnico, Jefe=jefe, Rev2=segundorevisor, coletilla=coletillaRV2, Titulacion=titulacion, M=m, Nota_Ad = nota_ad)
        
        output_path = f'C:\\Users\\{username}\\Incye\\France - Projets\\{codigo}\\07 Production\\2_Acta_Entrega\\{codigo}_ActaDeEntrega_{custom_date_format}.docx'
        document.write(output_path)
        pdf_path = output_path.replace(".docx", ".pdf")
        docx2pdf.convert(output_path, pdf_path)

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.To = email_cliente
        mail.Subject = f"{codigo} - {nombre_obra} -- ACTA DE ENTREGA INCYE -- {formatted_date}" 
        mail.CC = email_delegado
        mail.Body = f"Cher client. \n\nAprès l'achèvement de l'installation, veuillez trouver ci-joint le rapport de livraison de l'installation. \n\nMeilleures salutations. \n\n{tecnico} - Dpto. de Ingeniería INCYE "
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


