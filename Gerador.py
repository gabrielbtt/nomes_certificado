import os
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename
from PIL import Image, ImageDraw, ImageFont, ImageTk
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
import pandas as pd
import yagmail
import configparser

class EditCertificate:
    '''Classe que cria certificado dinamicamente e envia por email'''
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config_path = 'config.ini'
        self.check_and_create_config()
        self.create_gui()

    def check_and_create_config(self):
        if not os.path.exists(self.config_path):
            self.config['credentials'] = {
                'email': 'seu_email@gmail.com',
                'password': 'sua_senha'
            }
            with open(self.config_path, 'w') as configfile:
                self.config.write(configfile)
        else:
            self.config.read(self.config_path)

    def read_data(self, data_path):
        self.df = pd.read_excel(data_path)
        
    def send_email_with_certificate(self, data_path, template_path, sender_email, sender_password, subject, content):
        self.read_data(data_path)
        for _, row in self.df.iterrows():
            name = row['Nome']
            email = row['Email']
            certificate_number = row['Numero do Certificado']
            certificate_image = self.create_certificate(name, certificate_number, template_path)
            self.send_email_generic(name, email, certificate_image, sender_email, sender_password, subject, content)
            
    def get_font_path(self, font_name):
        fonts_dir = "C:\\Windows\\Fonts"
        for root, dirs, files in os.walk(fonts_dir):
            for file in files:
                if font_name.lower() in file.lower():
                    return os.path.join(root, file)
        raise FileNotFoundError(f"Fonte '{font_name}' não encontrada no diretório de fontes do Windows.")

    def create_certificate(self, name, certificate_number, template_path):
        template_extension = os.path.splitext(template_path)[1].lower()
        # Supondo que certificate_number venha da planilha
        certificate_number = str(certificate_number)  # Certifica-se de que é uma string
        num_digits = len(certificate_number)  # Conta quantos dígitos já existem
        certificate_number_padded = certificate_number.zfill(num_digits)  # Preenche com zeros à esquerda

        # Para garantir que o número tenha pelo menos 4 dígitos
        certificate_number_padded = certificate_number.zfill(max(4, num_digits))

        if template_extension == '.pdf':
            images = convert_from_path(template_path)
            writer = PdfWriter()

            for image in images:
                draw = ImageDraw.Draw(image)

                # Font and position for name
                font_name = self.font_combobox.get()
                font_size = int(self.font_size_entry.get())
                font_path = self.get_font_path(font_name)
                font = ImageFont.truetype(font_path, font_size)

                x = int(self.x_entry.get())
                y = int(self.y_entry.get())
                draw.text((x, y), name, font=font, fill=(0, 0, 0))

                # Font and position for certificate number
                font_cert_size = int(self.font_cert_size_entry.get())
                print(f"Tamanho da fonte do número do certificado: {font_cert_size}")  # Verificação do valor da fonte
                font_cert = ImageFont.truetype(font_path, font_cert_size)
                
                cert_x = int(self.cert_x_entry.get())
                cert_y = int(self.cert_y_entry.get())
                draw.text((cert_x, cert_y), certificate_number_padded, font=font_cert, fill=(0, 0, 0))

                temp_path = f"{name}_temp.png"
                image.save(temp_path)
                reader_temp = PdfReader(temp_path)
                page = reader_temp.pages[0]
                writer.add_page(page)

                os.remove(temp_path)

            certificate_image = f"{self.output_name_entry.get()}_{name}.pdf"
            with open(certificate_image, "wb") as output_pdf:
                writer.write(output_pdf)
        else:
            imagem = Image.open(template_path)
            draw = ImageDraw.Draw(imagem)

            # Font and position for name
            font_name = self.font_combobox.get()
            font_size = int(self.font_size_entry.get())
            font_path = self.get_font_path(font_name)
            font = ImageFont.truetype(font_path, font_size)

            x = int(self.x_entry.get())
            y = int(self.y_entry.get())
            draw.text((x, y), name, font=font, fill=(0, 0, 0))

            # Font and position for certificate number
            font_cert_size = int(self.font_cert_size_entry.get())
            print(f"Tamanho da fonte do número do certificado: {font_cert_size}")  # Verificação do valor da fonte
            font_cert = ImageFont.truetype(font_path, font_cert_size)
                
            cert_x = int(self.cert_x_entry.get())
            cert_y = int(self.cert_y_entry.get())
            draw.text((cert_x, cert_y), certificate_number_padded, font=font_cert, fill=(0, 0, 0))

            certificate_image = f"{self.output_name_entry.get()}_{name}.pdf"
            imagem.save(certificate_image)
        
        return certificate_image
    
    def send_email_generic(self, name, email, certificate_image, sender_email, sender_password, subject, content):
        usuario = yagmail.SMTP(user=sender_email, password=sender_password)
        
        personalized_content = content.replace("{name}", name)

        usuario.send(
            to=email,
            subject=subject,
            contents=personalized_content,
            attachments=certificate_image
        )
        
        print(f'Email enviado para {email} com sucesso!')
    
    def preview_certificate(self):
        name = "Pré-visualização"
        certificate_number = "0000"  # Preview example for certificate number
        template_extension = os.path.splitext(self.template_path)[1].lower()
        # Supondo que certificate_number venha da planilha
        certificate_number = str(certificate_number)  # Certifica-se de que é uma string
        num_digits = len(certificate_number)  # Conta quantos dígitos já existem
        certificate_number_padded = certificate_number.zfill(num_digits)  # Preenche com zeros à esquerda

        # Para garantir que o número tenha pelo menos 4 dígitos
        certificate_number_padded = certificate_number.zfill(max(4, num_digits))

        if template_extension == '.pdf':
            images = convert_from_path(self.template_path)
            preview_image = images[0]

            draw = ImageDraw.Draw(preview_image)

            # Font and position for name
            font_name = self.font_combobox.get()
            font_size = int(self.font_size_entry.get())
            font_path = self.get_font_path(font_name)
            font = ImageFont.truetype(font_path, font_size)

            x = int(self.x_entry.get())
            y = int(self.y_entry.get())
            draw.text((x, y), name, font=font, fill=(0, 0, 0))

            # Font and position for certificate number
            font_cert_size = int(self.font_cert_size_entry.get())
            print(f"Tamanho da fonte do número do certificado: {font_cert_size}")  # Verificação do valor da fonte
            font_cert = ImageFont.truetype(font_path, font_cert_size)
                
            cert_x = int(self.cert_x_entry.get())
            cert_y = int(self.cert_y_entry.get())
            draw.text((cert_x, cert_y), certificate_number_padded, font=font_cert, fill=(0, 0, 0))

        else:
            preview_image = Image.open(self.template_path)
            draw = ImageDraw.Draw(preview_image)

            # Font and position for name
            font_name = self.font_combobox.get()
            font_size = int(self.font_size_entry.get())
            font_path = self.get_font_path(font_name)
            font = ImageFont.truetype(font_path, font_size)

            x = int(self.x_entry.get())
            y = int(self.y_entry.get())
            draw.text((x, y), name, font=font, fill=(0, 0, 0))

            # Font and position for certificate number
            font_cert_size = int(self.font_cert_size_entry.get())
            print(f"Tamanho da fonte do número do certificado: {font_cert_size}")  # Verificação do valor da fonte
            font_cert = ImageFont.truetype(font_path, font_cert_size)
                
            cert_x = int(self.cert_x_entry.get())
            cert_y = int(self.cert_y_entry.get())
            draw.text((cert_x, cert_y), certificate_number_padded, font=font_cert, fill=(0, 0, 0))

        preview_window = Toplevel(self.window)
        preview_window.title("Pré-visualização do Certificado")
        
        preview_image.thumbnail((800, 600))
        img_tk = ImageTk.PhotoImage(preview_image)

        label = Label(preview_window, image=img_tk)
        label.image = img_tk
        label.pack()

    def load_font_families(self):
        fonts_dir = "C:\\Windows\\Fonts"
        self.font_families = {}
        for root, dirs, files in os.walk(fonts_dir):
            for file in files:
                if file.endswith('.ttf') or file.endswith('.TTF'):
                    font_name = os.path.splitext(file)[0]
                    self.font_families[font_name] = file
    
    def create_gui(self):
        self.window = Tk()
        self.window.title("Gerador de Certificados - Crea-Jr")
        self.window.geometry('1020x620')
        self.window.configure(bg='#1E3A5F')  # Azul do Crea-Jr
        self.load_font_families()

        style = ttk.Style(self.window)
        style.theme_use('clam')

        style.configure('TButton', font=('Helvetica', 12), padding=6, background='#00509E', foreground='#ffffff')
        style.map('TButton', background=[('active', '#003F7F')])
        style.configure('TLabel', font=('Helvetica', 12), background='#1E3A5F', foreground='#ffffff')
        style.configure('TEntry', font=('Helvetica', 12), fieldbackground='#3e3e3e', foreground='#ffffff')
        style.configure('TFrame', background='#1E3A5F')
        style.configure('TText', font=('Helvetica', 12), fieldbackground='#3e3e3e', foreground='#ffffff')

        main_frame = ttk.Frame(self.window, padding="10 10 10 10", style='TFrame')
        main_frame.grid(column=0, row=0, sticky=(N, W, E, S), padx=20, pady=20)

        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)

        ttk.Label(main_frame, text="Email:").grid(row=0, column=0, sticky=W, pady=5)
        self.email_entry = ttk.Entry(main_frame, width=30)
        self.email_entry.insert(0, self.config.get('credentials', 'email'))
        self.email_entry.grid(row=0, column=1, columnspan=4, sticky=(W, E), pady=5)

        ttk.Label(main_frame, text="Senha:").grid(row=1, column=0, sticky=W, pady=5)
        self.password_entry = ttk.Entry(main_frame, width=30, show="*")
        self.password_entry.insert(0, self.config.get('credentials', 'password'))
        self.password_entry.grid(row=1, column=1, columnspan=4, sticky=(W, E), pady=5)
        
        self.save_button = ttk.Button(main_frame, text="Atualizar", command=self.save_config)
        self.save_button.grid(row=0, column=5, sticky=W, padx=20, pady=10)

        ttk.Label(main_frame, text="Assunto:").grid(row=2, column=0, sticky=W, pady=5)
        self.subject_entry = ttk.Entry(main_frame, width=30)
        self.subject_entry.insert(0, 'Texto a ser inserido no campo Assunto')
        self.subject_entry.grid(row=2, column=1, columnspan=4, sticky=(W, E), pady=5)

        ttk.Label(main_frame, text="Conteúdo:").grid(row=3, column=0, sticky=W, pady=5)
        self.content_text = Text(main_frame, width=40, height=5)
        self.content_text.insert(1.0, 'Texto a ser inserido no campo conteudo (Utilize name entre chaves para escrever o nome da pessoa no campo)')
        self.content_text.grid(row=3, column=1, columnspan=4, sticky=(W, E), pady=5)

        ttk.Label(main_frame, text="X:").grid(row=4, column=0, sticky=W, pady=5)
        self.x_entry = ttk.Entry(main_frame, width=15)
        self.x_entry.insert(0, '200')
        self.x_entry.grid(row=4, column=1, sticky=W, pady=5)

        ttk.Label(main_frame, text="Y:").grid(row=4, column=2, sticky=W, pady=5)
        self.y_entry = ttk.Entry(main_frame, width=15)
        self.y_entry.insert(0, '1340')
        self.y_entry.grid(row=4, column=3, sticky=W, pady=5)
        
        ttk.Label(main_frame, text="Tamanho da Fonte:").grid(row=5, column=0, sticky=W, pady=5)
        self.font_size_entry = ttk.Entry(main_frame, width=15)
        self.font_size_entry.insert(0, '100')
        self.font_size_entry.grid(row=5, column=1, sticky=W, pady=5)
        
        ttk.Label(main_frame, text="X (Nº Certificado):").grid(row=6, column=0, sticky=W, pady=5)
        self.cert_x_entry = ttk.Entry(main_frame, width=15)
        self.cert_x_entry.insert(0, '570')
        self.cert_x_entry.grid(row=6, column=1, sticky=W, pady=5)

        ttk.Label(main_frame, text="Y (Nº Certificado):").grid(row=6, column=2, sticky=W, pady=5)
        self.cert_y_entry = ttk.Entry(main_frame, width=15)
        self.cert_y_entry.insert(0, '1930')
        self.cert_y_entry.grid(row=6, column=3, sticky=W, pady=5)

        ttk.Label(main_frame, text="Tamanho do Nº:").grid(row=7, column=0, sticky=W, pady=5)
        self.font_cert_size_entry = ttk.Entry(main_frame, width=15)
        self.font_cert_size_entry.insert(0, '60')
        self.font_cert_size_entry.grid(row=7, column=1, sticky=W, pady=5)

        ttk.Label(main_frame, text="Fonte:").grid(row=8, column=0, sticky=W, pady=5)
        self.font_combobox = ttk.Combobox(main_frame, values=list(self.font_families.keys()), width=27)
        self.font_combobox.set('Times')
        self.font_combobox.grid(row=8, column=1, columnspan=3, sticky=(W, E), pady=5)

        ttk.Label(main_frame, text="Modelo do Certificado:").grid(row=9, column=0, sticky=W, pady=5)
        self.template_button = ttk.Button(main_frame, text="Selecionar", command=self.select_template)
        self.template_button.grid(row=9, column=1, sticky=W, pady=5)

        ttk.Label(main_frame, text="Dados (Excel):").grid(row=9, column=3, sticky=W, pady=5)
        self.data_button = ttk.Button(main_frame, text="Selecionar", command=self.select_data_file)
        self.data_button.grid(row=9, column=4, sticky=W, pady=5)

        ttk.Label(main_frame, text="Nome do Certificado:").grid(row=10, column=0, sticky=W, pady=5)
        self.output_name_entry = ttk.Entry(main_frame, width=30)
        self.output_name_entry.insert(0, 'certificado')
        self.output_name_entry.grid(row=10, column=1, columnspan=3, sticky=(W, E), pady=5)

        self.preview_button = ttk.Button(main_frame, text="Pré-visualizar", command=self.preview_certificate)
        self.preview_button.grid(row=11, column=1, sticky=W, pady=5)

        self.send_button = ttk.Button(main_frame, text="Enviar", command=self.send_email)
        self.send_button.grid(row=11, column=4, sticky=W, pady=5)

        self.window.mainloop()

    def select_template(self):
        self.template_path = askopenfilename(title="Selecionar Modelo de Certificado", filetypes=[("Image Files", "*.png *.jpg *.jpeg"), ("PDF Files", "*.pdf")])

    def select_data_file(self):
        self.data_path = askopenfilename(title="Selecionar Arquivo de Dados", filetypes=[("Excel Files", "*.xlsx")])

    def save_config(self):
        self.config['credentials']['email'] = self.email_entry.get()
        self.config['credentials']['password'] = self.password_entry.get()

        with open(self.config_path, 'w') as configfile:
            self.config.write(configfile)

        print('Configurações salvas com sucesso!')

    def send_email(self):
        sender_email = self.email_entry.get()
        sender_password = self.password_entry.get()
        subject = self.subject_entry.get()
        content = self.content_text.get(1.0, END)

        self.send_email_with_certificate(
            self.data_path, 
            self.template_path, 
            sender_email, 
            sender_password, 
            subject, 
            content
        )

if __name__ == '__main__':
    EditCertificate()