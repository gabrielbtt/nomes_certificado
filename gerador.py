from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import yagmail
import configparser
import os

class EditCertificate:
    '''classe que cria certificado dinamicamente e envia por email'''
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

            # Gerar um certificado com o nome
            certificate_image = self.create_certificate(name, template_path)

            # Enviar email personalizado
            self.send_email_generic(name, email, certificate_image, sender_email, sender_password, subject, content)
            
    def create_certificate(self, name, template_path):
        # Abrir a imagem de modelo
        imagem = Image.open(template_path)
        
        draw = ImageDraw.Draw(imagem)

        font = ImageFont.truetype('calibrib.ttf', 55)

        # Definindo a cor da fonte como preto (0, 0, 0)
        draw.text((125, 770), name, font=font, fill=(0, 0, 0))

        path = ""
        certificate_image = f'{path}{name}_Certificate.pdf'
        imagem.save(certificate_image)
        return certificate_image
    
    def send_email_generic(self, name, email, certificate_image, sender_email, sender_password, subject, content):
        # Inicializar o Yagmail SMTP
        usuario = yagmail.SMTP(user=sender_email, password=sender_password)
        
        # Personalizar o conteúdo do email
        personalized_content = content.replace("{name}", name)

        # Anexar a imagem do certificado
        usuario.send(
            to=email,
            subject=subject,
            contents=personalized_content,
            attachments=certificate_image
        )
        
        print(f'Email enviado para {email} com sucesso!')
    
    def create_gui(self):
        self.window = Tk()
        self.window.title("Gerador de Certificados")
        self.window.geometry('850x400')
        self.window.configure(bg='#2e2e2e')

        style = ttk.Style(self.window)
        style.theme_use('clam')

        style.configure('TButton', font=('Helvetica', 12), padding=6, background='#444444', foreground='#ffffff')
        style.map('TButton', background=[('active', '#555555')])
        style.configure('TLabel', font=('Helvetica', 12), background='#2e2e2e', foreground='#ffffff')
        style.configure('TEntry', font=('Helvetica', 12), fieldbackground='#3e3e3e', foreground='#ffffff')
        style.configure('TFrame', background='#2e2e2e')
        style.configure('TText', font=('Helvetica', 12), fieldbackground='#3e3e3e', foreground='#ffffff')

        main_frame = ttk.Frame(self.window, padding="10 10 10 10", style='TFrame')
        main_frame.grid(column=0, row=0, sticky=(N, W, E, S))

        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)

        ttk.Label(main_frame, text="Email:").grid(row=0, column=0, sticky=W)
        self.email_entry = ttk.Entry(main_frame, width=30)
        self.email_entry.insert(0, self.config.get('credentials', 'email'))
        self.email_entry.grid(row=0, column=1, columnspan=3, sticky=(W, E))

        ttk.Label(main_frame, text="Senha:").grid(row=1, column=0, sticky=W)
        self.password_entry = ttk.Entry(main_frame, width=30, show="*")
        self.password_entry.insert(0, self.config.get('credentials', 'password'))
        self.password_entry.grid(row=1, column=1, columnspan=3, sticky=(W, E))

        ttk.Label(main_frame, text="Assunto:").grid(row=2, column=0, sticky=W)
        self.subject_entry = ttk.Entry(main_frame, width=30)
        self.subject_entry.insert(0, 'Texto a ser inserido no campo Assunto')
        self.subject_entry.grid(row=2, column=1, columnspan=3, sticky=(W, E))

        ttk.Label(main_frame, text="Conteúdo:").grid(row=3, column=0, sticky=W)
        self.content_text = Text(main_frame, width=40, height=10)
        self.content_text.insert(1.0, 'Texto a ser inserido no campo conteudo (Utilize name entre chaves para escrever o nome da pessoa no campo)')
        self.content_text.grid(row=3, column=1, columnspan=3, sticky=(W, E))

        button_width = 20  # Largura fixa dos botões

        self.btn_select_data = ttk.Button(main_frame, text="Selecionar Dados", command=self.select_data)
        self.btn_select_data.grid(row=4, column=0, sticky=(W, E), ipadx=button_width)

        self.btn_select_template = ttk.Button(main_frame, text="Selecionar Template", command=self.select_template)
        self.btn_select_template.grid(row=4, column=1, sticky=(W, E), ipadx=button_width)

        self.btn_generate_certificates = ttk.Button(main_frame, text="Gerar Certificados", state=DISABLED)
        self.btn_generate_certificates.grid(row=4, column=2, sticky=(W, E), ipadx=button_width)

        self.btn_update_credentials = ttk.Button(main_frame, text="Atualizar Credenciais", command=self.update_credentials)
        self.btn_update_credentials.grid(row=0, column=4, rowspan=1, sticky=(N, S, W, E), ipadx=button_width)

        for child in main_frame.winfo_children(): 
            child.grid_configure(padx=5, pady=5)

        self.window.mainloop()

    def select_data(self):
        data_path = askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if data_path:
            self.data_path = data_path

    def select_template(self):
        template_path = askopenfilename(filetypes=[("Image files", "*.png")])
        if template_path:
            self.template_path = template_path
            self.btn_generate_certificates.config(state=NORMAL, command=lambda: self.send_email_with_certificate(self.data_path, self.template_path, self.email_entry.get(), self.password_entry.get(), self.subject_entry.get(), self.content_text.get("1.0", END)))

    def update_credentials(self):
        self.config.set('credentials', 'email', self.email_entry.get())
        self.config.set('credentials', 'password', self.password_entry.get())
        with open(self.config_path, 'w') as configfile:
            self.config.write(configfile)
        print("Credenciais atualizadas com sucesso!")

start = EditCertificate()