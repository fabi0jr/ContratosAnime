import os
import sys
import customtkinter as ctk
from customtkinter import CTkImage
from tkinter import messagebox, Canvas
from services.google_sheets_service import save_data_to_sheet
from utils.docx_generator import generate_contract
from num2words import num2words
from datetime import datetime
from babel.dates import format_date
from PIL import Image, ImageTk

def resource_path(relative_path):
    """Obter o caminho absoluto para o recurso, funciona para dev e PyInstaller"""
    try:
        # PyInstaller cria uma pasta temporária e armazena o caminho nela
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Gerador de Contratos Automáticos")
        self.geometry("1200x700")

        # Criação dos campos de entrada
        self.create_widgets()

    def create_widgets(self):

        # Canvas para a imagem de fundo
        self.canvas = Canvas(self, width=1495, height=200, highlightthickness=0)
        self.canvas.place(x=0, y=0)
        # Carregar a imagem de fundo
        img_path = resource_path(os.path.join("assets", "main_img.png"))
        self.bg_image = Image.open(img_path)
        self.bg_image = self.bg_image.resize((1495, 200), Image.LANCZOS)
        self.bg_photo = ImageTk.PhotoImage(self.bg_image)
        self.canvas.create_image(0, 0, anchor="nw", image=self.bg_photo)

        colx1 = 300
        colx2 = 700
        widthEntry = 200

        # Criação dos campos de entrada usando a função
        self.entry_nome = self.create_label_entry("Nome", colx1, 210, widthEntry)

        self.entry_cpf = self.create_label_entry("CPF", colx1, 270, widthEntry)
        self.entry_cpf.bind('<KeyRelease>', self.format_cpf)

        self.entry_celular = self.create_label_entry("Celular", colx1, 330, widthEntry)
        self.entry_celular.bind('<KeyRelease>', self.format_cllr)

        self.entry_tipo_de_evento = self.create_label_entry("Tipo de Evento", colx1, 390, widthEntry)

        self.entry_data_do_evento = self.create_label_entry("Data do Evento", colx1, 450, widthEntry)
        self.entry_data_do_evento.bind('<KeyRelease>', self.format_data)

        self.entry_hora_do_evento = self.create_label_entry("Hora do Evento", colx1, 510, widthEntry)

        self.entry_local_do_evento = self.create_label_entry("Local do Evento", colx2, 210, widthEntry)

        self.entry_nome_do_aniversariante = self.create_label_entry("Nome do Aniversariante", colx2, 270, widthEntry)

        self.entry_tema_da_festa = self.create_label_entry("Tema da Festa", colx2, 330, widthEntry)

        self.entry_serviço_contratado = self.create_label_entry("Serviço Contratado", colx2, 390, widthEntry)
        
        self.entry_valor_do_contrato = self.create_label_entry("Valor do Contrato", colx2, 450, widthEntry)

        # Configurar o comportamento de foco
        self.configure_focus_order([
            self.entry_nome,
            self.entry_cpf,
            self.entry_celular,
            self.entry_tipo_de_evento,
            self.entry_data_do_evento,
            self.entry_hora_do_evento,
            self.entry_local_do_evento,
            self.entry_nome_do_aniversariante,
            self.entry_tema_da_festa,
            self.entry_serviço_contratado,
            self.entry_valor_do_contrato
        ])

        # Botão para gerar o contrato
        self.btn_gerar = ctk.CTkButton(self, text="Gerar Contrato", command=self.gerar_contrato)
        self.btn_gerar.place(x=550, y=600)

    def format_data(self, event):
        data = self.entry_data_do_evento.get()
        data = data.replace("/", "")
        if len(data) > 2:
            data = f"{data[:2]}/{data[2:4]}/{data[4:8]}"
        self.entry_data_do_evento.delete(0, "end")
        self.entry_data_do_evento.insert(0, data)

    def configure_focus_order(self, entries):
        for i in range(len(entries) - 1):
            entries[i].bind('<Return>', lambda event, next_entry=entries[i+1]: next_entry.focus())

    def format_cllr(self, event):
        cllr = self.entry_celular.get()
        cllr = cllr.replace("(", "").replace(")", "").replace(" ", "").replace("-", "")
        if len(cllr) > 0:
            cllr = f"({cllr[:2]}) {cllr[2:]}"
        if len(cllr) > 10:
            cllr = f"{cllr[:10]}-{cllr[10:]}"
        self.entry_celular.delete(0, "end")
        self.entry_celular.insert(0, cllr)

    def format_cpf(self, event):
        cpf = self.entry_cpf.get()
        cpf = cpf.replace(".", "").replace("-", "")
        if len(cpf) > 3:
            cpf = f"{cpf[:3]}.{cpf[3:]}"
        if len(cpf) > 7:
            cpf = f"{cpf[:7]}.{cpf[7:]}"
        if len(cpf) > 11:
            cpf = f"{cpf[:11]}-{cpf[11:]}"
        self.entry_cpf.delete(0, "end")
        self.entry_cpf.insert(0, cpf)

    def create_label_entry(self, text, x, y, width):
        ctk.CTkLabel(self, text=text).place(x=x, y=y)
        entry = ctk.CTkEntry(self, width=width)
        entry.place(x=x, y=y+30)
        return entry

    def gerar_contrato(self):
        try:
            dados_contrato = {
                "contratante": self.entry_nome.get(),
                "CPF": self.entry_cpf.get(),
                "fone": self.entry_celular.get(),
                "tipo_evento": self.entry_tipo_de_evento.get(),
                "data": self.entry_data_do_evento.get(),
                "hora": self.entry_hora_do_evento.get(),
                "local": self.entry_local_do_evento.get(),
                "aniversariante": self.entry_nome_do_aniversariante.get(),
                "tema": self.entry_tema_da_festa.get(),
                "servico": self.entry_serviço_contratado.get(),
                "valor": self.entry_valor_do_contrato.get()
            }

            valor_str = num2words(dados_contrato["valor"], lang='pt_BR').title()
            dados_contrato.update({"valor_str": valor_str})

            data_atual = datetime.now()
            dia = str(data_atual.day)
            mes = format_date(data_atual, 'MMMM', locale='pt_BR').title()
            ano = str(data_atual.year)
            dados_contrato.update({"dia": dia, "mes": mes, "ano": ano})

            contrato_path = generate_contract(dados_contrato)
            messagebox.showinfo("Sucesso", f"Contrato gerado com sucesso: {contrato_path}")

            spreadsheet_id = '1hrONt6OArCC_niXzzh3AA5m18TNdvSC-AS8e73wk4DU'
            range_name = 'Sheet1!A1'
            values = [[dados_contrato[key] for key in dados_contrato]]
            save_data_to_sheet(spreadsheet_id, range_name, values)
            messagebox.showinfo("Sucesso", "Dados salvos na planilha do Google com sucesso!")
            # Limpar todos os campos de entrada
            for entry in [
                self.entry_nome,
                self.entry_cpf,
                self.entry_celular,
                self.entry_tipo_de_evento,
                self.entry_data_do_evento,
                self.entry_hora_do_evento,
                self.entry_local_do_evento,
                self.entry_nome_do_aniversariante,
                self.entry_tema_da_festa,
                self.entry_serviço_contratado,
                self.entry_valor_do_contrato
            ]:
                entry.delete(0, "end")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

if __name__ == "__main__":
    app = App()
    app.mainloop()