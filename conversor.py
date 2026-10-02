import os
import glob
import subprocess
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading

# Importação condicional para evitar quebra se a lib não estiver instalada perfeitamente
try:
    import win32com.client as win32
    import pythoncom
    COM_DISPONIVEL = True
except ImportError:
    COM_DISPONIVEL = False

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class ConversorMapfre(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Conversor de Planilhas - Mapfre")
        self.geometry("500x420")
        self.resizable(False, False)

        self.pasta_origem = ""
        self.pasta_destino = ""
        self.motor_selecionado = ctk.StringVar(value="libreoffice") # Padrão definido para LibreOffice

        self.lbl_titulo = ctk.CTkLabel(self, text="Conversor XLSM para XLS", font=("Roboto", 16, "bold"))
        self.lbl_titulo.pack(pady=15)

        # Seleção do Motor
        self.frame_motor = ctk.CTkFrame(self)
        self.frame_motor.pack(pady=5, padx=20, fill="x")
        
        self.lbl_motor = ctk.CTkLabel(self.frame_motor, text="Selecione o programa instalado na máquina:", font=("Roboto", 12))
        self.lbl_motor.pack(pady=(5, 0))
        
        self.radio_libre = ctk.CTkRadioButton(self.frame_motor, text="LibreOffice Calc", variable=self.motor_selecionado, value="libreoffice")
        self.radio_libre.pack(side="left", padx=20, pady=10)
        
        self.radio_excel = ctk.CTkRadioButton(self.frame_motor, text="Microsoft Excel", variable=self.motor_selecionado, value="excel")
        self.radio_excel.pack(side="right", padx=20, pady=10)

        # Botões de Diretório
        self.btn_origem = ctk.CTkButton(self, text="1. Selecionar Pasta Origem (.xlsm)", command=self.selecionar_origem)
        self.btn_origem.pack(pady=10)

        self.btn_destino = ctk.CTkButton(self, text="2. Selecionar Pasta Destino", command=self.selecionar_destino)
        self.btn_destino.pack(pady=10)

        self.lbl_status = ctk.CTkLabel(self, text="Aguardando seleção de pastas...", text_color="gray")
        self.lbl_status.pack(pady=10)

        self.btn_converter = ctk.CTkButton(self, text="3. Iniciar Conversão", command=self.iniciar_thread_conversao, fg_color="green", hover_color="darkgreen")
        self.btn_converter.pack(pady=15)

    def selecionar_origem(self):
        self.pasta_origem = filedialog.askdirectory(title="Selecione a pasta Origem")
        if self.pasta_origem:
            self.atualizar_status(f"Origem: {os.path.basename(self.pasta_origem)}", "white")

    def selecionar_destino(self):
        self.pasta_destino = filedialog.askdirectory(title="Selecione a pasta Destino")
        if self.pasta_destino:
            self.atualizar_status(f"Destino: {os.path.basename(self.pasta_destino)}", "white")

    def iniciar_thread_conversao(self):
        if not self.pasta_origem or not self.pasta_destino:
            messagebox.showwarning("Aviso", "Selecione as pastas de origem e destino.")
            return
        
        self.btn_converter.configure(state="disabled", text="Convertendo...")
        threading.Thread(target=self.processar_conversao, daemon=True).start()

    def buscar_libreoffice(self):
        # Mapeia caminhos padrão de instalação do LibreOffice no Windows
        caminhos = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
        ]
        for caminho in caminhos:
            if os.path.exists(caminho):
                return caminho
        return None

    def processar_conversao(self):
        arquivos_xlsm = glob.glob(os.path.join(self.pasta_origem, "*.xlsm"))
        total = len(arquivos_xlsm)
        
        if total == 0:
            self.atualizar_status("Nenhum arquivo .xlsm encontrado.", "red", normalizar=True)
            return

        motor = self.motor_selecionado.get()
        sucesso = 0

        if motor == "excel":
            if not COM_DISPONIVEL:
                self.atualizar_status("Erro: Bibliotecas win32com ausentes.", "red", normalizar=True)
                return
            
            try:
                pythoncom.CoInitialize()
                excel = win32.DispatchEx('Excel.Application')
                excel.Visible = False
                excel.DisplayAlerts = False
                
                for index, caminho_xlsm in enumerate(arquivos_xlsm, 1):
                    nome = os.path.basename(caminho_xlsm)
                    self.atualizar_status(f"Excel processando ({index}/{total}): {nome}", "yellow")
                    
                    caminho_xls = os.path.join(self.pasta_destino, nome.replace(".xlsm", ".xls"))
                    wb = excel.Workbooks.Open(os.path.abspath(caminho_xlsm))
                    wb.SaveAs(os.path.abspath(caminho_xls), FileFormat=56)
                    wb.Close(SaveChanges=False)
                    sucesso += 1
            except Exception as e:
                self.atualizar_status(f"Falha COM: {str(e)}", "red", normalizar=True)
                return
            finally:
                excel.Quit()
                pythoncom.CoUninitialize()

        elif motor == "libreoffice":
            caminho_soffice = self.buscar_libreoffice()
            if not caminho_soffice:
                self.atualizar_status("Erro: LibreOffice (soffice.exe) não encontrado.", "red", normalizar=True)
                return
            
            for index, caminho_xlsm in enumerate(arquivos_xlsm, 1):
                nome = os.path.basename(caminho_xlsm)
                self.atualizar_status(f"LibreOffice processando ({index}/{total}): {nome}", "yellow")
                
                # Flag de criação para não piscar tela de terminal no Windows durante a conversão
                creation_flags = 0
                if os.name == 'nt':
                    creation_flags = subprocess.CREATE_NO_WINDOW
                
                comando = [
                    caminho_soffice,
                    "--headless",
                    "--convert-to", "xls",
                    "--outdir", self.pasta_destino,
                    caminho_xlsm
                ]
                
                try:
                    subprocess.run(comando, check=True, creationflags=creation_flags)
                    sucesso += 1
                except subprocess.CalledProcessError as e:
                    print(f"Erro ao converter {nome}: {e}")

        self.atualizar_status(f"Concluído! {sucesso} de {total} convertidos via {motor.upper()}.", "green", normalizar=True)

    def atualizar_status(self, mensagem, cor, normalizar=False):
        self.lbl_status.configure(text=mensagem, text_color=cor)
        if normalizar:
            self.btn_converter.configure(state="normal", text="3. Iniciar Conversão")

if __name__ == "__main__":
    app = ConversorMapfre()
    app.mainloop()
