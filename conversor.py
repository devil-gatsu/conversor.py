import os
import glob
import customtkinter as ctk
from tkinter import filedialog, messagebox
import win32com.client as win32
import pythoncom
import threading

# Configuração visual do CustomTkinter
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class ConversorMapfre(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Conversor de Planilhas - Mapfre")
        self.geometry("500x320")
        self.resizable(False, False)

        self.pasta_origem = ""
        self.pasta_destino = ""

        # Interface
        self.lbl_titulo = ctk.CTkLabel(self, text="Conversor XLSM para XLS (Formato 56)", font=("Roboto", 16, "bold"))
        self.lbl_titulo.pack(pady=15)

        self.btn_origem = ctk.CTkButton(self, text="1. Selecionar Pasta de Origem (.xlsm)", command=self.selecionar_origem)
        self.btn_origem.pack(pady=10)

        self.btn_destino = ctk.CTkButton(self, text="2. Selecionar Pasta de Destino", command=self.selecionar_destino)
        self.btn_destino.pack(pady=10)

        self.lbl_status = ctk.CTkLabel(self, text="Aguardando seleção de pastas...", text_color="gray")
        self.lbl_status.pack(pady=10)

        self.btn_converter = ctk.CTkButton(self, text="3. Iniciar Conversão", command=self.iniciar_thread_conversao, fg_color="green", hover_color="darkgreen")
        self.btn_converter.pack(pady=15)

    def selecionar_origem(self):
        self.pasta_origem = filedialog.askdirectory(title="Selecione a pasta com os arquivos XLSM")
        if self.pasta_origem:
            self.lbl_status.configure(text=f"Origem: {os.path.basename(self.pasta_origem)}", text_color="white")

    def selecionar_destino(self):
        self.pasta_destino = filedialog.askdirectory(title="Selecione a pasta onde os arquivos XLS serão salvos")
        if self.pasta_destino:
            self.lbl_status.configure(text=f"Destino: {os.path.basename(self.pasta_destino)}", text_color="white")

    def iniciar_thread_conversao(self):
        if not self.pasta_origem or not self.pasta_destino:
            messagebox.showwarning("Aviso", "Selecione as pastas de origem e destino antes de continuar.")
            return
        
        self.btn_converter.configure(state="disabled", text="Convertendo...")
        # Executa em thread separada para não congelar a interface
        threading.Thread(target=self.processar_conversao, daemon=True).start()

    def processar_conversao(self):
        # pythoncom.CoInitialize() é obrigatório ao usar win32com dentro de threads
        pythoncom.CoInitialize()
        
        arquivos_xlsm = glob.glob(os.path.join(self.pasta_origem, "*.xlsm"))
        total = len(arquivos_xlsm)
        
        if total == 0:
            self.atualizar_interface("Nenhum arquivo .xlsm encontrado na origem.", "red", normalizar_botao=True)
            return

        try:
            excel = win32.DispatchEx('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            
            sucesso = 0
            
            for index, caminho_xlsm in enumerate(arquivos_xlsm, 1):
                nome_arquivo = os.path.basename(caminho_xlsm)
                novo_nome = nome_arquivo.replace(".xlsm", ".xls")
                caminho_xls = os.path.join(self.pasta_destino, novo_nome)
                
                self.atualizar_interface(f"Processando ({index}/{total}): {novo_nome}", "yellow")
                
                wb = excel.Workbooks.Open(os.path.abspath(caminho_xlsm))
                wb.SaveAs(os.path.abspath(caminho_xls), FileFormat=56) # 56 = xlExcel8
                wb.Close(SaveChanges=False)
                
                sucesso += 1
                
        except Exception as e:
            self.atualizar_interface(f"Erro inesperado: {str(e)}", "red", normalizar_botao=True)
        finally:
            excel.Quit()
            pythoncom.CoUninitialize()
            self.atualizar_interface(f"Concluído! {sucesso} de {total} arquivos convertidos.", "green", normalizar_botao=True)

    def atualizar_interface(self, mensagem, cor, normalizar_botao=False):
        self.lbl_status.configure(text=mensagem, text_color=cor)
        if normalizar_botao:
            self.btn_converter.configure(state="normal", text="3. Iniciar Conversão")

if __name__ == "__main__":
    app = ConversorMapfre()
    app.mainloop()
