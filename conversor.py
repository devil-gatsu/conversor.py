import os
import glob
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox

try:
    import win32com.client as win32
    import pythoncom
except ImportError:
    pass # Tratamento para compilação

# Configuração visual do CustomTkinter (Branco e Azul)
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

class AppTotumseg(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Totumseg - Jet x Mapfre xls")
        self.geometry("600x550")
        self.resizable(False, False)

        # Variáveis de diretório
        self.pasta_modelos = ""
        self.pasta_origem = ""
        self.pasta_destino = ""

        # ================= INTERFACE =================
        self.lbl_titulo = ctk.CTkLabel(self, text="Transposição de Dados - Mapfre", font=("Segoe UI", 20, "bold"), text_color="#0047AB")
        self.lbl_titulo.pack(pady=15)

        # Frame de Seleção de Pastas
        self.frame_pastas = ctk.CTkFrame(self, fg_color="white", corner_radius=10)
        self.frame_pastas.pack(pady=10, padx=20, fill="x")

        # Botão 1: Modelos
        self.btn_modelos = ctk.CTkButton(self.frame_pastas, text="1. Selecionar Pasta de MODELOS", command=self.selecionar_modelos, fg_color="#0066cc")
        self.btn_modelos.grid(row=0, column=0, padx=15, pady=10, sticky="w")
        self.lbl_modelos = ctk.CTkLabel(self.frame_pastas, text="Nenhuma pasta selecionada", text_color="gray")
        self.lbl_modelos.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        # Botão 2: Origem
        self.btn_origem = ctk.CTkButton(self.frame_pastas, text="2. Selecionar Pasta de ORIGEM (.xlsm)", command=self.selecionar_origem, fg_color="#0066cc")
        self.btn_origem.grid(row=1, column=0, padx=15, pady=10, sticky="w")
        self.lbl_origem = ctk.CTkLabel(self.frame_pastas, text="Nenhuma pasta selecionada", text_color="gray")
        self.lbl_origem.grid(row=1, column=1, padx=10, pady=10, sticky="w")

        # Botão 3: Destino
        self.btn_destino = ctk.CTkButton(self.frame_pastas, text="3. Selecionar Pasta de DESTINO", command=self.selecionar_destino, fg_color="#0066cc")
        self.btn_destino.grid(row=2, column=0, padx=15, pady=10, sticky="w")
        self.lbl_destino = ctk.CTkLabel(self.frame_pastas, text="Nenhuma pasta selecionada", text_color="gray")
        self.lbl_destino.grid(row=2, column=1, padx=10, pady=10, sticky="w")

        # Log e Progresso
        self.txt_log = ctk.CTkTextbox(self, width=560, height=150, fg_color="#f0f4f8", text_color="black")
        self.txt_log.pack(pady=10)
        self.txt_log.insert("0.0", "Aguardando inicialização...\n")
        self.txt_log.configure(state="disabled")

        self.barra_progresso = ctk.CTkProgressBar(self, width=560, progress_color="#0047AB")
        self.barra_progresso.pack(pady=5)
        self.barra_progresso.set(0)

        # Botão Processar
        self.btn_processar = ctk.CTkButton(self, text="INICIAR TRANSPOSIÇÃO", command=self.iniciar_processo, font=("Segoe UI", 14, "bold"), fg_color="#28a745", hover_color="#218838")
        self.btn_processar.pack(pady=15)

    # ================= FUNÇÕES DE SELEÇÃO =================
    def selecionar_modelos(self):
        self.pasta_modelos = filedialog.askdirectory(title="Selecione a pasta com os Modelos XLS")
        if self.pasta_modelos: self.lbl_modelos.configure(text=os.path.basename(self.pasta_modelos), text_color="black")

    def selecionar_origem(self):
        self.pasta_origem = filedialog.askdirectory(title="Selecione a pasta Origem")
        if self.pasta_origem: self.lbl_origem.configure(text=os.path.basename(self.pasta_origem), text_color="black")

    def selecionar_destino(self):
        self.pasta_destino = filedialog.askdirectory(title="Selecione a pasta Destino")
        if self.pasta_destino: self.lbl_destino.configure(text=os.path.basename(self.pasta_destino), text_color="black")

    def log(self, mensagem):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", mensagem + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    # ================= LÓGICA CORE =================
    def iniciar_processo(self):
        if not all([self.pasta_modelos, self.pasta_origem, self.pasta_destino]):
            messagebox.showwarning("Atenção", "Por favor, selecione as três pastas antes de iniciar.")
            return
        
        self.btn_processar.configure(state="disabled", text="PROCESSANDO...")
        threading.Thread(target=self.transpor_dados, daemon=True).start()

    def identificar_modelo(self, nome_arquivo):
        nome_upper = nome_arquivo.upper()
        ramo = ""
        valor = ""

        # Identificar Ramo
        if "MADEIRA" in nome_upper: ramo = "CASA_MADEIRA"
        elif "CASA" in nome_upper: ramo = "CASA"
        elif "APARTAMENTO" in nome_upper or "APTO" in nome_upper: ramo = "APTO"
        elif "COMERCIO" in nome_upper or "COMER" in nome_upper: ramo = "COMER"
        elif "ESCRITORIO" in nome_upper or "ESCRIT" in nome_upper: ramo = "ESCRIT"

        # Identificar Valor
        if "200000" in nome_upper: valor = "200"
        elif "300000" in nome_upper: valor = "300"
        elif "400000" in nome_upper: valor = "400"
        elif "600000" in nome_upper: valor = "600"
        
        # 3. Retornar a junção exata que deve bater com o nome do arquivo na Pasta de Modelos
        if ramo and valor:
            return f"{ramo}_{valor}.xls"
        return None

   def transpor_dados(self):
        pythoncom.CoInitialize()
        arquivos_origem = glob.glob(os.path.join(self.pasta_origem, "*.xlsm"))
        total = len(arquivos_origem)

        if total == 0:
            self.log("Nenhum arquivo .xlsm encontrado na pasta de origem.")
            self.btn_processar.configure(state="normal", text="INICIAR TRANSPOSIÇÃO")
            return

        try:
            excel = win32.DispatchEx('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3 
            excel.EnableEvents = False

            sucesso = 0
            for index, caminho_xlsm in enumerate(arquivos_origem, 1):
                nome_arq = os.path.basename(caminho_xlsm)
                modelo_necessario = self.identificar_modelo(nome_arq)

                if not modelo_necessario:
                    self.log(f"[X] Ignorado: {nome_arq} (Não foi possível identificar o modelo pelo nome)")
                    continue

                caminho_modelo = os.path.join(self.pasta_modelos, modelo_necessario)
                if not os.path.exists(caminho_modelo):
                    self.log(f"[X] Erro: Modelo {modelo_necessario} não encontrado.")
                    continue

                self.log(f"-> Transpondo: {nome_arq} => Usando template {modelo_necessario}")

                wb_origem = None
                wb_modelo = None
                try:
                    # Abre Origem
                    wb_origem = excel.Workbooks.Open(os.path.abspath(caminho_xlsm), ReadOnly=True, UpdateLinks=False)
                    ws_origem = wb_origem.Sheets("Formulario")
                    
                    ultima_linha = ws_origem.Cells(ws_origem.Rows.Count, "A").End(-4162).Row
                    ultima_coluna = ws_origem.UsedRange.Columns.Count

                    if ultima_linha >= 6:
                        # Abre Template Base
                        wb_modelo = excel.Workbooks.Open(os.path.abspath(caminho_modelo), UpdateLinks=False)
                        ws_modelo = wb_modelo.Sheets("Formulario")

                        # Transposição Cirúrgica: Pula células protegidas pela seguradora
                        for linha in range(6, ultima_linha + 1):
                            for coluna in range(1, ultima_coluna + 1):
                                celula_destino = ws_modelo.Cells(linha, coluna)
                                
                                # Só injeta o dado se a célula da Mapfre estiver destravada para digitação
                                if not celula_destino.Locked:
                                    valor_origem = ws_origem.Cells(linha, coluna).Value
                                    if valor_origem is not None:
                                        celula_destino.Value = valor_origem

                        novo_nome = nome_arq.replace(".xlsm", ".xls")
                        caminho_final = os.path.join(self.pasta_destino, novo_nome)
                        
                        wb_modelo.SaveAs(os.path.abspath(caminho_final), FileFormat=56)
                        wb_modelo.Close(SaveChanges=False)
                        sucesso += 1
                    else:
                        self.log(f"[!] Aviso: Nenhuma linha de dado encontrada a partir da linha 6 em {nome_arq}")
                        
                    wb_origem.Close(SaveChanges=False)

                except Exception as e_interno:
                    self.log(f"[X] Erro ao manipular o arquivo {nome_arq}: {e_interno}")
                    if wb_origem: wb_origem.Close(SaveChanges=False)
                    if wb_modelo: wb_modelo.Close(SaveChanges=False)

                progresso = index / total
                self.barra_progresso.set(progresso)

            self.log(f"\n[OK] Processo finalizado! {sucesso} arquivos transpostos com sucesso.")

        except Exception as e:
            self.log(f"[ERRO FATAL] Ocorreu um problema de comunicação com o Excel: {str(e)}")
        finally:
            excel.Quit()
            pythoncom.CoUninitialize()
            self.btn_processar.configure(state="normal", text="INICIAR TRANSPOSIÇÃO")

if __name__ == "__main__":
    app = AppTotumseg()
    app.mainloop()
