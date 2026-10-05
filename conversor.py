import os
import glob
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox

try:
    import win32com.client as win32
    import pythoncom
except ImportError:
    pass # Tratamento para compilação via GitHub Actions

# Configuração Base do CustomTkinter
ctk.set_appearance_mode("light")

class AppTotumseg(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Totumseg | Integração Mapfre")
        self.geometry("700x650")
        self.resizable(False, False)
        # Fundo geral da janela num tom de cinza web ultra leve
        self.configure(fg_color="#F0F2F5") 

        # ================= HEADER =================
        self.header = ctk.CTkFrame(self, fg_color="#003366", corner_radius=0, height=80)
        self.header.pack(fill="x", side="top")
        
        self.lbl_titulo = ctk.CTkLabel(self.header, text="Integração Mapfre", font=("Segoe UI", 24, "bold"), text_color="white")
        self.lbl_titulo.place(relx=0.05, rely=0.5, anchor="w")
        
        self.lbl_subtitulo = ctk.CTkLabel(self.header, text="Transposição Automática de Dados", font=("Segoe UI", 12), text_color="#A9C2E3")
        self.lbl_subtitulo.place(relx=0.95, rely=0.5, anchor="e")

        # ================= CARD PRINCIPAL =================
        self.card = ctk.CTkFrame(self, fg_color="white", corner_radius=15, border_width=1, border_color="#E1E5EB")
        self.card.pack(pady=25, padx=25, fill="both", expand=True)

        self.lbl_instrucao = ctk.CTkLabel(self.card, text="Configuração de Diretórios", font=("Segoe UI", 16, "bold"), text_color="#1E293B")
        self.lbl_instrucao.pack(pady=(20, 15), padx=25, anchor="w")

        # ================= SELEÇÃO DE PASTAS =================
        self.frame_pastas = ctk.CTkFrame(self.card, fg_color="transparent")
        self.frame_pastas.pack(fill="x", padx=25)

        # Paleta de Cores e Fontes dos Botões
        btn_font = ("Segoe UI", 13, "bold")
        lbl_font = ("Segoe UI", 13)
        cor_btn = "#2563EB"          # Azul vibrante corporativo
        cor_hover = "#1D4ED8"        # Azul escuro no hover
        cor_texto_vazio = "#94A3B8"  # Cinza claro para placeholder
        cor_texto_ok = "#0F172A"     # Quase preto para pasta selecionada

        # 1. Modelos
        self.btn_modelos = ctk.CTkButton(self.frame_pastas, text="1. Pasta de Modelos", command=self.selecionar_modelos,
                                         font=btn_font, fg_color=cor_btn, hover_color=cor_hover, width=220, height=42)
        self.btn_modelos.grid(row=0, column=0, pady=10, sticky="w")
        self.lbl_modelos = ctk.CTkLabel(self.frame_pastas, text=" Nenhuma pasta selecionada", font=lbl_font, text_color=cor_texto_vazio)
        self.lbl_modelos.grid(row=0, column=1, padx=15, sticky="w")

        # 2. Origem
        self.btn_origem = ctk.CTkButton(self.frame_pastas, text="2. Pasta de Origem (.xlsx)", command=self.selecionar_origem,
                                         font=btn_font, fg_color=cor_btn, hover_color=cor_hover, width=220, height=42)
        self.btn_origem.grid(row=1, column=0, pady=10, sticky="w")
        self.lbl_origem = ctk.CTkLabel(self.frame_pastas, text=" Nenhuma pasta selecionada", font=lbl_font, text_color=cor_texto_vazio)
        self.lbl_origem.grid(row=1, column=1, padx=15, sticky="w")

        # 3. Destino
        self.btn_destino = ctk.CTkButton(self.frame_pastas, text="3. Pasta de Destino", command=self.selecionar_destino,
                                         font=btn_font, fg_color=cor_btn, hover_color=cor_hover, width=220, height=42)
        self.btn_destino.grid(row=2, column=0, pady=10, sticky="w")
        self.lbl_destino = ctk.CTkLabel(self.frame_pastas, text=" Nenhuma pasta selecionada", font=lbl_font, text_color=cor_texto_vazio)
        self.lbl_destino.grid(row=2, column=1, padx=15, sticky="w")

        # ================= LOG E PROGRESSO =================
        self.frame_console = ctk.CTkFrame(self.card, fg_color="#F8FAFC", corner_radius=10, border_width=1, border_color="#E2E8F0")
        self.frame_console.pack(fill="both", expand=True, padx=25, pady=(20, 10))

        self.txt_log = ctk.CTkTextbox(self.frame_console, fg_color="transparent", text_color="#334155", font=("Consolas", 12))
        self.txt_log.pack(fill="both", expand=True, padx=10, pady=10)
        self.txt_log.insert("0.0", "Sistema pronto. Aguardando configuração...\n")
        self.txt_log.configure(state="disabled")

        self.barra_progresso = ctk.CTkProgressBar(self.card, progress_color="#10B981", fg_color="#E2E8F0", height=6)
        self.barra_progresso.pack(fill="x", padx=25, pady=(0, 20))
        self.barra_progresso.set(0)

        # ================= AÇÃO =================
        self.btn_processar = ctk.CTkButton(self, text="INICIAR TRANSPOSIÇÃO", command=self.iniciar_processo, 
                                           font=("Segoe UI", 15, "bold"), fg_color="#059669", hover_color="#047857", height=50)
        self.btn_processar.pack(fill="x", padx=25, pady=(0, 25))

    # ================= FUNÇÕES DE SELEÇÃO =================
    def selecionar_modelos(self):
        self.pasta_modelos = filedialog.askdirectory(title="Selecione a pasta com os Modelos XLS da Mapfre")
        if self.pasta_modelos: 
            self.lbl_modelos.configure(text=f"  {os.path.basename(self.pasta_modelos)}", text_color="#0F172A")

    def selecionar_origem(self):
        self.pasta_origem = filedialog.askdirectory(title="Selecione a pasta com as planilhas do seu sistema")
        if self.pasta_origem: 
            self.lbl_origem.configure(text=f"  {os.path.basename(self.pasta_origem)}", text_color="#0F172A")

    def selecionar_destino(self):
        self.pasta_destino = filedialog.askdirectory(title="Selecione a pasta onde salvar os arquivos finais")
        if self.pasta_destino: 
            self.lbl_destino.configure(text=f"  {os.path.basename(self.pasta_destino)}", text_color="#0F172A")

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
        
        self.btn_processar.configure(state="disabled", text="PROCESSANDO...", fg_color="#64748B")
        threading.Thread(target=self.transpor_dados, daemon=True).start()

    def identificar_modelo(self, nome_arquivo):
        nome_upper = nome_arquivo.upper()
        ramo = ""
        valor = ""

        if "MADEIRA" in nome_upper: ramo = "CASA_MADEIRA"
        elif "CASA" in nome_upper: ramo = "CASA"
        elif "APARTAMENTO" in nome_upper or "APTO" in nome_upper: ramo = "APTO"
        elif "COMERCIO" in nome_upper or "COMER" in nome_upper: ramo = "COMER"
        elif "ESCRITORIO" in nome_upper or "ESCRIT" in nome_upper: ramo = "ESCRIT"

        if "200000" in nome_upper: valor = "200"
        elif "300000" in nome_upper: valor = "300"
        elif "400000" in nome_upper: valor = "400"
        elif "600000" in nome_upper: valor = "600"

        if ramo and valor:
            return f"{ramo}_{valor}.xls"
        return None

    def transpor_dados(self):
        pythoncom.CoInitialize()
        
        # Lê tanto os arquivos .xlsm quanto os arquivos brutos .xlsx
        arquivos_origem = glob.glob(os.path.join(self.pasta_origem, "*.xlsm"))
        arquivos_origem.extend(glob.glob(os.path.join(self.pasta_origem, "*.xlsx")))
        
        total = len(arquivos_origem)

        if total == 0:
            self.log("[!] Nenhum arquivo .xlsm ou .xlsx encontrado na pasta de origem.")
            self.btn_processar.configure(state="normal", text="INICIAR TRANSPOSIÇÃO", fg_color="#059669")
            return

        try:
            excel = win32.DispatchEx('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3 # Bloqueia macros durante o processo para não travar
            excel.EnableEvents = False

            sucesso = 0
            for index, caminho_arquivo in enumerate(arquivos_origem, 1):
                nome_arq = os.path.basename(caminho_arquivo)
                modelo_necessario = self.identificar_modelo(nome_arq)

                if not modelo_necessario:
                    self.log(f"[X] Ignorado: {nome_arq} (Modelo não identificado)")
                    continue

                caminho_modelo = os.path.join(self.pasta_modelos, modelo_necessario)
                if not os.path.exists(caminho_modelo):
                    self.log(f"[X] Erro: Template {modelo_necessario} não encontrado.")
                    continue

                self.log(f"-> Injetando: {nome_arq}")

                wb_origem = None
                wb_modelo = None
                try:
                    # Abre Origem e Modelo
                    wb_origem = excel.Workbooks.Open(os.path.abspath(caminho_arquivo), ReadOnly=True, UpdateLinks=False)
                    ws_origem = wb_origem.Sheets(1)
                    
                    wb_modelo = excel.Workbooks.Open(os.path.abspath(caminho_modelo), UpdateLinks=False)
                    ws_modelo = wb_modelo.Sheets("Formulario")

                    linha = 6
                    linhas_processadas = 0

                    while True:
                        # Para de varrer se a coluna B estiver vazia
                        valor_verificador = ws_origem.Cells(linha, 2).Value
                        if valor_verificador is None or str(valor_verificador).strip() == "":
                            break

                        # Varre apenas da coluna B (2) até a coluna AN (40)
                        for coluna in range(2, 41):
                            celula_destino = ws_modelo.Cells(linha, coluna)
                            
                            # Injeta os valores APENAS se a célula do modelo permitir digitação
                            if not celula_destino.Locked:
                                valor_origem = ws_origem.Cells(linha, coluna).Value
                                if valor_origem is not None:
                                    celula_destino.Value = valor_origem
                        
                        linha += 1
                        linhas_processadas += 1

                    if linhas_processadas > 0:
                        # Salva o resultado final no formato aceito (Excel 97-2003)
                        novo_nome = nome_arq.replace(".xlsm", ".xls").replace(".xlsx", ".xls")
                        caminho_final = os.path.join(self.pasta_destino, novo_nome)
                        
                        wb_modelo.SaveAs(os.path.abspath(caminho_final), FileFormat=56)
                        sucesso += 1
                    else:
                        self.log(f"[!] Aviso: Nenhuma linha com dados encontrada em {nome_arq}")
                        
                    wb_origem.Close(SaveChanges=False)
                    wb_modelo.Close(SaveChanges=False)

                except Exception as e_interno:
                    self.log(f"[X] Falha no arquivo {nome_arq}: {e_interno}")
                    if wb_origem: wb_origem.Close(SaveChanges=False)
                    if wb_modelo: wb_modelo.Close(SaveChanges=False)

                progresso = index / total
                self.barra_progresso.set(progresso)

            self.log(f"\n[OK] Concluído! {sucesso} arquivos gerados com sucesso.")

        except Exception as e:
            self.log(f"[ERRO FATAL] Falha de comunicação com o Excel: {str(e)}")
        finally:
            excel.Quit()
            pythoncom.CoUninitialize()
            self.btn_processar.configure(state="normal", text="INICIAR TRANSPOSIÇÃO", fg_color="#059669")

if __name__ == "__main__":
    app = AppTotumseg()
    app.mainloop()
