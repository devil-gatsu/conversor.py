"""
Totumseg - Jet x Mapfre xls
---------------------------
Le as planilhas (.xlsm/.xlsx) da pasta base, escolhe o modelo .xls da Mapfre
pelo nome do arquivo (tipo do imovel + valor), cola os dados no modelo e salva
na pasta destino com o mesmo nome do arquivo original (extensao .xls).
"""
import glob
import json
import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import unicodedata
import re

import customtkinter as ctk

try:
    import pythoncom
    import win32com.client as win32
    TEM_EXCEL_COM = True
except ImportError:
    TEM_EXCEL_COM = False

NOME_APP = "Totumseg | Integração Mapfre"

MODELOS_POR_RAMO = {
    "APTO":         ["200", "300", "400"],
    "CASA_MADEIRA": ["200", "300", "400"],
    "CASA":         ["200", "300", "400"],
    "COMER":        ["200", "400", "600"],
    "ESCRIT":       ["200", "300", "600"],
}
MODELOS_VALIDOS = [f"{r}_{v}" for r, vs in MODELOS_POR_RAMO.items() for v in vs]

BRANCO = "#FFFFFF"
FUNDO_SUAVE = "#F0F2F5"
AZUL = "#2563EB"
AZUL_ESCURO = "#1D4ED8"
AZUL_CLARO = "#EFF6FF"
BORDA = "#E2E8F0"
TEXTO = "#0F172A"
TEXTO_SUAVE = "#64748B"
CINZA_CHIP = "#F1F5F9"
CINZA_CHIP_TEXTO = "#94A3B8"
VERDE = "#059669"
VERMELHO = "#B91C1C"
LARANJA = "#B45309"

ARQ_CONFIG = os.path.join(os.path.expanduser("~"), ".totumseg_jet_mapfre.json")

ctk.set_appearance_mode("light")

# --------------------------------------------------------------------------
# Funções de Tratamento de Dados
# --------------------------------------------------------------------------
def identificar_modelo(nome_arquivo):
    nome = nome_arquivo.upper()
    ramo = ""
    valor = ""

    if "MADEIRA" in nome: ramo = "CASA_MADEIRA"
    elif "CASA" in nome: ramo = "CASA"
    elif "APART" in nome or "APTO" in nome: ramo = "APTO"
    elif "COMER" in nome: ramo = "COMER"
    elif "ESCRIT" in nome: ramo = "ESCRIT"

    for v in ("200", "300", "400", "600"):
        if f"{v}000" in nome:
            valor = v
            break

    if ramo and valor:
        return f"{ramo}_{valor}"
    return None

def sanitizar_padrao(valor):
    """Ajustes universais que a Mapfre exige (país e máscara de documentos)."""
    if isinstance(valor, str):
        texto = valor.strip()
        if texto.upper() == "BRASIL":
            return "1-BRASIL"
        if len(texto) == 11 and texto.isdigit():
            return f"{texto[:3]}.{texto[3:6]}.{texto[6:9]}-{texto[9:]}"
        if len(texto) == 14 and texto.isdigit():
            return f"{texto[:2]}.{texto[2:5]}.{texto[5:8]}/{texto[8:12]}-{texto[12:]}"
    return valor

def remover_caracteres_especiais(texto):
    """Remove acentos, ç, e pontuações, mantendo apenas letras, números e espaços."""
    if not isinstance(texto, str):
        return texto
    
    # Substitui acentos e caracteres especiais (ex: Ç vira C, Á vira A)
    texto_sem_acento = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    # Remove qualquer coisa que não seja letra (a-z), número (0-9) ou espaço (\s)
    texto_limpo = re.sub(r'[^a-zA-Z0-9\s]', '', texto_sem_acento)
    
    return texto_limpo

def encurtar(caminho, limite=48):
    if len(caminho) <= limite:
        return caminho
    return "…" + caminho[-(limite - 1):]

# --------------------------------------------------------------------------
# Interface (Layout Compacto)
# --------------------------------------------------------------------------
class AppTotumseg(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(NOME_APP)
        self.geometry("600x550")
        self.resizable(False, False)
        self.configure(fg_color=FUNDO_SUAVE)

        self.modelos = {}          
        self.pasta_origem = ""
        self.pasta_destino = ""
        self.fila = queue.Queue()  
        self.processando = False

        self._montar_cabecalho()

        corpo = ctk.CTkFrame(self, fg_color=BRANCO, corner_radius=10, border_width=1, border_color=BORDA)
        corpo.pack(fill="both", expand=True, padx=20, pady=(15, 15))

        ctk.CTkLabel(corpo, text="Configuração de Diretórios", font=("Segoe UI", 14, "bold"), text_color=TEXTO).pack(anchor="w", padx=20, pady=(10, 5))

        self.lbl_modelos = self._montar_linha_selecao(corpo, "1. Pasta de Modelos", " Nenhuma pasta selecionada", self.selecionar_modelos)
        self.lbl_origem = self._montar_linha_selecao(corpo, "2. Pasta de Origem (.xlsx)", " Nenhuma pasta selecionada", self.selecionar_origem)
        self.lbl_destino = self._montar_linha_selecao(corpo, "3. Pasta de Destino", " Nenhuma pasta selecionada", self.selecionar_destino)

        self._montar_log_e_progresso(corpo)

        self.btn_processar = ctk.CTkButton(
            self, text="INICIAR TRANSPOSIÇÃO", command=self.iniciar_processo,
            font=("Segoe UI", 14, "bold"), height=40, corner_radius=6,
            fg_color=VERDE, hover_color="#047857", text_color=BRANCO)
        self.btn_processar.pack(fill="x", padx=20, pady=(0, 20))

        self._carregar_config()
        self._atualizar_pastas()
        self._mensagem_inicial()
        self.after(100, self._ler_fila)

    def _montar_cabecalho(self):
        topo = ctk.CTkFrame(self, fg_color="#003366", corner_radius=0, height=60)
        topo.pack(fill="x", side="top")
        topo.pack_propagate(False)
        ctk.CTkLabel(topo, text="Integração Mapfre", font=("Segoe UI", 20, "bold"), text_color=BRANCO).pack(side="left", padx=20, pady=15)
        ctk.CTkLabel(topo, text="Transposição Automática", font=("Segoe UI", 12), text_color="#A9C2E3").pack(side="right", padx=20, pady=20)

    def _montar_linha_selecao(self, pai, texto_btn, texto_lbl, comando):
        frame = ctk.CTkFrame(pai, fg_color="transparent")
        frame.pack(fill="x", padx=20, pady=5)
        ctk.CTkButton(frame, text=texto_btn, command=comando, width=180, height=32,
                      font=("Segoe UI", 12, "bold"), fg_color=AZUL, hover_color=AZUL_ESCURO).pack(side="left")
        lbl = ctk.CTkLabel(frame, text=texto_lbl, font=("Segoe UI", 12), text_color=TEXTO_SUAVE)
        lbl.pack(side="left", padx=10)
        lbl._descricao_padrao = texto_lbl
        return lbl

    def _montar_log_e_progresso(self, pai):
        frame_console = ctk.CTkFrame(pai, fg_color=FUNDO_SUAVE, corner_radius=8, border_width=1, border_color=BORDA)
        frame_console.pack(fill="both", expand=True, padx=20, pady=(15, 10))

        self.txt_log = ctk.CTkTextbox(frame_console, fg_color="transparent", text_color=TEXTO, font=("Consolas", 11))
        self.txt_log.pack(fill="both", expand=True, padx=10, pady=5)
        self.txt_log.configure(state="disabled")

        self.barra = ctk.CTkProgressBar(pai, height=6, progress_color="#10B981", fg_color=BORDA)
        self.barra.pack(fill="x", padx=20, pady=(0, 15))
        self.barra.set(0)

    def _atualizar_pastas(self):
        for lbl, caminho in ((self.lbl_modelos, self.pasta_modelos), (self.lbl_origem, self.pasta_origem), (self.lbl_destino, self.pasta_destino)):
            if caminho:
                lbl.configure(text=f"  {encurtar(os.path.basename(caminho))}", text_color=TEXTO)
            else:
                lbl.configure(text=lbl._descricao_padrao, text_color=TEXTO_SUAVE)

        # Mapeia os arquivos de modelos encontrados na pasta
        if self.pasta_modelos:
            arquivos_xls = glob.glob(os.path.join(self.pasta_modelos, "*.xls"))
            for caminho in arquivos_xls:
                nome = os.path.splitext(os.path.basename(caminho))[0].upper()
                if nome in MODELOS_VALIDOS:
                    self.modelos[nome] = caminho

    def _mensagem_inicial(self):
        if not TEM_EXCEL_COM:
            self.log("[ERRO] pywin32 não encontrado. Este programa precisa do Excel.")
        else:
            self.log("Sistema pronto. Aguardando configuração...")

    def log(self, mensagem):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", mensagem + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    # ---------------- selecoes ----------------
    def selecionar_modelos(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta com os Modelos XLS")
        if pasta:
            self.pasta_modelos = pasta
            self.modelos.clear()
            self._atualizar_pastas()
            self._salvar_config()

    def selecionar_origem(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta base (dados)")
        if pasta:
            self.pasta_origem = pasta
            self._atualizar_pastas()
            self._salvar_config()

    def selecionar_destino(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta destino")
        if pasta:
            self.pasta_destino = pasta
            self._atualizar_pastas()
            self._salvar_config()

    # ---------------- configuracao ----------------
    def _salvar_config(self):
        try:
            with open(ARQ_CONFIG, "w", encoding="utf-8") as f:
                json.dump({"pasta_modelos": self.pasta_modelos, "origem": self.pasta_origem,
                           "destino": self.pasta_destino}, f, ensure_ascii=False)
        except OSError:
            pass

    def _carregar_config(self):
        try:
            with open(ARQ_CONFIG, encoding="utf-8") as f:
                cfg = json.load(f)
        except (OSError, ValueError):
            return
        if os.path.isdir(cfg.get("pasta_modelos", "")): self.pasta_modelos = cfg["pasta_modelos"]
        if os.path.isdir(cfg.get("origem", "")): self.pasta_origem = cfg["origem"]
        if os.path.isdir(cfg.get("destino", "")): self.pasta_destino = cfg["destino"]

    # ---------------- execucao ----------------
    def iniciar_processo(self):
        if self.processando: return
        
        if not all([self.pasta_modelos, self.pasta_origem, self.pasta_destino]):
            messagebox.showwarning("Atenção", "Selecione as três pastas antes de iniciar.")
            return

        self.processando = True
        self.btn_processar.configure(state="disabled", text="PROCESSANDO...", fg_color=TEXTO_SUAVE)
        self.barra.set(0)
        self.log("\n— Novo processamento iniciado —")
        threading.Thread(
            target=self._processar,
            args=(dict(self.modelos), self.pasta_origem, self.pasta_destino),
            daemon=True).start()

    def _ui(self, tipo, a=None):
        self.fila.put((tipo, a))

    def _ler_fila(self):
        try:
            while True:
                tipo, a = self.fila.get_nowait()
                if tipo == "log":
                    self.log(a)
                elif tipo == "progresso":
                    self.barra.set(a)
                elif tipo == "fim":
                    self.processando = False
                    self.btn_processar.configure(state="normal", text="INICIAR TRANSPOSIÇÃO", fg_color=VERDE)
        except queue.Empty:
            pass
        self.after(100, self._ler_fila)

    def _processar(self, modelos, origem, destino):
        pythoncom.CoInitialize()
        excel = None
        try:
            arquivos = glob.glob(os.path.join(origem, "*.xlsm")) + glob.glob(os.path.join(origem, "*.xlsx"))
            arquivos = sorted(a for a in arquivos if not os.path.basename(a).startswith("~$"))
            total = len(arquivos)
            
            if total == 0:
                self._ui("log", "[!] Nenhum arquivo .xlsm ou .xlsx encontrado na pasta base.")
                return

            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3
            excel.EnableEvents = False
            excel.ScreenUpdating = False # Otimização de velocidade

            sucesso = 0
            for i, caminho in enumerate(arquivos, 1):
                nome = os.path.basename(caminho)
                chave = identificar_modelo(nome)

                if not chave:
                    self._ui("log", f"[X] Ignorado: {nome} (Modelo não identificado)")
                    self._ui("progresso", i / total)
                    continue
                if chave not in modelos:
                    self._ui("log", f"[X] Erro: Template {chave}.xls não foi encontrado na pasta de modelos.")
                    self._ui("progresso", i / total)
                    continue

                self._ui("log", f"-> Injetando: {nome}")
                
                wb_origem = wb_modelo = None
                try:
                    wb_origem = excel.Workbooks.Open(os.path.abspath(caminho), ReadOnly=True, UpdateLinks=False)
                    ws_origem = wb_origem.Sheets(1)
                    
                    wb_modelo = excel.Workbooks.Open(os.path.abspath(modelos[chave]), UpdateLinks=False)
                    ws_modelo = wb_modelo.Sheets("Formulario")

                    linha = 6
                    processadas = 0
                    
                    # Colunas do Excel que não podem ter pontuação ou acentos: G(7), N(14), P(16), W(23), Y(25)
                    colunas_sem_especiais = {7, 14, 16, 23, 25}

                    while True:
                        verificador = ws_origem.Cells(linha, 2).Value
                        if verificador is None or str(verificador).strip() == "":
                            break
                        
                        # OTIMIZAÇÃO: Fotografa a linha toda da B(2) até AN(40) para a memória em uma única batida
                        valores_linha = ws_origem.Range(ws_origem.Cells(linha, 2), ws_origem.Cells(linha, 40)).Value[0]

                        # Itera sobre os valores que já estão na memória do Python
                        for indice_memoria, valor_bruto in enumerate(valores_linha):
                            coluna = indice_memoria + 2 # Ajuste para bater com o índice do Excel (B = 2)
                            destino_cel = ws_modelo.Cells(linha, coluna)
                            
                            if not destino_cel.Locked:
                                valor_tratado = sanitizar_padrao(valor_bruto)
                                
                                # Aplica o trator de limpeza apenas nas colunas restritas
                                if coluna in colunas_sem_especiais:
                                    valor_tratado = remover_caracteres_especiais(valor_tratado)
                                
                                if valor_tratado is not None:
                                    destino_cel.Value = valor_tratado
                                    
                        linha += 1
                        processadas += 1

                    if processadas > 0:
                        novo_nome = os.path.splitext(nome)[0] + ".xls"
                        caminho_final = os.path.join(destino, novo_nome)
                        wb_modelo.SaveAs(os.path.abspath(caminho_final), FileFormat=56)
                        sucesso += 1
                    else:
                        self._ui("log", f"[!] Aviso: Nenhuma linha com dados em {nome}")

                except Exception as e_interno:
                    self._ui("log", f"[X] Falha no arquivo {nome}: {e_interno}")
                finally:
                    for wb in (wb_origem, wb_modelo):
                        if wb is not None:
                            try: wb.Close(SaveChanges=False)
                            except: pass

                self._ui("progresso", i / total)

            self._ui("log", f"\n[OK] Concluído! {sucesso} de {total} arquivo(s) gerado(s).")
            
        except Exception as e:
            self._ui("log", f"[ERRO FATAL] Falha com Excel: {str(e)}")
        finally:
            if excel is not None:
                try: excel.Quit()
                except: pass
            pythoncom.CoUninitialize()
            self._ui("fim", None)

if __name__ == "__main__":
    app = AppTotumseg()
    app.mainloop()
