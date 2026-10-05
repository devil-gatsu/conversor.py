"""
Totumseg - Jet x Mapfre xls
---------------------------
Le as planilhas (.xlsm/.xlsx) da pasta base, escolhe o modelo .xls da Mapfre
pelo nome do arquivo (tipo do imovel + valor), cola os dados no modelo e salva
na pasta destino com o mesmo nome do arquivo original (extensao .xls).

Dependencias (requirements.txt):
    customtkinter
    pywin32
"""
import glob
import json
import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

try:
    import pythoncom
    import win32com.client as win32
    TEM_EXCEL_COM = True
except ImportError:  # fora do Windows / sem pywin32
    TEM_EXCEL_COM = False

NOME_APP = "Totumseg - Jet x Mapfre xls"

# --------------------------------------------------------------------------
# Modelos aceitos (os arquivos precisam ter exatamente estes nomes + .xls)
# --------------------------------------------------------------------------
MODELOS_POR_RAMO = {
    "APTO":         ["200", "300", "400"],
    "CASA_MADEIRA": ["200", "300", "400"],
    "CASA":         ["200", "300", "400"],
    "COMER":        ["200", "400", "600"],
    "ESCRIT":       ["200", "300", "600"],
}
MODELOS_VALIDOS = [f"{r}_{v}" for r, vs in MODELOS_POR_RAMO.items() for v in vs]

# --------------------------------------------------------------------------
# Paleta: predominantemente branco, com azul de destaque
# --------------------------------------------------------------------------
BRANCO = "#FFFFFF"
FUNDO_SUAVE = "#F8FAFC"
AZUL = "#2563EB"
AZUL_ESCURO = "#1D4ED8"
AZUL_CLARO = "#EFF6FF"
AZUL_BORDA = "#BFDBFE"
BORDA = "#E2E8F0"
TEXTO = "#0F172A"
TEXTO_SUAVE = "#64748B"
CINZA_CHIP = "#F1F5F9"
CINZA_CHIP_TEXTO = "#94A3B8"
VERDE = "#15803D"
VERMELHO = "#B91C1C"
LARANJA = "#B45309"

ARQ_CONFIG = os.path.join(os.path.expanduser("~"), ".totumseg_jet_mapfre.json")

ctk.set_appearance_mode("light")


# --------------------------------------------------------------------------
# Regras de negocio (mesma logica do programa original)
# --------------------------------------------------------------------------
def identificar_modelo(nome_arquivo):
    """Devolve 'RAMO_VALOR' (ex.: 'APTO_300') ou None se nao identificar."""
    nome = nome_arquivo.upper()
    ramo = ""
    valor = ""

    if "MADEIRA" in nome:
        ramo = "CASA_MADEIRA"
    elif "CASA" in nome:
        ramo = "CASA"
    elif "APART" in nome or "APTO" in nome:
        ramo = "APTO"
    elif "COMER" in nome:
        ramo = "COMER"
    elif "ESCRIT" in nome:
        ramo = "ESCRIT"

    for v in ("200", "300", "400", "600"):
        if f"{v}000" in nome:
            valor = v
            break

    if ramo and valor:
        return f"{ramo}_{valor}"
    return None


def sanitizar(valor):
    """Ajustes que a Mapfre exige (pais e mascara de CPF/CNPJ)."""
    if isinstance(valor, str):
        texto = valor.strip()
        if texto.upper() == "BRASIL":
            return "1-BRASIL"
        if len(texto) == 11 and texto.isdigit():
            return f"{texto[:3]}.{texto[3:6]}.{texto[6:9]}-{texto[9:]}"
        if len(texto) == 14 and texto.isdigit():
            return f"{texto[:2]}.{texto[2:5]}.{texto[5:8]}/{texto[8:12]}-{texto[12:]}"
    return valor


def encurtar(caminho, limite=52):
    if len(caminho) <= limite:
        return caminho
    return "…" + caminho[-(limite - 1):]


# --------------------------------------------------------------------------
# Interface
# --------------------------------------------------------------------------
class AppTotumseg(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(NOME_APP)
        self.geometry("760x720")
        self.minsize(680, 640)
        self.configure(fg_color=BRANCO)

        self.modelos = {}          # 'APTO_300' -> caminho do .xls
        self.pasta_origem = ""
        self.pasta_destino = ""
        self.fila = queue.Queue()  # comunicacao thread -> interface
        self.processando = False

        self._montar_cabecalho()

        corpo = ctk.CTkFrame(self, fg_color=BRANCO)
        corpo.pack(fill="both", expand=True, padx=24, pady=(14, 18))
        corpo.grid_columnconfigure(0, weight=1)
        corpo.grid_rowconfigure(4, weight=1)

        self._montar_card_modelos(corpo)
        self.lbl_origem = self._montar_card_pasta(
            corpo, 1, "2", "Pasta base",
            "Planilhas (.xlsm / .xlsx) com os dados", "Selecionar pasta", self.selecionar_origem)
        self.lbl_destino = self._montar_card_pasta(
            corpo, 2, "3", "Pasta destino",
            "Onde os arquivos .xls serão salvos", "Selecionar pasta", self.selecionar_destino)
        self._montar_progresso(corpo)
        self._montar_log(corpo)

        self.btn_processar = ctk.CTkButton(
            corpo, text="Iniciar processamento", command=self.iniciar_processo,
            font=("Segoe UI", 14, "bold"), height=44, corner_radius=8,
            fg_color=AZUL, hover_color=AZUL_ESCURO, text_color=BRANCO,
            text_color_disabled="#DBEAFE")
        self.btn_processar.grid(row=5, column=0, sticky="ew", pady=(12, 0))

        self._carregar_config()
        self._atualizar_modelos()
        self._atualizar_pastas()
        self._mensagem_inicial()
        self.after(100, self._ler_fila)

    # ---------------- construcao da tela ----------------
    def _montar_cabecalho(self):
        topo = ctk.CTkFrame(self, fg_color=BRANCO, corner_radius=0)
        topo.pack(fill="x")
        ctk.CTkLabel(topo, text=NOME_APP, font=("Segoe UI", 22, "bold"),
                     text_color=AZUL).pack(anchor="w", padx=24, pady=(18, 0))
        ctk.CTkLabel(topo, text="Transposição automática de planilhas para o modelo da Mapfre",
                     font=("Segoe UI", 12), text_color=TEXTO_SUAVE).pack(anchor="w", padx=24, pady=(0, 14))
        ctk.CTkFrame(topo, fg_color=AZUL, height=3, corner_radius=0).pack(fill="x")

    def _card(self, pai, linha):
        card = ctk.CTkFrame(pai, fg_color=BRANCO, corner_radius=10,
                            border_width=1, border_color=BORDA)
        card.grid(row=linha, column=0, sticky="ew", pady=(0, 8))
        card.grid_columnconfigure(1, weight=1)
        return card

    def _selo(self, card, numero):
        ctk.CTkLabel(card, text=numero, width=30, height=30, corner_radius=15,
                     fg_color=AZUL_CLARO, text_color=AZUL,
                     font=("Segoe UI", 13, "bold")).grid(row=0, column=0, rowspan=2, padx=(14, 10), pady=12)

    def _botao(self, card, texto, comando):
        b = ctk.CTkButton(card, text=texto, command=comando, width=150, height=34,
                          corner_radius=8, font=("Segoe UI", 12, "bold"),
                          fg_color=BRANCO, hover_color=AZUL_CLARO, text_color=AZUL,
                          border_width=1, border_color=AZUL)
        b.grid(row=0, column=2, rowspan=2, padx=14)
        return b

    def _montar_card_modelos(self, pai):
        card = self._card(pai, 0)
        self._selo(card, "1")
        ctk.CTkLabel(card, text="Modelos da Mapfre", font=("Segoe UI", 14, "bold"),
                     text_color=TEXTO, anchor="w").grid(row=0, column=1, sticky="sw", pady=(12, 0))
        self.lbl_modelos = ctk.CTkLabel(card, text="", font=("Segoe UI", 11),
                                        text_color=TEXTO_SUAVE, anchor="w")
        self.lbl_modelos.grid(row=1, column=1, sticky="nw", pady=(0, 10))
        self.btn_modelos = self._botao(card, "Carregar modelos", self.selecionar_modelos)

        # grade de "etiquetas": uma coluna por ramo, uma etiqueta por valor
        grade = ctk.CTkFrame(card, fg_color=FUNDO_SUAVE, corner_radius=8)
        grade.grid(row=2, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 12))
        self.chips = {}
        for col, (ramo, valores) in enumerate(MODELOS_POR_RAMO.items()):
            grade.grid_columnconfigure(col, weight=1, uniform="ramo")
            ctk.CTkLabel(grade, text=ramo.replace("_", " "), font=("Segoe UI", 10, "bold"),
                         text_color=TEXTO_SUAVE).grid(row=0, column=col, pady=(6, 2))
            for i, v in enumerate(valores, start=1):
                chip = ctk.CTkLabel(grade, text=v, width=54, height=20, corner_radius=6,
                                    font=("Segoe UI", 11, "bold"),
                                    fg_color=CINZA_CHIP, text_color=CINZA_CHIP_TEXTO)
                chip.grid(row=i, column=col, pady=2)
                self.chips[f"{ramo}_{v}"] = chip
            ctk.CTkFrame(grade, fg_color="transparent", height=4).grid(row=4, column=col)

    def _montar_card_pasta(self, pai, linha, numero, titulo, descricao, texto_botao, comando):
        card = self._card(pai, linha)
        self._selo(card, numero)
        ctk.CTkLabel(card, text=titulo, font=("Segoe UI", 14, "bold"),
                     text_color=TEXTO, anchor="w").grid(row=0, column=1, sticky="sw", pady=(12, 0))
        lbl = ctk.CTkLabel(card, text=descricao, font=("Segoe UI", 11),
                           text_color=TEXTO_SUAVE, anchor="w")
        lbl.grid(row=1, column=1, sticky="nw", pady=(0, 12))
        self._botao(card, texto_botao, comando)
        lbl._descricao_padrao = descricao
        return lbl

    def _montar_progresso(self, pai):
        quadro = ctk.CTkFrame(pai, fg_color=BRANCO)
        quadro.grid(row=3, column=0, sticky="ew", pady=(6, 6))
        quadro.grid_columnconfigure(0, weight=1)
        self.lbl_status = ctk.CTkLabel(quadro, text="Aguardando", font=("Segoe UI", 12),
                                       text_color=TEXTO_SUAVE, anchor="w")
        self.lbl_status.grid(row=0, column=0, sticky="w")
        self.lbl_percentual = ctk.CTkLabel(quadro, text="0%", font=("Segoe UI", 12, "bold"),
                                           text_color=AZUL)
        self.lbl_percentual.grid(row=0, column=1, sticky="e")
        self.barra = ctk.CTkProgressBar(quadro, height=10, corner_radius=5,
                                        fg_color=AZUL_CLARO, progress_color=AZUL)
        self.barra.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        self.barra.set(0)

    def _montar_log(self, pai):
        self.txt_log = ctk.CTkTextbox(pai, fg_color=FUNDO_SUAVE, text_color=TEXTO,
                                      font=("Consolas", 11), corner_radius=10,
                                      border_width=1, border_color=BORDA, wrap="word",
                                      height=140)
        self.txt_log.grid(row=4, column=0, sticky="nsew")
        self.txt_log.tag_config("ok", foreground=VERDE)
        self.txt_log.tag_config("erro", foreground=VERMELHO)
        self.txt_log.tag_config("aviso", foreground=LARANJA)
        self.txt_log.tag_config("info", foreground=AZUL)
        self.txt_log.configure(state="disabled")

    # ---------------- estado da tela ----------------
    def _atualizar_modelos(self):
        for nome, chip in self.chips.items():
            if nome in self.modelos:
                chip.configure(fg_color=AZUL, text_color=BRANCO)
            else:
                chip.configure(fg_color=CINZA_CHIP, text_color=CINZA_CHIP_TEXTO)
        n = len(self.modelos)
        total = len(MODELOS_VALIDOS)
        if n == 0:
            self.lbl_modelos.configure(text=f"Nenhum modelo carregado (0/{total})", text_color=TEXTO_SUAVE)
        elif n == total:
            self.lbl_modelos.configure(text=f"Todos os modelos carregados ({n}/{total})", text_color=VERDE)
        else:
            self.lbl_modelos.configure(text=f"{n}/{total} modelos carregados", text_color=AZUL)

    def _atualizar_pastas(self):
        for lbl, caminho in ((self.lbl_origem, self.pasta_origem), (self.lbl_destino, self.pasta_destino)):
            if caminho:
                lbl.configure(text=encurtar(caminho), text_color=TEXTO)
            else:
                lbl.configure(text=lbl._descricao_padrao, text_color=TEXTO_SUAVE)

    def _mensagem_inicial(self):
        if not TEM_EXCEL_COM:
            self.log("Biblioteca pywin32 não encontrada. Este programa precisa do Windows com Excel e do pywin32 instalado.", "erro")
        else:
            self.log("Pronto. Carregue os modelos, escolha as pastas e clique em Iniciar processamento.", "info")

    def log(self, mensagem, tipo=None):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", mensagem + "\n", tipo)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    # ---------------- selecoes ----------------
    def selecionar_modelos(self):
        arquivos = filedialog.askopenfilenames(
            title="Selecione os modelos (.xls) da Mapfre",
            filetypes=[("Planilhas Excel 97-2003", "*.xls")])
        if not arquivos:
            return
        recusados = []
        for caminho in arquivos:
            nome, ext = os.path.splitext(os.path.basename(caminho))
            if ext.lower() == ".xls" and nome.upper() in MODELOS_VALIDOS:
                self.modelos[nome.upper()] = caminho
            else:
                recusados.append(os.path.basename(caminho))
        self._atualizar_modelos()
        self._salvar_config()
        if recusados:
            messagebox.showwarning(
                "Alguns arquivos foram ignorados",
                "Os modelos precisam ter exatamente um destes nomes (.xls):\n\n"
                + ", ".join(MODELOS_VALIDOS)
                + "\n\nIgnorados:\n" + "\n".join(recusados))

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

    # ---------------- configuracao salva ----------------
    def _salvar_config(self):
        try:
            with open(ARQ_CONFIG, "w", encoding="utf-8") as f:
                json.dump({"modelos": self.modelos, "origem": self.pasta_origem,
                           "destino": self.pasta_destino}, f, ensure_ascii=False)
        except OSError:
            pass

    def _carregar_config(self):
        try:
            with open(ARQ_CONFIG, encoding="utf-8") as f:
                cfg = json.load(f)
        except (OSError, ValueError):
            return
        self.modelos = {k: v for k, v in cfg.get("modelos", {}).items()
                        if k in MODELOS_VALIDOS and os.path.isfile(v)}
        if os.path.isdir(cfg.get("origem", "")):
            self.pasta_origem = cfg["origem"]
        if os.path.isdir(cfg.get("destino", "")):
            self.pasta_destino = cfg["destino"]

    # ---------------- execucao ----------------
    def iniciar_processo(self):
        if self.processando:
            return
        faltando = []
        if not self.modelos:
            faltando.append("carregar os modelos")
        if not self.pasta_origem:
            faltando.append("selecionar a pasta base")
        if not self.pasta_destino:
            faltando.append("selecionar a pasta destino")
        if faltando:
            messagebox.showwarning("Falta configurar", "Antes de iniciar, é preciso:\n\n• " + "\n• ".join(faltando))
            return
        if not TEM_EXCEL_COM:
            messagebox.showerror("Excel não disponível",
                                 "Não foi possível usar o Excel (pywin32 não encontrado).")
            return

        self.processando = True
        self.btn_processar.configure(state="disabled", text="Processando…")
        self.barra.set(0)
        self.lbl_percentual.configure(text="0%")
        self.lbl_status.configure(text="Iniciando…")
        self.log("— Novo processamento —", "info")
        threading.Thread(
            target=self._processar,
            args=(dict(self.modelos), self.pasta_origem, self.pasta_destino),
            daemon=True).start()

    # A thread nunca mexe na interface direto: manda mensagens pela fila.
    def _ui(self, tipo, a=None, b=None):
        self.fila.put((tipo, a, b))

    def _ler_fila(self):
        try:
            while True:
                tipo, a, b = self.fila.get_nowait()
                if tipo == "log":
                    self.log(a, b)
                elif tipo == "progresso":
                    self.barra.set(a)
                    self.lbl_percentual.configure(text=f"{int(a * 100)}%")
                    self.lbl_status.configure(text=b)
                elif tipo == "fim":
                    self.processando = False
                    self.btn_processar.configure(state="normal", text="Iniciar processamento")
                    self.lbl_status.configure(text=a)
        except queue.Empty:
            pass
        self.after(100, self._ler_fila)

    def _processar(self, modelos, origem, destino):
        pythoncom.CoInitialize()
        excel = None
        resumo = "Concluído"
        try:
            arquivos = glob.glob(os.path.join(origem, "*.xlsm")) + glob.glob(os.path.join(origem, "*.xlsx"))
            arquivos = sorted(a for a in arquivos if not os.path.basename(a).startswith("~$"))
            total = len(arquivos)
            if total == 0:
                self._ui("log", "Nenhum arquivo .xlsm/.xlsx encontrado na pasta base.", "aviso")
                resumo = "Nenhum arquivo encontrado"
                return

            self._ui("log", f"{total} arquivo(s) encontrado(s). Abrindo o Excel…", "info")
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3
            excel.EnableEvents = False

            sucesso = falhas = 0
            for i, caminho in enumerate(arquivos, 1):
                nome = os.path.basename(caminho)
                self._ui("progresso", (i - 1) / total, f"Processando {i} de {total}: {nome}")
                resultado = self._processar_arquivo(excel, caminho, modelos, destino)
                if resultado == "ok":
                    sucesso += 1
                else:
                    falhas += 1
                self._ui("progresso", i / total, f"Processado {i} de {total}")

            self._ui("log", f"Concluído: {sucesso} gerado(s), {falhas} com problema.", "ok" if falhas == 0 else "aviso")
            resumo = f"Concluído: {sucesso} de {total} arquivo(s) gerado(s)"
        except Exception as e:
            self._ui("log", f"Erro fatal com o Excel: {e}", "erro")
            resumo = "Erro no processamento"
        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()
            self._ui("fim", resumo)

    def _processar_arquivo(self, excel, caminho, modelos, destino):
        """Devolve 'ok' ou 'falha'. Mesma logica do programa original."""
        nome = os.path.basename(caminho)
        chave = identificar_modelo(nome)

        if not chave:
            self._ui("log", f"Ignorado: {nome} (modelo não identificado pelo nome)", "erro")
            return "falha"
        if chave not in MODELOS_VALIDOS:
            self._ui("log", f"Ignorado: {nome} (o modelo {chave} não existe)", "erro")
            return "falha"
        if chave not in modelos:
            self._ui("log", f"Ignorado: {nome} (modelo {chave}.xls não foi carregado)", "erro")
            return "falha"

        self._ui("log", f"{nome}  →  modelo {chave}")
        wb_origem = wb_modelo = None
        try:
            wb_origem = excel.Workbooks.Open(os.path.abspath(caminho), ReadOnly=True, UpdateLinks=False)
            ws_origem = wb_origem.Sheets(1)
            wb_modelo = excel.Workbooks.Open(os.path.abspath(modelos[chave]), UpdateLinks=False)
            ws_modelo = wb_modelo.Sheets("Formulario")

            linha = 6
            processadas = 0
            while True:
                verificador = ws_origem.Cells(linha, 2).Value
                if verificador is None or str(verificador).strip() == "":
                    break
                for coluna in range(2, 41):
                    destino_cel = ws_modelo.Cells(linha, coluna)
                    if not destino_cel.Locked:
                        valor = sanitizar(ws_origem.Cells(linha, coluna).Value)
                        if valor is not None:
                            destino_cel.Value = valor
                linha += 1
                processadas += 1

            if processadas == 0:
                self._ui("log", f"Aviso: nenhuma linha com dados em {nome}", "aviso")
                return "falha"

            novo_nome = os.path.splitext(nome)[0] + ".xls"
            wb_modelo.SaveAs(os.path.abspath(os.path.join(destino, novo_nome)), FileFormat=56)
            self._ui("log", f"   salvo: {novo_nome} ({processadas} linha(s))", "ok")
            return "ok"
        except Exception as e:
            self._ui("log", f"Falha em {nome}: {e}", "erro")
            return "falha"
        finally:
            for wb in (wb_origem, wb_modelo):
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass


if __name__ == "__main__":
    app = AppTotumseg()
    app.mainloop()
