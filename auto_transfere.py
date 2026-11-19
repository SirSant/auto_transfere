import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from openpyxl import load_workbook
import pyautogui
import time
import threading
import keyboard
import pytesseract
from PIL import ImageOps, ImageEnhance
import re
import json
import os
import sys
import tempfile
import subprocess
import urllib.request
import urllib.error
import shutil
import hashlib


# ==============================
# CONFIGURAÇÕES
# ==============================

ABA_PLANILHA = "Listagem do Browse"

COLUNA_CODIGO = 2
COLUNA_ORIGEM = 3
COLUNA_DESTINO = 4
COLUNA_QUANTIDADE = 9

LINHA_INICIAL = 2

TEMPO_PREPARO = 3
DELAY_ENTRE_ITENS = 0.3

executando = False
pausado = False
linha_retomada = None
arquivo_excel = None

linha_atual = None
codigo_atual = ""
pausa_por_erro = False
texto_erro_detectado = ""
IDIOMA_OCR = "por"
if getattr(sys, "frozen", False):
    DIRETORIO_TESSERACT = os.path.join(sys._MEIPASS, "tesseract")
else:
    DIRETORIO_PROJETO = os.path.dirname(os.path.abspath(__file__))
    DIRETORIO_TESSERACT = os.path.join(DIRETORIO_PROJETO, "runtime", "tesseract")
    if not os.path.isfile(os.path.join(DIRETORIO_TESSERACT, "tesseract.exe")):
        DIRETORIO_TESSERACT = r"C:\Program Files\Tesseract-OCR"
TESSERACT_CMD = os.path.join(DIRETORIO_TESSERACT, "tesseract.exe")
TESSDATA_DIR = os.path.join(DIRETORIO_TESSERACT, "tessdata")
os.environ["TESSDATA_PREFIX"] = DIRETORIO_TESSERACT
pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

# Distribuição pública e atualização automática via GitHub Releases
VERSAO_ATUAL = "1.0.0"
GITHUB_REPOSITORIO = "SirSant/auto_transfere"
NOME_INSTALADOR = "AutoTransfere_Setup.exe"
URL_RELEASES_API = f"https://api.github.com/repos/{GITHUB_REPOSITORIO}/releases/latest"

pyautogui.PAUSE = 0.2
pyautogui.FAILSAFE = True


# ==============================
# FUNÇÕES AUXILIARES
# ==============================

def registrar_log(mensagem):
    """Atualiza o log pela thread principal do Tkinter."""
    def inserir(msg=mensagem):
        try:
            log_text.configure(state="normal")
            log_text.insert(tk.END, msg + "\n")
            log_text.see(tk.END)
            log_text.configure(state="disabled")
        except tk.TclError:
            pass
    log_text.after(0, inserir)


def normalizar(valor):
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)

    texto = str(valor).strip()

    # Armazéns de um dígito ficam no formato 01, 02...;
    # valores que já têm dois ou mais caracteres não ganham zeros extras.
    if texto.isdigit():
        texto = texto.zfill(2)

    return texto


def atualizar_campo_linha(linha):
    def atualizar():
        campo_linha.delete(0, tk.END)
        campo_linha.insert(0, str(linha))
    janela.after(0, atualizar)


def ler_erro_tela():
    """Lê a janela de aviso do TOTVS por OCR e pausa se encontrar erro."""
    global pausado, pausa_por_erro, linha_retomada, texto_erro_detectado

    if not executando or pausa_por_erro:
        return pausa_por_erro

    try:
        imagem = pyautogui.screenshot()
        imagem = ImageOps.grayscale(imagem)
        imagem = ImageEnhance.Contrast(imagem).enhance(1.8)
        config_ocr = f'--psm 6 --tessdata-dir "{TESSDATA_DIR}"'
        texto = pytesseract.image_to_string(imagem, lang=IDIOMA_OCR, config=config_ocr)
        baixo = texto.lower()

        marcadores = (
            "help:", "problema:", "solução:", "solucao:",
            "não existe registro", "nao existe registro",
            "registro relacionado"
        )
        if any(m in baixo for m in marcadores):
            linhas = [" ".join(x.split()) for x in texto.splitlines() if x.strip()]
            inicio = next((i for i, x in enumerate(linhas)
                           if "problema" in x.lower()), None)
            fim = next((i for i, x in enumerate(linhas)
                        if "solução" in x.lower() or "solucao" in x.lower()), None)

            if inicio is not None:
                trecho = linhas[inicio:fim] if fim is not None and fim > inicio else linhas[inicio:inicio+4]
                texto_erro_detectado = " ".join(trecho)
            else:
                texto_erro_detectado = " ".join(linhas[:8])

            texto_erro_detectado = texto_erro_detectado[:500]
            pausa_por_erro = True
            pausado = True
            linha_retomada = linha_atual or LINHA_INICIAL
            atualizar_campo_linha(linha_retomada)
            registrar_log(
                f'AUTOMAÇÃO PAUSADA NO ITEM {codigo_atual}, LINHA {linha_retomada}, '
                f'devido ao erro "{texto_erro_detectado}". '
                "Solucione antes de retomar a automação."
            )
            return True

    except Exception as erro:
        registrar_log(f"AVISO: falha na leitura OCR: {erro}")
    return False


def digitar(texto):
    """Digita o texto rapidamente. Retorna False se a automação foi encerrada."""
    texto = str(texto)

    for char in texto:
        if not executando:
            return False
        pyautogui.write(char)

    return executando


def pausar_automacao():
    """Solicita pausa; ela será aplicada após a transferência atual."""
    global pausado

    if executando and not pausado:
        pausado = True
        registrar_log("PAUSA SOLICITADA. A transferência atual será concluída antes da pausa.")


def retomar_automacao():
    global pausado, linha_retomada, pausa_por_erro

    if not executando:
        return

    if not pausado:
        registrar_log("A automação não está pausada.")
        return

    try:
        linha = int(campo_linha.get())

        if linha < LINHA_INICIAL:
            raise ValueError

        linha_retomada = linha
        registrar_log("RETOMANDO EM 3 SEGUNDOS...")

        for segundo in range(3, 0, -1):
            registrar_log(str(segundo))
            time.sleep(1)

        pausa_por_erro = False
        pausado = False
        registrar_log(f"RETOMANDO A PARTIR DA LINHA {linha}")

    except ValueError:
        messagebox.showwarning(
            "Atenção",
            f"Informe um número de linha válido (a partir de {LINHA_INICIAL})."
        )


# ==============================
# SELECIONAR PLANILHA
# ==============================

def selecionar_arquivo():
    global arquivo_excel

    arquivo_excel = filedialog.askopenfilename(
        title="Selecionar planilha",
        filetypes=[("Planilhas Excel", "*.xlsx")]
    )

    if arquivo_excel:
        lbl_arquivo.config(text=arquivo_excel)


# ==============================
# REALIZAR TRANSFERÊNCIA
# ==============================

def realizar_transferencia(codigo, origem, destino, quantidade):
    if not executando or (pausado and pausa_por_erro):
        return False

    def checar():
        return executando and not ler_erro_tela() and not (pausado and pausa_por_erro)

    if not digitar(codigo):
        return False
    if len(codigo) != 15:
        pyautogui.press("tab")
        if not checar(): return False

    time.sleep(0.5)
    if not checar(): return False
    pyautogui.press("tab")
    if not checar(): return False

    time.sleep(0.5)
    pyautogui.press("right", presses=2)
    if not checar(): return False

    time.sleep(0.5)
    if not digitar(origem): return False
    pyautogui.press("right", presses=4)
    if not checar(): return False

    if not digitar(destino): return False
    pyautogui.press("right", presses=6)
    if not checar(): return False

    if not digitar(quantidade): return False
    pyautogui.press("down", presses=2)
    time.sleep(0.4)
    if ler_erro_tela(): return False

    return executando and not (pausado and pausa_por_erro)


def executar_automacao():
    global executando, pausado, linha_retomada, linha_atual, codigo_atual, pausa_por_erro, texto_erro_detectado

    workbook = None

    try:
        executando = True
        pausado = False
        linha_retomada = None
        linha_atual = None
        codigo_atual = ""
        pausa_por_erro = False
        texto_erro_detectado = ""

        if not arquivo_excel:
            raise Exception("Selecione uma planilha primeiro.")

        workbook = load_workbook(arquivo_excel, data_only=True)

        if ABA_PLANILHA not in workbook.sheetnames:
            raise Exception(f"A aba '{ABA_PLANILHA}' não foi encontrada.")

        planilha = workbook[ABA_PLANILHA]
        total_linhas = planilha.max_row

        registrar_log(f"Planilha carregada: {ABA_PLANILHA}")
        registrar_log(f"Total de linhas: {max(0, total_linhas - LINHA_INICIAL + 1)}")
        registrar_log(f"Você tem {TEMPO_PREPARO} segundos para se posicionar no sistema.")

        # Contagem regressiva
        for t in range(TEMPO_PREPARO, 0, -1):
            if not executando:
                break
            registrar_log(f"Iniciando em {t} segundos...")
            time.sleep(1)

        if not executando:
            registrar_log("Automação encerrada antes do início.")
            return

        registrar_log("--------------------------------")
        registrar_log("INICIANDO PROCESSAMENTO")
        registrar_log("--------------------------------")

        realizadas = 0
        ignoradas = 0

        # Posicionar no primeiro campo do sistema apenas uma vez.
        pyautogui.doubleClick(x=70, y=410)
        time.sleep(0.5)

        linha = LINHA_INICIAL

        while linha <= total_linhas:
            if not executando:
                break

            # Aguarda aqui, entre uma transferência e outra.
            while pausado and executando:
                time.sleep(0.1)

            if not executando:
                break

            # Se o usuário informou outra linha ao retomar, usa essa linha.
            if linha_retomada is not None:
                linha = linha_retomada
                linha_retomada = None
                registrar_log(f"Continuando pela linha {linha}")

            if linha > total_linhas:
                break

            # Leitura dos dados
            linha_atual = linha
            codigo = planilha.cell(linha, COLUNA_CODIGO).value
            origem = planilha.cell(linha, COLUNA_ORIGEM).value
            destino = planilha.cell(linha, COLUNA_DESTINO).value
            quantidade = planilha.cell(linha, COLUNA_QUANTIDADE).value

            # Ignorar linhas completamente vazias
            valores = [codigo, origem, destino, quantidade]
            if all(valor is None for valor in valores):
                ignoradas += 1
                linha += 1
                continue

            # Ignorar linhas com algum campo obrigatório ausente
            if any(valor is None for valor in valores):
                ignoradas += 1
                registrar_log(f"Linha {linha}: dados incompletos. Ignorada.")
                linha += 1
                continue

            # Normalizar os dados
            codigo = normalizar(codigo)
            codigo_atual = codigo
            origem = normalizar(origem)
            destino = normalizar(destino)
            quantidade = normalizar(quantidade)

            # Ignorar transferência entre o mesmo armazém
            if origem == destino:
                ignoradas += 1
                registrar_log(f"Linha {linha}: origem e destino iguais. Ignorada.")
                linha += 1
                continue

            registrar_log(
                f"Linha {linha} | Código: {codigo} | Origem: {origem} | "
                f"Destino: {destino} | Quantidade: {quantidade}"
            )

            sucesso = realizar_transferencia(codigo, origem, destino, quantidade)

            if not sucesso:
                if pausa_por_erro:
                    linha_retomada = linha
                    atualizar_campo_linha(linha)
                    while pausado and executando:
                        time.sleep(0.1)
                    if not executando:
                        break
                    if linha_retomada is not None:
                        linha = linha_retomada
                        linha_retomada = None
                    continue
                break

            realizadas += 1
            linha += 1
            atualizar_campo_linha(linha)

            if pausado and not pausa_por_erro:
                linha_retomada = linha
                registrar_log(
                    f"Linha anterior concluída com sucesso. Próxima linha para retomada: {linha}."
                )
                while pausado and executando:
                    time.sleep(0.1)
                if not executando:
                    break
                if linha_retomada is not None:
                    linha = linha_retomada
                    linha_retomada = None

            time.sleep(DELAY_ENTRE_ITENS)

        registrar_log("--------------------------------")

        if executando:
            registrar_log("PROCESSAMENTO FINALIZADO!")
        else:
            registrar_log("AUTOMAÇÃO ENCERRADA!")

        registrar_log(f"Operações enviadas: {realizadas}")
        registrar_log(f"Linhas ignoradas: {ignoradas}")

    except pyautogui.FailSafeException:
        executando = False
        pausado = False
        registrar_log("AUTOMAÇÃO INTERROMPIDA PELO FAILSAFE!")

    except Exception as erro:
        mensagem = str(erro)
        registrar_log(f"ERRO: {mensagem}")
        janela.after(
            0,
            lambda msg=mensagem: messagebox.showerror("Erro", msg)
        )

    finally:
        if workbook:
            workbook.close()

        if not pausado:
            executando = False
            linha_atual = None
            codigo_atual = ""

        botao_iniciar.after(
            0,
            lambda: botao_iniciar.config(state=tk.NORMAL)
        )


# ==============================
# INICIAR AUTOMAÇÃO
# ==============================

def iniciar_automacao():
    global executando

    if executando:
        return

    if not arquivo_excel:
        messagebox.showwarning("Atenção", "Selecione uma planilha primeiro.")
        return

    botao_iniciar.config(state=tk.DISABLED)
    botao_retomar.config(state=tk.NORMAL)

    threading.Thread(
        target=executar_automacao,
        daemon=True
    ).start()


# ==============================
# PARAR AUTOMAÇÃO DEFINITIVAMENTE
# ==============================

def parar_automacao():
    global executando, pausado

    executando = False
    pausado = False
    registrar_log("SOLICITAÇÃO DE ENCERRAMENTO enviada.")



# ==============================
# ATUALIZAÇÃO AUTOMÁTICA
# ==============================

def versao_numerica(valor):
    try:
        partes = [int(p) for p in valor.lstrip("vV").split(".")]
        return tuple((partes + [0, 0, 0])[:3])
    except (ValueError, AttributeError):
        return (0, 0, 0)


def verificar_atualizacao():
    """Consulta a última release do GitHub sem bloquear a interface."""
    try:
        requisicao = urllib.request.Request(
            URL_RELEASES_API,
            headers={"Accept": "application/vnd.github+json", "User-Agent": "AutoTransfere-Updater"}
        )
        with urllib.request.urlopen(requisicao, timeout=10) as resposta:
            release = json.loads(resposta.read().decode("utf-8"))

        versao_remota = release.get("tag_name", "").lstrip("vV")
        if not versao_remota or versao_numerica(versao_remota) <= versao_numerica(VERSAO_ATUAL):
            return

        asset = next(
            (a for a in release.get("assets", []) if a.get("name") == NOME_INSTALADOR),
            None
        )
        if not asset:
            registrar_log(f"Versão {versao_remota} encontrada, mas o instalador não está anexado.")
            return

        def oferecer_atualizacao():
            if not messagebox.askyesno(
                "Atualização disponível",
                f"Versão instalada: {VERSAO_ATUAL}\nNova versão: {versao_remota}\n\nDeseja atualizar agora?"
            ):
                registrar_log("Atualização adiada pelo usuário.")
                return
            try:
                registrar_log("Baixando atualização...")
                caminho = os.path.join(tempfile.gettempdir(), NOME_INSTALADOR)
                download = urllib.request.Request(
                    asset["browser_download_url"],
                    headers={"User-Agent": "AutoTransfere-Updater"}
                )
                with urllib.request.urlopen(download, timeout=180) as origem:
                    with open(caminho, "wb") as destino:
                        shutil.copyfileobj(origem, destino)

                digest = asset.get("digest")
                if digest and digest.startswith("sha256:"):
                    hash_calculado = hashlib.sha256()
                    with open(caminho, "rb") as arquivo:
                        for bloco in iter(lambda: arquivo.read(1024 * 1024), b""):
                            hash_calculado.update(bloco)
                    if hash_calculado.hexdigest().lower() != digest.split(":", 1)[1].lower():
                        os.remove(caminho)
                        raise ValueError("A verificação de integridade do instalador falhou.")

                messagebox.showinfo(
                    "Atualização pronta",
                    "O instalador será aberto para atualizar o AutoTransfere. O programa será fechado."
                )
                subprocess.Popen([
                    caminho, "/VERYSILENT", "/SUPPRESSMSGBOXES",
                    "/CLOSEAPPLICATIONS", "/NORESTART"
                ])
                ao_fechar()
            except Exception as erro:
                messagebox.showerror("Falha na atualização", f"Não foi possível atualizar:\n{erro}")

        janela.after(0, oferecer_atualizacao)
    except urllib.error.HTTPError as erro:
        if erro.code != 404:
            registrar_log(f"Não foi possível verificar atualizações: HTTP {erro.code}")
    except Exception as erro:
        registrar_log(f"Não foi possível verificar atualizações: {erro}")


# ==============================
# INTERFACE GRÁFICA - TEMA MODERNO
# ==============================

from tkinter import ttk

janela = tk.Tk()
janela.title("Winner Motors | Automação de Transferências")
janela.geometry("900x720")
janela.minsize(820, 650)
janela.configure(bg="#111827")

# Paleta visual
BG = "#111827"
PANEL = "#1f2937"
PANEL_ALT = "#273449"
TEXT = "#f9fafb"
MUTED = "#9ca3af"
ACCENT = "#0d9488"
ACCENT_HOVER = "#0f766e"
DANGER = "#dc2626"
WARNING = "#d97706"
BLUE = "#2563eb"
BORDER = "#374151"

style = ttk.Style()
style.theme_use("clam")
style.configure("TFrame", background=BG)
style.configure("TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 10))
style.configure("Title.TLabel", background=BG, foreground=TEXT,
                font=("Segoe UI", 22, "bold"))
style.configure("Subtitle.TLabel", background=BG, foreground=MUTED,
                font=("Segoe UI", 10))
style.configure("Panel.TFrame", background=PANEL)
style.configure("PanelTitle.TLabel", background=PANEL, foreground=TEXT,
                font=("Segoe UI", 11, "bold"))
style.configure("Muted.TLabel", background=PANEL, foreground=MUTED,
                font=("Segoe UI", 9))
style.configure("Accent.TButton", background=ACCENT, foreground="white",
                font=("Segoe UI", 10, "bold"), padding=(16, 11), borderwidth=0)
style.map("Accent.TButton",
          background=[("active", ACCENT_HOVER), ("disabled", "#4b5563")])
style.configure("Danger.TButton", background=DANGER, foreground="white",
                font=("Segoe UI", 10, "bold"), padding=(16, 11), borderwidth=0)
style.map("Danger.TButton",
          background=[("active", "#b91c1c"), ("disabled", "#4b5563")])
style.configure("Warning.TButton", background=WARNING, foreground="white",
                font=("Segoe UI", 10, "bold"), padding=(14, 10), borderwidth=0)
style.map("Warning.TButton", background=[("active", "#b45309")])
style.configure("Blue.TButton", background=BLUE, foreground="white",
                font=("Segoe UI", 10, "bold"), padding=(14, 10), borderwidth=0)
style.map("Blue.TButton", background=[("active", "#1d4ed8")])
style.configure("Secondary.TButton", background=PANEL_ALT, foreground=TEXT,
                font=("Segoe UI", 9), padding=(12, 9), borderwidth=0)
style.map("Secondary.TButton", background=[("active", BORDER)])

container = ttk.Frame(janela, padding=24)
container.pack(fill="both", expand=True)

# Cabeçalho
cabecalho = ttk.Frame(container)
cabecalho.pack(fill="x", pady=(0, 18))

ttk.Label(cabecalho, text="Automação de Transferências",
          style="Title.TLabel").pack(anchor="w")
ttk.Label(
    cabecalho,
    text="Winner Motors  •  Processamento de transferências entre armazéns",
    style="Subtitle.TLabel"
).pack(anchor="w", pady=(4, 0))

# Painel da planilha
painel_arquivo = tk.Frame(container, bg=PANEL,
                          highlightthickness=1, highlightbackground=BORDER)
painel_arquivo.pack(fill="x", pady=(0, 12))

tk.Label(painel_arquivo, text="PLANILHA DE TRANSFERÊNCIAS",
         bg=PANEL, fg=TEXT, font=("Segoe UI", 10, "bold")).pack(
             anchor="w", padx=16, pady=(13, 8))

linha_arquivo = tk.Frame(painel_arquivo, bg=PANEL)
linha_arquivo.pack(fill="x", padx=16, pady=(0, 14))

btn_arquivo = ttk.Button(
    linha_arquivo, text="Selecionar arquivo Excel",
    style="Secondary.TButton", command=selecionar_arquivo
)
btn_arquivo.pack(side="left")

lbl_arquivo = tk.Label(
    linha_arquivo, text="Nenhum arquivo selecionado",
    bg=PANEL, fg=MUTED, font=("Segoe UI", 9),
    anchor="w", justify="left", wraplength=650
)
lbl_arquivo.pack(side="left", padx=12, fill="x", expand=True)

# Painel de controles
painel_controles = tk.Frame(container, bg=PANEL,
                            highlightthickness=1, highlightbackground=BORDER)
painel_controles.pack(fill="x", pady=(0, 12))

tk.Label(painel_controles, text="CONTROLE DA AUTOMAÇÃO",
         bg=PANEL, fg=TEXT, font=("Segoe UI", 10, "bold")).pack(
             anchor="w", padx=16, pady=(13, 10))

linha_controles = tk.Frame(painel_controles, bg=PANEL)
linha_controles.pack(fill="x", padx=16, pady=(0, 14))

coluna_linha = tk.Frame(linha_controles, bg=PANEL)
coluna_linha.pack(side="left", padx=(0, 22))

tk.Label(coluna_linha, text="Linha para retomar",
         bg=PANEL, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w")

campo_linha = tk.Entry(
    coluna_linha, width=12, justify="center",
    bg=PANEL_ALT, fg=TEXT, insertbackground=TEXT,
    relief="flat", font=("Segoe UI", 11), highlightthickness=1,
    highlightbackground=BORDER, highlightcolor=ACCENT
)
campo_linha.insert(0, str(LINHA_INICIAL))
campo_linha.pack(anchor="w", ipady=7, pady=(5, 0))

coluna_acoes = tk.Frame(linha_controles, bg=PANEL)
coluna_acoes.pack(side="left", fill="x", expand=True, pady=(17, 0))

botao_iniciar = ttk.Button(
    coluna_acoes, text="▶  INICIAR",
    style="Accent.TButton", command=iniciar_automacao
)
botao_iniciar.pack(side="left", padx=(0, 8))

botao_pausar = ttk.Button(
    coluna_acoes, text="Ⅱ  PAUSAR (ESC)",
    style="Warning.TButton", command=pausar_automacao
)
botao_pausar.pack(side="left", padx=8)

botao_retomar = ttk.Button(
    coluna_acoes, text="↻  RETOMAR",
    style="Blue.TButton", command=retomar_automacao,
    state=tk.DISABLED
)
botao_retomar.pack(side="left", padx=8)

botao_parar = ttk.Button(
    coluna_acoes, text="■  ENCERRAR",
    style="Danger.TButton", command=parar_automacao
)
botao_parar.pack(side="left", padx=8)

# Orientação de uso
tk.Label(
    container,
    text="Posicione o cursor no primeiro campo do sistema antes de iniciar. "
         "Pressione ESC ou use PAUSAR para interromper entre transferências.",
    bg=BG, fg=MUTED, font=("Segoe UI", 9), anchor="w",
    wraplength=850, justify="left"
).pack(fill="x", pady=(0, 12))

# Painel de log
painel_log = tk.Frame(container, bg=PANEL,
                      highlightthickness=1, highlightbackground=BORDER)
painel_log.pack(fill="both", expand=True)

cabecalho_log = tk.Frame(painel_log, bg=PANEL)
cabecalho_log.pack(fill="x", padx=16, pady=(12, 6))

tk.Label(cabecalho_log, text="LOG DE EXECUÇÃO",
         bg=PANEL, fg=TEXT, font=("Segoe UI", 10, "bold")).pack(side="left")

def limpar_log():
    log_text.configure(state="normal")
    log_text.delete("1.0", tk.END)
    log_text.configure(state="disabled")

ttk.Button(cabecalho_log, text="Limpar", style="Secondary.TButton",
           command=limpar_log).pack(side="right")

log_text = scrolledtext.ScrolledText(
    painel_log,
    bg="#0b1220",
    fg="#d1d5db",
    insertbackground=TEXT,
    selectbackground=ACCENT,
    relief="flat",
    borderwidth=0,
    font=("Consolas", 9),
    wrap="word",
    height=15,
    padx=12,
    pady=10
)
log_text.pack(fill="both", expand=True, padx=12, pady=(0, 12))
log_text.configure(state="disabled")

# ESC solicita pausa (não encerra definitivamente).
keyboard.add_hotkey("esc", pausar_automacao)

def ao_fechar():
    global executando, pausado
    executando = False
    pausado = False
    try:
        keyboard.remove_hotkey("esc")
    except Exception:
        pass
    janela.destroy()

janela.protocol("WM_DELETE_WINDOW", ao_fechar)
registrar_log(f"AutoTransfere versão {VERSAO_ATUAL}")
janela.after(1200, lambda: threading.Thread(target=verificar_atualizacao, daemon=True).start())
janela.mainloop()