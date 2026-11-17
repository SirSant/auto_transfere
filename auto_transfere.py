import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from openpyxl import load_workbook
import pyautogui
import time
import threading
import keyboard


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

pyautogui.PAUSE = 0.2
pyautogui.FAILSAFE = True


# ==============================
# FUNÇÕES AUXILIARES
# ==============================

def registrar_log(mensagem):
    """Atualiza o log pela thread principal do Tkinter."""
    log_text.after(
        0,
        lambda msg=mensagem: (
            log_text.insert(tk.END, msg + "\n"),
            log_text.see(tk.END)
        )
    )


def normalizar(valor):
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)

    texto = str(valor).strip()

    # Armazéns de um dígito ficam no formato 01, 02...;
    # valores que já têm dois ou mais caracteres não ganham zeros extras.
    if texto.isdigit():
        texto = texto.zfill(2)

    return texto


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
    global pausado, linha_retomada

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
    """Preenche uma transferência. A pausa só ocorre entre linhas."""
    if not executando:
        return False

    # Código do produto
    if not digitar(codigo):
        return False

    # O sistema avança automaticamente para códigos com 15 caracteres.
    if len(codigo) != 15:
        pyautogui.press("tab")

    time.sleep(2)
    if not executando:
        return False

    pyautogui.press("tab")

    time.sleep(2)
    if not executando:
        return False

    pyautogui.press("right")
    pyautogui.press("right")

    time.sleep(2)
    if not executando:
        return False

    if not digitar(origem):
        return False

    pyautogui.press("right", presses=4)

    if not executando:
        return False

    if not digitar(destino):
        return False

    pyautogui.press("right", presses=6)

    if not executando:
        return False

    if not digitar(quantidade):
        return False

    pyautogui.press("down", presses=2)

    return executando


# ==============================
# EXECUTAR AUTOMAÇÃO
# ==============================

def executar_automacao():
    global executando, pausado, linha_retomada

    workbook = None

    try:
        executando = True
        pausado = False
        linha_retomada = None

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
                break

            realizadas += 1
            time.sleep(DELAY_ENTRE_ITENS)
            linha += 1

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

        executando = False
        pausado = False
        linha_retomada = None

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
# INTERFACE GRÁFICA
# ==============================

janela = tk.Tk()
janela.title("Automação de Transferências")
janela.geometry("650x600")

bg = "gray27"
janela.configure(bg=bg)


tk.Label(
    janela,
    text="AUTOMAÇÃO DE TRANSFERÊNCIAS",
    font=("Arial", 14, "bold"),
    bg=bg,
    fg="white"
).pack(pady=15)

btn_arquivo = tk.Button(
    janela,
    text="Selecionar Planilha",
    command=selecionar_arquivo,
    width=25
)
btn_arquivo.pack(pady=5)

lbl_arquivo = tk.Label(
    janela,
    text="Nenhum arquivo selecionado",
    bg=bg,
    fg="white",
    wraplength=600
)
lbl_arquivo.pack(pady=5)

botao_iniciar = tk.Button(
    janela,
    text="INICIAR AUTOMAÇÃO",
    command=iniciar_automacao,
    bg="#4CAF50",
    fg="white",
    width=25
)
botao_iniciar.pack(pady=10)

botao_parar = tk.Button(
    janela,
    text="PARAR",
    command=parar_automacao,
    bg="#d32f2f",
    fg="white",
    width=25
)
botao_parar.pack(pady=5)

# Campo para escolher a linha de retomada
tk.Label(
    janela,
    text="Linha para retomar:",
    bg=bg,
    fg="white"
).pack(pady=5)

campo_linha = tk.Entry(
    janela,
    width=15,
    justify="center"
)
campo_linha.insert(0, str(LINHA_INICIAL))
campo_linha.pack(pady=5)

botao_pausar = tk.Button(
    janela,
    text="PAUSAR (ESC)",
    command=pausar_automacao,
    bg="#FF9800",
    fg="white",
    width=25
)
botao_pausar.pack(pady=5)

botao_retomar = tk.Button(
    janela,
    text="RETOMAR",
    command=retomar_automacao,
    bg="#2196F3",
    fg="white",
    width=25
)
botao_retomar.pack(pady=5)

log_text = scrolledtext.ScrolledText(
    janela,
    width=75,
    height=15
)
log_text.pack(padx=10, pady=10)

# ESC solicita pausa (não encerra definitivamente).
keyboard.add_hotkey("esc", pausar_automacao)

# Remover o atalho global ao fechar a janela.
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
janela.mainloop()
