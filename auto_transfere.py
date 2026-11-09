
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from openpyxl import load_workbook
import pyautogui
import time
import threading


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
DELAY_ENTRE_ITENS = 0.4

executando = False
arquivo_excel = None

pyautogui.PAUSE = 0.2
pyautogui.FAILSAFE = True


# ==============================
# FUNÇÕES AUXILIARES
# ==============================

def registrar_log(mensagem):

    log_text.after(
        0,
        lambda: (
            log_text.insert(tk.END, mensagem + "\n"),
            log_text.see(tk.END)
        )
    )


def normalizar(valor):

    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor)).strip()

    return str(valor).strip()


def digitar(texto):

    for char in str(texto):

        if not executando:
            return

        pyautogui.write(char)


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

        lbl_arquivo.config(
            text=arquivo_excel
        )


# ==============================
# REALIZAR TRANSFERÊNCIA
# ==============================

def realizar_transferencia(codigo, origem, destino, quantidade):

    if not executando:
        return False


    digitar(codigo)

    pyautogui.press("right")
    pyautogui.press("right")

    digitar(origem)

    pyautogui.press("right")
    pyautogui.press("right")
    pyautogui.press("right")
    pyautogui.press("right")

    digitar(destino)

    pyautogui.press("right")
    pyautogui.press("right")
    pyautogui.press("right")
    pyautogui.press("right")
    pyautogui.press("right")
    pyautogui.press("right")

    digitar(quantidade)

    pyautogui.press("down")
    pyautogui.press("down")

    return True


# ==============================
# EXECUTAR AUTOMAÇÃO
# ==============================

def executar_automacao():

    global executando

    workbook = None

    try:

        executando = True

        if not arquivo_excel:
            raise Exception("Selecione uma planilha primeiro.")

        workbook = load_workbook(
            arquivo_excel,
            data_only=True
        )

        if ABA_PLANILHA not in workbook.sheetnames:

            raise Exception(
                f"A aba '{ABA_PLANILHA}' não foi encontrada."
            )

        planilha = workbook[ABA_PLANILHA]

        total_linhas = planilha.max_row

        registrar_log(
            f"Planilha carregada: {ABA_PLANILHA}"
        )

        registrar_log(
            f"Total de linhas: {total_linhas - LINHA_INICIAL + 1}"
        )

        registrar_log(
            f"Você tem {TEMPO_PREPARO} segundos para se posicionar no sistema."
        )

        # CONTAGEM REGRESSIVA

        for t in range(TEMPO_PREPARO, 0, -1):

            if not executando:
                break

            registrar_log(f"Iniciando em {t} segundos...")

            time.sleep(1)

        if not executando:

            registrar_log("Automação interrompida.")
            return

        registrar_log("--------------------------------")
        registrar_log("INICIANDO PROCESSAMENTO")
        registrar_log("--------------------------------")

        realizadas = 0
        ignoradas = 0

        pyautogui.doubleClick(x=50, y=280)

        time.sleep(0.5)

        # PERCORRER PLANILHA

        for linha in range(LINHA_INICIAL, total_linhas + 1):

            if not executando:
                break

            codigo = planilha.cell(
                linha, COLUNA_CODIGO
            ).value

            origem = planilha.cell(
                linha, COLUNA_ORIGEM
            ).value

            destino = planilha.cell(
                linha, COLUNA_DESTINO
            ).value

            quantidade = planilha.cell(
                linha, COLUNA_QUANTIDADE
            ).value

            # IGNORAR LINHAS VAZIAS

            if all(
                valor is None
                for valor in [codigo, origem, destino, quantidade]
            ):

                ignoradas += 1
                continue

            # VALIDAR CAMPOS OBRIGATÓRIOS

            if None in [codigo, origem, destino, quantidade]:

                ignoradas += 1

                registrar_log(
                    f"Linha {linha}: dados incompletos. Ignorada."
                )

                continue

            # NORMALIZAR DADOS

            codigo = normalizar(codigo)
            origem = normalizar(origem)
            destino = normalizar(destino)
            quantidade = normalizar(quantidade)

            # IGNORAR MESMO ARMAZÉM

            if origem == destino:

                ignoradas += 1

                registrar_log(
                    f"Linha {linha}: origem e destino iguais. Ignorada."
                )

                continue

            # EXIBIR DADOS

            registrar_log(
                f"Linha {linha} | Código: {codigo} | "
                f"Origem: {origem} | Destino: {destino} | "
                f"Quantidade: {quantidade}"
            )

            # REALIZAR TRANSFERÊNCIA

            if realizar_transferencia(
                codigo,
                origem,
                destino,
                quantidade
            ):

                realizadas += 1

            else:
                break

            time.sleep(DELAY_ENTRE_ITENS)

        # RESULTADO FINAL

        registrar_log("--------------------------------")

        if executando:

            registrar_log("PROCESSAMENTO FINALIZADO!")

        else:

            registrar_log("AUTOMAÇÃO INTERROMPIDA!")

        registrar_log(
            f"Operações enviadas: {realizadas}"
        )

        registrar_log(
            f"Linhas ignoradas: {ignoradas}"
        )

    except pyautogui.FailSafeException:

        executando = False

        registrar_log(
            "AUTOMAÇÃO INTERROMPIDA PELO FAILSAFE!"
        )

    except Exception as erro:

        mensagem = str(erro)

        registrar_log(
            f"ERRO: {mensagem}"
        )

        janela.after(
            0,
            lambda msg=mensagem: messagebox.showerror(
                "Erro",
                msg
            )
        )

    finally:

        if workbook:
            workbook.close()

        executando = False

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

        messagebox.showwarning(
            "Atenção",
            "Selecione uma planilha primeiro."
        )

        return

    botao_iniciar.config(
        state=tk.DISABLED
    )

    threading.Thread(
        target=executar_automacao,
        daemon=True
    ).start()


# ==============================
# PARAR AUTOMAÇÃO
# ==============================

def parar_automacao():

    global executando

    executando = False

    registrar_log(
        "Solicitando interrupção..."
    )


# ==============================
# INTERFACE GRÁFICA
# ==============================

janela = tk.Tk()

janela.title("Automação de Transferências")
janela.geometry("650x450")

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


log_text = scrolledtext.ScrolledText(
    janela,
    width=75,
    height=15
)

log_text.pack(
    padx=10,
    pady=10
)


janela.mainloop()