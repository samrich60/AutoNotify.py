from tkinter import Tk, Button, Label, Toplevel
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
from datetime import datetime, timedelta
from tkinter import filedialog

executando = True
erros = []
sucessos = []


def fecha_whats_desk():
    pyautogui.hotkey('alt', 'f4')


def parar_execucao():
    global executando
    executando = False
    print("Execução interrompida pelo usuário.")


def rodar_script(arquivo_planilha):

    global executando
    executando = True

    data_atual = datetime.now().date()

    webbrowser.open('https://web.whatsapp.com/')
    sleep(20)

    webbrowser.open("whatsapp://send?phone=")
    sleep(5)

    workbook = openpyxl.load_workbook(arquivo_planilha)
    pag_cliente = workbook.active

    for linha in pag_cliente.iter_rows(min_row=2):
        if not executando:
            print("Processo interrompido.")
            break

        nome = linha[0].value
        telefone = linha[1].value
        vencimento = linha[2].value

        if isinstance(vencimento, datetime):
            vencimento = vencimento.date()

        linkp_what = f'https://wa.me/{telefone}'
        webbrowser.open(linkp_what)
        sleep(10)

        if vencimento - timedelta(days=1) == data_atual:
            msg = f'Olá {nome}, o prazo de uso do container vence amanhã ({vencimento.strftime("%d/%m/%Y")}). Favor entrar em contato com o número abaixo: 61955555555'

            try:
                linkp_what = f'https://wa.me/{telefone}?text={quote(msg)}'
                webbrowser.open(linkp_what)
                sleep(10)

                seta_pos = pyautogui.locateCenterOnScreen('seta.png')

                if seta_pos:
                    pyautogui.click(seta_pos)
                    sleep(5)
                    sucessos.append((nome, telefone, vencimento))
                else:
                    print(
                        f"Botão de envio não encontrado para {nome} ({telefone}).")
                    erros.append((nome, telefone, vencimento))

            except Exception as e:
                print(
                    f"Erro ao enviar mensagem para {nome} ({telefone}): {str(e)}")
                erros.append((nome, telefone, vencimento))
        else:
            print(f"Vencimento não é amanhã para {nome} ({telefone}).")

        sleep(2)
        pyautogui.hotkey('ctrl', 'w')
        fecha_whats_desk()
        sleep(2)

    if erros:
        print(f"Envio concluído com {len(erros)} erro(s).")
    else:
        print("Envio concluído com sucesso, sem erros.")


def selecionar_arquivo():
    arquivo_planilha = filedialog.askopenfilename(
        title="Selecione a planilha", filetypes=[("Excel Files", "*.xlsx")]
    )
    if arquivo_planilha:
        rodar_script(arquivo_planilha)
        mostrar_resultado()


def mostrar_resultado():
    resultado_window = Toplevel(root)
    resultado_window.title("Resultado da Execução")
    resultado_window.geometry("600x250")

    if erros:
        Label(resultado_window,
              text=f"Envio concluído com {len(erros)} erro(s).").pack(pady=10)
        Button(resultado_window, text="Mostrar Erros",
               command=mostrar_lista_erros).pack(pady=10)
    else:
        Label(resultado_window,
              text="Envio concluído com sucesso, sem erros.").pack(pady=10)


def mostrar_lista_erros():
    erros_window = Toplevel(root)
    erros_window.title("Lista de Erros")
    erros_window.geometry("600x400")

    Label(erros_window, text="Contatos com erro:").pack(pady=10)

    if erros:
        for erro in erros:
            nome, telefone, vencimento = erro
            Label(
                erros_window, text=f"Nome: {nome} | Telefone: {telefone} | Vencimento: {vencimento.strftime('%d/%m/%Y')}").pack(pady=5)
    else:
        Label(erros_window, text="Nenhum erro ao enviar. Todos os contatos foram enviados com sucesso.").pack(pady=5)

    Button(erros_window, text="Fechar",
           command=erros_window.destroy).pack(pady=10)


root = Tk()
root.title("AutoNotify - Início")
root.geometry("380x250")

Label(root, text="Selecione a planilha para enviar as mensagens:").pack(pady=10)
Button(root, text="Selecionar Planilha",
       command=selecionar_arquivo).pack(pady=10)
Button(root, text="Parar Execução",
       command=parar_execucao, fg="red").pack(pady=10)

root.mainloop()
