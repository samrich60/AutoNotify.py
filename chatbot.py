from tkinter import Tk, Button, Label, Toplevel
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
from datetime import datetime, timedelta
from tkinter import Tk, filedialog, Button, Label

# Variável de controle para parar
executando = True

def parar_execucao():
    global executando
    executando = False
    print("Execução interrompida pelo usuário.")

def rodar_script(arquivo_planilha):
    global executando
    executando = True  # Reseta a variável ao iniciar o processo

    data_atual = datetime.now().date()

    if not arquivo_planilha:
        print("Nenhuma planilha selecionada. Encerrando o programa.")
        return

    # Abrir WhatsApp Web
    webbrowser.open('https://web.whatsapp.com/')
    sleep(20)

    # Tenta abrir o WhatsApp Desktop
    webbrowser.open("whatsapp://send?phone=")
    sleep(5)

    # Fecha a aba do WhatsApp Web
    pyautogui.hotkey('ctrl', 'w')
    sleep(2)

    workbook = openpyxl.load_workbook(arquivo_planilha)
    pag_cliente = workbook.active

    for linha in pag_cliente.iter_rows(min_row=2):
        if not executando:  # Se o usuário parar, interrompe o loop
            print("Processo interrompido.")
            break

        nome = linha[0].value
        telefone = linha[1].value
        vencimento = linha[2].value

        # Verifica se a data é válida
        if isinstance(vencimento, datetime):
            vencimento = vencimento.date()  # Converte para formato de data

        linkp_what = f'https://wa.me/{telefone}'
        webbrowser.open(linkp_what)
        sleep(10)

        if vencimento - timedelta(days=1) == data_atual:
            msg = f'Olá {nome}, o prazo de uso do container vence amanhã ({vencimento.strftime("%d/%m/%Y")}). Favor entrar em contato com o número abaixo: 61955555555'

            try:
                # Adiciona a mensagem no link
                linkp_what = f'https://wa.me/{telefone}?text={quote(msg)}'
                webbrowser.open(linkp_what)
                sleep(10)

                # Clica no botão de enviar mensagem
                pyautogui.click(pyautogui.locateCenterOnScreen('seta.png'))
                sleep(5)

            except Exception as e:
                print(
                    f"Erro ao enviar mensagem para {nome} ({telefone}): {str(e)}")
                with open('error.csv', 'a', newline='', encoding='utf-8') as arquivo:
                    arquivo.write(f'{nome},{telefone}\n')

        # Fecha a aba do WhatsApp Web antes de passar para o próximo cliente
        pyautogui.hotkey('ctrl', 'w')
        sleep(2)

    print("Envio de mensagens concluído.")

def selecionar_arquivo():
    arquivo_planilha = filedialog.askopenfilename(
        title="Selecione a planilha", filetypes=[("Excel Files", "*.xlsx")]
    )
    if arquivo_planilha:
        rodar_script(arquivo_planilha)

# Função para exibir a tela de observação
def mostrar_observacao():
    observacao_window = Toplevel(root)  # Cria uma nova janela
    observacao_window.title("Observação")
    observacao_window.geometry("600x250")

    Label(observacao_window, text="Regras de Observação:").pack(pady=10)
    Label(observacao_window, text="1 - Não mexa o mouse durante o envio").pack(pady=5)
    Label(observacao_window, text="2 - Caso ocorra erro, entre em contato o criador do programa:\nEmail: emailficticio@gmail.com\ntel: 61955555555").pack(pady=5)

    Button(observacao_window, text="Fechar",
        command=observacao_window.destroy).pack(pady=10)

def voltar_inicial():
    root.deiconify()  # Mostra novamente a tela inicial

    # Fecha a tela de observação (caso tenha sido aberta)
    try:
        observacao_window.destroy()
    except NameError:
        pass

# Interface
root = Tk()
root.title("AutoNotify - Inicio")
root.geometry("380x250")

Label(root, text="Selecione a planilha para enviar as mensagens:").pack(pady=10)

Button(root, text="Selecionar Planilha",
    command=selecionar_arquivo).pack(pady=10)

Button(root, text="Parar Execução", command=parar_execucao,
    fg="red").pack(pady=10)  # Botão para parar o código


# Botão para mostrar a janela de observação
Button(root, text="Observação", command=mostrar_observacao).pack(pady=10)

root.mainloop()