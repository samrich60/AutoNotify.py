# 🤖 AutoNotify

AutoNotify é um programa automatizado que envia mensagens via WhatsApp com base em uma planilha do Excel. Ele foi desenvolvido para facilitar o envio de lembretes automáticos, como avisos de vencimento, para uma lista de contatos.

**⚠️ O PROGRAMA AINDA ESTÁ EM DESENVOLVIMENTO** ⚠️

## 📌 Funcionalidades
- Seleciona uma planilha Excel com os dados dos clientes.
- Abre o WhatsApp Web ou WhatsApp Desktop para envio das mensagens.
- Envia mensagens automáticas personalizadas.
- Permite interromper a execução a qualquer momento.
- Registra erros de envio em um arquivo CSV.

## 🛠 Tecnologias Utilizadas
- Python
- Tkinter (Interface Gráfica)
- OpenPyXL (Manipulação de planilhas Excel)
- Webbrowser (Abertura de links no navegador)
- PyAutoGUI (Automatiza interações com a interface)
- Datetime (Manipulação de datas)

## 🚀 Como Usar
1. Certifique-se de que o WhatsApp Web está logado no navegador.
2. Execute o programa Python.
3. Clique em "Selecionar Planilha" e escolha um arquivo `.xlsx` com os dados.
4. O programa iniciará o envio das mensagens automaticamente.
5. Caso precise interromper o processo, clique em "Parar Execução".
6. Em caso de erro, os contatos afetados serão registrados no arquivo `error.csv`.

## 📋 Formato da Planilha
A planilha deve conter os seguintes campos na primeira linha:
- **Nome** (Coluna A)
- **Telefone** (Coluna B)
- **Data de Vencimento** (Coluna C)

Exemplo:
| Nome      | Telefone     | Vencimento  |
|-----------|-------------|-------------|
| Cliente 1 | 55119999999 | 2025-03-10  |
| Cliente 2 | 55119888888 | 2025-03-11  |

## ⚠️ Observações Importantes durante o uso do programa
- **Não mexa o mouse** durante o envio das mensagens, pois o PyAutoGUI controla a interação.
