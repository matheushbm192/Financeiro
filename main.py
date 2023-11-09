import openpyxl
import tkinter as tk
from tkinter import Entry, Button, Label, StringVar

arquivo = openpyxl.load_workbook('BaseDeDados.xlsx')
planilha = arquivo['base de dados']  # Abre a planilha Excel
# Função para adicionar dados à planilha Excel
def adicionar_dados():
    data = data_var.get()  # Obtem a data da entrada do usuário
    tipo = int(tipo_var.get())  # Obtém a coluna da entrada do usuário
    valor = float(valor_var.get())  # Obtém o valor da entrada do usuário



    linha = 2
    while planilha[f'A{linha}'].value is not None:
        linha += 1

    # Preenche as células com os dados
    planilha[f'A{linha}'].value = data
    planilha[f'B{linha}'].value = tipo_valor(tipo, valor)
    planilha[f'C{linha}'].value = tipo_nome(tipo)
    planilha[f'D2'].value = soma(linha)
    arquivo.save('BaseDeDados.xlsx')  # Salva as alterações na planilha

    # Limpa os campos de entrada após o envio
    data_var.set('')
    tipo_var.set('')
    valor_var.set('')

def soma(linha):
    acumulate = 0.0
    for lin in range(2, linha + 1):
        acumulate += planilha[f'B{lin}'].value
    print(acumulate)
    return acumulate

# Função para mapear o número da coluna para a letra da coluna
def tipo_valor(tipo, valor):
    match tipo:
        case 1:
            return valor
        case 2:
            return valor * (-1)
        case 3:
            return valor * (-1)
        case 4:
            return valor * (-1)
def tipo_nome(tipo):
    match tipo:
        case 1:
            return 'Receita'
        case 2:
            return 'Fixa'
        case 3:
            return 'Variavel'
        case 4:
            return 'Arbritaria'
def verificar_campos():
    if data_var.get() and tipo_var.get() and valor_var.get():
        adicionar_dados()
# Cria a interface gráfica com o Tkinter
root = tk.Tk()
root.title("Inserir Dados no Excel")

data_var = StringVar()  # Variável para a entrada da data
tipo_var = StringVar()  # Variável para a entrada da coluna
valor_var = StringVar()  # Variável para a entrada do valor

# Elementos da interface gráfica
label_data = Label(root, text="Digite uma data no formato DD/MM/AAAA:")
label_data.pack()
entry_data = Entry(root, textvariable=data_var)
entry_data.pack()

label_coluna = Label(root, text="1.Receita | 2.Fixas | 3.Variaveis | 4.Arbitrarias:")
label_coluna.pack()
entry_coluna = Entry(root, textvariable=tipo_var)
entry_coluna.pack()

label_valor = Label(root, text="Digite o valor:")
label_valor.pack()
entry_valor = Entry(root, textvariable=valor_var)
entry_valor.pack()

#verifica se o campo esta preenchido
entry_data.bind('<Return>', verificar_campos)
entry_coluna.bind('<Return>', verificar_campos)
entry_valor.bind('<Return>', verificar_campos)

button_adicionar = Button(root, text="Adicionar Dados", command=verificar_campos)
button_adicionar.pack()

root.mainloop()  # Inicia a interface gráfica