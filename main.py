# Importe a biblioteca pandas para trabalhar com arquivos Excel.
import pandas as pd
import os
from tkinter import filedialog
from tkinter import simpledialog, Tk
import tkinter.font as font
 
def buscar_arquivo():
    # Abre uma janela de diálogo para permitir que o usuário escolha um arquivo Excel (.xlsx). 
    file_path = filedialog.askopenfilename()
    if not file_path.endswith('.xlsx'):
        raise ValueError("Por favor, selecione um arquivo Excel (.xlsx)")
    return file_path

# Leia o arquivo Excel de entrada e armazene-o em um DataFrame (tabela)
df = pd.read_excel(buscar_arquivo())

# Criar uma instância do Tkinter
root = Tk()

# Criar uma fonte customizada com tamanho maior
custom_font = font.nametofont("TkDefaultFont")
custom_font.configure(size=20)

# Aplicar a fonte customizada à caixa de diálogo
root.option_add('*Dialog.msg.font', custom_font)

# Criar novamente a caixa de diálogo com o tamanho de fonte ajustado
qtd_linhas = simpledialog.askinteger("dato® - Quebra", "Digite a quantidade de linhas para a quebra:", initialvalue=0)

# Fechar a instância do Tkinter
root.destroy()

# Divida os dados em pedaços de acordo com a quantidade de linhas definida pelo usuário
pedacos = [df[i:i+qtd_linhas] for i in range(0, len(df), qtd_linhas)]

# Solicitar ao usuário que escolha o nome e diretório de saída uma única vez
arquivo_base = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])

# Iterar sobre cada pedaço definido pela quantidade de linhas
for i, pedaco in enumerate(pedacos):
    # Acrescentar o índice como sufixo ao nome do arquivo
    arquivo_saida = f"{arquivo_base}_{i+1}.xlsx"
    pedaco.to_excel(arquivo_saida, index=False)  # O parâmetro `index=False` evita que seja criada uma coluna para índices no arquivo de saída

print(f"Arquivos Excel criados com sucesso.")