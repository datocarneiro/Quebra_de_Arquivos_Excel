# Importe a biblioteca pandas para trabalhar com arquivos Excel.
import pandas as pd
import os
from tkinter import filedialog
from tkinter import simpledialog

def buscar_arquivo():
    file_path = filedialog.askopenfilename()
    if not file_path.endswith('.xlsx'):
        raise ValueError("Por favor, selecione um arquivo Excel (.xlsx)")
    return file_path

# Leia o arquivo Excel de entrada e armazene-o em um DataFrame (tabela)
df = pd.read_excel(buscar_arquivo())

# Solicitar ao usuário a quantidade de linhas desejada
qtd_linhas = simpledialog.askinteger("Quantidade de Linhas", "Digite a quantidade de linhas para a quebra:", initialvalue=900)

# Divida os dados em pedaços de acordo com a quantidade de linhas definida pelo usuário
pedacos = [df[i:i+qtd_linhas] for i in range(0, len(df), qtd_linhas)]


# Solicitar ao usuário que escolha o nome e diretório de saída uma única vez
arquivo_base = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])

# Iterar sobre cada pedaço
for i, pedaco in enumerate(pedacos):
    # Acrescentar o índice como sufixo ao nome do arquivo
    arquivo_saida = f"{arquivo_base}_{i+1}.xlsx"
    pedaco.to_excel(arquivo_saida, index=False)  # O parâmetro `index=False` evita que seja criada uma coluna para índices no arquivo de saída

print(f"Arquivos Excel criados com sucesso.")
