# Documentação do Programa de Quebra de Arquivos Excel

## Visão Geral

Este programa foi desenvolvido para facilitar a quebra de grandes conjuntos de dados armazenados em arquivos Excel (.xlsx) em pedaços menores. Ele usa a biblioteca Pandas para manipulação de dados e a interface gráfica do Tkinter para interagir com o usuário.

## Requisitos de Sistema

    - Python 3.x
    - Biblioteca Pandas
    - Biblioteca Tkinter

## Instalação

Certifique-se de ter o Python instalado em seu sistema. Instale as bibliotecas necessárias utilizando o seguinte comando:

    pip install pandas

# Uso
Execute o programa utilizando um interpretador Python.

Uma janela de diálogo será aberta, permitindo que você selecione um arquivo Excel (.xlsx). Certifique-se de escolher um arquivo compatível.

Insira a quantidade desejada de linhas para a quebra. Isso definirá o tamanho dos pedaços em que o arquivo será dividido.

Escolha o nome e o diretório para o arquivo de saída. O programa gerará vários arquivos, cada um representando um pedaço do arquivo original.

Aguarde até que o programa conclua o processo de quebra. Você será notificado quando os arquivos Excel forem criados com sucesso.

# Funções Principais

buscar_arquivo()
Abre uma janela de diálogo para permitir que o usuário escolha um arquivo Excel (.xlsx).
Leitura do Arquivo e Quebra
Leitura do Arquivo:

Utiliza a biblioteca Pandas para ler o arquivo Excel escolhido e armazenar os dados em um DataFrame.
Quantidade de Linhas para Quebra:

Pergunta ao usuário a quantidade desejada de linhas para a quebra do arquivo.
Quebra em Pedaços:

Divide o DataFrame em pedaços de acordo com a quantidade de linhas especificada pelo usuário.
Escolha de Nome e Diretório de Saída
Pede ao usuário para escolher o nome e o diretório de saída para os arquivos que serão gerados.
Iteração e Salvamento
Itera sobre cada pedaço do DataFrame e o salva como um novo arquivo Excel, adicionando o índice como sufixo ao nome do arquivo.
Exceções
Se o usuário escolher um arquivo que não é um arquivo Excel (.xlsx), uma exceção ValueError será levantada, informando sobre o erro.
Considerações Finais
Este programa é uma ferramenta simples e eficiente para a quebra de arquivos Excel em pedaços menores. Certifique-se de seguir as instruções apresentadas na interface gráfica para garantir o uso adequado do programa.

Nota: Este é um exemplo de documentação e pode ser necessário ajustá-lo com base nas necessidades específicas do seu projeto.