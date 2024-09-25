#Decidi usar dict comprehension

import os
from tkinter.filedialog import askdirectory

# Abre um diálogo para selecionar uma pasta
caminho = askdirectory(title="Selecione uma pasta")

# Lista todos os arquivos na pasta selecionada
lista_arquivos = os.listdir(caminho)

# Dicionário que mapeia extensões de arquivos para pastas
locais = {
    "imagens": [".png", ".jpg"],
    "planilhas": [".xlsx"],
    "pdf": [".pdf"],
    "csv": [".csv"]
}

# Cria um dicionário que mapeia arquivos para seus diretórios de destino
arquivo_para_pasta = {
    arquivo: pasta
    for arquivo in lista_arquivos
    for pasta, extensoes in locais.items()
    if os.path.splitext(f"{caminho}/{arquivo}")[1].lower() in extensoes
}

# Move cada arquivo para o diretório correspondente
for arquivo, pasta in arquivo_para_pasta.items():
    destino = f"{caminho}/{pasta}"
    if not os.path.exists(destino):  # Verifica se a pasta destino existe
        os.mkdir(destino)  # Cria a pasta se não existir
    os.rename(f"{caminho}/{arquivo}", f"{destino}/{arquivo}")  # Move o arquivo para a pasta correta

#Dict Comprehension:

#O dict comprehension cria um dicionário arquivo_para_pasta onde as chaves são os nomes dos arquivos e os valores são as pastas de destino, baseados na extensão do arquivo.
#Exemplo de mapeamento: Se arquivo é "foto.png", e "foto.png" termina em ".png", arquivo_para_pasta["foto.png"] será "imagens".
#For Loop:

#Depois de construir o dicionário, o loop final move cada arquivo para a pasta correspondente.