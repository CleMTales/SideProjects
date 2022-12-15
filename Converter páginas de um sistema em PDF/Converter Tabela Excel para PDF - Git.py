#Programa para salvar (em PDF) um conjunto de páginas presentes em um sistema
#através de listagem feita em uma planilha Excel
import pdfkit
import pandas as pd


#Definir caminho para wkhtmltopdf.exe
path_to_wkhtmltopdf = r'E:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'

#Configurando caminha para wkhtmltopdf.exe nas configurações do PDFKIT
config = pdfkit.configuration(wkhtmltopdf=path_to_wkhtmltopdf)

#Número de linhas para converter
totalRows = 7998
#Número de linhas para pular
rowsSkipped = 1042
numberRows = totalRows - rowsSkipped

#Especificar endereço da planilha Excel
excel =r'Tabela Excel.xlsx'

#Definir colunas
lerColunas = ['Id', 'Nome', 'Link']

#Lendo Excel para o dataframe
df = pd.read_excel(excel, usecols= lerColunas, nrows= numberRows, skiprows=[*range(1, rowsSkipped-1)])


#df = df[['Id', 'Nome', 'Link']]
#Coluna do Nome
names = df['Nome']
#Coluna da ID
ids = df['Id']
#Coluna do URL
links = df['Link']
i = 0
numNome = 0
idAnt = 0
for x in names:
    nome = x
    #Evitar que python considere id como float
    id = int(ids[i])
    id = str(id)
    #Contador para arquivos de mesmo nome
    #Evita que um arquivo sobreescreva o anterior caso tenham o mesmo nome
    if (id == idAnt):
        numNome += 1
    else:
        numNome = 0
    idAnt = id
    url = links[i]
    #Definir nome do arquivo de saída
    if (numNome > 0):
        filename = nome + " - " + id + "_" + str(numNome) + ".pdf"
    else:
        filename = nome + " - " + id + ".pdf"
    pdfkit.from_url(url, output_path=filename, configuration=config)
    #Imprimir nome do arquivo para checar progresso
    print (filename)
    i = i+1
