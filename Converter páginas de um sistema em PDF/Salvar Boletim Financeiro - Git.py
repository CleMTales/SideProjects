#Programa para salvar (em HTML) um conjunto de páginas presentes em uma planilha
#Excel
import pandas as pd
import requests

#Iniciando uma sessão no site do sistema
url = #'http://.../source/checar.asp?&senha=321564&ok.x=0&ok.y=0'
s = requests.Session()
r = s.get(url)


#Definir número de linhas
totalRows = 5409
rowsSkipped = 0#Número de linhas para pular
numberRows = totalRows - rowsSkipped
#Definir colunas que serão lidas
lerColunas = ['Aluno','Curso','Link', 'Matricula']

#Especificar endereço da planilha Excel
excel = r'C:\Users\Tales\Desktop\HTML to PDF\V1 - Financeiro\Tabela Excel.xlsx'

#Ler Excel para dataframe
df = pd.read_excel(excel, sheet_name = 2, usecols= lerColunas, nrows= numberRows, skiprows=[*range(1, rowsSkipped-1)])


#df = df[['Aluno', 'Curso', 'Link', 'Matricula']]
names = df['Aluno']
cursos = df['Curso']
link = df['Link']
matriculas = df['Matricula']
i = 0

for x in names:
    nome = x
    matricula = int(matriculas[i])
    matricula = str(matricula)
    curso = str(cursos[i])
    f = s.get(link[i])
    #Definir nome do arquivo de saída
    filename = nome + "." + matricula + " - " + curso + ".html"
    arquivo = open(filename, "w")
    arquivo.write(f.text)
    arquivo.close()
    #Imprimir nome do arquivo para checar progresso
    print (filename)
    i = i+1
