#Converter um conjunto de arquivos HTML em PDF
import pdfkit
import pandas as pd


#Definir caminho para wkhtmltopdf.exe
path_to_wkhtmltopdf = r'E:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'

#Apontar configuração do pdfkit para wkhtmltopdf.exe
config = pdfkit.configuration(wkhtmltopdf=path_to_wkhtmltopdf)

#Definir número de linhas
totalRows = 5409
rowsSkipped = 0#1042
numberRows = totalRows - rowsSkipped
#Definir colunas que serão lidas
lerColunas = ['Aluno','Curso', 'Matricula']

#Especificar endereço da planilha Excel
excel = r'C:\Users\Tales\Desktop\HTML to PDF\V1 - Financeiro\Tabela Excel.xlsx'

#Lendo Excel para dataframe
df = pd.read_excel(excel, sheet_name = 2, usecols= lerColunas, nrows= numberRows, skiprows=[*range(1, rowsSkipped-1)])


#df = df[['Aluno', 'Curso', 'Matricula']]
names = df['Aluno']
cursos = df['Curso']
matriculas = df['Matricula']
i = 0

for x in names:
    nome = x
    matricula = int(matriculas[i])
    matricula = str(matricula)
    curso = str(cursos[i])
    #Nome do arquivo em HTML
    filenamehtml = nome + "." + matricula + " - " + curso + ".html"
    #Nome do arquivo de saida em PDF
    filenamepdf = nome + "." + matricula + " - " + curso + ".pdf"
    pdfkit.from_file(filenamehtml, output_path=filenamepdf, configuration=config)
    #Imprimir nome do arquivo para checar progresso
    print (filenamepdf)
    i = i+1
    

