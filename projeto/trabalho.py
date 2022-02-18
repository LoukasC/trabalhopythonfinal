import os, openpyxl
from openpyxl.styles import Alignment 
import pandas as pd
import matplotlib.pyplot as plt

from fpdf import FPDF

print ("Verificando planilha")

# Criando a planilha

planilha = openpyxl.Workbook ()

page = planilha ['Sheet']
page.title = "dados covid"

page = planilha.active

# Adicionando dados

page['A1'].value = "Data" 
page['B1'].value = "Novos casos" 
page['C1'].value = "Mortes" 

page['A2'].value = "20/02" 
page['B2'].value = "57472" 
page['C2'].value = "1212" 

page['A3'].value = "21/02" 
page['B3'].value = "29026" 
page['C3'].value = "527" 

page['A4'].value = "22/02" 
page['B4'].value = "26986" 
page['C4'].value = "639"

page['A5'].value = "23/02"
page['B5'].value = "62715"
page['C5'].value = "1386"

page['A6'].value = "24/02"
page['B6'].value = "66588"
page['C6'].value = "1428"

page['A7'].value = "25/02"
page['B7'].value = "65998"
page['C7'].value = "1541"

page['A8'].value = "26/02"
page['B8'].value = "65169"
page['C8'].value = "1337"

page['A9'].value = "27/02"
page['B9'].value = "61602"
page['C9'].value = "1386"


planilha.save ("dados.xlsx")
print ("")
print ("Aguarde um instante...")
print ("")
print ("carregando...")

# Fazendo os gráficos

planilha = pd.read_excel ("dados.xlsx")

dia = planilha ['Data']
casos = planilha ['Novos casos']
mortes = planilha ['Mortes']

plt.title ("Casos novos de covid por data de notificação no Brasil - 20/02 a 27/02")
plt.bar (dia, mortes, color = 'orange', width = 0.5)
plt.grid ()
plt.savefig ("casos.png")
plt.show ()


plt.title ("Mortes de covid por data de notificação no Brasil - 20/02 a 27/02")
plt.bar (dia, mortes, color = 'black', width = 0.5)
plt.grid ()
plt.savefig ("mortes.png")
plt.show ()
print ("Gerando os gráficos")

# Gerando o pdf

pdf = FPDF ('P', 'mm', 'A4')

pdf.add_page ()
pdf.set_font ('Arial', '', 10)
pdf.multi_cell (w = 0, h = 5, txt = "Lucas Cardoso \n\n  fonte: https://covid.saude.gov.br/'", align='J')

pdf.image (name = "casos.png", x = 0, y = 50, w = 100)
 
pdf.add_page ()
pdf.image (name = "mortes.png", x = 0, y = 50, w = 100)

pdf.output ("relatorios.pdf")
print ("PDF criado")

os.system ("pause")