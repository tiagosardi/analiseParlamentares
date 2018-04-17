# Exportar resultados para nova planilha

# Plotar grafico barra de gastos e quantidade de propostas
#LER textos das propostas e identificar quantidade que vao pras categorias: educacao, saude, seguranca publica, gestao, reforma, etc
import xlrd

import matplotlib.pyplot as grafico
grafico.rcdefaults()
import numpy as np
import matplotlib.pyplot as grafico

import plotly.plotly as py

from datetime import datetime
import xlsxwriter


import os




os.system('play --no-show-progress --null --channels 1 synth %s sine %f' % ( .05, 1000))

book = xlrd.open_workbook("dados_entrada.xlsx")
print "Numero de abas: ", book.nsheets
print "Nomes das Planilhas:", book.sheet_names()
print("folha 1: 2010")
transparencia2010 = book.sheet_by_index(0)
print("folha 2: 2011")
transparencia2011 = book.sheet_by_index(1)
print("folha 3: 2012")
transparencia2012 = book.sheet_by_index(2)
print("folha 4: 2013")
transparencia2013 = book.sheet_by_index(3)
print("folha 5: 2014")
transparencia2014 = book.sheet_by_index(4)
print("folha 6: PARTIDOS")
partidos = book.sheet_by_index(5)
print("folha 7: EXTRACAO")
extracao = book.sheet_by_index(6)
print("folha 8:PROPOSICOES")
proposicoes = book.sheet_by_index(7)

# #CRIACAO DA LISTA DE PARTIDOS, PEGANDO DA PLANILHA DE PARTIDOS
listaPartidos = []
for i in range(partidos.nrows-1):
	listaPartidos.append(partidos.row(i+1)[0].value)

print(listaPartidos)

os.system('play --no-show-progress --null --channels 1 synth %s sine %f' % ( .08, 1000))
# 3 minutos de processamento


#CALCULA OS GASTOS POR CADA ANO E DEPOIS SOMA TUDO -> SAO GASTOS POR PARTIDO

cont=0
listaGastosPartido2010 = []
for j in range(partidos.nrows-1):
	for i in range(transparencia2010.nrows-1):
		if transparencia2010.row(i+1)[1].value == 	partidos.row(j+1)[0].value:
			cont=cont+transparencia2010.row(i+1)[7].value
	listaGastosPartido2010.insert(j,cont)
	cont=0
os.system('play --no-show-progress --null --channels 1 synth %s sine %f' % ( .5, 1000))
# print("\n2010")
# print(listaGastosPartido2010)

cont=0
listaGastosPartido2011 = []
for j in range(partidos.nrows-1):
    for i in range(transparencia2011.nrows-1):
        if transparencia2011.row(i+1)[1].value ==   partidos.row(j+1)[0].value:
            cont=cont+transparencia2011.row(i+1)[7].value
    listaGastosPartido2011.insert(j,cont)
    cont=0
os.system('play --no-show-progress --null --channels 1 synth %s sine %f' % ( .5, 1000))
# print("\n2011")
# print(listaGastosPartido2011)

cont=0
listaGastosPartido2012 = []
for j in range(partidos.nrows-1):
    for i in range(transparencia2012.nrows-1):
        if transparencia2012.row(i+1)[1].value ==   partidos.row(j+1)[0].value:
            cont=cont+transparencia2012.row(i+1)[7].value
    listaGastosPartido2012.insert(j,cont)
    cont=0
os.system('play --no-show-progress --null --channels 1 synth %s sine %f' % ( .5, 1000))
# print("\n2012")
# print(listaGastosPartido2012)

cont=0
listaGastosPartido2013 = []
for j in range(partidos.nrows-1):
    for i in range(transparencia2013.nrows-1):
        if transparencia2013.row(i+1)[1].value ==   partidos.row(j+1)[0].value:
            cont=cont+transparencia2013.row(i+1)[7].value
    listaGastosPartido2013.insert(j,cont)
    cont=0
os.system('play --no-show-progress --null --channels 1 synth %s sine %f' % ( .5, 1000))
# print("\n2013")
# print(listaGastosPartido2013)

cont=0
listaGastosPartido2014 = []
for j in range(partidos.nrows-1):
    for i in range(transparencia2014.nrows-1):
        if transparencia2014.row(i+1)[1].value ==   partidos.row(j+1)[0].value:
            cont=cont+transparencia2014.row(i+1)[7].value
    listaGastosPartido2014.insert(j,cont)
    cont=0
os.system('play --no-show-progress --null --channels 1 synth %s sine %f' % ( .5, 1000))
# print("\n2014")
# print(listaGastosPartido2014)

#ESSE EH O RESULTADO DA SOMA DE TODOS OS GASTOS DA LEGISLATURA 54 SEPARADO POR PARTIDO
#fazer grafico disso
print("\n\nSOMA GASTOS DA LEGISLATURA\n\n")
listaGastosTotal = [listaGastosPartido2010[i]+listaGastosPartido2011[i]+listaGastosPartido2012[i]+listaGastosPartido2013[i]+listaGastosPartido2014[i] for i in xrange(len(listaGastosPartido2012)-1)]
# print(listaGastosTotal)



# GERA A LISTA DE SUBCOTAS SEM REPETICAO
listaSubCotas = []
for i in range(transparencia2010.nrows-1):
    if transparencia2010.row(i+1)[3].value not in listaSubCotas:
        listaSubCotas.append(transparencia2010.row(i+1)[3].value)
# print(listaSubCotas)


#QUANTIDADE DE SUBCOTAS -> COM O QUE OS PARLAMENTARES MAIS GASTARAM
qtdeSubCota = [0,0,0,0,0,0,0,0,0,0,0,0,0]
for i in range(extracao.nrows-1):
    for j in range(transparencia2010.nrows-1):
        if extracao.row(i+1)[0].value == transparencia2010.row(j+1)[3].value:
            qtdeSubCota[i] = qtdeSubCota[i] + 1

# print("qtdeSubCota\n\n")
# print(qtdeSubCota)

nomesSubcotas = ['a','a','a','a','a','a','a','a','a','a','a','a','a']

#CAPTURA O NOME DAS SUBCOTAS COM A MESMA ORDEM DE QTDESUBCOTAS E LISTSUBCOTAS
for i in range(extracao.nrows-1):
    nomesSubcotas[i] = extracao.row(i+1)[1].value
    
# print(nomesSubcotas)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


# listaCNPJ = []
# for i in range(transparencia.nrows-1):
#   if transparencia.row(i+1)[6].value not in listaCNPJ:
#       listaCNPJ.append(transparencia.row(i+1)[6].value)


# CRIA UMA LISTA COM TODOS OS GASTOS DE CADA PARTIDO
'''
cont=0
#CRIA UMA LISTA COM A QUANTIDADE DE PROPOSICOES QUE SEJA DA LEGISLATURA 54
listaQtdeProposicoes = []
for j in range(partidos.nrows-1):
	for i in range(proposicoes.nrows-1):
            if proposicoes.row(i+1)[0].value == 54 and proposicoes.row(i+1)[1].value == partidos.row(j+1)[0].value:
                cont = cont + 1
	listaQtdeProposicoes.insert(j,cont)
	cont=0


'''
#CRIA UMA INTERCALACAO PARA MOSTRAR PARTIDO E EM SEGUIDA O GASTO
listaResultado = []
for i in range(partidos.nrows-1):
        listaResultado.append(partidos.row(i+1)[0].value)
        listaResultado.append(listaGastosPartido[i])
'''
#####################################################################################
print("\n============================================")
print ("QUANTIDADE DE PROPOSICOES DE CADA PARTIDO")
print("============================================")
# print (listaQtdeProposicoes)
# print("\n============================================")
# print ("LISTA DE GASTOS CAUSADOS POR PARTIDO")
# print("============================================")
# print(listaResultado)

########################################################################################


# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('resultados.xlsx')
worksheet = workbook.add_worksheet()

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': 1})

# Add a number format for cells with money.
money_format = workbook.add_format({'num_format': 'R$#,##00'})

# Add an Excel date format.
#date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})

# Adjust the column width.
worksheet.set_column(1, 1, 15)

# Write some data headers.
worksheet.write('A1', 'Partidos', bold)
worksheet.write('B1', 'Proposicoes', bold)
worksheet.write('C1', 'Gastos2010', bold)
worksheet.write('D1', 'Gastos2011', bold)
worksheet.write('E1', 'Gastos2012', bold)
worksheet.write('F1', 'Gastos2013', bold)
worksheet.write('G1', 'Gastos2014', bold)
worksheet.write('H1', 'GastosTotais', bold)
worksheet.write('J1', 'SubCotas', bold)
worksheet.write('K1', 'Nome_SubCota', bold)

for i in range(partidos.nrows-1):
    worksheet.write(i+1, 0, listaPartidos[i])
    worksheet.write(i+1, 1, listaQtdeProposicoes[i])
    worksheet.write(i+1, 2, listaGastosPartido2010[i], money_format)
    worksheet.write(i+1, 3, listaGastosPartido2011[i], money_format)
    worksheet.write(i+1, 4, listaGastosPartido2012[i], money_format)
    worksheet.write(i+1, 5, listaGastosPartido2013[i], money_format)
    worksheet.write(i+1, 6, listaGastosPartido2014[i], money_format)


for i in range(len(listaGastosTotal)):
    worksheet.write(i+1, 7, listaGastosTotal[i], money_format)



#O TAMANHO EH IGUAL A QUANTIDADE DE SUBCOTAS EXISTENTES
for i in range(len(listaSubCotas)):
    worksheet.write(i+1, 9, qtdeSubCota[i])
    worksheet.write(i+1, 10, nomesSubcotas[i])

workbook.close()



##########################################################################################
import numpy as np
import matplotlib.pyplot as plt

data = (listaQtdeProposicoes[:len(listaQtdeProposicoes)/2], ('r', 'g', '#00FF33'), listaPartidos[:len(listaPartidos)/2])
xPositions = np.arange(len(data[0]))
barWidth = 0.40  # Largura da barra

_ax = plt.axes()  # Cria axes

# bar(left, height, width=0.8, bottom=None, hold=None, **kwargs)
_chartBars = plt.bar(xPositions, data[0], barWidth, color=data[1],
                     yerr=5)  # Gera barras

for bars in _chartBars:
    # text(x, y, s, fontdict=None, withdash=False, **kwargs)
    _ax.text(bars.get_x() + (bars.get_width()/2), bars.get_height()+5,
             bars.get_height(), ha='center')  # Label acima das barras

_ax.set_xticks(xPositions)
_ax.set_xticklabels(data[2])

plt.xlabel('PARTIDOS')
plt.ylabel('PROPOSICOES')
plt.grid(True)
plt.legend(_chartBars, data[2])


###################################################################################################


plt.figure(2)

data = (listaQtdeProposicoes[len(listaQtdeProposicoes)/2:], ('r', 'g', '#00FF33'), listaPartidos[len(listaPartidos)/2:])
xPositions = np.arange(len(data[0]))
barWidth = 0.40  # Largura da barra

_ax = plt.axes()  # Cria axes

# bar(left, height, width=0.8, bottom=None, hold=None, **kwargs)
_chartBars = plt.bar(xPositions, data[0], barWidth, color=data[1],
                     yerr=5)  # Gera barras

for bars in _chartBars:
    # text(x, y, s, fontdict=None, withdash=False, **kwargs)
    _ax.text(bars.get_x() + (bars.get_width()/2), bars.get_height()+5,
             bars.get_height(), ha='center')  # Label acima das barras

_ax.set_xticks(xPositions)
_ax.set_xticklabels(data[2])

plt.xlabel('PARTIDOS')
plt.ylabel('PROPOSICOES')
plt.grid(True)
plt.legend(_chartBars, data[2])

######################################################################################################

plt.figure(3)

data = (listaGastosTotal[:len(listaGastosTotal)/2], ('#FFCC00', '#FF00FF', '#CCFFCC'), listaPartidos[:len(listaPartidos)/2])
xPositions = np.arange(len(data[0]))
barWidth = 0.40  # Largura da barra

_ax = plt.axes()  # Cria axes

# bar(left, height, width=0.8, bottom=None, hold=None, **kwargs)
_chartBars = plt.bar(xPositions, data[0], barWidth, color=data[1],
                     yerr=5)  # Gera barras

for bars in _chartBars:
    # text(x, y, s, fontdict=None, withdash=False, **kwargs)
    _ax.text(bars.get_x() + (bars.get_width()/2), bars.get_height()+5,
             bars.get_height(), ha='center')  # Label acima das barras

_ax.set_xticks(xPositions)
_ax.set_xticklabels(data[2])

plt.xlabel('PARTIDOS')
plt.ylabel('GASTOS')
plt.grid(True)
plt.legend(_chartBars, data[2])


########################################################################################################



plt.figure(4)

data = (listaGastosTotal[len(listaGastosTotal)/2:], ('#FFCC00', '#FF00FF', '#CCFFCC'), listaPartidos[len(listaPartidos)/2:])
xPositions = np.arange(len(data[0]))
barWidth = 0.40  # Largura da barra

_ax = plt.axes()  # Cria axes

# bar(left, height, width=0.8, bottom=None, hold=None, **kwargs)
_chartBars = plt.bar(xPositions, data[0], barWidth, color=data[1],
                     yerr=5)  # Gera barras

for bars in _chartBars:
    # text(x, y, s, fontdict=None, withdash=False, **kwargs)
    _ax.text(bars.get_x() + (bars.get_width()/2), bars.get_height()+5,
             bars.get_height(), ha='center')  # Label acima das barras

_ax.set_xticks(xPositions)
_ax.set_xticklabels(data[2])

plt.xlabel('PARTIDOS')
plt.ylabel('GASTOS')
plt.grid(True)
plt.legend(_chartBars, data[2])

######################################################################################################
#GRAFICO DE QUANTIDADE DE COTAS E SEUS NOMES
#GRAFICO DE PIZZA

grafico.figure(5)
#porcentagens = [porcentagemClienteDificil , porcentagemManutencao,porcentagemCosmeticosHomens,porcentagemProdutoBarba]
#labels = ['CLIENTE DIFICIL.COM' , 'MANUTENCAO BARBA','COSMESTICOS PARA\n HOMENS LTDA','PRODUTO BARBA STU']
grafico.axis('equal')
grafico.title("Com o que gastam")

#explode = (0,0,0)
#tamanho, explode (qual vai se deslocar),  
grafico.pie(qtdeSubCota, labels=nomesSubcotas,autopct='%1.0f%%',shadow=True, startangle=90)




#########################################################################################################






os.system('play --no-show-progress --null --channels 1 synth %s sine %f' % ( .3, 1000))
os.system('play --no-show-progress --null --channels 1 synth %s sine %f' % ( .3, 1000))
os.system('play --no-show-progress --null --channels 1 synth %s sine %f' % ( .3, 1000))
os.system('play --no-show-progress --null --channels 1 synth %s sine %f' % ( .6, 1000))


plt.show()
