import pandas as pd
writer = pd.ExcelWriter('BaseIPDOcarga.xlsx', engine= 'openpyxl') #para usar essa função, é necessário instalar o pacote "xlsxwriter" ou "openpyxl"
pd.set_option('display.width', 300)
indexSubsistema = ['Norte', 'Nordeste', 'Sul', 'Sudeste']

#FUNÇÕES UTILIZADAS:
#determinar o nome do arquivo a ser analisado:
def setFile(a,m):
    if m == 1:
        d = 31
        mes = 'JAN'
        mesPasta = 'Janeiro'
    elif m == 2:
        d = 28
        mes = 'FEV'
        mesPasta = 'Fevereiro'
    elif m == 3:
        d = 31
        mes = 'MAR'
        mesPasta = 'Março'
    elif m == 4:
        d = 30
        mes = 'ABR'
        mesPasta = 'Abril'
    elif m == 5:
        d = 31
        mes = 'MAI'
        mesPasta = 'Maio'
    elif m == 6:
        d = 30
        mes = 'JUN'
        mesPasta = 'Junho'
    elif m == 7:
        d = 31
        mes = 'JUL'
        mesPasta = 'Julho'
    elif m == 8:
        d = 31
        mes = 'AGO'
        mesPasta = 'Agosto'
    elif m == 9:
        d = 30
        mes = 'SET'
        mesPasta = 'Setembro'
    elif m == 10:
        d = 31
        mes = 'OUT'
        mesPasta = 'Outubro'
    elif m == 11:
        d = 30
        mes = 'NOV'
        mesPasta = 'Novembro'
    elif m == 12:
        d = 31
        mes = 'DEC'
        mesPasta = 'Dezembro'

    excel = 'IPDO_' + str(d) + mes + str(a) + '.xlsx'
    return(excel, mesPasta,d)

#acessar a planilha a ser analisada:
#há diferença na formatação da planilha, e portanto, são necessárias duas funções
def cargaMensal2018(ano, mes):
    cargaNorte = 0
    cargaNordeste = 0
    cargaSul = 0
    cargaSudeste = 0
    (excel, mesPasta, dia) = setFile(ano, mes)
    #dadosTotais = pd.read_excel(excel, "02-Balanço de Energia Acumulado")
    dadosCargaS = pd.read_excel(excel, "03-Dados Diários Acumulados S")
    dadosCargaSE = pd.read_excel(excel, "04-Dados Diários Acumulados SE")
    dadosCargaN = pd.read_excel(excel, "06-Dados Diários Acumulados N")
    dadosCargaNE = pd.read_excel(excel, "05-Dados Diários Acumulados NE")
    #print(excel, dadosCargaN)
#define o ultimo dia do mês:
    dmax = dia + 6
    cargaNorte = dadosCargaN.iloc[6:dmax, 7].sum()
    #print("cargaNorte ", cargaNorte)
    cargaNordeste = dadosCargaNE.iloc[6:dmax, 7].sum()
    #print("cargaNordeste ", cargaNordeste)
    cargaSul = dadosCargaS.iloc[6:dmax, 7].sum()
    #print("cargaSul ", cargaSul)
    cargaSudeste = dadosCargaSE.iloc[6:dmax,7].sum()
    #print("cargaSudeste ", cargaSudeste)

    # #cria uma série com os novos dados
    dadosCarga = [cargaNorte, cargaNordeste, cargaSul, cargaSudeste]
    cargaMensal = pd.Series(dadosCarga, name= str(mesPasta))
    cargaMensal.set_axis(indexSubsistema, inplace=True)
    return(cargaMensal)
def cargaMensal2017(ano, mes):
    cargaNorte = 0
    cargaNordeste = 0
    cargaSul = 0
    cargaSudeste = 0

    (excel, mesPasta, dia) = setFile(ano, mes)
    # dadosTotais = pd.read_excel(excel, "02-Balanço de Energia Acumulado")
    dadosCargaS = pd.read_excel(excel, "03-Dados Diários Acumulados S")
    dadosCargaSE = pd.read_excel(excel, "04-Dados Diários Acumulados SE")
    dadosCargaNE = pd.read_excel(excel, "05-Dados Diários Acumulados NE")
    dadosCargaN = pd.read_excel(excel, "06-Dados Diários Acumulados N")
    #print(excel, dadosCargaN)


    # define o ultimo dia do mês:
    dmax = dia + 6
    cargaNorte = dadosCargaN.iloc[6:dmax, 6].sum()
    #print("cargaNorte ", cargaNorte)
    cargaNordeste = dadosCargaNE.iloc[6:dmax, 6].sum()
    #print("cargaNordeste ", cargaNordeste)
    cargaSul = dadosCargaS.iloc[6:dmax, 6].sum()
    #print("cargaSul ", cargaSul)
    cargaSudeste = dadosCargaSE.iloc[6:dmax, 6].sum()
    #print("cargaSudeste ", cargaSudeste)

    # #cria uma série com os novos dados
    dadosCarga = [cargaNorte, cargaNordeste, cargaSul, cargaSudeste]
    cargaMensal = pd.Series(dadosCarga, name= str(mesPasta))
    cargaMensal.set_axis(indexSubsistema, inplace=True)
    return (cargaMensal)

#as funções abaixo coletam os dados de cada ano:
def ano2018():
    print("1/2018")
    jan18 = cargaMensal2018(2018, 1)
    print("2/2018")
    fev18 = cargaMensal2018(2018, 2)
    print("3/2018")
    mar18 = cargaMensal2018(2018, 3)

    ano2018 = pd.DataFrame([jan18, fev18, mar18])
    return(ano2018)
def ano2017():
    print("1/2017")
    jan17 = cargaMensal2017(2017, 1)
    print("2/2017")
    fev17 = cargaMensal2017(2017, 2)
    print("3/2017")
    mar17 = cargaMensal2017(2017, 3)
    print("4/2017")
    abr17 = cargaMensal2017(2017, 4)
    print("5/2017")
    mai17 = cargaMensal2017(2017, 5)
    print("6/2017")
    jun17 = cargaMensal2017(2017, 6)
    print("7/2017")
    jul17 = cargaMensal2017(2017, 7)
    print("8/2017")
    ago17 = cargaMensal2017(2017, 8)
    print("9/2017")
    set17 = cargaMensal2018(2017, 9)
    print("10/2017")
    out17 = cargaMensal2018(2017, 10)
    print("11/2017")
    nov17 = cargaMensal2018(2017, 11)
    print("12/2017")
    dez17 = cargaMensal2018(2017, 12)

    ano2017 = pd.DataFrame([jan17, fev17, mar17, abr17, mai17, jun17, jul17, ago17, set17, out17, nov17, dez17])
    return(ano2017)

# (excel, mesPasta,d) = setFile(2017, 2)
# dadosTotais = pd.read_excel(excel, "03-Dados Diários Acumulados S")
# print(dadosTotais.head(20))
# print(dadosTotais.iloc[6,6])
# print(dadosTotais.iloc[6:31,6].sum())

df2018 = pd.DataFrame(ano2018())
df2018.to_excel(writer, sheet_name= 'carga2018', index= True, startcol = 1, startrow = 1)
writer.save()
df2017 = pd.DataFrame(ano2017())
df2017.to_excel(writer,sheet_name= 'carga2017', index= True, startcol= 1, startrow = 1)
writer.save()
