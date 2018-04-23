import pandas as pd
writer = pd.ExcelWriter('BaseIPDO.xlsx', engine= 'openpyxl') #para usar essa função, é necessário instalar o pacote "xlsxwriter" ou "openpyxl"
pd.set_option('display.width', 100)
#determinação de constantes, para melhor legibilidade do código:
#definições relativas às tabelas em análise para indexação
linNorteENA = 2
linNordesteENA = 3
linSulENA = 4
linSudesteENA = 5

#FUNÇÕES UTILIZADAS:
#determinar o nome do arquivo a ser analisado:
def setFile(d,a,m):
    if m == 1:
        ultimoDia = 31
        mes = 'JAN'
        mesPasta = 'Janeiro'
    elif m == 2:
        ultimoDia = 28
        mes = 'FEV'
        mesPasta = 'Fevereiro'
    elif m == 3:
        ultimoDia = 31
        mes = 'MAR'
        mesPasta = 'Março'
    elif m == 4:
        ultimoDia = 30
        mes = 'ABR'
        mesPasta = 'Abril'
    elif m == 5:
        ultimoDia = 31
        mes = 'MAI'
        mesPasta = 'Maio'
    elif m == 6:
        ultimoDia = 30
        mes = 'JUN'
        mesPasta = 'Junho'
    elif m == 7:
        ultimoDia = 31
        mes = 'JUL'
        mesPasta = 'Julho'
    elif m == 8:
        ultimoDia = 31
        mes = 'AGO'
        mesPasta = 'Agosto'
    elif m == 9:
        ultimoDia = 30
        mes = 'SET'
        mesPasta = 'Setembro'
    elif m == 10:
        ultimoDia = 31
        mes = 'OUT'
        mesPasta = 'Outubro'
    elif m == 11:
        ultimoDia = 30
        mes = 'NOV'
        mesPasta = 'Novembro'
    elif m == 12:
        ultimoDia = 31
        mes = 'DEC'
        mesPasta = 'Dezembro'
    excel = 'IPDO_' + str(d) + mes + str(a) + '.xlsx'
    fim = (ultimoDia == d)
    return(excel, mesPasta, fim)

#acessar a planilha a ser analisada:
#a ordem das planilhas no arquivo excel é diferente nos intervalos fev-abr/mai-ago/set-atualmente.
#por isso, é necessário definir três funções diferentes.
def ENAdiaria17a(ano, mes):
    #Inicialização das series:
    enaNorte = pd.Series(name='ENA')
    enaNordeste = pd.Series(name='ENA')
    enaSul = pd.Series(name='ENA')
    enaSudeste = pd.Series(name='ENA')

    #Descrição do loop diário
    dia = 1
    ultimoDia = False
    while ultimoDia == False:
        (excel, mesPasta, ultimoDia) = setFile(dia, ano, mes)
        energiaNaturalAfluente = pd.read_excel(excel, "19-Energia Natural Afluente")
        enaN = energiaNaturalAfluente.iloc[linNorteENA, 3]
        enaNE = energiaNaturalAfluente.iloc[linNordesteENA, 3]
        enaS = energiaNaturalAfluente.iloc[linSulENA, 3]
        enaSE = energiaNaturalAfluente.iloc[linSudesteENA, 3]

        #Alocação do dado diário na série
        enaNorte.append(enaN)
        enaNordeste.append(enaNE)
        enaSul.append(enaS)
        enaSudeste.append(enaSE)

def ENAdiaria17b(ano, mes):
    (excel, mesPasta) = setFile(ano, mes)
    #ENA
    energiaNaturalAfluente = pd.read_excel(excel, "20-Energia Natural Afluente")
    enaNorte = energiaNaturalAfluente.iloc[linNorteENA, 4] #ENA bruta do mês
    enaNordeste = energiaNaturalAfluente.iloc[linNordesteENA, 4]
    enaSul = energiaNaturalAfluente.iloc[linSulENA, 4]
    enaSudeste = energiaNaturalAfluente.iloc[linSudesteENA,4]

    #cria uma série com os novos dados
    dadosENA = [enaNorte, enaNordeste, enaSul, enaSudeste]
    ENAmensal = pd.Series(dadosENA, name= "ENA " + str(mesPasta) + "2018")
    ENAmensal.set_axis(indexSubsistema, inplace=True)
    return(ENAmensal)
def ENAmensal18(ano, mes):
    (excel, mesPasta) = setFile(ano, mes)
    #ENA
    energiaNaturalAfluente = pd.read_excel(excel, "21-Energia Natural Afluente")
    enaNorte = energiaNaturalAfluente.iloc[linNorteENA, 4] #ENA bruta do mês
    enaNordeste = energiaNaturalAfluente.iloc[linNordesteENA, 4]
    enaSul = energiaNaturalAfluente.iloc[linSulENA, 4]
    enaSudeste = energiaNaturalAfluente.iloc[linSudesteENA,4]

    #cria uma série com os novos dados
    dadosENA = [enaNorte, enaNordeste, enaSul, enaSudeste]
    ENAmensal = pd.Series(dadosENA, name= "ENA " + str(mesPasta) + "2018")
    ENAmensal.set_axis(indexSubsistema, inplace=True)
    return(ENAmensal)

#as funções abaixo coletam os dados de cada ano:
def ano2018():
    jan18 = ENAmensal18(2018, 1)
    fev18 = ENAmensal18(2018, 2)
    mar18 = ENAmensal18(2018, 3)

    ano2018 = pd.DataFrame([jan18, fev18, mar18])
    return(ano2018)
def ano2017():
    fev17 = ENAmensal17a(2017, 2)
    mar17 = ENAmensal17a(2017, 3)
    abr17 = ENAmensal17a(2017, 4)
    mai17 = ENAmensal17b(2017, 5)
    jun17 = ENAmensal17b(2017, 6)
    jul17 = ENAmensal17b(2017, 7)
    ago17 = ENAmensal17b(2017, 8)
    set17 = ENAmensal18(2017, 9)
    out17 = ENAmensal18(2017, 10)
    nov17 = ENAmensal18(2017, 11)
    dez17 = ENAmensal18(2017, 12)

    ano2017 = pd.DataFrame([fev17, mar17, abr17, mai17, jun17, jul17, ago17, set17, out17, nov17, dez17])
    return(ano2017)

df2018 = pd.DataFrame(ano2018())
df2018.to_excel(writer, sheet_name= 'ena2018', index= True, startcol = 1, startrow = 1)
writer.save()
df2017 = pd.DataFrame(ano2017())
df2017.to_excel(writer,sheet_name= 'ena2017', index= True, startcol= 1, startrow = 1)
writer.save()