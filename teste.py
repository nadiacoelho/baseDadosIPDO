import pandas as pd
writer = pd.ExcelWriter('BaseIPDO.xlsx', engine= 'openpyxl') #para usar essa função, é necessário instalar o pacote "xlsxwriter" ou "openpyxl"
pd.set_option('display.width', 100)
#determinação de constantes, para melhor legibilidade do código:
#definições relativas às tabelas em análise para indexação
linNorteENA = 2
linNordesteENA = 3
linSulENA = 4
linSudesteENA = 5

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

def ENAdiaria17a():
    #Inicialização dos dataframes:
    enaNorte = pd.DataFrame()
    enaNordeste = pd.DataFrame()
    enaSul = pd.DataFrame()
    enaSudeste = pd.DataFrame()

    for mes in range (1,3):
        #Os dados diários são guardados em listas mensais. Ao fim de cada mês, as listas são adicionadas ao Dataframe do subsitema correspondente.
        #A cada mês as listas são zeradas:
        norte = []
        nordeste = []
        sul = []
        sudeste = []

        # Descrição do loop diário:
        dia = 1
        ultimoDia = False
        while ultimoDia == False:
            (excel, mesPasta, ultimoDia) = setFile(dia, 2017, mes)
            print(excel)
            energiaNaturalAfluente = pd.read_excel(excel, "19-Energia Natural Afluente")
            enaN = float(energiaNaturalAfluente.iloc[linNorteENA, 3])
            norte.append(enaN)
            enaNE = float(energiaNaturalAfluente.iloc[linNordesteENA, 3])
            nordeste.append(enaNE)
            enaS = float(energiaNaturalAfluente.iloc[linSulENA, 3])
            sul.append(enaS)
            enaSE = float(energiaNaturalAfluente.iloc[linSudesteENA, 3])
            sudeste.append(enaSE)
            dia += 1

        # Passagem das listas para os Dataframes:
        auxN = pd.Series(norte, name= "Norte")
        enaNorte = pd.concat([enaNorte,auxN])
        auxNE = pd.Series(nordeste, name= "Nordeste")
        enaNordeste = pd.concat([enaNordeste,auxNE])
        auxS = pd.Series(sul, name= "Sul")
        enaSul = pd.concat([enaSul,auxS])
        auxSE = pd.Series(sudeste, name= "Sudeste")
        enaSudeste = pd.concat([enaSudeste, auxSE])

        # Passagem dos dataframes para planilhas de um arquivo excel:
        enaNorte.to_excel(writer, sheet_name='Norte', index = True, startcol= 1, startrow= 1)
        writer.save()
        enaNordeste.to_excel(writer, sheet_name= 'Nordeste', index= True, startcol = 1, startrow = 1)
        writer.save()
        enaSul.to_excel(writer, sheet_name= 'Sul', index= True, startcol = 1, startrow = 1)
        writer.save()
        enaSudeste.to_excel(writer, sheet_name= 'Sudeste', index= True, startcol = 1, startrow = 1)
        writer.save()

        mes += 1

ENAdiaria17a()

#fazer outras funções com o acesso à planilha de interessa adaptado para:
#   1 - coletar a ENA dos mess seguintes (mai - ago/ set - 2018)
#   2 - coletar outros dados (EA, carga)