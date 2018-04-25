import pandas as pd
writer = pd.ExcelWriter('BaseIPDO2017.xlsx', engine= 'openpyxl') #para usar essa função, é necessário instalar o pacote "xlsxwriter" ou "openpyxl"
pd.set_option('display.width', 100)
#determinação de constantes, para melhor legibilidade do código:
#definições relativas às tabelas em análise para indexação
linNorteENA = 2
linNordesteENA = 3
linSulENA = 4
linSudesteENA = 5

colNorteEAR = 3
colNordesteEAR = 4
colSulEAR = 1
colSudesteEAR = 2

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

def ENAdiaria2017():
    #Inicializar os dataframes:
    enaNorte = pd.DataFrame()
    enaNordeste = pd.DataFrame()
    enaSul = pd.DataFrame()
    enaSudeste = pd.DataFrame()

    print(enaNorte)
    for mes in range (1,13):
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
            if (mes in range(1, 5) and dia in range(1, 32)):
                energiaNaturalAfluente = pd.read_excel(excel, "19-Energia Natural Afluente")
            if ((mes==5) and dia in range (16,31)) or (mes in range (6,9) and dia in range (1,32)):
                energiaNaturalAfluente = pd.read_excel(excel, "20-Energia Natural Afluente")
            if mes in range(10,13):
                energiaNaturalAfluente = pd.read_excel(excel, "21-Energia Natural Afluente")

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
        auxN = pd.Series(norte, name= "ENA")
        auxNE = pd.Series(nordeste, name="ENA")
        auxS = pd.Series(sul, name="ENA")
        auxSE = pd.Series(sudeste, name="ENA")

        #Tentativa com concat
        enaNorte = pd.concat([enaNorte,auxN], axis=0)
        enaNordeste = pd.concat([enaNordeste,auxNE], axis=0)
        enaSul = pd.concat([enaSul,auxS], axis=0)
        enaSudeste = pd.concat([enaSudeste, auxSE], axis=0)

        mes += 1

    enaNorte = pd.Series(enaNorte.iloc[:,0], name="ENA")
    enaNordeste = pd.Series(enaNordeste.iloc[:,0], name="ENA")
    enaSul = pd.Series(enaSul.iloc[:,0], name="ENA")
    enaSudeste = pd.Series(enaSudeste.iloc[:,0], name="ENA")

    # Passagem dos dataframes para planilhas de um arquivo excel:
    enaNorte.to_excel(writer, sheet_name='Norte', index = True, startcol= 1, startrow= 1)
    writer.save()
    enaNordeste.to_excel(writer, sheet_name= 'Nordeste', index= True, startcol = 1, startrow = 1)
    writer.save()
    enaSul.to_excel(writer, sheet_name= 'Sul', index= True, startcol = 1, startrow = 1)
    writer.save()
    enaSudeste.to_excel(writer, sheet_name= 'Sudeste', index= True, startcol = 1, startrow = 1)
    writer.save()

def EARdiaria2017():
    # Inicializar os dataframes:
    earNorte = pd.DataFrame()
    earNordeste = pd.DataFrame()
    earSul = pd.DataFrame()
    earSudeste = pd.DataFrame()

    for mes in range(1,13):
    # Os dados diários são guardados em listas mensais. Ao fim de cada mês, as listas são adicionadas ao Dataframe do subsitema correspondente.
    # A cada mês as listas são zeradas:
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
            if (mes in range(1, 5) and dia in range(1, 32)):
                energiaArmazenada = pd.read_excel(excel, "18-Variação Energia Armazenada")
            if ((mes == 5) and dia in range(16, 31)) or (mes in range(6, 9) and dia in range(1, 32)):
                energiaArmazenada = pd.read_excel(excel, "19-Variação Energia Armazenada")
            if mes in range(10, 13):
                energiaArmazenada = pd.read_excel(excel, "20-Variação Energia Armazenada")

            earN = float(energiaArmazenada.iloc[5, colNorteEAR])
            norte.append(earN)
            earNE = float(energiaArmazenada.iloc[5, colNordesteEAR])
            nordeste.append(earNE)
            earS = float(energiaArmazenada.iloc[5, colSulEAR])
            sul.append(earS)
            earSE = float(energiaArmazenada.iloc[5, colSudesteEAR])
            sudeste.append(earSE)
            dia += 1

    # Passagem das listas para os Dataframes:
        auxN = pd.Series(norte, name="Energia Armazenada")
        auxNE = pd.Series(nordeste, name="Energia Armazenada")
        auxS = pd.Series(sul, name="Energia Armazenada")
        auxSE = pd.Series(sudeste, name="Energia Armazenada")

    # Tentativa com concat
        earNorte = pd.concat([earNorte, auxN], axis=0)
        earNordeste = pd.concat([earNordeste, auxNE], axis=0)
        earSul = pd.concat([earSul, auxS], axis=0)
        earSudeste = pd.concat([earSudeste, auxSE], axis=0)

        mes += 1

    earNorte = pd.Series(earNorte.iloc[:, 0], name="Energia Armazenada (%)")
    earNordeste = pd.Series(earNordeste.iloc[:, 0], name="Energia Armazenada (%)")
    earSul = pd.Series(earSul.iloc[:, 0], name="Energia Armazenada(%)")
    earSudeste = pd.Series(earSudeste.iloc[:, 0], name="Energia Armazenada(%)")

     # Passagem dos dataframes para planilhas de um arquivo excel:
    earNorte.to_excel(writer, sheet_name='Norte', index=True, startcol=3, startrow=1)
    writer.save()
    earNordeste.to_excel(writer, sheet_name='Nordeste', index=True, startcol=3, startrow=1)
    writer.save()
    earSul.to_excel(writer, sheet_name='Sul', index=True, startcol=3, startrow=2)
    writer.save()
    earSudeste.to_excel(writer, sheet_name='Sudeste', index=True, startcol=3, startrow=1)
    writer.save()

ENAdiaria2017()
EARdiaria2017()

#fazer outras funções com o acesso à planilha de interessa adaptado para:
#   1 - coletar a ENA dos mess seguintes (mai - ago/ set - 2018) (check)
#   2 - coletar outros dados (EA, carga)