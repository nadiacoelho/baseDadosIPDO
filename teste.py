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
    fim = (ultimoDia == d)
    excel = 'IPDO_' + str(d) + mes + str(a) + '.xlsx'
    return(excel, mesPasta, fim)

def ENAdiaria2017():
    norte = pd.DataFrame(columns=['Dia', 'ENA'])
    nordeste = pd.DataFrame(columns=['Dia', 'ENA'])
    sul = pd.DataFrame(columns=['Dia', 'ENA'])
    sudeste = pd.DataFrame(columns=['Dia', 'ENA'])

    enaN = []
    enaNE = []
    enaS = []
    enaSE = []
    data = []

    for mes in range (1,13):
        # Descrição do loop diário:
        dia = 1
        ultimoDia = False
        while ultimoDia == False:
            data.append(str(dia) + "/" + str(mes) + "/2017")
            (excel, mesPasta, ultimoDia) = setFile(dia, 2017, mes)
            print(excel)
            if (mes in range(1, 5) and dia in range(1, 32)):
                energiaNaturalAfluente = pd.read_excel(excel, "19-Energia Natural Afluente")
            if ((mes==5) and dia in range (16,31)) or (mes in range (6,9) and dia in range (1,32)):
                energiaNaturalAfluente = pd.read_excel(excel, "20-Energia Natural Afluente")
            if mes in range(10,13):
                energiaNaturalAfluente = pd.read_excel(excel, "21-Energia Natural Afluente")

            enaN.append(float(energiaNaturalAfluente.iloc[linNorteENA, 3]))
            enaNE.append(float(energiaNaturalAfluente.iloc[linNordesteENA, 3]))
            enaS.append(float(energiaNaturalAfluente.iloc[linSulENA, 3]))
            enaSE.append(float(energiaNaturalAfluente.iloc[linSudesteENA, 3]))

            dia += 1

        mes += 1
    norte["Dia"] = data
    norte["ENA"] = enaN
    norte.set_index("Dia")
    nordeste["Dia"] = data
    nordeste["ENA"] = enaNE
    nordeste.set_index("Dia")
    sul["Dia"] = data
    sul["ENA"] = enaS
    sul.set_index("Dia")
    sudeste["Dia"] = data
    sudeste["ENA"] = enaSE
    sudeste.set_index("Dia")

    return(norte, nordeste, sul, sudeste)

def EARdiaria2017():
    norte = pd.DataFrame(columns=['Dia', 'EAr'])
    nordeste = pd.DataFrame(columns=['Dia', 'EAr'])
    sul = pd.DataFrame(columns=['Dia', 'EAr'])
    sudeste = pd.DataFrame(columns=['Dia', 'EAr'])

    earN = []
    earNE = []
    earS = []
    earSE = []
    data = []
    for mes in range(1,13):
        # Descrição do loop diário:
        dia = 1
        ultimoDia = False
        while ultimoDia == False:
            data.append(str(dia) + "/" + str(mes) + "/2017")
            (excel, mesPasta, ultimoDia) = setFile(dia, 2017, mes)
            print(excel)
            if (mes in range(1, 5) and dia in range(1, 32)):
                energiaArmazenada = pd.read_excel(excel, "18-Variação Energia Armazenada")
            if ((mes == 5) and dia in range(16, 31)) or (mes in range(6, 9) and dia in range(1, 32)):
                energiaArmazenada = pd.read_excel(excel, "19-Variação Energia Armazenada")
            if mes in range(10, 13):
                energiaArmazenada = pd.read_excel(excel, "20-Variação Energia Armazenada")

            earN.append(float(energiaArmazenada.iloc[5, colNorteEAR]))
            earNE.append(float(energiaArmazenada.iloc[5, colNordesteEAR]))
            earS.append(float(energiaArmazenada.iloc[5, colSulEAR]))
            earSE.append(float(energiaArmazenada.iloc[5, colSudesteEAR]))

            dia += 1

        mes += 1
    norte["Dia"] = data
    norte["EAr"] = earN
    norte.set_index("Dia")
    nordeste["Dia"] = data
    nordeste["EAr"] = earNE
    nordeste.set_index("Dia")
    sul["Dia"] = data
    sul["EAr"] = earS
    sul.set_index("Dia")
    sudeste["Dia"] = data
    sudeste["EAr"] = earSE
    sudeste.set_index("Dia")

    return(norte, nordeste, sul, sudeste)

(enaNorte, enaNordeste, enaSul, enaSudeste) = ENAdiaria2017()
(earNorte, earNordeste, earSul, earSudeste) = EARdiaria2017()
enaNorte.set_index("Dia")
enaNordeste.set_index("Dia")
enaSul.set_index("Dia")
enaSudeste.set_index("Dia")
earNorte.set_index("Dia")
earNordeste.set_index("Dia")
earSul.set_index("Dia")
earSudeste.set_index("Dia")

norte = pd.merge(right=enaNorte, left=earNorte, how='left', on=["Dia"])
nordeste = pd.merge(right = enaNordeste,left = earNordeste, how='left', on=["Dia"] )
sul = pd.merge(enaSul, earSul, how='left', on=["Dia"])
sudeste = pd.merge(enaSudeste, earSudeste, how='left', on=["Dia"])


norte.to_excel(writer, sheet_name='Norte', index = True)
writer.save()
nordeste.to_excel(writer, sheet_name= 'Nordeste', index= True)
writer.save()
sul.to_excel(writer, sheet_name= 'Sul', index= True)
writer.save()
sudeste.to_excel(writer, sheet_name= 'Sudeste', index= True)
writer.save()