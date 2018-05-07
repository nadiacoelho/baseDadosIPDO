import pandas as pd
writer = pd.ExcelWriter('AnaliseTermicas.xlsx', engine= 'openpyxl') #####Modificar aqui o nome do arquivo para a fonte em análise
pd.set_option('display.width', 400)

def setFile(a,m):
    if m == 1:
        d = 31
        mes = 'JAN'
    elif m == 2:
        d = 28
        mes = 'FEV'
    elif m == 3:
        d = 31
        mes = 'MAR'
    elif m == 4:
        d = 30
        mes = 'ABR'
    elif m == 5:
        d = 31
        mes = 'MAI'
    elif m == 6:
        d = 30
        mes = 'JUN'
    elif m == 7:
        d = 31
        mes = 'JUL'
    elif m == 8:
        d = 31
        mes = 'AGO'
    elif m == 9:
        d = 30
        mes = 'SET'
    elif m == 10:
        d = 31
        mes = 'OUT'
    elif m == 11:
        d = 30
        mes = 'NOV'
    elif m == 12:
        d = 31
        mes = 'DEC'

    excel = 'IPDO_' + str(d) + mes + str(a) + '.xlsx'
    return(excel, d)

df = pd.DataFrame(columns=['Dia', 'Sudeste', 'Sul', 'Norte', 'Nordeste'])

N = []
NE = []
S = []
SE = []
data = []
for mes in range(3,5): ##### Definir aqui os meses para análise
     (excel, ultimoDia) = setFile(2018, mes)
     print(excel)
     for d in range(1,ultimoDia):
         tabelaS = pd.read_excel(excel, "03-Dados Diários Acumulados S", skiprows=6)  #####Modificar aqui para coletar outras fontes
         tabelaSE = pd.read_excel(excel, "04-Dados Diários Acumulados SE", skiprows=6)  #####Modificar aqui para coletar outras fontes
         tabelaNE = pd.read_excel(excel, "05-Dados Diários Acumulados NE", skiprows=6)  #####Modificar aqui para coletar outras fontes
         tabelaN = pd.read_excel(excel, "06-Dados Diários Acumulados N", skiprows=6)  #####Modificar aqui para coletar outras fontes

         data.append(tabelaS.iloc[d, 0])
         N.append(float(tabelaN.iloc[d, 3]))
         NE.append(float(tabelaNE.iloc[d, 3]))
         S.append(float(tabelaS.iloc[d, 3]))
         SE.append(float(tabelaSE.iloc[d, 3]))

df["Dia"] = data
df["Norte"] = N
df["Nordeste"] = NE
df["Sul"] = S
df["Sudeste"] = SE

df.set_index("Dia")

df.to_excel(writer, sheet_name='Geração Térmica', index = True) #####Modificar o nome da planilha
writer.save()