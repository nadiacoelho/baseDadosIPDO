import urllib.request

a = 2017
m = 6
for m in range (6,8):
    if m == 1:
        dmax = 32
        mes = 'JAN'
        mesPasta = 'Janeiro'
    elif m == 2:
        dmax = 29
        mes = 'FEV'
        mesPasta = 'Fevereiro'
    elif m == 3:
        dmax = 32
        mes = 'MAR'
        mesPasta = 'Março'
    elif m == 4:
        dmax = 31
        mes = 'ABR'
        mesPasta = 'Abril'
    elif m == 5:
        dmax = 32
        mes = 'MAI'
        mesPasta = 'Maio'
    elif m == 6:
        dmax = 31
        mes = 'JUN'
        mesPasta = 'Junho'
    elif m == 7:
        dmax = 32
        mes = 'JUL'
        mesPasta = 'Julho'
    elif m == 8:
        dmax = 32
        mes = 'AGO'
        mesPasta = 'Agosto'
    elif m == 9:
        dmax = 31
        mes = 'SET'
        mesPasta = 'Setembro'
    elif m == 10:
        dmax = 32
        mes = 'OUT'
        mesPasta = 'Outubro'
    elif m == 11:
        dmax = 31
        mes = 'NOV'
        mesPasta = 'Novembro'
    elif m == 12:
        dmax = 32
        mes = 'DEC'
        mesPasta = 'Dezembro'

    baseIPDO = 'http://sdro.ons.org.br/SDRO/DIARIO/'

    for d in range(1, dmax):
        if m < 10:
            if d < 10:
                enderecoIPDO = baseIPDO + str(a) + '_0' + str(m) + '_0' + str(d) + '/Html/DIARIO_0' + str(d) + '-0' + str(m) + '-' + str(a) + '.xlsx'
            else:
                enderecoIPDO = baseIPDO + str(a) + '_0' + str(m) + '_' + str(d) + '/Html/DIARIO_' + str(d) + '-0' + str(m) + '-' + str(a) + '.xlsx'
        else:
            if d < 10:
                enderecoIPDO = baseIPDO + str(a) + '_' + str(m) + '_0' + str(d) + '/Html/DIARIO_0' + str(d) + '-' + str(m) + '-' + str(a) + '.xlsx'
            else:
                enderecoIPDO = baseIPDO + str(a) + '_' + str(m) + '_' + str(d) + '/Html/DIARIO_' + str(d) + '-' + str(m) + '-' + str(a) + '.xlsx'

        destino = 'IPDO_' + str(d) + mes + str(a) + '.xlsx'
        urllib.request.urlretrieve(enderecoIPDO, destino)
    m+=1