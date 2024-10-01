import pandas as pd

MRE = pd.read_excel('MRE.xlsx')
CIEE = pd.read_excel('CIEE.xlsx')
SCE = pd.read_excel('SCE.xlsx')

CpfCiee = []
CpfMre = []
CpfSce = []
Diferenca = []

def conversorCpf(aux):
    if len(aux) != 11 and aux[0] != '0' and '.' not in aux:
        aux = '0' + aux
        if len(aux) != 11:
            pass
        else:
            return aux
    elif len(aux) != 11:
            return aux[:3] + aux[4:7] + aux[8:11] + aux[12:]
    else:
        return aux

def switch(n, c, d):
    if d == 'nan':
        return f'{n} de cpf: {c}, continua ativo(a).'
    elif n and c != 0:
        return f'{n} de cpf: {c}, foi desligado em {d}.'
    else:
        return f'O cpf {d} não foi localizado na base de dados.'

for a in range(len(CIEE.iloc[:,4])):
    CpfCiee.append(conversorCpf(str(CIEE.iloc[a,4])))

for b in range(len(MRE.iloc[:,3])):
    CpfMre.append(conversorCpf(str(MRE.iloc[b,3])))

for c in range(len(CpfCiee)):
    if CpfCiee[c] not in CpfMre:
        Diferenca.append(CpfCiee[c])

for d in range(len(SCE.iloc[:,6])):
    CpfSce.append(conversorCpf(str(SCE.iloc[d,6])))

with open('ResultadoAnaliseDados.txt','w') as resultado:
    for e in range(len(Diferenca)):
        if Diferenca[e] not in CpfSce:
            resultado.write(switch(0,0,Diferenca[e]) + '\n')
        for f in range(len(SCE.iloc[:,6])):
            if Diferenca[e] == conversorCpf(str(SCE.iloc[f,6])):
                resultado.write(switch(str(SCE.iloc[f,4]),str(SCE.iloc[f,6]),str(SCE.iloc[f,28])) + '\n')