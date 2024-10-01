import pandas as pd
from datetime import datetime

SCE = pd.read_excel('SCE.xlsx')

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

with open('ResultadoBancos.txt','w') as resultado:
    for a in range (len(SCE.iloc[:,28])):
        if str(SCE.iloc[a,23]) != 'nan':
            if str(SCE.iloc[a,28]) == 'nan' and datetime.strptime(str(SCE.iloc[a,23]).split(' a ')[-1],"%d/%m/%Y") >= datetime.strptime(str(datetime.now().strftime("%d/%m/%Y")),"%d/%m/%Y"):
                resultado.write(f'{SCE.iloc[a,4]} | {conversorCpf(str(SCE.iloc[a,6]))}\n')