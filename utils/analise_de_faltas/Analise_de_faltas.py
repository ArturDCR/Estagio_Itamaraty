import pandas as pd
import os

from utils.analise_de_faltas import Hob

class Analise_de_faltas:
    def __init__(self):
        self.__FALTAS = pd.read_excel(os.path.join("utils/analise_de_faltas/dados","Faltas.xlsx"))
        self.__SAIDA = open('utils/analise_de_faltas/dados/saida.txt', 'w')
        self.__SCE = pd.read_excel(os.path.join('utils/analise_de_faltas/dados/Sce.xlsx'))

    def __swicth (self, aux):
        if len(str(aux)) == 2:
            return f'000000{aux},00'
        elif len(str(aux)) == 3:
            return f'00000{aux},00'
        elif len(str(aux)) == 4:
            return f'0000{aux},00'
        elif len(str(aux)) == 5:
            return f'000{aux},00'
        elif len(str(aux)) == 6:
            return f'00{aux},00'
        elif len(str(aux)) == 7:
            return f'0{aux},00'
        elif len(str(aux)) == 8:
            return f'{aux},00'

    def __conversor_de_cpf(self, cpf):
            if len(cpf) != 11 and cpf[0] != '0' and '.' not in cpf:
                cpf = '0' + cpf
                if len(cpf) != 11:
                    pass
                else:
                    return str(cpf)
            elif len(cpf) != 11:
                    return str(cpf[:3] + cpf[4:7] + cpf[8:11] + cpf[12:])
            else:
                return str(cpf)

    def __gerar_dados(self, escolha, mes, ano):
        if escolha == 'VT':
            for a in range(len(self.__FALTAS.iloc[:,0])):
                if str(self.__FALTAS.iloc[a,2]) != 'nan':
                    self.__SAIDA.write(f'{(self.__FALTAS.iloc[a,3])},D,82695,6,I,{str(mes)+str(ano)},"{self.__swicth(int(self.__FALTAS.iloc[a,2])*10)}","SCE,Folha de ponto e ficha F","{self.__FALTAS.iloc[a,4]}"\n')
        else:
            for c in range(len(self.__FALTAS.iloc[:,0])):
                if str(self.__FALTAS.iloc[c,2]) != 'nan':
                    for d in range(len(self.__SCE.iloc[:,6])):
                        if self.__conversor_de_cpf(str(self.__SCE.iloc[d,6])) == self.__conversor_de_cpf(str(self.__FALTAS.iloc[c,1])):
                            self.__SAIDA.write(f'{(self.__FALTAS.iloc[c,3])},D,83172,6,I,{str(mes)+str(ano)},"{self.__swicth(int(self.__FALTAS.iloc[c,2])*float(str(self.__SCE.iloc[d,22]).replace(',','.')))}","SCE,Folha de ponto e ficha F","{self.__FALTAS.iloc[c,4]}"\n')

    def __gerar_saida(self):
        self.__SAIDA.close()

        os.rename('utils/analise_de_faltas/dados/saida.txt','utils/analise_de_faltas/dados/dadosFPATMOVFIN_V3_REF.csv')

        Hob.Hob()
    
    def __limpar_dados(self):
        os.remove('utils/analise_de_faltas/dados/dadosFPATMOVFIN_V3_REF.csv')
        os.remove('utils/analise_de_faltas/dados/Faltas.xlsx')
    
    def iniciar(self, escolha, mes, ano):
        self.__gerar_dados(escolha, mes, ano)
        self.__gerar_saida()
        self.__limpar_dados()