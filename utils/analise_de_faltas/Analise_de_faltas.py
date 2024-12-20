import pandas as pd
import os

from utils.analise_de_faltas import Gerador_de_faltas as gerador
from utils.analise_de_faltas import Hob

class Analise_de_faltas:
    def __init__(self):

        self.__FALTAS = pd.read_excel(os.path.join("utils/analise_de_faltas/dados","Faltas.xlsx"))
        self.__SAIDA = open('utils/analise_de_faltas/dados/saida.txt', 'w')

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

    def __gerar_dados(self, escolha, mes, ano):
        if escolha == 'VT':
            for a in range(len(self.__FALTAS.iloc[:,0])):
                if str(self.__FALTAS.iloc[a,2]) != 'nan' and str(self.__FALTAS.iloc[a,3]) != 'nan':
                    self.__SAIDA.write(f'{(self.__FALTAS.iloc[a,4])},D,82695,6,I,{str(mes)+str(ano)},"{self.__swicth((int(self.__FALTAS.iloc[a,2])*10) + int(self.__FALTAS.iloc[a,3])*10)}","SCE,Folha de ponto e ficha F","Dias"\n')
                elif str(self.__FALTAS.iloc[a,2]) != 'nan':
                    self.__SAIDA.write(f'{(self.__FALTAS.iloc[a,4])},D,82695,6,I,{str(mes)+str(ano)},"{self.__swicth(int(self.__FALTAS.iloc[a,2])*10)}","SCE,Folha de ponto e ficha F","Dias"\n')
                else:
                    self.__SAIDA.write(f'{(self.__FALTAS.iloc[a,4])},D,82695,6,I,{str(mes)+str(ano)},"{self.__swicth(int(self.__FALTAS.iloc[a,3])*10)}","SCE,Folha de ponto e ficha F","Dias"\n')
        else:
            if str(self.__FALTAS.iloc[a,3]) != 'nan':
                self.__SAIDA.write(f'{(self.__FALTAS.iloc[a,4])},D,83172,6,I,{str(mes)+str(ano)},"{self.__swicth(int(self.__FALTAS.iloc[a,3])*10)}","SCE,Folha de ponto e ficha F","Dias"\n')
    
    def __gerar_saida(self):
        self.__SAIDA.close()

        os.rename('utils/analise_de_faltas/dados/saida.txt','utils/analise_de_faltas/dados/dadosFPATMOVFIN_V3_REF.csv')

        Hob.Hob()
    
    def iniciar(self, escolha, mes, ano):
        gerador.Gerador_de_faltas(mes,ano)
        self.__gerar_dados(escolha, mes, ano)
        self.__gerar_saida()