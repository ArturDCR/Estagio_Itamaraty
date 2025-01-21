import pandas as pd
import os

from utils.analise_de_faltas import Hob

class Analise_de_faltas:
    def __init__(self):
        self.__FALTAS = pd.read_excel(os.path.join("utils/analise_de_faltas/dados","Faltas.xlsx"))
        self.__SAIDA = open('utils/analise_de_faltas/dados/saida.txt', 'w')

    def __swicth (self, aux):
        if str(aux).__contains__('.'):
            if len(str(aux).split('.')[-1]) == 2:
                return f'000000{str(aux).replace('.',',')}'
            elif len(str(aux).split('.')[-1]) == 3:
                return f'00000{str(aux).replace('.',',')}'
            elif len(str(aux).split('.')[-1]) == 4:
                return f'0000{str(aux).replace('.',',')}'
            elif len(str(aux).split('.')[-1]) == 5:
                return f'000{str(aux).replace('.',',')}'
            elif len(str(aux).split('.')[-1]) == 6:
                return f'00{str(aux).replace('.',',')}'
            elif len(str(aux).split('.')[-1]) == 7:
                return f'0{str(aux).replace('.',',')}'
            elif len(str(aux).split('.')[-1]) == 8:
                return f'{str(aux).replace('.',',')}'
        else:
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
                if str(self.__FALTAS.iloc[a,1]) != 'NÃO ENCONTRADO':
                    if str(self.__FALTAS.iloc[a,2]) != 'nan':
                        self.__SAIDA.write(f'{(self.__FALTAS.iloc[a,3])},D,82695,6,I,{str(mes)[:3]+str(ano)},"{self.__swicth(self.__FALTAS.iloc[a,5])}","SCE,Folha de ponto e ficha F","{self.__FALTAS.iloc[a,4]}"\n')
                else:
                    pass       
        else:
            for c in range(len(self.__FALTAS.iloc[:,0])):
                if str(self.__FALTAS.iloc[c,1]) != 'NÃO ENCONTRADO':
                    if str(self.__FALTAS.iloc[c,2]) != 'nan':
                        self.__SAIDA.write(f'{(self.__FALTAS.iloc[c,3])},D,83172,6,I,{str(mes)[:3]+str(ano)},"{self.__swicth(self.__FALTAS.iloc[c,5])}","SCE,Folha de ponto e ficha F","{self.__FALTAS.iloc[c,4]}"\n')
                else:
                    pass

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