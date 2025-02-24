import pandas as pd
from datetime import datetime
import os
import shutil

class Gerador_de_lote():
    def __init__(self):
        self.__saida = open('utils/data/lote.txt', 'w')
        self.__mre = pd.read_excel('utils/data/Mre.xlsx')
        self.__cpf = None
        self.__siape = []

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

    def __swicth (self, aux):
        while len(str(aux)) != 6:
            aux = str(0)+str(aux)
        return aux

    def __gerar_dados(self, cpf, ano_inicial, ano_final):
        self.__cpf = self.__conversor_de_cpf(cpf)

        for siape in range(len(self.__mre.iloc[:,0])):
            self.__siape.append(self.__mre.iloc[siape,1])
        
        self.__saida.write(f'035000{ano_final}{ datetime.now().month:02d}000\n')

        for inserir in self.__siape:
            self.__saida.write(f'1{self.__cpf}3500{inserir}{ano_inicial}{ano_final}1\n')
        
        self.__saida.write(f'9{self.__swicth(len(self.__siape))}')

        self.__saida.close()
    
    def __gerar_saida(self):
        shutil.copy('utils/data/lote.txt', os.path.join(os.path.expanduser("~"), 'Downloads'))
        os.remove('utils/data/lote.txt')
    
    def __limpar_dados(self):
        self.__cpf = None
        self.__siape.clear()

    def iniciar(self, cpf, ano_inicial, ano_final):
        self.__gerar_dados(cpf, ano_inicial, ano_final)
        self.__gerar_saida()
        self.__limpar_dados()