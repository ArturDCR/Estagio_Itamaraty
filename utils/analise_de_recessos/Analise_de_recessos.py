import pandas as pd
import os
from datetime import datetime

class Analise_de_recessos:
    def __init__(self):
        self.__RECESSOS = pd.read_excel(os.path.join("utils/data","Recessos.xlsx"))
        self.__EXIT_PATH = os.path.join(os.path.join(os.path.expanduser("~"), 'Downloads'), f"Resultado Analise de Recessos {datetime.now().strftime('%Y_%m_%d')}.xlsx")

        self.__dados = {
            'nome': [],
            'cpf': [],
            'dias_de_recesso': []
        }

        self.__auxiliar = []

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
    
    def __gerar_dados(self, cpf):
        for a in range(len(self.__RECESSOS.iloc[:,7])):
            if str(self.__RECESSOS.iloc[a,7]).__contains__('CPF'):
                if self.__conversor_de_cpf(str(self.__RECESSOS.iloc[a,7]).split(': ')[-1]) == self.__conversor_de_cpf(str(cpf)):
                    if self.__RECESSOS.iloc[a+2,5] == 'Aprovado':
                        self.__dados['nome'].append(str(self.__RECESSOS.iloc[a,1]).split(' : ')[-1])
                        self.__dados['cpf'].append(self.__conversor_de_cpf(self.__conversor_de_cpf(str(self.__RECESSOS.iloc[a,7]).split(': ')[-1])))
                        aux = a+2
                        while not str(self.__RECESSOS.iloc[aux,1]).__contains__('Total'):
                            diferenca = datetime.strptime(self.__RECESSOS.iloc[aux,8], "%d/%m/%Y") - datetime.strptime(self.__RECESSOS.iloc[aux,7], "%d/%m/%Y")
                            self.__auxiliar.append(diferenca.days)
                            aux += 1
                        if len(self.__auxiliar) > 1:
                            deux = 0
                            for a in self.__auxiliar:
                                deux += a
                            self.__dados['dias_de_recesso'].append(deux)
                        else:
                            self.__dados['dias_de_recesso'].append(self.__auxiliar[0])

    def __gerar_saida(self):
        pd.DataFrame(self.__dados).to_excel(self.__EXIT_PATH, index=False)
    
    def iniciar(self, cpf):
        self.__gerar_dados(cpf)
        self.__gerar_saida()