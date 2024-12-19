import pandas as pd
from datetime import datetime
import os

class Analise_de_vigencia:
    def __init__(self):
        self.__SCE = pd.read_excel(os.path.join("utils/analise_de_vigencias/dados","Sce.xlsx"))
        self.__VIGENCIA = pd.read_excel(os.path.join("utils/analise_de_vigencias/dados","Vigentes.xlsx"))
        self.__EXIT_PATH = os.path.join(os.path.join(os.path.expanduser("~"), 'Downloads'), f"Resultado Analise de Vigencias {datetime.now().strftime('%Y_%m_%d')}.xlsx")

        self.__dados = {
            'Nome' : [],
            'CPF' : []
        }

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

    def __gerador_de_dados(self):
        for a in range (len(self.__SCE.iloc[:,28])):
            if str(self.__SCE.iloc[a,23]) != 'nan' and str(self.__SCE.iloc[a,23]) != 'Vigência do  Contrato':
                if str(self.__SCE.iloc[a,28]) == 'nan' and datetime.strptime(str(self.__SCE.iloc[a,23]).split(' a ')[-1],"%d/%m/%Y") >= datetime.strptime(str(datetime.now().strftime("%d/%m/%Y")),"%d/%m/%Y"):
                    self.__dados['Nome'].append(self.__SCE.iloc[a,4])
                    self.__dados['CPF'].append(self.__conversor_de_cpf(str(self.__SCE.iloc[a,6])))
                else:
                    if str(self.__SCE.iloc[a,13]) == 'Ativo' and str(self.__SCE.iloc[a,28]) == 'nan':
                        self.__dados['Nome'].append(self.__SCE.iloc[a,4])
                        self.__dados['CPF'].append(self.__conversor_de_cpf(str(self.__SCE.iloc[a,6])))

    def __gerador_de_saida(self):
        pd.DataFrame(self.__dados).to_excel(self.__EXIT_PATH, index=False)

    def __limpar_listas(self):
        for chave in self.__dados:
            self.__dados[chave].clear()
    
    def __limpar_duplicados(self):
        for a in self.__dados['CPF']:
            index = self.__dados['CPF'].index(a)
            if self.__dados['CPF'].count(a) > 1:
                self.__dados['Nome'].pop(index)
                self.__dados['CPF'].pop(index)
                self.__limpar_duplicados()
    
    def iniciar(self):
        self.__gerador_de_dados()
        self.__limpar_duplicados()
        self.__gerador_de_saida()
        self.__limpar_listas()