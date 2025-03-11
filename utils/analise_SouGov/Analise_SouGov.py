import pandas as pd
import os
from datetime import datetime

class Analise_SouGov:
    def __init__(self):
        self.__TA = pd.read_excel(os.path.join("utils/data","Tas.xlsx"))
        self.__EMAILS = pd.read_excel(os.path.join("utils/data","Emails.xlsx"))
        self.__SCE = pd.read_excel(os.path.join("utils/data","Sce.xlsx"))
        self.__EXIT_PATH = os.path.join(os.path.join(os.path.expanduser("~"), 'Downloads'), f"Resultado Analise SouGov {datetime.now().strftime('%Y_%m_%d')}.xlsx")

        self.__aux = []

        self.__dados = {
            'nome' : [],
            'cpf' : [],
            'email' : []
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

    def __gerar_dados(self, dia, mes, ano):    
        for j in range(len(self.__EMAILS.iloc[:,0])):
            self.__aux.append(str(self.__EMAILS.iloc[j,0]).split('"')[0].lower())

        for a in range(len(self.__TA.iloc[:,1])):
            if 'Dt. Cadastro' in str(self.__TA.iloc[a,1]) and int(str(self.__TA.iloc[a+1,1])[6:]) >= int(ano) and int(str(self.__TA.iloc[a+1,1])[3:5]) >= mes:
                if int(str(self.__TA.iloc[a+1,1])[3:5]) == mes:
                    if int(str(self.__TA.iloc[a+1,1])[:2]) >= int(dia):
                        if str(self.__TA.iloc[a-1,1][6:]).lower() in self.__aux:
                            for b in range(len(self.__EMAILS.iloc[:,0])):
                                if str(self.__EMAILS.iloc[b,0]).split('"')[0].lower() == str(self.__TA.iloc[a-1,1][6:]).lower():
                                    for atividade in range(len(self.__SCE.iloc[:,6])):
                                        if self.__conversor_de_cpf(str(self.__SCE.iloc[atividade,6])) == self.__conversor_de_cpf(str(self.__TA.iloc[a-1,5][5:])):
                                            if str(self.__SCE.iloc[atividade,13]) == 'Ativo':
                                                self.__dados['nome'].append(str(self.__TA.iloc[a-1,1][6:]))
                                                self.__dados['cpf'].append(str(self.__conversor_de_cpf(str(self.__TA.iloc[a-1,5][5:]))))
                                                self.__dados['email'].append(str(self.__EMAILS.iloc[b,0]).split('<')[-1])
                        else:
                            for atividade in range(len(self.__SCE.iloc[:,6])):
                                if self.__conversor_de_cpf(str(self.__SCE.iloc[atividade,6])) == self.__conversor_de_cpf(str(self.__TA.iloc[a-1,5][5:])):
                                    if str(self.__SCE.iloc[atividade,13]) == 'Ativo':
                                        self.__dados['nome'].append(str(self.__TA.iloc[a-1,1][6:]))
                                        self.__dados['cpf'].append(str(self.__conversor_de_cpf(str(self.__TA.iloc[a-1,5][5:]))))
                                        self.__dados['email'].append('Email não encontrado')
                else:
                    if str(self.__TA.iloc[a-1,1][6:]).lower() in self.__aux:
                        for b in range(len(self.__EMAILS.iloc[:,0])):
                            if str(self.__EMAILS.iloc[b,0]).split('"')[0].lower() == str(self.__TA.iloc[a-1,1][6:]).lower():
                                for atividade in range(len(self.__SCE.iloc[:,6])):
                                        if self.__conversor_de_cpf(str(self.__SCE.iloc[atividade,6])) == self.__conversor_de_cpf(str(self.__TA.iloc[a-1,5][5:])):
                                            if str(self.__SCE.iloc[atividade,13]) == 'Ativo':
                                                self.__dados['nome'].append(str(self.__TA.iloc[a-1,1][6:]))
                                                self.__dados['cpf'].append(str(self.__conversor_de_cpf(str(self.__TA.iloc[a-1,5][5:]))))
                                                self.__dados['email'].append(str(self.__EMAILS.iloc[b,0]).split('<')[-1])
                    else:
                        for atividade in range(len(self.__SCE.iloc[:,6])):
                            if self.__conversor_de_cpf(str(self.__SCE.iloc[atividade,6])) == self.__conversor_de_cpf(str(self.__TA.iloc[a-1,5][5:])):
                                if str(self.__SCE.iloc[atividade,13]) == 'Ativo':
                                    self.__dados['nome'].append(str(self.__TA.iloc[a-1,1][6:]))
                                    self.__dados['cpf'].append(str(self.__conversor_de_cpf(str(self.__TA.iloc[a-1,5][5:]))))
                                    self.__dados['email'].append('Email não encontrado')

    def __gerar_saida(self):                    
        pd.DataFrame(self.__dados).to_excel(self.__EXIT_PATH,index=False)

    def __limpar_listas(self):
        for chave in self.__dados:
            self.__dados[chave].clear()

        self.__aux.clear()
    
    def iniciar(self, dia, mes, ano):
        self.__gerar_dados(dia, mes, ano)
        self.__gerar_saida()
        self.__limpar_listas()
