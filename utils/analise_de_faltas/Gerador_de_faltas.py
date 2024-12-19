import pandas as pd
import os
import shutil

class Gerador_de_faltas:
    def __init__(self, mes, ano):
        self.__FORMS = pd.read_excel(os.path.join('utils/analise_de_faltas/dados/Forms.xlsx'))
        self.__SCE = pd.read_excel(os.path.join('utils/analise_de_faltas/dados/Sce.xlsx'))
        self.__MRE = pd.read_excel(os.path.join('utils/analise_de_faltas/dados/Mre.xlsx'))

        self.__dados = {
            'nome': [],
            'cpf': [],
            'dias_justificados': [],
            'dias_injustificados': [],
            'siape': []
        }

        self.__gerar_dados(mes, ano)
        self.__gerar_saida()
        self.__limpar_listas()

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
            
    def __gerar_dados(self, mes, ano):
        for p in range(len(self.__FORMS.iloc[:,17])):
            if str(self.__FORMS.iloc[p,8]).__contains__(mes) and str(self.__FORMS.iloc[p,8]).__contains__(ano):
                if str(self.__FORMS.iloc[p,23]) != 'nan':
                    self.__dados['nome'].append(str(self.__FORMS.iloc[p,11]).split('|')[0])
                    self.__dados['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[p,11]).split(' | ')[-1]))
                    self.__dados['dias_justificados'].append(len(self.__FORMS.iloc[p,23].split(';')))
                    if str(self.__FORMS.iloc[p,32]) != 'nan':
                        self.__dados['dias_injustificados'].append(len(self.__FORMS.iloc[p,32].split(';')))
                    else:
                        self.__dados['dias_injustificados'].append(None)
                else:
                    if str(self.__FORMS.iloc[p,32]) != 'nan':               
                        self.__dados['nome'].append(str(self.__FORMS.iloc[p,11]).split('|')[0])
                        self.__dados['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[p,11]).split(' | ')[-1]))
                        self.__dados['dias_justificados'].append(None)
                        self.__dados['dias_injustificados'].append(len(self.__FORMS.iloc[p,32].split(';')))

        for c in self.__dados['cpf']:
                for b in range(len(self.__MRE.iloc[:,4])):
                    if c == self.__conversor_de_cpf(str(self.__MRE.iloc[b,4])):
                        self.__dados['siape'].append(self.__MRE.iloc[b,2])

        for e in self.__dados['cpf']:
            for d in range(len(self.__SCE.iloc[:,6])):
                if self.__conversor_de_cpf(str(self.__SCE.iloc[d,6])) == e:
                    if str(self.__SCE.iloc[d,28]) != 'nan':
                        index = self.__dados['cpf'].index(e)
                        self.__dados['nome'].pop(index)
                        self.__dados['cpf'].pop(index)
                        self.__dados['siape'].pop(index)
                        self.__dados['dias_justificados'].pop(index)
                        self.__dados['dias_injustificados'].pop(index)
    
    def __gerar_saida(self):
        pd.DataFrame(self.__dados).to_excel('utils/analise_de_faltas/dados/Faltas.xlsx',index=False)
        shutil.copy('utils/analise_de_faltas/dados/Faltas.xlsx', os.path.join(os.path.expanduser("~"), 'Downloads'))

    def __limpar_listas(self):
        for chave in self.__dados:
            self.__dados[chave].clear()