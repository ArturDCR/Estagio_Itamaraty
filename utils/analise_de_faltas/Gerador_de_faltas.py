import pandas as pd
import os
import shutil

from utils.analise_de_faltas.Analise_de_faltas import Analise_de_faltas

class Gerador_de_faltas:
    def __init__(self):
        self.__FORMS = pd.read_excel(os.path.join('utils/analise_de_faltas/dados/Forms.xlsx'))
        self.__SCE = pd.read_excel(os.path.join('utils/analise_de_faltas/dados/Sce.xlsx'))
        self.__MRE = pd.read_excel(os.path.join('utils/analise_de_faltas/dados/Mre.xlsx'))

        self.__dados_VT = {
            'nome': [],
            'cpf': [],
            'dias_justificados': [],
            'siape': []
        }
        
        self.__dados_BE = {
            'nome': [],
            'cpf': [],
            'dias_injustificados': [],
            'siape': []
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
            
    def __gerar_dados(self,escolha, mes, ano):
        for p in range(len(self.__FORMS.iloc[:,17])):
            if str(self.__FORMS.iloc[p,8]).__contains__(mes) and str(self.__FORMS.iloc[p,8]).__contains__(ano):
                if escolha == 'VT':
                    if str(self.__FORMS.iloc[p,23]) != 'nan':
                        if str(self.__FORMS.iloc[p,11]) == 'NÃO ENCONTRADO':
                            self.__dados_VT['nome'].append('NÃO ENCONTRADO')
                            self.__dados_VT['cpf'].append('NÃO ENCONTRADO')
                        else:
                            self.__dados_VT['nome'].append(str(self.__FORMS.iloc[p,11]).split('|')[0])
                            self.__dados_VT['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[p,11]).split(' | ')[-1]))
                        if str(self.__FORMS.iloc[p,23]).__contains__(';'):
                            self.__dados_VT['dias_justificados'].append(len(self.__FORMS.iloc[p,23].split(';')))
                        else:
                            self.__dados_VT['dias_justificados'].append(self.__FORMS.iloc[p,23])
                else:
                    if str(self.__FORMS.iloc[p,32]) != 'nan':               
                        if str(self.__FORMS.iloc[p,11]) == 'NÃO ENCONTRADO':
                            self.__dados_BE['nome'].append('NÃO ENCONTRADO')
                            self.__dados_BE['cpf'].append('NÃO ENCONTRADO')
                        else:
                            self.__dados_BE['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[p,11]).split(' | ')[-1]))
                            self.__dados_BE['nome'].append(str(self.__FORMS.iloc[p,11]).split('|')[0])
                        if str(self.__FORMS.iloc[p,32]).__contains__(';'):
                            self.__dados_BE['dias_injustificados'].append(len(self.__FORMS.iloc[p,32].split(';')))
                        else:
                            self.__dados_BE['dias_injustificados'].append(self.__FORMS.iloc[p,32])

        for c in self.__dados_VT['cpf']:
                for b in range(len(self.__MRE.iloc[:,3])):
                    if c == self.__conversor_de_cpf(str(self.__MRE.iloc[b,3])):
                        self.__dados_VT['siape'].append(self.__MRE.iloc[b,1])
                    elif c == 'NÃO ENCONTRADO':
                        self.__dados_VT['siape'].append('xxx')
                        
        for c in self.__dados_BE['cpf']:
                for b in range(len(self.__MRE.iloc[:,3])):
                    if c == self.__conversor_de_cpf(str(self.__MRE.iloc[b,3])):
                        self.__dados_BE['siape'].append(self.__MRE.iloc[b,1])
                    elif c == 'NÃO ENCONTRADO':
                        self.__dados_BE['siape'].append('xxx')

        for e in self.__dados_VT['cpf']:
            for d in range(len(self.__SCE.iloc[:,6])):
                if self.__conversor_de_cpf(str(self.__SCE.iloc[d,6])) == e:
                    if str(self.__SCE.iloc[d,28]) != 'nan':
                        index = self.__dados_VT['cpf'].index(e)
                        self.__dados_VT['nome'].pop(index)
                        self.__dados_VT['cpf'].pop(index)
                        self.__dados_VT['siape'].pop(index)
                        self.__dados_VT['dias_justificados'].pop(index)
                        self.__dados_VT['dias_injustificados'].pop(index)
        
        for e in self.__dados_BE['cpf']:
            for d in range(len(self.__SCE.iloc[:,6])):
                if self.__conversor_de_cpf(str(self.__SCE.iloc[d,6])) == e:
                    if str(self.__SCE.iloc[d,28]) != 'nan':
                        index = self.__dados_BE['cpf'].index(e)
                        self.__dados_BE['nome'].pop(index)
                        self.__dados_BE['cpf'].pop(index)
                        self.__dados_BE['siape'].pop(index)
                        self.__dados_BE['dias_justificados'].pop(index)
                        self.__dados_BE['dias_injustificados'].pop(index)
    
    def __gerar_saida(self, escolha):
        if escolha == 'VT':
            pd.DataFrame(self.__dados_VT).to_excel('utils/analise_de_faltas/dados/Faltas.xlsx',index=False)
            shutil.copy('utils/analise_de_faltas/dados/Faltas.xlsx', os.path.join(os.path.expanduser("~"), 'Downloads'))
        else:
            pd.DataFrame(self.__dados_BE).to_excel('utils/analise_de_faltas/dados/Faltas.xlsx',index=False)
            shutil.copy('utils/analise_de_faltas/dados/Faltas.xlsx', os.path.join(os.path.expanduser("~"), 'Downloads'))           

    def __limpar_listas(self):
        for chave in self.__dados_BE:
            self.__dados_BE[chave].clear()
        for key in self.__dados_VT:
            self.__dados_VT[chave].clear()
    
    def iniciar(self, escolha, mes, ano):
        self.__gerar_dados(escolha, mes, ano)
        self.__gerar_saida(escolha)
        self.__limpar_listas()
        analise_faltas = Analise_de_faltas()
        analise_faltas.iniciar(escolha, mes, ano)