import pandas as pd
import os
from datetime import datetime

class Conferencia_ciee:
    def __init__(self):
        self.__MRE = pd.read_excel(os.path.join('utils/conferencia_ciee/dados', 'Mre.xlsx'))
        self.__SCE = pd.read_excel(os.path.join('utils/conferencia_ciee/dados', 'Sce.xlsx'))
        self.__CIEE = pd.read_excel(os.path.join('utils/conferencia_ciee/dados', 'Ciee.xlsx'))
        self.__EXIT_PATH = os.path.join(os.path.join(os.path.expanduser('~'), 'Downloads'), f"Resultado conferencia CIEE {datetime.now().strftime('%Y_%m_%d')}.xlsx")

        self.__cpf_ciee = []
        self.__cpf_mre = []
        self.__cpf_sce = []
            
        self.__dados = {
            'nome' :[],
            'cpf': [],
            'estado': []
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
            
    def __switch(self, estado, cpf):
        if estado == 'fora da base':
            for x in range(len(self.__CIEE.iloc[:,4])):
                if self.__conversor_de_cpf(str(self.__CIEE.iloc[x,4])) == cpf:
                    self.__dados['nome'].append(self.__CIEE.iloc[x,3])
                    self.__dados['cpf'].append(cpf)
                    self.__dados['estado'].append('Fora da base de dados')
        elif estado == 'desligado':
            for z in range(len(self.__SCE.iloc[:,6])):
                if self.__conversor_de_cpf(str(self.__SCE.iloc[z,6])) == cpf:
                    for c in range(len((self.__CIEE.iloc[:,4]))):
                        if self.__conversor_de_cpf(str(self.__CIEE.iloc[c,4])) == cpf:
                            self.__dados['nome'].append(self.__CIEE.iloc[c,3])
                            self.__dados['cpf'].append(cpf)
                            self.__dados['estado'].append(f"Desligado em {self.__SCE.iloc[z,28]}")
        elif estado == 'inicio':
            for z in range(len(self.__SCE.iloc[:,6])):
                if self.__conversor_de_cpf(str(self.__SCE.iloc[z,6])) == cpf:
                    for d in range(len((self.__CIEE.iloc[:,4]))):
                        if self.__conversor_de_cpf(str(self.__CIEE.iloc[d,4])) == cpf:
                            self.__dados['nome'].append(self.__CIEE.iloc[d,3])
                            self.__dados['cpf'].append(cpf)
                            self.__dados['estado'].append(f"Inicio em {str(self.__SCE.iloc[z,23]).split(' a ')[0]}")                    
        else:
            for y in range(len(self.__SCE.iloc[:,6])):
                if self.__conversor_de_cpf(str(self.__SCE.iloc[y,6])) == cpf:
                    for c in range(len((self.__CIEE.iloc[:,4]))):
                        if self.__conversor_de_cpf(str(self.__CIEE.iloc[c,4])) == cpf:
                            self.__dados['nome'].append(self.__CIEE.iloc[c,3])
                            self.__dados['cpf'].append(cpf)
                            self.__dados['estado'].append('Ativo')

    def __gerar_dados(self):
        for a in range(len(self.__CIEE.iloc[:,4])):
            if str(self.__CIEE.iloc[a,4]) != 'nan' and str(self.__CIEE.iloc[a,4]) != 'CPF':
                self.__cpf_ciee.append(self.__conversor_de_cpf(str(self.__CIEE.iloc[a,4])))
        
        for b in range(len(self.__MRE.iloc[:,3])):
            if str(self.__MRE.iloc[b,3]) != 'nan' and str(self.__MRE.iloc[b,3]) != 'CPF':
                self.__cpf_mre.append(self.__conversor_de_cpf(str(self.__MRE.iloc[b,3])))
        
        for c in range(len(self.__SCE.iloc[:,6])):
            if str(self.__SCE.iloc[c,6]) != 'nan' and str(self.__SCE.iloc[c,6]) != 'CPF':
                self.__cpf_sce.append(self.__conversor_de_cpf(str(self.__SCE.iloc[c,6])))
        
        for d in self.__cpf_ciee:
            if d in self.__cpf_mre:
                self.__switch('ativo',d)
            elif d in self.__cpf_sce:
                for a in range(len(self.__SCE.iloc[:,6])):
                    if d == self.__conversor_de_cpf(str(self.__SCE.iloc[a,6])):
                        if str(self.__SCE.iloc[a,28]).__contains__('/'):
                            self.__switch('desligado',d)
                        else:
                            self.__switch('inicio',d)
            else:
                self.__switch('fora da base',d)

    def __gerar_saida(self):
        pd.DataFrame(self.__dados).to_excel(self.__EXIT_PATH, index=False)

    def __limpar_listas(self):
        for chave in self.__dados:
            self.__dados[chave].clear()
        
        self.__cpf_ciee.clear()
        self.__cpf_mre.clear()
        self.__cpf_sce.clear()
    
    def __limpar_duplicados(self):
        for k in self.__dados['cpf']:
            index = self.__dados['cpf'].index(k)
            if self.__dados['cpf'].count(k) > 1:
                self.__dados['nome'].pop(index)
                self.__dados['cpf'].pop(index)
                self.__dados['estado'].pop(index)
                self.__limpar_duplicados()
    
    def iniciar(self):
        self.__gerar_dados()
        self.__limpar_duplicados()
        self.__gerar_saida()
        self.__limpar_listas()