import pandas as pd
import os
from datetime import datetime

class Conferencia_ciee:
    def __init__(self, caminho_mre, caminho_sce, caminho_ciee):
        self.__MRE = pd.read_excel(caminho_mre)
        self.__SCE = pd.read_excel(caminho_sce)
        self.__CIEE = pd.read_excel(caminho_ciee)

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        nome_arquivo = f"Resultado_conferencia_CIEE_{timestamp}.xlsx"
        self.__EXIT_PATH = os.path.join(os.path.expanduser('~'), 'Downloads', nome_arquivo)

        self.__cpf_ciee = []
        self.__cpf_mre = []
        self.__cpf_sce = []
            
        self.__dados = {
            'nome' :[],
            'cpf': [],
            'estado': []
        }
    
    def __conversor_de_cpf(self, cpf):
        cpf_limpo = ''.join(filter(str.isdigit, str(cpf)))
        return cpf_limpo.zfill(11)
            
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
        df_saida = pd.DataFrame(self.__dados)
        df_saida = df_saida.drop_duplicates(subset=['cpf']) 
        df_saida.to_excel(self.__EXIT_PATH, index=False)

    def __limpar_listas(self):
        for chave in self.__dados:
            self.__dados[chave].clear()
        
        self.__cpf_ciee.clear()
        self.__cpf_mre.clear()
        self.__cpf_sce.clear()
    
    def iniciar(self):
        self.__gerar_dados()
        self.__gerar_saida()
        self.__limpar_listas()