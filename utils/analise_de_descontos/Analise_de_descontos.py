import pandas as pd
import os
import PyPDF2
import re
from datetime import datetime

class Analise_de_descontos:
    def __init__(self):
        self.__DESCONTOS = open(os.path.join('utils/data', 'Descontos.pdf'), 'rb')
        self.__PDF = PyPDF2.PdfReader(self.__DESCONTOS)
        self.__EXIT_PATH = os.path.join(os.path.join(os.path.expanduser('~'), 'Downloads'), f"Resultado Analise de Descontos {datetime.now().strftime('%Y_%m_%d')}.xlsx")

        self.__dados = {
            'siape_e_nome': [],
            'desconto_AT' : [],
            'desconto_BE' :[]
        }

        self.__paginas = []
    
    def __alimentar_paginas(self):
        for a in range(len(self.__PDF.pages)):
            pagina = self.__PDF.pages[a]

            texto = pagina.extract_text()

            linhas = texto.split('\n')

            for b in linhas:
                if b.__contains__(f'FICHA FINANCEIRA REFERENTE A'):
                    self.__paginas.append(a)

    def __gerar_dados(self):
        for c in self.__paginas:
            nova_pagina = self.__PDF.pages[int(c)]

            novo_texto = nova_pagina.extract_text()

            novas_linhas = novo_texto.split('\n')

            nome_e_siape = r"(?<=:)\s*(.*?)\s*(?=BANCO)"

            for x in novas_linhas:
                if x.__contains__("BANCO"):
                    if re.search(nome_e_siape,x).group(1) not in self.__dados['siape_e_nome']:
                        self.__dados['siape_e_nome'].append(re.search(nome_e_siape,x).group(1))
                        self.__dados['desconto_AT'].append(0)
                        self.__dados['desconto_BE'].append(0)
                        for r in novas_linhas:
                            if r.__contains__('D E S C O N T O S'):
                                for w in r[41:].split():
                                    if float(w.split(',')[-1]) != 00:
                                        self.__dados['desconto_BE'][int(self.__dados['siape_e_nome'].index(re.search(nome_e_siape,x).group(1)))] += float(w.replace(',','.'))
                                    else:
                                        self.__dados['desconto_AT'][int(self.__dados['siape_e_nome'].index(re.search(nome_e_siape,x).group(1)))] += float(w.replace(',','.'))
                    else:
                        for p in novas_linhas:
                            if p.__contains__('D E S C O N T O S'):
                                for b in p[41:].split():
                                    if float(b.split(',')[-1]) != 00:
                                        self.__dados['desconto_BE'][int(self.__dados['siape_e_nome'].index(re.search(nome_e_siape,x).group(1)))] += float(b.replace(',','.'))
                                    else:
                                        self.__dados['desconto_AT'][int(self.__dados['siape_e_nome'].index(re.search(nome_e_siape,x).group(1)))] += float(b.replace(',','.'))
                                                        
    def __gerar_saida(self):
        pd.DataFrame(self.__dados).to_excel(self.__EXIT_PATH, index=False)
        self.__DESCONTOS.close()

    def __limpar_listas(self):
        for chave in self.__dados:
            self.__dados[chave].clear()
        self.__paginas.clear()
    
    def iniciar(self):
        self.__alimentar_paginas()
        self.__gerar_dados()
        self.__gerar_saida()
        self.__limpar_listas()