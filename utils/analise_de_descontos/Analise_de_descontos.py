import pandas as pd
import os
import PyPDF2
from datetime import datetime

class Analise_de_descontos:
    def __init__(self):
        self.__DESCONTOS = open(os.path.join('utils/analise_de_descontos/dados', 'Descontos.pdf'), 'rb')
        self.__PDF = PyPDF2.PdfReader(self.__DESCONTOS)
        self.__EXIT_PATH = os.path.join(os.path.join(os.path.expanduser('~'), 'Downloads'), f"Resultado Analise de Descontos {datetime.now().strftime('%Y_%m_%d')}.xlsx")

        self.__dados = {
            'siape_e_nome': [],
            'ano' : [],
            'emissao' : [],
            'desconto' : []
        }

        self.__paginas = []
    
    def __alimentar_paginas(self, ano):
        for a in range(len(self.__PDF.pages)):
            pagina = self.__PDF.pages[a]

            texto = pagina.extract_text()

            linhas = texto.split('\n')

            for b in linhas:
                if b.__contains__(f'FICHA FINANCEIRA REFERENTE A {ano}'):
                    self.__paginas.append(a)

    def __gerar_dados(self):
        for c in self.__paginas:
            nova_pagina = self.__PDF.pages[int(c)]

            novo_texto = nova_pagina.extract_text()

            novas_linhas = novo_texto.split('\n')
            print(novas_linhas)

            for d in novas_linhas:
                if d.__contains__('D E S C O N T O S'):
                    if d[41:].split() != []:
                        descontos = d[41:].split()
                        aux = 0
                        for e in range(len(descontos)):
                            aux = aux + float(descontos[e].replace(',','.'))
                        if str(novas_linhas[6])[15:79] in self.__dados['siape_e_nome']:
                            index = self.__dados['siape_e_nome'].index(str(novas_linhas[6])[15:79])
                            self.__dados['desconto'][index] += aux
                        else:
                            self.__dados['siape_e_nome'].append(str(novas_linhas[6])[15:79])
                            self.__dados['ano'].append(str(novas_linhas[2][:33]))
                            self.__dados['emissao'].append(novas_linhas[1][109:])
                            self.__dados['desconto'].append(aux)
    def __gerar_saida(self):
        pd.DataFrame(self.__dados).to_excel(self.__EXIT_PATH, index=False)
        self.__DESCONTOS.close()

    def __limpar_listas(self):
        for chave in self.__dados:
            self.__dados[chave].clear()

        self.__paginas.clear()
    
    def iniciar(self, ano):
        self.__alimentar_paginas(ano)
        self.__gerar_dados()
        self.__gerar_saida()
        self.__limpar_listas()