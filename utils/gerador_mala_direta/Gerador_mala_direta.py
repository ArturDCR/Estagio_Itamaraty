from docx import Document as dc
import pandas as pd
import os
from datetime import datetime as dt

class Gerador_mala_direta():
    def __init__(self):
        self.__modelo = dc(os.path.join('utils/gerador_mala_direta/dados', 'Modelo.docx'))
        self.__SCE = pd.read_excel(os.path.join('utils/gerador_mala_direta/dados', 'Sce.xlsx'))
    
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
    
    def __gerar_dados(self, cpf, dia, mes , ano):
        for dados in range(len(self.__SCE.iloc[:,6])):
            if self.__conversor_de_cpf(str(self.__SCE.iloc[dados,6])) == self.__conversor_de_cpf(str(cpf)):
                for linhas in self.__modelo.paragraphs:
                    if '«divisão»' and '«DATA»' and '«NOME»' and '«CPF»' and '«SUPERVISOR»' in linhas.text:
                        linhas.text = linhas.text.replace('«divisão»', self.__SCE.iloc[dados,2])
                        linhas.text = linhas.text.replace('«DATA»', f'{dia}/{mes}/{ano}')
                        linhas.text = linhas.text.replace('«NOME»', self.__SCE.iloc[dados,4])
                        linhas.text = linhas.text.replace('«CPF»', self.__SCE.iloc[dados,6])
                        linhas.text = linhas.text.replace('«SUPERVISOR»', self.__SCE.iloc[dados,19])
                    elif 'DD' and 'MM' and 'AA' in linhas.text:
                        linhas.text = linhas.text.replace('DD', str(dt.now().day))
                        linhas.text = linhas.text.replace('MM', str(dt.now().strftime('%m')))
                        linhas.text = linhas.text.replace('AA', str(dt.now().year))
                    elif 'Ano_atual' in linhas.text:
                        linhas.text = linhas.text.replace('Ano_atual', str(dt.now().year))
                        self.__modelo.save(os.path.join(os.path.join(os.path.expanduser('~'), 'Downloads'), f'{self.__SCE.iloc[dados,4]}.docx'))
                        break

    def iniciar(self, cpf, dia, mes, ano):
        self.__gerar_dados(cpf, dia, mes, ano)