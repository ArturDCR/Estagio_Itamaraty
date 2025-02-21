from docx import Document as dc
import os
import pandas as pd
import datetime

class Gerador_de_declaracao:
    def __init__(self):
        self.__modelo = dc(os.path.join('utils/gerador_de_declaracao/dados', 'Modelo.docx'))
        self.__SCE = pd.read_excel(os.path.join('utils/gerador_de_declaracao/dados', 'Sce.xlsx'))
 

        self.__dia = datetime.datetime.now().day
        self.__mes = datetime.datetime.now().strftime('%B')
        self.__ano = datetime.datetime.now().year

        self.__meses = {'January': 'janeiro',
                        'February': 'fevereiro',
                        'March': 'março',
                        'April': 'abril',
                        'May': 'maio',
                        'June': 'junho',
                        'July': 'julho',
                        'August': 'agosto',
                        'September': 'setembro',
                        'October': 'outubro',
                        'November': 'novembro', 
                        'December': 'dezembro'
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

    def __gerar_dados(self, cpf):
        for dados in range(len(self.__SCE.iloc[:,6])):
            if self.__conversor_de_cpf(str(self.__SCE.iloc[dados,6])) == self.__conversor_de_cpf(str(cpf)):
                for linhas in self.__modelo.paragraphs:
                    if 'NOME' and 'CPFZ' and 'FACULDADE' and 'SETOR' and 'CURSO' and 'DATA' and 'CH' and 'CS' in linhas.text:
                        linhas.text = linhas.text.replace('NOME', self.__SCE.iloc[dados,4])
                        linhas.text = linhas.text.replace('CPFZ', self.__SCE.iloc[dados,6])
                        linhas.text = linhas.text.replace('FACULDADE', self.__SCE.iloc[dados,12])
                        linhas.text = linhas.text.replace('SETOR', self.__SCE.iloc[dados,2])
                        linhas.text = linhas.text.replace('CURSO', self.__SCE.iloc[dados,11])
                        if not str(self.__SCE.iloc[dados,28]).__contains__('/'):
                            linhas.text = linhas.text.replace('DATA', self.__SCE.iloc[dados,23])
                        else:
                            linhas.text = linhas.text.replace('DATA', f'{self.__SCE.iloc[dados,23].split("a")[0]}a {self.__SCE.iloc[dados,28]}')
                        linhas.text = linhas.text.replace('CH', self.__SCE.iloc[dados,16].replace('H',''))
                        if int(self.__SCE.iloc[dados,16].replace('H','')) == 6:
                            linhas.text = linhas.text.replace('CS', str(30))
                        else:
                            linhas.text = linhas.text.replace('CS', str(20))
                    elif 'ATUAL' in linhas.text:
                        linhas.text = linhas.text.replace('ATUAL', f'{self.__dia} de {self.__meses[self.__mes]} de {self.__ano}')
                        self.__modelo.save(os.path.join(os.path.join(os.path.expanduser('~'), 'Downloads'), f'{self.__SCE.iloc[dados,4]}.docx'))
                        break

    def iniciar(self, cpf):
        self.__gerar_dados(cpf)