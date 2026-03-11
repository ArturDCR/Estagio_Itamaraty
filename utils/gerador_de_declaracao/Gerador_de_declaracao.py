from docx import Document as dc
import os
import pandas as pd
import datetime

from utils.gerenciador_caminhos.gerenciador_caminhos import GerenciadorCaminhos

class Gerador_de_declaracao:
    # Recebendo a planilha via parâmetro (Categoria 2)
    def __init__(self, caminho_planilha_sce):
        self.__caminho_modelo = GerenciadorCaminhos.obter_caminho_recurso(os.path.join('utils', 'data', 'Modelo.docx'))
        
        self.__SCE = pd.read_excel(caminho_planilha_sce)

        agora = datetime.datetime.now()
        self.__dia = agora.day
        self.__mes = agora.strftime('%B')
        self.__ano = agora.year

        self.__meses = {
            'January': 'janeiro', 'February': 'fevereiro', 'March': 'março', 'April': 'abril',
            'May': 'maio', 'June': 'junho', 'July': 'julho', 'August': 'agosto',
            'September': 'setembro', 'October': 'outubro', 'November': 'novembro', 'December': 'dezembro'
        }

    def __conversor_de_cpf(self, cpf):
        cpf_limpo = ''.join(filter(str.isdigit, str(cpf)))
        return cpf_limpo.zfill(11)

    def __gerar_dados(self, cpf):
        cpf_buscado = self.__conversor_de_cpf(cpf)
        
        for dados in range(len(self.__SCE.iloc[:,6])):
            if self.__conversor_de_cpf(str(self.__SCE.iloc[dados,6])) == cpf_buscado:
                
                modelo_atual = dc(self.__caminho_modelo)
                
                for linhas in modelo_atual.paragraphs:
                    linhas.text = linhas.text.replace('NOME', str(self.__SCE.iloc[dados,4]))
                    linhas.text = linhas.text.replace('CPFZ', str(self.__SCE.iloc[dados,6]))
                    linhas.text = linhas.text.replace('FACULDADE', str(self.__SCE.iloc[dados,12]))
                    linhas.text = linhas.text.replace('SETOR', str(self.__SCE.iloc[dados,2]))
                    linhas.text = linhas.text.replace('CURSO', str(self.__SCE.iloc[dados,11]))
                    
                    if not '/' in str(self.__SCE.iloc[dados,28]):
                        linhas.text = linhas.text.replace('DATA', str(self.__SCE.iloc[dados,23]))
                    else:
                        data_formatada = f'{str(self.__SCE.iloc[dados,23]).split("a")[0]}a {self.__SCE.iloc[dados,28]}'
                        linhas.text = linhas.text.replace('DATA', data_formatada)
                    
                    ch_valor = str(self.__SCE.iloc[dados,16]).replace('H','')
                    linhas.text = linhas.text.replace('CH', ch_valor)
                    
                    if ch_valor.isdigit() and int(ch_valor) == 6:
                        linhas.text = linhas.text.replace('CS', '30')
                    else:
                        linhas.text = linhas.text.replace('CS', '20')
                        
                    data_atual_texto = f'{self.__dia} de {self.__meses.get(self.__mes, self.__mes)} de {self.__ano}'
                    linhas.text = linhas.text.replace('ATUAL', data_atual_texto)

                nome_funcionario = str(self.__SCE.iloc[dados,4])
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                nome_arquivo = f"{nome_funcionario}_{timestamp}.docx"
                
                caminho_salvamento = os.path.join(os.path.expanduser('~'), 'Downloads', nome_arquivo)
                
                modelo_atual.save(caminho_salvamento)
                break

    def iniciar(self, cpf):
        self.__gerar_dados(cpf)