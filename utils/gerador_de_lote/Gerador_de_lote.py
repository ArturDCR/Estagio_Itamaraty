import pandas as pd
from datetime import datetime
import os

class Gerador_de_lote():
    def __init__(self, caminho_planilha_mre):
        self.__mre = pd.read_excel(caminho_planilha_mre)
        self.__siape = []

    def __conversor_de_cpf(self, cpf):
        cpf_limpo = ''.join(filter(str.isdigit, str(cpf)))
        return cpf_limpo.zfill(11)

    def __gerar_dados(self, cpf, ano_inicial, ano_final):
        cpf_limpo = self.__conversor_de_cpf(cpf)

        self.__siape = self.__mre.iloc[:, 1].tolist()
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"lote_{timestamp}.txt"
        caminho_saida = os.path.join(os.path.expanduser("~"), 'Downloads', nome_arquivo)

        with open(caminho_saida, 'w') as arquivo_saida:
            arquivo_saida.write(f'035000{ano_final}{datetime.now().month:02d}000\n')

            for inserir in self.__siape:
                arquivo_saida.write(f'1{cpf_limpo}35000{inserir}{ano_inicial}{ano_final}1\n')
            
            quantidade_siape = str(len(self.__siape)).zfill(6)
            arquivo_saida.write(f'9{quantidade_siape}')

    def __limpar_dados(self):
        self.__siape.clear()

    def iniciar(self, cpf, ano_inicial, ano_final):
        self.__gerar_dados(cpf, ano_inicial, ano_final)
        self.__limpar_dados()