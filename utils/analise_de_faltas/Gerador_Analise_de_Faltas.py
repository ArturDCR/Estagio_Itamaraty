import pandas as pd
import os
import tempfile
from datetime import datetime

from utils.analise_de_faltas.Hob import Hob

class Gerador_Analise_Faltas:
    def __init__(self, caminho_forms, caminho_sce, caminho_mre):
        self.__FORMS = pd.read_excel(caminho_forms)
        self.__SCE = pd.read_excel(caminho_sce)
        self.__MRE = pd.read_excel(caminho_mre)

        self.__dados_VT = {
            'nome': [],
            'cpf': [],
            'valor_dias': [],
            'siape': [],
            'dias': [],
            'valor_total': []
        }
        
        self.__dados_BE = {
            'nome': [],
            'cpf': [],
            'valor_dias': [],
            'siape': [],
            'dias': [],
            'valor_total': [],
            'salario' : []
        }

    def __conversor_de_cpf(self, cpf):
        cpf_limpo = ''.join(filter(str.isdigit, str(cpf)))
        return cpf_limpo.zfill(11)

    def __swicth(self, aux):
        # Formatador de valores monetários para a Macro Hob
        aux_str = str(aux)
        if '.' in aux_str:
            if len(aux_str.split('.')[-1]) == 2:
                return f'000000{aux_str.replace(".",",")}'
            elif len(aux_str.split('.')[-1]) == 3:
                return f'00000{aux_str.replace(".",",")}'
            elif len(aux_str.split('.')[-1]) == 4:
                return f'0000{aux_str.replace(".",",")}'
            elif len(aux_str.split('.')[-1]) == 5:
                return f'000{aux_str.replace(".",",")}'
            elif len(aux_str.split('.')[-1]) == 6:
                return f'00{aux_str.replace(".",",")}'
            elif len(aux_str.split('.')[-1]) == 7:
                return f'0{aux_str.replace(".",",")}'
            elif len(aux_str.split('.')[-1]) == 8:
                return f'{aux_str.replace(".",",")}'
            else:
                return aux_str.replace(".",",")
        else:
            if len(aux_str) == 2:
                return f'000000{aux_str},00'
            elif len(aux_str) == 3:
                return f'00000{aux_str},00'
            elif len(aux_str) == 4:
                return f'0000{aux_str},00'
            elif len(aux_str) == 5:
                return f'000{aux_str},00'
            elif len(aux_str) == 6:
                return f'00{aux_str},00'
            elif len(aux_str) == 7:
                return f'0{aux_str},00'
            elif len(aux_str) == 8:
                return f'{aux_str},00'
            else:
                return f'{aux_str},00'

    def __inserir_siape(self, cpf, escolha):
        cpfs = []
        for aux in range(len(self.__MRE.iloc[:,0])):
            cpfs.append(self.__conversor_de_cpf(str(self.__MRE.iloc[aux,3])))
        for siape in range(len(self.__MRE.iloc[:,0])):
            if self.__conversor_de_cpf(self.__MRE.iloc[siape,3]) == cpf:
                if escolha == 'VT':
                    self.__dados_VT['siape'].append(self.__MRE.iloc[siape,1])
                    break
                else:
                    self.__dados_BE['siape'].append(self.__MRE.iloc[siape,1])
                    break
            elif cpf not in cpfs:
                self.__dados_VT['siape'].append('Não consta no MRE')
                self.__dados_BE['siape'].append('Não consta no MRE')
                break
        cpfs.clear()

    def __gerar_dados(self, escolha, mes, ano):
        if escolha == 'VT':
            for VT in range(len(self.__FORMS.iloc[:,0])):
                if str(self.__FORMS.iloc[VT,8]).split()[0] == mes and str(self.__FORMS.iloc[VT,8]).split()[-1] == ano:
                    if str(self.__FORMS.iloc[VT,26]) != 'nan' and str(str(self.__FORMS.iloc[VT,33])) != 'nan':
                        if len(self.__FORMS.iloc[VT,26].split(';')) > 1 and len(str(str(self.__FORMS.iloc[VT,33])).split(';')) > 1:
                            if str(self.__FORMS.iloc[VT,11]) == 'Não Encontrado':
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,14]))
                                self.__dados_VT['cpf'].append('Não encontrado')
                                self.__dados_VT['siape'].append('xxx')
                            else:
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,11]).split(' | ')[0])
                                self.__dados_VT['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]))         
                                self.__inserir_siape(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]),escolha)
                            self.__dados_VT['valor_dias'].append(len(self.__FORMS.iloc[VT,26].split(';')) + len(str(self.__FORMS.iloc[VT,33]).split(';')))
                            self.__dados_VT['dias'].append(f'{str(self.__FORMS.iloc[VT,26]).split(";")} e {str(self.__FORMS.iloc[VT,33]).split(";")} de {mes[:3]}')
                            self.__dados_VT['valor_total'].append((len(self.__FORMS.iloc[VT,26].split(';')) + len(str(self.__FORMS.iloc[VT,33]).split(';')))*10)
                        
                        elif len(self.__FORMS.iloc[VT,26].split(';')) > 1 and len(str(self.__FORMS.iloc[VT,33]).split()) == 1:
                            if str(self.__FORMS.iloc[VT,11]) == 'Não Encontrado':
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,14]))
                                self.__dados_VT['cpf'].append('Não encontrado')
                                self.__dados_VT['siape'].append('xxx')
                            else:
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,11]).split(' | ')[0])
                                self.__dados_VT['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]))                            
                                self.__inserir_siape(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]),escolha)
                            self.__dados_VT['valor_dias'].append(len(self.__FORMS.iloc[VT,26].split(';')) + len(str(self.__FORMS.iloc[VT,33]).split()))
                            self.__dados_VT['dias'].append(f'{str(self.__FORMS.iloc[VT,26]).split(";")} e {str(self.__FORMS.iloc[VT,33]).split()} de {mes[:3]}')
                            self.__dados_VT['valor_total'].append((len(self.__FORMS.iloc[VT,26].split(';')) + len(str(self.__FORMS.iloc[VT,33]).split()))*10)
                        
                        elif len(self.__FORMS.iloc[VT,26].split()) == 1 and len(str(self.__FORMS.iloc[VT,33]).split(';')) > 1:
                            if str(self.__FORMS.iloc[VT,11]) == 'Não Encontrado':
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,14]))
                                self.__dados_VT['cpf'].append('Não encontrado')
                                self.__dados_VT['siape'].append('xxx')
                            else:
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,11]).split(' | ')[0])
                                self.__dados_VT['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]))                        
                                self.__inserir_siape(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]),escolha)
                            self.__dados_VT['valor_dias'].append(len(self.__FORMS.iloc[VT,26].split()) + len(str(self.__FORMS.iloc[VT,33]).split(';')))
                            self.__dados_VT['dias'].append(f'{str(self.__FORMS.iloc[VT,26]).split()} e {str(self.__FORMS.iloc[VT,33]).split(";")} de {mes[:3]}')
                            self.__dados_VT['valor_total'].append((len(self.__FORMS.iloc[VT,26].split()) + len(str(self.__FORMS.iloc[VT,33]).split(';')))*10)
                        
                        elif len(self.__FORMS.iloc[VT,26].split()) == 1 and len(str(self.__FORMS.iloc[VT,33]).split()) == 1:
                            if str(self.__FORMS.iloc[VT,11]) == 'Não Encontrado':
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,14]))
                                self.__dados_VT['cpf'].append('Não encontrado')
                                self.__dados_VT['siape'].append('xxx')
                            else:
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,11]).split(' | ')[0])
                                self.__dados_VT['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]))                          
                                self.__inserir_siape(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]),escolha)
                            self.__dados_VT['valor_dias'].append(len(self.__FORMS.iloc[VT,26].split()) + len(str(self.__FORMS.iloc[VT,33]).split()))
                            self.__dados_VT['dias'].append(f'{str(self.__FORMS.iloc[VT,26]).split()} e {str(self.__FORMS.iloc[VT,33]).split()} de {mes[:3]}')
                            self.__dados_VT['valor_total'].append((len(self.__FORMS.iloc[VT,26].split()) + len(str(self.__FORMS.iloc[VT,33]).split()))*10)
                    
                    elif str(self.__FORMS.iloc[VT,26]) != 'nan' and str(str(self.__FORMS.iloc[VT,33])) == 'nan':
                        if len(self.__FORMS.iloc[VT,26].split(';')) > 1:
                            if str(self.__FORMS.iloc[VT,11]) == 'Não Encontrado':
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,14]))
                                self.__dados_VT['cpf'].append('Não encontrado')
                                self.__dados_VT['siape'].append('xxx')
                            else:
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,11]).split(' | ')[0])
                                self.__dados_VT['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]))                         
                                self.__inserir_siape(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]),escolha)
                            self.__dados_VT['valor_dias'].append(len(self.__FORMS.iloc[VT,26].split(';')))
                            self.__dados_VT['dias'].append(f'{str(self.__FORMS.iloc[VT,26]).split(";")} de {mes[:3]}')
                            self.__dados_VT['valor_total'].append(len(self.__FORMS.iloc[VT,26].split(';'))*10)
                        else:
                            if str(self.__FORMS.iloc[VT,11]) == 'Não Encontrado':
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,14]))
                                self.__dados_VT['cpf'].append('Não encontrado')
                                self.__dados_VT['siape'].append('xxx')
                            else:
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,11]).split(' | ')[0])
                                self.__dados_VT['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]))                      
                                self.__inserir_siape(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]),escolha)
                            self.__dados_VT['valor_dias'].append(len(self.__FORMS.iloc[VT,26].split()))
                            self.__dados_VT['dias'].append(f'{str(self.__FORMS.iloc[VT,26]).split()} de {mes[:3]}')
                            self.__dados_VT['valor_total'].append(len(self.__FORMS.iloc[VT,26].split())*10)
                    
                    elif str(self.__FORMS.iloc[VT,26]) == 'nan' and str(str(self.__FORMS.iloc[VT,33])) != 'nan':
                        if len(str(self.__FORMS.iloc[VT,33]).split(';')) > 1:
                            if str(self.__FORMS.iloc[VT,11]) == 'Não Encontrado':
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,14]))
                                self.__dados_VT['cpf'].append('Não encontrado')
                                self.__dados_VT['siape'].append('xxx')
                            else:
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,11]).split(' | ')[0])
                                self.__dados_VT['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]))                    
                                self.__inserir_siape(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]),escolha)
                            self.__dados_VT['valor_dias'].append(len(str(self.__FORMS.iloc[VT,33]).split(';')))
                            self.__dados_VT['dias'].append(f'{str(str(self.__FORMS.iloc[VT,33]).split(";"))} de {mes[:3]}')
                            self.__dados_VT['valor_total'].append(len(str(self.__FORMS.iloc[VT,33]).split(';'))*10)
                        else:
                            if str(self.__FORMS.iloc[VT,11]) == 'Não Encontrado':
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,14]))
                                self.__dados_VT['cpf'].append('Não encontrado')
                                self.__dados_VT['siape'].append('xxx')
                            else:
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,11]).split(' | ')[0])
                                self.__dados_VT['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]))
                                self.__inserir_siape(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]),escolha)
                            self.__dados_VT['valor_dias'].append(len(str(self.__FORMS.iloc[VT,33]).split()))
                            self.__dados_VT['dias'].append(f'{str(str(self.__FORMS.iloc[VT,33]).split())} de {mes[:3]}')
                            self.__dados_VT['valor_total'].append(len(str(self.__FORMS.iloc[VT,33]).split())*10)
        else:
            for BE in range(len(self.__FORMS.iloc[:,0])):
                if str(self.__FORMS.iloc[BE,8]).split()[0] == mes and str(self.__FORMS.iloc[BE,8]).split()[-1] == ano:
                    if str(self.__FORMS.iloc[BE,33]) != 'nan':
                        if len(str(self.__FORMS.iloc[BE,33]).split(';')) > 1:
                            if str(self.__FORMS.iloc[BE,11]) == 'Não Encontrado':
                                self.__dados_BE['nome'].append(str(self.__FORMS.iloc[BE,14]))
                                self.__dados_BE['cpf'].append('Não encontrado')
                                self.__dados_BE['siape'].append('xxx')
                                self.__dados_BE['valor_total'].append('Indisponivel')
                                self.__dados_BE['salario'].append('Indisponivel')
                            else:
                                self.__dados_BE['nome'].append(str(self.__FORMS.iloc[BE,11]).split(' | ')[0])
                                self.__dados_BE['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[BE,11]).split(' | ')[-1]))
                                self.__inserir_siape(self.__conversor_de_cpf(str(self.__FORMS.iloc[BE,11]).split(' | ')[-1]),escolha)
                                for salario in range(len(self.__SCE.iloc[:,22])):
                                    if self.__conversor_de_cpf(str(self.__SCE.iloc[salario,6])) == self.__conversor_de_cpf(str(self.__FORMS.iloc[BE,11]).split(' | ')[-1]):
                                        self.__dados_BE['valor_total'].append(round(len(self.__FORMS.iloc[BE,33].split(';'))*(float(str(self.__SCE.iloc[salario,22]).replace(',','.'))/30),2))
                                        self.__dados_BE['salario'].append(str(self.__SCE.iloc[salario,22]))
                                        break
                            self.__dados_BE['valor_dias'].append(len(self.__FORMS.iloc[BE,33].split(';')))
                            self.__dados_BE['dias'].append(f'{str(self.__FORMS.iloc[BE,33].split(";"))} de {mes[:3]}')
                        else:
                            if str(self.__FORMS.iloc[BE,11]) == 'NÃO ENCONTRADO':
                                self.__dados_BE['nome'].append(str(self.__FORMS.iloc[BE,14]))
                                self.__dados_BE['cpf'].append('Não encontrado')
                                self.__dados_BE['siape'].append('xxx')
                                self.__dados_BE['valor_total'].append('Indisponivel')
                                self.__dados_BE['salario'].append('Indisponivel')
                            else:
                                self.__dados_BE['nome'].append(str(self.__FORMS.iloc[BE,11]).split(' | ')[0])
                                self.__dados_BE['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[BE,11]).split(' | ')[-1]))
                                self.__inserir_siape(self.__conversor_de_cpf(str(self.__FORMS.iloc[BE,11]).split(' | ')[-1]),escolha)
                                for salario in range(len(self.__SCE.iloc[:,22])):
                                    if self.__conversor_de_cpf(str(self.__SCE.iloc[salario,6])) == self.__conversor_de_cpf(str(self.__FORMS.iloc[BE,11]).split(' | ')[-1]):
                                        self.__dados_BE['valor_total'].append(round(len(str(self.__FORMS.iloc[BE,33]).split())*(float(str(self.__SCE.iloc[salario,22]).replace(',','.'))/30),2))
                                        self.__dados_BE['salario'].append(str(self.__SCE.iloc[salario,22]))
                                        break
                            self.__dados_BE['valor_dias'].append(len(str(self.__FORMS.iloc[BE,33]).split()))
                            self.__dados_BE['dias'].append(f'{str(self.__FORMS.iloc[BE,33]).split()} de {mes[:3]}')

        # Limpeza de dados (Inativos, Não encontrados e Duplicados)
        for deletar_vt in range(5):
            for cpf_vt in self.__dados_VT['cpf']:
                if cpf_vt == 'Não encontrado':
                    indexa = self.__dados_VT['cpf'].index(cpf_vt)
                    for chave in self.__dados_VT:
                        self.__dados_VT[chave].pop(indexa)
                else:
                    for cpf_sce in range(len(self.__SCE.iloc[:,6])):
                        if self.__conversor_de_cpf(str(self.__SCE.iloc[cpf_sce,6])) == cpf_vt:
                            if str(self.__SCE.iloc[cpf_sce,13]) == 'Inativo':
                                indexb = self.__dados_VT['cpf'].index(cpf_vt)
                                for chave in self.__dados_VT:
                                    self.__dados_VT[chave].pop(indexb)
        
        for deletar_BE in range(5):
            for cpf_be in self.__dados_BE['cpf']:
                if cpf_be == 'Não encontrado':
                    indexc = self.__dados_BE['cpf'].index(cpf_be)
                    for chave in self.__dados_BE:
                        self.__dados_BE[chave].pop(indexc)
                else:
                    for cpf_sceb in range(len(self.__SCE.iloc[:,6])):
                        if self.__conversor_de_cpf(str(self.__SCE.iloc[cpf_sceb,6])) == cpf_be:
                            if str(self.__SCE.iloc[cpf_sceb,13]) == 'Inativo':
                                indexd = self.__dados_BE['cpf'].index(cpf_be)
                                for chave in self.__dados_BE:
                                    self.__dados_BE[chave].pop(indexd)

        for r in range(5):
            for cpf in self.__dados_VT['cpf']:
                while self.__dados_VT['cpf'].count(cpf) > 1:
                    index = self.__dados_VT['cpf'].index(cpf)
                    for chave in self.__dados_VT:
                        self.__dados_VT[chave].pop(index)

        for p in range(5):
            for cpf in self.__dados_BE['cpf']:
                while self.__dados_BE['cpf'].count(cpf) > 1:
                    index = self.__dados_BE['cpf'].index(cpf)
                    for chave in self.__dados_BE:
                        self.__dados_BE[chave].pop(index)

    def iniciar(self, escolha, mes, ano):
        self.__gerar_dados(escolha, mes, ano)
        
        dados_finais = self.__dados_VT if escolha == 'VT' else self.__dados_BE
        df_faltas = pd.DataFrame(dados_finais)

        if df_faltas.empty:
            raise ValueError(f"Não foram encontrados registros válidos de faltas para {escolha} no período selecionado.")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        pasta_downloads = os.path.join(os.path.expanduser("~"), 'Downloads')

        caminho_excel = os.path.join(pasta_downloads, f'Analise_Faltas_{escolha}_{timestamp}.xlsx')
        df_faltas.to_excel(caminho_excel, index=False)

        pasta_temp = tempfile.gettempdir()
        caminho_csv_temp = os.path.join(pasta_temp, f'dados_macro_{timestamp}.csv')

        with open(caminho_csv_temp, 'w', encoding='utf-8') as f_csv:
            for index, row in df_faltas.iterrows():
                if str(row['cpf']) != 'Não encontrado' and str(row['siape']) != 'xxx':
                    valor_formatado = self.__swicth(row['valor_total'])
                    dias_formatados = str(row['dias']).replace('[','').replace(']','')
                    rubrica = '82695' if escolha == 'VT' else '83172'
                    
                    linha = f"{row['siape']},D,{rubrica},6,I,{str(mes)[:3]+str(ano)},\"{valor_formatado}\",\"SCE,Folha de ponto e ficha F\",\"{dias_formatados}\"\n"
                    f_csv.write(linha)

        try:
            hob = Hob(caminho_csv_temp)
        finally:
            if os.path.exists(caminho_csv_temp):
                os.remove(caminho_csv_temp)