import pandas as pd
import os
import shutil

from utils.analise_de_faltas.Analise_de_faltas import Analise_de_faltas

class Gerador_de_faltas:
    def __init__(self):
        self.__FORMS = pd.read_excel(os.path.join('utils/data/Forms.xlsx'))
        self.__SCE = pd.read_excel(os.path.join('utils/data/Sce.xlsx'))
        self.__MRE = pd.read_excel(os.path.join('utils/data/Mre.xlsx'))

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
                    #Decisão que lida com justificada e injustificada != vazio
                    if str(self.__FORMS.iloc[VT,26]) != 'nan' and str(self.__FORMS.iloc[VT,33]) != 'nan':
                        #Decisão que verifica injustificada e justificada > 1
                        if len(self.__FORMS.iloc[VT,26].split(';')) > 1 and len(self.__FORMS.iloc[VT,33].split(';')) > 1:
                            #Bloco que adiciona valores ao dicionario de acordo com injustificada e justificada > 1
                            #Verifica se nome e cpf existem
                            if str(self.__FORMS.iloc[VT,11]) == 'Não Encontrado':
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,14]))
                                self.__dados_VT['cpf'].append('Não encontrado')
                                self.__dados_VT['siape'].append('xxx')
                            else:
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,11]).split(' | ')[0])
                                self.__dados_VT['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]))         
                                self.__inserir_siape(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]),escolha)
                            self.__dados_VT['valor_dias'].append(len(self.__FORMS.iloc[VT,26].split(';')) + len(self.__FORMS.iloc[VT,33].split(';')))
                            self.__dados_VT['dias'].append(f'{str(self.__FORMS.iloc[VT,26]).split(";") + self.__FORMS.iloc[VT,33].split(";")} de {mes[:3]}')
                            self.__dados_VT['valor_total'].append((len(self.__FORMS.iloc[VT,26].split(';')) + len(self.__FORMS.iloc[VT,33].split(';')))*10)
                        #Decisão que lida com justificada > 1 e injustificada = 1
                        elif len(self.__FORMS.iloc[VT,26].split(';')) > 1 and len(self.__FORMS.iloc[VT,33].split()) == 1:
                            #Bloco que adiciona valores ao dicionario de acordo com justificada > 1 e injustificada = 1
                            #Verifica se nome e cpf existem
                            if str(self.__FORMS.iloc[VT,11]) == 'Não Encontrado':
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,14]))
                                self.__dados_VT['cpf'].append('Não encontrado')
                                self.__dados_VT['siape'].append('xxx')
                            else:
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,11]).split(' | ')[0])
                                self.__dados_VT['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]))                            
                                self.__inserir_siape(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]),escolha)
                            self.__dados_VT['valor_dias'].append(len(self.__FORMS.iloc[VT,26].split(';')) + len(self.__FORMS.iloc[VT,33].split()))
                            self.__dados_VT['dias'].append(f'{str(self.__FORMS.iloc[VT,26]).split(";") + self.__FORMS.iloc[VT,33].split()} de {mes[:3]}')
                            self.__dados_VT['valor_total'].append((len(self.__FORMS.iloc[VT,26].split(';')) + len(self.__FORMS.iloc[VT,33].split()))*10)
                        #Decisão que lida com justificada = 1 e injustificada > 1
                        elif len(self.__FORMS.iloc[VT,26].split()) == 1 and len(self.__FORMS.iloc[VT,33].split(';')) > 1:
                            #Bloco que adiciona valores ao dicionario de acordo com justificada = 1 e injustificada > 1
                            #Verifica se nome e cpf existem
                            if str(self.__FORMS.iloc[VT,11]) == 'Não Encontrado':
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,14]))
                                self.__dados_VT['cpf'].append('Não encontrado')
                                self.__dados_VT['siape'].append('xxx')
                            else:
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,11]).split(' | ')[0])
                                self.__dados_VT['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]))                        
                                self.__inserir_siape(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]),escolha)
                            self.__dados_VT['valor_dias'].append(len(self.__FORMS.iloc[VT,26].split()) + len(self.__FORMS.iloc[VT,33].split(';')))
                            self.__dados_VT['dias'].append(f'{str(self.__FORMS.iloc[VT,26]).split() + self.__FORMS.iloc[VT,33].split(";")} de {mes[:3]}')
                            self.__dados_VT['valor_total'].append((len(self.__FORMS.iloc[VT,26].split()) + len(self.__FORMS.iloc[VT,33].split(';')))*10)
                        #Decisão que lida com justificada = 1 e injustificada = 1
                        elif len(self.__FORMS.iloc[VT,26].split()) == 1 and len(self.__FORMS.iloc[VT,33].split()) == 1:
                            #Bloco que adiciona valores ao dicionario de acordo com justificada = 1 e injustificada = 1
                            #Verifica se nome e cpf existem
                            if str(self.__FORMS.iloc[VT,11]) == 'Não Encontrado':
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,14]))
                                self.__dados_VT['cpf'].append('Não encontrado')
                                self.__dados_VT['siape'].append('xxx')
                            else:
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,11]).split(' | ')[0])
                                self.__dados_VT['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]))                          
                                self.__inserir_siape(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]),escolha)
                            self.__dados_VT['valor_dias'].append(len(self.__FORMS.iloc[VT,26].split()) + len(self.__FORMS.iloc[VT,33].split()))
                            self.__dados_VT['dias'].append(f'{str(self.__FORMS.iloc[VT,26]).split() + self.__FORMS.iloc[VT,33].split()} de {mes[:3]}')
                            self.__dados_VT['valor_total'].append((len(self.__FORMS.iloc[VT,26].split()) + len(self.__FORMS.iloc[VT,33].split()))*10)
                    #Decisão que lida com justificada != vazio e injustificada == vazio
                    elif str(self.__FORMS.iloc[VT,26]) != 'nan' and str(self.__FORMS.iloc[VT,33]) == 'nan':
                        #Decisão que lida com justificada > 1
                        if len(self.__FORMS.iloc[VT,26].split(';')) > 1:
                            #Bloco que adiciona valores ao dicionario de acordo com justificada > 1
                            #Verifica se nome e cpf existem
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
                        #Decisão que lida com justificada == 1
                        else:
                            #Bloco que adiciona valores ao dicionario de acordo com justificada > 1
                            #Verifica se nome e cpf existem
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
                    #Decisão que lida com justificada == vazio e injustificada != vazio
                    elif str(self.__FORMS.iloc[VT,26]) == 'nan' and str(self.__FORMS.iloc[VT,33]) != 'nan':
                        #Decisão que verifica se injustificada > 1
                        if len(self.__FORMS.iloc[VT,33].split(';')) > 1:
                            #Bloco que adiciona valores ao dicionario de acordo com injustificada > 1
                            #Verifica se nome e cpf existem
                            if str(self.__FORMS.iloc[VT,11]) == 'Não Encontrado':
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,14]))
                                self.__dados_VT['cpf'].append('Não encontrado')
                                self.__dados_VT['siape'].append('xxx')
                            else:
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,11]).split(' | ')[0])
                                self.__dados_VT['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]))                    
                                self.__inserir_siape(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]),escolha)
                            self.__dados_VT['valor_dias'].append(len(self.__FORMS.iloc[VT,33].split(';')))
                            self.__dados_VT['dias'].append(f'{str(self.__FORMS.iloc[VT,33].split(";"))} de {mes[:3]}')
                            self.__dados_VT['valor_total'].append(len(self.__FORMS.iloc[VT,33].split(';'))*10)
                        else:
                            #Bloco que adiciona valores ao dicionario de acordo com injustificada == 1
                            #Verifica se nome e cpf existem
                            if str(self.__FORMS.iloc[VT,11]) == 'Não Encontrado':
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,14]))
                                self.__dados_VT['cpf'].append('Não encontrado')
                                self.__dados_VT['siape'].append('xxx')
                            else:
                                self.__dados_VT['nome'].append(str(self.__FORMS.iloc[VT,11]).split(' | ')[0])
                                self.__dados_VT['cpf'].append(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]))
                                self.__inserir_siape(self.__conversor_de_cpf(str(self.__FORMS.iloc[VT,11]).split(' | ')[-1]),escolha)
                            self.__dados_VT['valor_dias'].append(len(self.__FORMS.iloc[VT,33].split()))
                            self.__dados_VT['dias'].append(f'{str(self.__FORMS.iloc[VT,33].split())} de {mes[:3]}')
                            self.__dados_VT['valor_total'].append(len(self.__FORMS.iloc[VT,33].split())*10)
        # Decisão caso seja BT
        else:
            for BE in range(len(self.__FORMS.iloc[:,0])):
                if str(self.__FORMS.iloc[BE,8]).split()[0] == mes and str(self.__FORMS.iloc[BE,8]).split()[-1] == ano:
                    if str(self.__FORMS.iloc[BE,33]) != 'nan':
                        #Decisão que verifica se injustificada > 1
                        if len(str(self.__FORMS.iloc[BE,33]).split(';')) > 1:
                            #Bloco que adiciona valores ao dicionario de acordo com injustificada > 1
                            #Verifica se nome e cpf existem
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
                                #Laço que encontra o salário pelo CPF
                                for salario in range(len(self.__SCE.iloc[:,22])):
                                    if self.__conversor_de_cpf(str(self.__SCE.iloc[salario,6])) == self.__conversor_de_cpf(str(self.__FORMS.iloc[BE,11]).split(' | ')[-1]):
                                        self.__dados_BE['valor_total'].append(round(len(self.__FORMS.iloc[BE,33].split(';'))*(float(str(self.__SCE.iloc[salario,22]).replace(',','.'))/30),2))
                                        self.__dados_BE['salario'].append(str(self.__SCE.iloc[salario,22]))
                                        break
                            self.__dados_BE['valor_dias'].append(len(self.__FORMS.iloc[BE,33].split(';')))
                            self.__dados_BE['dias'].append(f'{str(self.__FORMS.iloc[BE,33].split(";"))} de {mes[:3]}')
                        #Decisão que verifica se injustificada = 1
                        else:
                            #Bloco que adiciona valores ao dicionario de acordo com injustificada = 1
                            #Verifica se nome e cpf existem
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
                                #Laço que encontra o salário pelo CPF
                                for salario in range(len(self.__SCE.iloc[:,22])):
                                    if self.__conversor_de_cpf(str(self.__SCE.iloc[salario,6])) == self.__conversor_de_cpf(str(self.__FORMS.iloc[BE,11]).split(' | ')[-1]):
                                        self.__dados_BE['valor_total'].append(round(len(str(self.__FORMS.iloc[BE,33]).split())*(float(str(self.__SCE.iloc[salario,22]).replace(',','.'))/30),2))
                                        self.__dados_BE['salario'].append(str(self.__SCE.iloc[salario,22]))
                                        break
                            self.__dados_BE['valor_dias'].append(len(str(self.__FORMS.iloc[BE,33]).split()))
                            self.__dados_BE['dias'].append(f'{str(self.__FORMS.iloc[BE,33]).split()} de {mes[:3]}')

        for z in self.__dados_VT:
            print(len(self.__dados_VT[z]))

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

        for r in range (5):
            for cpf in self.__dados_VT['cpf']:
                while self.__dados_VT['cpf'].count(cpf) > 1:
                    index = self.__dados_VT['cpf'].index(cpf)
                    for chave in self.__dados_VT:
                        self.__dados_VT[chave].pop(index)

        for p in range (5):
            for cpf in self.__dados_BE['cpf']:
                while self.__dados_BE['cpf'].count(cpf) > 1:
                    index = self.__dados_BE['cpf'].index(cpf)
                    for chave in self.__dados_BE:
                        self.__dados_BE[chave].pop(index)


    def __gerar_saida(self, escolha):
        if escolha == 'VT':
            pd.DataFrame(self.__dados_VT).to_excel('utils/data/Faltas.xlsx',index=False)
            shutil.copy('utils/data/Faltas.xlsx', os.path.join(os.path.expanduser("~"), 'Downloads'))
        else:
            pd.DataFrame(self.__dados_BE).to_excel('utils/data/Faltas.xlsx',index=False)
            shutil.copy('utils/data/Faltas.xlsx', os.path.join(os.path.expanduser("~"), 'Downloads'))           

    def __limpar_listas(self):
        for chave in self.__dados_BE:
            self.__dados_BE[chave].clear()
        for key in self.__dados_VT:
            self.__dados_VT[key].clear()
    
    def iniciar(self, escolha, mes, ano):
        self.__gerar_dados(escolha, mes, ano)
        self.__gerar_saida(escolha)
        # self.__limpar_listas()
        # analise_faltas = Analise_de_faltas()
        # analise_faltas.iniciar(escolha, mes, ano)