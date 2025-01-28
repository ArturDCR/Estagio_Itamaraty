import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import shutil
import os
import subprocess
import threading
import time
from datetime import datetime

from utils.gerador_mala_direta.Gerador_mala_direta import Gerador_mala_direta

class Interface_mala_direta:
    def __init__(self, root):
            self.__frame_botoes = tk.Frame(root)
            self.__frame_botoes.pack(pady=20)

            self.__upload_sce_button = tk.Button(self.__frame_botoes, text='Upload SCE', command=lambda: self.__upload_file_Gerador_mala_direta('SCE'))
            self.__upload_sce_button.pack(pady=10)

            self.__upload_modelo_button = tk.Button(self.__frame_botoes, text='Upload Modelo', command=lambda: self.__upload_file_Gerador_mala_direta('Modelo'))
            self.__upload_modelo_button.pack(pady=10)

            self.__variavel_dia = None

            self.__dias = ['1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31']
            self.__meses = ['Janeiro','Fevereiro','Março','Abril','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
            self.__anos = [str(ano) for ano in range(2023, datetime.now().year + 1)]

            self.__variavel_escolha_dia = tk.StringVar(self.__frame_botoes)
            self.__variavel_escolha_dia.set('Escolha o dia')
            
            self.__variavel_escolha_dia.trace_add("write", self.__Set_dia)

            self.__caixa_opcoes = tk.OptionMenu(self.__frame_botoes, self.__variavel_escolha_dia, *self.__dias)
            self.__caixa_opcoes.pack(pady=10)

            self.__variavel_mes = None

            self.__variavel_escolha_Mes = tk.StringVar(self.__frame_botoes)
            self.__variavel_escolha_Mes.set('Escolha um Mês')
            
            self.__variavel_escolha_Mes.trace_add("write", self.__Set_mes)

            self.__caixa_opcoes = tk.OptionMenu(self.__frame_botoes, self.__variavel_escolha_Mes, *self.__meses)
            self.__caixa_opcoes.pack(pady=10)

            self.__variavel_ano = None

            self.__variavel_escolha_ano = tk.StringVar(self.__frame_botoes)
            self.__variavel_escolha_ano.set('Escolha um Ano')
            
            self.__variavel_escolha_ano.trace_add("write", self.__Set_ano)

            self.__caixa_opcoes = tk.OptionMenu(self.__frame_botoes, self.__variavel_escolha_ano, *self.__anos)
            self.__caixa_opcoes.pack(pady=10)

            self.__entrada_cpf = tk.Entry(self.__frame_botoes, font=('Arial', 14))
            self.__entrada_cpf.pack(pady=20)
            self.__entrada_cpf.insert(0, 'Insira um CPF')
            self.__entrada_cpf.bind("<Button-1>", self.__limpar_texto_cpf)
            self.__entrada_cpf.bind("<KeyRelease> ", self.__formatar_cpf) 

            self.__analyze_button = tk.Button(self.__frame_botoes, text='Resultado Mala Direta', command= self.__run_analyzer_Gerador_mala_direta)
            self.__analyze_button.pack(pady=10)

            self.__barra_progresso = ttk.Progressbar(self.__frame_botoes, orient="horizontal", length=300, mode="determinate")
            self.__barra_progresso.pack(pady=20)

            self.__analyze_button.bind("<Button-1>", self.__inserir_texto)

            self.__frame_botoes.mainloop()
    
    def __Set_dia(self, *args):
        self.__variavel_dia = self.__variavel_escolha_dia.get()
    
    def __Set_ano(self, *agrs):
        self.__variavel_ano = self.__variavel_escolha_ano.get()
    
    def __Set_mes(self, *args):
        self.__variavel_mes = self.__variavel_escolha_Mes.get()

    def __inserir_texto(self, event):
        if self.__entrada_cpf.get() == "":
            self.__entrada_cpf.insert(0, "Insira um CPF")
            self.__frame_botoes.focus()

    def __limpar_texto_cpf(self, event):
        if self.__entrada_cpf.get() == "Insira um CPF":
            self.__entrada_cpf.delete(0, tk.END)

    def __formatar_cpf(self, event):
        cpf = self.__entrada_cpf.get()
        
        cpf = ''.join(filter(str.isdigit, cpf))

        if len(cpf) <= 3:
            cpf = cpf
        elif len(cpf) <= 6:
            cpf = f'{cpf[:3]}.{cpf[3:]}'
        elif len(cpf) <= 9:
            cpf = f'{cpf[:3]}.{cpf[3:6]}.{cpf[6:]}'
        else:
            cpf = f'{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:11]}'

        self.__entrada_cpf.delete(0, tk.END)
        self.__entrada_cpf.insert(0, cpf)
     
    def __upload_file_Gerador_mala_direta(self, upload_type):
        file_path = filedialog.askopenfilename()
        if file_path:
            print(f'Arquivo selecionado para {upload_type}: {file_path}')

            destination_directory = 'utils/gerador_mala_direta/dados'

            if upload_type == 'SCE':
                new_file_name = 'Sce.xlsx'
            elif upload_type == 'Modelo':
                new_file_name = 'Modelo.docx'
            else:
                new_file_name = os.path.basename(file_path)

            destination_path = os.path.join(destination_directory, new_file_name)
            shutil.copy(file_path, destination_path)
            print(f'Arquivo copiado e renomeado para: {destination_path}')

            self.__frame_botoes.mainloop()

    def __run_analyzer_Gerador_mala_direta(self):
        try:
            mala_direta = Gerador_mala_direta()
            if str(self.__entrada_cpf.get()) != '' and str(self.__entrada_cpf.get()) != 'Insira um CPF' and len(self.__entrada_cpf.get()) == 14:
                self.__start_task(mala_direta.iniciar(str(self.__entrada_cpf.get()), self.__variavel_dia, self.__meses.index(self.__variavel_mes)+1, self.__variavel_ano))
                messagebox.showinfo('Sucesso', 'Verifique sua pasta de Downloads')
                print('Gerador de Mala Direta executado com sucesso.')
            else:
                messagebox.showerror('Erro', 'Digite um CPF válido')
        except subprocess.CalledProcessError:
            messagebox.showerror('Erro', 'Erro ao executar o Gerador de Mala Direta.')

    def __start_task(self, func):
        thread = threading.Thread(target=func)
        thread.daemon = True
        thread.start()

        progress = 0
        while progress < 100:
            progress += 10
            self.__barra_progresso['value'] = progress
            self.__frame_botoes.update_idletasks()
            time.sleep(1)
