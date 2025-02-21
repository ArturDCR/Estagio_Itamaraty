import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import shutil
import os
import subprocess
from datetime import datetime
import threading
import time

from . .analise_de_faltas.Gerador_de_faltas import Gerador_de_faltas

class Interface_analise_de_faltas():
    def __init__(self, root):
        self.__frame_botoes = tk.Frame(root)
        self.__frame_botoes.pack(pady=20)

        self.__variavel_desconto = None

        self.__descontos = ['VT','BE']
        self.__meses = ['Janeiro','Fevereiro','Março','Abril','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
        self.__anos = [str(ano) for ano in range(2023, datetime.now().year + 1)]

        self.__variavel_escolha_desconto = tk.StringVar(self.__frame_botoes)
        self.__variavel_escolha_desconto.set('Escolha o desconto')
        
        self.__variavel_escolha_desconto.trace_add("write", self.__Set_desconto)

        self.__caixa_opcoes = tk.OptionMenu(self.__frame_botoes, self.__variavel_escolha_desconto, *self.__descontos)
        self.__caixa_opcoes.pack(pady=10)

        self.__variavel_ano = None

        self.__variavel_escolha_ano = tk.StringVar(self.__frame_botoes)
        self.__variavel_escolha_ano.set('Escolha um Ano')
        
        self.__variavel_escolha_ano.trace_add("write", self.__Set_ano)

        self.__caixa_opcoes = tk.OptionMenu(self.__frame_botoes, self.__variavel_escolha_ano, *self.__anos)
        self.__caixa_opcoes.pack(pady=10)

        self.__variavel_mes = None

        self.__variavel_escolha_Mes = tk.StringVar(self.__frame_botoes)
        self.__variavel_escolha_Mes.set('Escolha um Mês')
        
        self.__variavel_escolha_Mes.trace_add("write", self.__Set_mes)

        self.__caixa_opcoes = tk.OptionMenu(self.__frame_botoes, self.__variavel_escolha_Mes, *self.__meses)
        self.__caixa_opcoes.pack(pady=10)

        self.__upload_ciee_button = tk.Button(self.__frame_botoes, text='Upload Forms', command=lambda: self.__upload_file_Analise_de_Faltas('Forms'))
        self.__upload_ciee_button.pack(pady=10)

        self.__upload_mre_button = tk.Button(self.__frame_botoes, text='Upload MRE', command=lambda: self.__upload_file_Analise_de_Faltas('MRE'))
        self.__upload_mre_button.pack(pady=10)

        self.__upload_sce_button = tk.Button(self.__frame_botoes, text='Upload SCE', command=lambda: self.__upload_file_Analise_de_Faltas('SCE'))
        self.__upload_sce_button.pack(pady=10)

        self.__analyze_button = tk.Button(self.__frame_botoes, text='Resultado da Análise', command= self.__run_analyzer_Analise_de_Faltas)
        self.__analyze_button.pack(pady=10)

        self.__barra_progresso = ttk.Progressbar(self.__frame_botoes, orient="horizontal", length=300, mode="determinate")
        self.__barra_progresso.pack(pady=20)

        self.__frame_botoes.mainloop()

    def __Set_desconto(self, *args):
        self.__variavel_desconto = self.__variavel_escolha_desconto.get()
    
    def __Set_ano(self, *agrs):
        self.__variavel_ano = self.__variavel_escolha_ano.get()
    
    def __Set_mes(self, *args):
        self.__variavel_mes = self.__variavel_escolha_Mes.get()
    
    def __upload_file_Analise_de_Faltas(self, upload_type):
        file_path = filedialog.askopenfilename()
        if file_path:
            print(f'Arquivo selecionado para {upload_type}: {file_path}')
            
            destination_directory = 'utils/analise_de_faltas/dados'
            
            if upload_type == 'Forms':
                new_file_name = 'Forms.xlsx'
            elif upload_type == 'MRE':
                new_file_name = 'Mre.xlsx'
            elif upload_type == 'SCE':
                new_file_name = 'Sce.xlsx'
            else:
                new_file_name = os.path.basename(file_path)

            destination_path = os.path.join(destination_directory, new_file_name)
            shutil.copy(file_path, destination_path)
            print(f'Arquivo copiado e renomeado para: {destination_path}')
    
    def __run_analyzer_Analise_de_Faltas(self):
        try:
            if self.__variavel_mes != None and self.__variavel_ano != None and self.__variavel_desconto != None:
                faltas = Gerador_de_faltas()
                self.__start_task(faltas.iniciar(self.__variavel_desconto, self.__variavel_mes, self.__variavel_ano))
                messagebox.showinfo('Sucesso', 'Verifique sua pasta de Downloads')
                print('Analisador de Faltas executado com sucesso.')
                self.__variavel_ano = None
                self.__variavel_escolha_ano.set('Escolha um Ano')
                self.__variavel_mes = None
                self.__variavel_escolha_Mes.set('Escolha um Mês')
                self.__variavel_desconto = None
                self.__variavel_escolha_desconto.set('Escolha o desconto')
            else:
                messagebox.showerror('Erro', 'Escolha um Mês e um Ano e o tipo de Desconto')
                self.__variavel_ano = None
                self.__variavel_escolha_ano.set('Escolha um Ano')
                self.__variavel_mes = None
                self.__variavel_escolha_Mes.set('Escolha um Mês')
                self.__variavel_desconto = None
                self.__variavel_escolha_desconto.set('Escolha o desconto')
        except subprocess.CalledProcessError:
            messagebox.showerror('Erro', 'Erro ao executar o Analisador de Faltas.')
            self.__variavel_ano = None
            self.__variavel_escolha_ano.set('Escolha um Ano')
            self.__variavel_mes = None
            self.__variavel_escolha_Mes.set('Escolha um Mês')
            self.__variavel_desconto = None
            self.__variavel_escolha_desconto.set('Escolha o desconto')
    
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