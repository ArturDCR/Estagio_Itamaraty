import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import shutil
import os
import subprocess
import threading
import time
from datetime import datetime

from utils.analise_SouGov.Analise_SouGov import Analise_SouGov

class Interface_SouGov:
    def __init__(self, root):
        self.__frame_botoes = tk.Frame(root)
        self.__frame_botoes.pack(pady=20)

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

        self.__upload_folhas_button = tk.Button(self.__frame_botoes, text='Upload Emails', command=lambda: self.__upload_file_Analise_SouGov('Emails'))
        self.__upload_folhas_button.pack(pady=10)

        self.__upload_folhas_button = tk.Button(self.__frame_botoes, text='Upload TA', command=lambda: self.__upload_file_Analise_SouGov('TA'))
        self.__upload_folhas_button.pack(pady=10)
    
        self.__upload_sce_button = tk.Button(self.__frame_botoes, text='Upload Sce', command=lambda: self.__upload_file_Analise_SouGov('Sce'))
        self.__upload_sce_button.pack(pady=10)

        self.__analyze_button = tk.Button(self.__frame_botoes, text='Resultado da Análise', command= self.__run_analyzer_Analise_SouGov)
        self.__analyze_button.pack(pady=10)

        self.__barra_progresso = ttk.Progressbar(self.__frame_botoes, orient="horizontal", length=300, mode="determinate")
        self.__barra_progresso.pack(pady=20)

        self.__text_area = tk.Text(root, height=30, width=90)
        self.__text_area.pack(pady=10)
        self.__ler_arquivo()

        self.__frame_botoes.mainloop()
    
    def __ler_arquivo(self):
        caminho_arquivo = "utils/interface_grafica/dados/informacoes_SouGov.txt"
        try:
            with open(caminho_arquivo, 'r') as arquivo:
                conteudo = arquivo.read()

            self.__text_area.delete(1.0, tk.END)
            self.__text_area.insert(tk.END, conteudo)
            self.__text_area.config(state=tk.DISABLED)
        except FileNotFoundError:
            self.__text_area.insert(tk.END, "Arquivo não encontrado.")

    def __Set_dia(self, *args):
        self.__variavel_dia = self.__variavel_escolha_dia.get()
    
    def __Set_ano(self, *agrs):
        self.__variavel_ano = self.__variavel_escolha_ano.get()
    
    def __Set_mes(self, *args):
        self.__variavel_mes = self.__variavel_escolha_Mes.get()

    def __upload_file_Analise_SouGov(self, upload_type):
        file_path = filedialog.askopenfilename()
        if file_path:
            print(f'Arquivo selecionado para {upload_type}: {file_path}')

            destination_directory = 'utils/analise_SouGov/dados'

            if upload_type == 'Emails':
                new_file_name = 'Emails.xlsx'
            elif upload_type == 'TA':
                new_file_name = 'Tas.xlsx'
            elif upload_type == 'Sce':
                new_file_name = 'Sce.xlsx'
            else:
                new_file_name = os.path.basename(file_path)

            destination_path = os.path.join(destination_directory, new_file_name)
            shutil.copy(file_path, destination_path)
            print(f'Arquivo copiado e renomeado para: {destination_path}')
            self.__frame_botoes.mainloop()

    def __run_analyzer_Analise_SouGov(self):
        try:
            if self.__variavel_ano != None and self.__variavel_mes != None and self.__variavel_dia != None:
                SouGov = Analise_SouGov()
                self.__start_task(SouGov.iniciar(self.__variavel_dia, self.__meses.index(self.__variavel_mes)+1, self.__variavel_ano))
                messagebox.showinfo('Sucesso', 'Verifique sua pasta de Downloads')
                print('Analisador SouGov executado com sucesso.')
            else:
                messagebox.showerror('Erro', 'Escolha um Dia um Mês e um Ano')
                self.__variavel_ano = None
                self.__variavel_mes = None
                self.__variavel_dia = None
        except subprocess.CalledProcessError:
            messagebox.showerror('Erro', 'Erro ao executar o Analisador SouGov.')
            self.__variavel_ano = None
            self.__variavel_mes = None
            self.__variavel_dia = None

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