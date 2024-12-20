import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import shutil
import os
import subprocess
from datetime import datetime
import threading
import time
from utils.analise_de_descontos.Analise_de_descontos import Analise_de_descontos

class Interface_descontos:
    def __init__(self, root):
        self.__frame_botoes = tk.Frame(root)
        self.__frame_botoes.pack(pady=20)

        self.__variavel_ano = None

        self.__anos = [str(ano) for ano in range(2023, datetime.now().year + 1)]

        self.__variavel_escolha_ano = tk.StringVar(self.__frame_botoes)
        self.__variavel_escolha_ano.set('Escolha um ano')
        
        self.__variavel_escolha_ano.trace_add("write", self.__Set_ano)

        caixa_opcoes = tk.OptionMenu(self.__frame_botoes, self.__variavel_escolha_ano, *self.__anos)
        caixa_opcoes.pack(pady=10)
        
        self.__upload_folhas_button = tk.Button(self.__frame_botoes, text='Upload Descontos', command=lambda: self.__upload_file_Analise_de_Descontos('Descontos'))
        self.__upload_folhas_button.pack(pady=10)

        self.__analyze_button = tk.Button(self.__frame_botoes, text='Resultado da Análise', command= self.__run_analyzer_Analise_de_Descontos)
        self.__analyze_button.pack(pady=10)

        self.__barra_progresso = ttk.Progressbar(self.__frame_botoes, orient="horizontal", length=300, mode="determinate")
        self.__barra_progresso.pack(pady=20)

        self.__text_area = tk.Text(root, height=30, width=90)
        self.__text_area.pack(pady=10)
        self.__ler_arquivo()

        self.__frame_botoes.mainloop()
    
    def __ler_arquivo(self):
        caminho_arquivo = "utils/interface_grafica/dados/informacoes_analise_de_descontos.txt"
        try:
            with open(caminho_arquivo, 'r') as arquivo:
                conteudo = arquivo.read()

            self.__text_area.delete(1.0, tk.END)
            self.__text_area.insert(tk.END, conteudo)
            self.__text_area.config(state=tk.DISABLED)
        except FileNotFoundError:
            self.__text_area.insert(tk.END, "Arquivo não encontrado.")

    def __Set_ano(self, *agrs):
        self.__variavel_ano = self.__variavel_escolha_ano.get()

    def __upload_file_Analise_de_Descontos(self, upload_type):
        file_path = filedialog.askopenfilename()
        if file_path:
            print(f'Arquivo selecionado para {upload_type}: {file_path}')
            
            destination_directory = 'utils/analise_de_descontos/dados'
            
            if upload_type == 'Descontos':
                new_file_name = 'Descontos.pdf'
            else:
                new_file_name = os.path.basename(file_path)

            destination_path = os.path.join(destination_directory, new_file_name)
            shutil.copy(file_path, destination_path)
            print(f'Arquivo copiado e renomeado para: {destination_path}')

    def __run_analyzer_Analise_de_Descontos(self):
        try:
            if self.__variavel_ano != None:
                descontos = Analise_de_descontos()
                self.__start_task(descontos.iniciar(self.__variavel_ano))
                messagebox.showinfo('Sucesso', 'Verifique sua pasta de Downloads')
                print('Analisador de Descontos executado com sucesso.')
                self.__variavel_ano = None
            else:
                messagebox.showerror('Erro', 'Escolha um ano.')
                self.__variavel_ano = None
        except subprocess.CalledProcessError:
            messagebox.showerror('Erro', 'Erro ao executar o Analisador de Descontos.')
            self.__variavel_ano = None

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