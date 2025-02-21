import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import shutil
import os
import subprocess
import threading
import time

from utils.gerador_de_declaracao.Gerador_de_declaracao import Gerador_de_declaracao

class Interface_declaracao:
    def __init__(self, root):
            self.__frame_botoes = tk.Frame(root)
            self.__frame_botoes.pack(pady=20)

            self.__upload_sce_button = tk.Button(self.__frame_botoes, text='Upload SCE', command=lambda: self.__upload_file_Gerador_de_declaracao('SCE'))
            self.__upload_sce_button.pack(pady=10)

            self.__upload_modelo_button = tk.Button(self.__frame_botoes, text='Upload Modelo', command=lambda: self.__upload_file_Gerador_de_declaracao('Modelo'))
            self.__upload_modelo_button.pack(pady=10)

            self.__entrada_cpf = tk.Entry(self.__frame_botoes, font=('Arial', 14))
            self.__entrada_cpf.pack(pady=20)
            self.__entrada_cpf.insert(0, 'Insira um CPF')
            self.__entrada_cpf.bind("<Button-1>", self.__limpar_texto_cpf)
            self.__entrada_cpf.bind("<KeyRelease> ", self.__formatar_cpf) 

            self.__analyze_button = tk.Button(self.__frame_botoes, text='Resultado da Declaração', command= self.__run_analyzer_Gerador_de_declaracao)
            self.__analyze_button.pack(pady=10)

            self.__barra_progresso = ttk.Progressbar(self.__frame_botoes, orient="horizontal", length=300, mode="determinate")
            self.__barra_progresso.pack(pady=20)

            self.__analyze_button.bind("<Button-1>", self.__inserir_texto)

            self.__frame_botoes.mainloop()
    
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
     
    def __upload_file_Gerador_de_declaracao(self, upload_type):
        file_path = filedialog.askopenfilename()
        if file_path:
            print(f'Arquivo selecionado para {upload_type}: {file_path}')

            destination_directory = 'utils/gerador_de_declaracao/dados'

            if upload_type == 'SCE':
                new_file_name = 'Sce.xlsx'
            elif upload_type == 'Modelo':
                new_file_name = 'Modelo.docx'
            else:
                new_file_name = os.path.basename(file_path)

            destination_path = os.path.join(destination_directory, new_file_name)
            shutil.copy(file_path, destination_path)
            print(f'Arquivo copiado e renomeado para: {destination_path}')

    def __run_analyzer_Gerador_de_declaracao(self):
        try:
            declaracao = Gerador_de_declaracao()
            if str(self.__entrada_cpf.get()) != '' and str(self.__entrada_cpf.get()) != 'Insira um CPF' and len(self.__entrada_cpf.get()) == 14:
                self.__start_task(declaracao.iniciar(str(self.__entrada_cpf.get())))
                messagebox.showinfo('Sucesso', 'Verifique sua pasta de Downloads')
                print('Gerador de Declarações executado com sucesso.')
            else:
                messagebox.showerror('Erro', 'Digite um CPF válido')
        except subprocess.CalledProcessError:
            messagebox.showerror('Erro', 'Erro ao executar o Gerador de Declarações.')

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
