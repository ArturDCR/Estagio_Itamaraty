import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import shutil
import os
import subprocess
import threading
import time

from utils.gerador_de_desligamentos.Gerador_de_desligamentos import Gerador_de_desligamentos

class Interface_desligamentos:
    def __init__(self, root):
        self.__frame_botoes = tk.Frame(root)
        self.__frame_botoes.pack(pady=20)

        self.__upload_mre_button = tk.Button(self.__frame_botoes, text='Upload MRE', command=lambda: self.__upload_file_Calculadora_de_Desligamentos('MRE'))
        self.__upload_mre_button.pack(pady=10)

        self.__upload_sce_button = tk.Button(self.__frame_botoes, text='Upload SCE', command=lambda: self.__upload_file_Calculadora_de_Desligamentos('SCE'))
        self.__upload_sce_button.pack(pady=10)

        self.__upload_recessos_button = tk.Button(self.__frame_botoes, text='Upload Recessos', command=lambda: self.__upload_file_Calculadora_de_Desligamentos('Recessos'))
        self.__upload_recessos_button.pack(pady=10)

        self.__entrada_cpf = tk.Entry(self.__frame_botoes, font=('Arial', 14))
        self.__entrada_cpf.pack(pady=20)
        self.__entrada_cpf.insert(0, 'Insira um CPF')
        self.__entrada_cpf.bind("<Button-1>", self.__limpar_texto)
        self.__entrada_cpf.bind("<KeyRelease> ", self.__formatar_cpf)

        self.__analyze_button = tk.Button(self.__frame_botoes, text='Resultado do Cálculo', command= self.__run_analyzer_Calculadora_de_Desligamentos)
        self.__analyze_button.pack(pady=10)

        self.__barra_progresso = ttk.Progressbar(self.__frame_botoes, orient="horizontal", length=300, mode="determinate")
        self.__barra_progresso.pack(pady=20)

        self.__text_area = tk.Text(root, height=30, width=90)
        self.__text_area.pack(pady=10)
        self.__ler_arquivo()

        self.__analyze_button.bind("<Button-1>", self.__inserir_texto)

        self.__frame_botoes.mainloop()
    
    def __inserir_texto(self, event):
        if self.__entrada_cpf.get() == "":
            self.__entrada_cpf.insert(0, "Insira um CPF")
            self.__frame_botoes.focus()

    def __limpar_texto(self, event):
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
        
    def __ler_arquivo(self):
        caminho_arquivo = "utils/interface_grafica/dados/informacoes_gerador_de_desligamentos.txt"
        try:
            with open(caminho_arquivo, 'r') as arquivo:
                conteudo = arquivo.read()

            self.__text_area.delete(1.0, tk.END)
            self.__text_area.insert(tk.END, conteudo)
            self.__text_area.config(state=tk.DISABLED)
        except FileNotFoundError:
            self.__text_area.insert(tk.END, "Arquivo não encontrado.")

    def __upload_file_Calculadora_de_Desligamentos(self, upload_type):
        file_path = filedialog.askopenfilename()
        if file_path:
            print(f'Arquivo selecionado para {upload_type}: {file_path}')

            destination_directory = 'utils/gerador_de_desligamentos/dados'

            if upload_type == 'MRE':
                new_file_name = 'Mre.xlsx'
            elif upload_type == 'SCE':
                new_file_name = 'Sce.xlsx'
            elif upload_type == 'Recessos':
                new_file_name = 'Recessos.xlsx'
            else:
                new_file_name = os.path.basename(file_path)

            destination_path = os.path.join(destination_directory, new_file_name)
            shutil.copy(file_path, destination_path)
            print(f'Arquivo copiado e renomeado para: {destination_path}')

    def __run_analyzer_Calculadora_de_Desligamentos(self):
        try:
            desligamentos = Gerador_de_desligamentos()
            self.__start_task(desligamentos.iniciar(str(self.__entrada_cpf.get())))
            messagebox.showinfo('Sucesso', 'Verifique sua pasta de Downloads')
            print('Calculadora de Desligamentos executado com sucesso.')
        except subprocess.CalledProcessError:
            messagebox.showerror('Erro', 'Erro ao executar o Calculo de Desligamento.')

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
