import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import shutil
import os
import subprocess
import threading
import time

from ..conferencia_ciee.Conferencia_ciee import Conferencia_ciee

class Interface_conferencia_ciee():
    def __init__(self, root):
        self.__frame_botoes = tk.Frame(root)
        self.__frame_botoes.pack(pady=20)

        self.__upload_ciee_button = tk.Button(self.__frame_botoes, text='Upload CIEE', command=lambda: self.__upload_file_conferencia_ciee('CIEE'))
        self.__upload_ciee_button.pack(pady=10)

        self.__upload_mre_button = tk.Button(self.__frame_botoes, text='Upload MRE', command=lambda: self.__upload_file_conferencia_ciee('MRE'))
        self.__upload_mre_button.pack(pady=10)

        self.__upload_sce_button = tk.Button(self.__frame_botoes, text='Upload SCE', command=lambda: self.__upload_file_conferencia_ciee('SCE'))
        self.__upload_sce_button.pack(pady=10)

        self.__analyze_button = tk.Button(self.__frame_botoes, text='Resultado da Conferência', command= self.__run_analyzer_conferencia_ciee)
        self.__analyze_button.pack(pady=10)

        self.__barra_progresso = ttk.Progressbar(self.__frame_botoes, orient="horizontal", length=300, mode="determinate")
        self.__barra_progresso.pack(pady=20)

        self.__text_area = tk.Text(root, height=10, width=90)
        self.__text_area.pack(pady=10)
        self.__ler_arquivo()

        self.__frame_botoes.mainloop()
    
    def __ler_arquivo(self):
        caminho_arquivo = "utils/interface_grafica/dados/informacoes_conferencia_ciee.txt"
        try:
            with open(caminho_arquivo, 'r') as arquivo:
                conteudo = arquivo.read()

            self.__text_area.delete(1.0, tk.END)
            self.__text_area.insert(tk.END, conteudo)
            self.__text_area.config(state=tk.DISABLED)
        except FileNotFoundError:
            self.__text_area.insert(tk.END, "Arquivo não encontrado.")

    def __upload_file_conferencia_ciee(self, upload_type):
        file_path = filedialog.askopenfilename()
        if file_path:
            print(f'Arquivo selecionado para {upload_type}: {file_path}')
            
            destination_directory = 'utils/conferencia_ciee/dados'
            
            if upload_type == 'CIEE':
                new_file_name = 'Ciee.xlsx'
            elif upload_type == 'MRE':
                new_file_name = 'Mre.xlsx'
            elif upload_type == 'SCE':
                new_file_name = 'Sce.xlsx'
            else:
                new_file_name = os.path.basename(file_path)

            destination_path = os.path.join(destination_directory, new_file_name)
            shutil.copy(file_path, destination_path)
            print(f'Arquivo copiado e renomeado para: {destination_path}')
            self.__frame_botoes.mainloop()
    
    def __run_analyzer_conferencia_ciee(self):
        try:
            conferencia_ciee = Conferencia_ciee()
            self.__start_task(conferencia_ciee.iniciar())
            messagebox.showinfo('Sucesso', 'Verifique sua pasta de Downloads')
            print('Conferencia CIEE executado com sucesso.')
        except subprocess.CalledProcessError:
            messagebox.showerror('Erro', 'Erro ao executar a Conferencia CIEE.')
    
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