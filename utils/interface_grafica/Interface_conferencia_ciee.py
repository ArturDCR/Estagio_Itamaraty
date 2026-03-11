import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import threading

from ..conferencia_ciee.Conferencia_ciee import Conferencia_ciee

class Interface_conferencia_ciee():
    def __init__(self, root):
        self.__frame_botoes = tk.Frame(root)
        self.__frame_botoes.pack(pady=20)

        self.__caminho_ciee = None
        self.__caminho_mre = None
        self.__caminho_sce = None

        self.__upload_ciee_button = tk.Button(self.__frame_botoes, text='Upload CIEE', command=lambda: self.__confirm_upload('CIEE'))
        self.__upload_ciee_button.pack(pady=10)

        self.__upload_mre_button = tk.Button(self.__frame_botoes, text='Upload MRE', command=lambda: self.__confirm_upload('MRE'))
        self.__upload_mre_button.pack(pady=10)

        self.__upload_sce_button = tk.Button(self.__frame_botoes, text='Upload SCE', command=lambda: self.__confirm_upload('SCE'))
        self.__upload_sce_button.pack(pady=10)

        self.__analyze_button = tk.Button(self.__frame_botoes, text='Resultado da Conferência', command=self.__run_analyzer_conferencia_ciee)
        self.__analyze_button.pack(pady=10)

        self.__barra_progresso = ttk.Progressbar(self.__frame_botoes, orient="horizontal", length=300, mode="determinate")
        self.__barra_progresso.pack(pady=20)

    def __confirm_upload(self, tipo):
        resposta = messagebox.askyesno("Confirmação", "Deseja selecionar a planilha no seu computador?")
        if resposta:
            self.__upload_file_conferencia_ciee(tipo)
    
    def __upload_file_conferencia_ciee(self, upload_type):
        file_path = filedialog.askopenfilename(
            title=f"Selecione a planilha {upload_type}",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        if file_path:
            if upload_type == 'CIEE':
                self.__caminho_ciee = file_path
            elif upload_type == 'MRE':
                self.__caminho_mre = file_path
            elif upload_type == 'SCE':
                self.__caminho_sce = file_path
                
            messagebox.showinfo('Sucesso', f'Arquivo {upload_type} carregado:\n{os.path.basename(file_path)}')
    
    def __run_analyzer_conferencia_ciee(self):
        if not (self.__caminho_ciee and self.__caminho_mre and self.__caminho_sce):
            messagebox.showwarning('Aviso', 'Por favor, faça o Upload das 3 planilhas (CIEE, MRE e SCE) antes de gerar o resultado.')
            return
        
        def tarefa_em_background():
            try:
                conferencia_ciee = Conferencia_ciee(self.__caminho_mre, self.__caminho_sce, self.__caminho_ciee)
                conferencia_ciee.iniciar()
                messagebox.showinfo('Sucesso', 'Conferência concluída! Verifique sua pasta de Downloads.')
            except Exception as e:
                messagebox.showerror('Erro', f'Ocorreu um erro durante a análise: {e}')
            finally:
                self.__barra_progresso['value'] = 0

        self.__start_task(tarefa_em_background)
    
    def __start_task(self, func):
        thread = threading.Thread(target=func)
        thread.daemon = True
        thread.start()

        self.__barra_progresso['value'] = 0
        def atualizar_barra():
            if thread.is_alive():
                if self.__barra_progresso['value'] < 90:
                    self.__barra_progresso['value'] += 5
                self.__frame_botoes.after(500, atualizar_barra)
            else:
                self.__barra_progresso['value'] = 100

        atualizar_barra()