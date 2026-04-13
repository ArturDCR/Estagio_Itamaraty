import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import threading

from utils.gerador_de_declaracao.Gerador_de_declaracao import Gerador_de_declaracao

class Interface_declaracao:
    def __init__(self, root):
        self.__frame_botoes = tk.Frame(root)
        self.__frame_botoes.pack(pady=20)

        self.__caminho_sce = None

        self.__upload_sce_button = tk.Button(self.__frame_botoes, text='Upload SCE', command=lambda: self.__upload_file_Gerador_de_declaracao('SCE'))
        self.__upload_sce_button.pack(pady=10)

        self.__entrada_cpf = tk.Entry(self.__frame_botoes, font=('Arial', 14))
        self.__entrada_cpf.pack(pady=20)
        self.__entrada_cpf.insert(0, 'Insira um CPF')
        self.__entrada_cpf.bind("<Button-1>", self.__limpar_texto_cpf)
        self.__entrada_cpf.bind("<KeyRelease>", self.__formatar_cpf) 

        self.__analyze_button = tk.Button(self.__frame_botoes, text='Resultado da Declaração', command=self.__run_analyzer_Gerador_de_declaracao)
        self.__analyze_button.pack(pady=10)

        self.__barra_progresso = ttk.Progressbar(self.__frame_botoes, orient="horizontal", length=300, mode="determinate")
        self.__barra_progresso.pack(pady=20)

        self.__analyze_button.bind("<Button-1>", self.__inserir_texto)
    
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
            pass
        elif len(cpf) <= 6:
            cpf = f'{cpf[:3]}.{cpf[3:]}'
        elif len(cpf) <= 9:
            cpf = f'{cpf[:3]}.{cpf[3:6]}.{cpf[6:]}'
        else:
            cpf = f'{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:11]}'

        self.__entrada_cpf.delete(0, tk.END)
        self.__entrada_cpf.insert(0, cpf)
     
    def __upload_file_Gerador_de_declaracao(self, upload_type):
        file_path = filedialog.askopenfilename(
            title="Selecione a planilha SCE",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        if file_path:
            self.__caminho_sce = file_path
            messagebox.showinfo('Sucesso', f'Arquivo {upload_type} carregado:\n{os.path.basename(file_path)}')

    def __run_analyzer_Gerador_de_declaracao(self):
        if not self.__caminho_sce:
            messagebox.showwarning('Aviso', 'Por favor, faça o Upload da planilha SCE antes de gerar a declaração.')
            return

        cpf_texto = str(self.__entrada_cpf.get())
        
        if cpf_texto != '' and cpf_texto != 'Insira um CPF' and len(cpf_texto) == 14:
            declaracao = Gerador_de_declaracao(caminho_planilha_sce=self.__caminho_sce)
            
            def tarefa_em_background():
                try:
                    declaracao.iniciar(cpf_texto)
                    messagebox.showinfo('Sucesso', 'Declaração gerada! Verifique sua pasta de Downloads.')
                except Exception as e:
                    messagebox.showerror('Erro', f'Ocorreu um erro: {e}')
                finally:
                    self.__barra_progresso['value'] = 0
            
            self.__start_task(tarefa_em_background)
        else:
            messagebox.showerror('Erro', 'Digite um CPF válido')

    def __start_task(self, func):
        thread = threading.Thread(target=func)
        thread.daemon = True
        thread.start()

        self.__barra_progresso['value'] = 0
        def atualizar_barra():
            if thread.is_alive():
                if self.__barra_progresso['value'] < 90:
                    self.__barra_progresso['value'] += 10
                self.__frame_botoes.after(500, atualizar_barra)
            else:
                self.__barra_progresso['value'] = 100

        atualizar_barra()