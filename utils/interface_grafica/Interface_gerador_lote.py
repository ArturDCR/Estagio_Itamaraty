import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import threading
from datetime import datetime

from ..gerador_de_lote.Gerador_de_lote import Gerador_de_lote

class Interface_gerador_lote():
    def __init__(self, root):
        self.__frame_botoes = tk.Frame(root)
        self.__frame_botoes.pack(pady=20)

        self.__caminho_mre = None 

        self.__anos = [str(ano) for ano in range(2020, datetime.now().year + 1)]

        self.__variavel_ano_inicio = None
        self.__variavel_ano_final = None

        self.__variavel_escolha_ano_inicio = tk.StringVar(self.__frame_botoes)
        self.__variavel_escolha_ano_inicio.set('Escolha o ano inicial')
        self.__variavel_escolha_ano_inicio.trace_add("write", self.__Set_ano)

        self.__caixa_opcoes_ano_inicio = tk.OptionMenu(self.__frame_botoes, self.__variavel_escolha_ano_inicio, *self.__anos)
        self.__caixa_opcoes_ano_inicio.pack(pady=10)

        self.__variavel_escolha_ano_final = tk.StringVar(self.__frame_botoes)
        self.__variavel_escolha_ano_final.set('Escolha o ano final')
        self.__variavel_escolha_ano_final.trace_add("write", self.__Set_ano)

        self.__caixa_opcoes_ano_final = tk.OptionMenu(self.__frame_botoes, self.__variavel_escolha_ano_final, *self.__anos)
        self.__caixa_opcoes_ano_final.pack(pady=10)

        self.__entrada_cpf = tk.Entry(self.__frame_botoes, font=('Arial', 14))
        self.__entrada_cpf.pack(pady=20)
        self.__entrada_cpf.insert(0, 'Insira um CPF')
        self.__entrada_cpf.bind("<Button-1>", self.__limpar_texto_cpf)
        self.__entrada_cpf.bind("<KeyRelease> ", self.__formatar_cpf)

        self.__upload_mre_button = tk.Button(self.__frame_botoes, text='Upload MRE', command=lambda: self.__confirm_upload('MRE'))
        self.__upload_mre_button.pack(pady=10)

        self.__analyze_button = tk.Button(self.__frame_botoes, text='Resultado da Análise', command=self.__run_analyzer_gerador_lote)
        self.__analyze_button.pack(pady=10)

        self.__barra_progresso = ttk.Progressbar(self.__frame_botoes, orient="horizontal", length=300, mode="determinate")
        self.__barra_progresso.pack(pady=20)

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

    def __limpar_texto_cpf(self, event):
        if self.__entrada_cpf.get() == "Insira um CPF":
            self.__entrada_cpf.delete(0, tk.END)
    
    def __ajustar_cpf_erro(self):
        self.__entrada_cpf.delete(0, tk.END)
        self.__entrada_cpf.insert(0, 'Insira um CPF')

    def __Set_ano(self, *agrs):
        self.__variavel_ano_inicio = self.__variavel_escolha_ano_inicio.get()
        self.__variavel_ano_final = self.__variavel_escolha_ano_final.get()

    def __confirm_upload(self, tipo):
        resposta = messagebox.askyesno("Confirmação", "Deseja selecionar a planilha no seu computador?")
        if resposta:
            self.__upload_file_gerador_lote(tipo)
    
    def __upload_file_gerador_lote(self, upload_type):
        file_path = filedialog.askopenfilename(
            title="Selecione a planilha MRE",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        if file_path:
            self.__caminho_mre = file_path
            messagebox.showinfo('Sucesso', f'Arquivo carregado:\n{os.path.basename(file_path)}')
    
    def __run_analyzer_gerador_lote(self):
        if not self.__caminho_mre:
            messagebox.showwarning('Aviso', 'Por favor, faça o Upload do MRE antes de gerar.')
            return

        cpf_texto = str(self.__entrada_cpf.get())
        cpf_valido = cpf_texto != '' and cpf_texto != 'Insira um CPF' and len(cpf_texto) == 14
        anos_validos = self.__variavel_ano_inicio is not None and self.__variavel_ano_final is not None

        if cpf_valido and anos_validos:
            lote = Gerador_de_lote(caminho_planilha_mre=self.__caminho_mre)
            
            def tarefa_em_background():
                try:
                    lote.iniciar(cpf_texto, self.__variavel_ano_inicio, self.__variavel_ano_final)
                    messagebox.showinfo('Sucesso', 'Lote gerado com sucesso! Verifique sua pasta de Downloads.')
                except Exception as e:
                    messagebox.showerror('Erro', f'Erro ao processar: {e}')
                finally:
                    self.__variavel_ano_inicio = None
                    self.__variavel_escolha_ano_inicio.set('Escolha o ano inicial')
                    self.__variavel_ano_final = None
                    self.__variavel_escolha_ano_final.set('Escolha o ano final')
                    self.__ajustar_cpf_erro()
                    self.__barra_progresso['value'] = 0

            self.__start_task(tarefa_em_background)
        else:
            messagebox.showerror('Erro', 'Verifique os campos de CPF e os Anos selecionados.')

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