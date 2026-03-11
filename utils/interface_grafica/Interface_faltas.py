import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import threading
from datetime import datetime

from utils.analise_de_faltas.Gerador_Analise_de_Faltas import Gerador_Analise_Faltas

class Interface_analise_de_faltas():
    def __init__(self, root):
        self.__frame_botoes = tk.Frame(root)
        self.__frame_botoes.pack(pady=20)

        self.__caminho_forms = None
        self.__caminho_mre = None
        self.__caminho_sce = None

        self.__variavel_desconto = None
        self.__descontos = ['VT','BE']
        self.__meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
        self.__anos = [str(ano) for ano in range(2023, datetime.now().year + 1)]

        self.__variavel_escolha_desconto = tk.StringVar(self.__frame_botoes)
        self.__variavel_escolha_desconto.set('Escolha o desconto')
        self.__variavel_escolha_desconto.trace_add("write", self.__Set_desconto)

        self.__caixa_opcoes_desconto = tk.OptionMenu(self.__frame_botoes, self.__variavel_escolha_desconto, *self.__descontos)
        self.__caixa_opcoes_desconto.pack(pady=10)

        self.__variavel_ano = None
        self.__variavel_escolha_ano = tk.StringVar(self.__frame_botoes)
        self.__variavel_escolha_ano.set('Escolha um Ano')
        self.__variavel_escolha_ano.trace_add("write", self.__Set_ano)

        self.__caixa_opcoes_ano = tk.OptionMenu(self.__frame_botoes, self.__variavel_escolha_ano, *self.__anos)
        self.__caixa_opcoes_ano.pack(pady=10)

        self.__variavel_mes = None
        self.__variavel_escolha_Mes = tk.StringVar(self.__frame_botoes)
        self.__variavel_escolha_Mes.set('Escolha um Mês')
        self.__variavel_escolha_Mes.trace_add("write", self.__Set_mes)

        self.__caixa_opcoes_mes = tk.OptionMenu(self.__frame_botoes, self.__variavel_escolha_Mes, *self.__meses)
        self.__caixa_opcoes_mes.pack(pady=10)

        self.__upload_forms_button = tk.Button(self.__frame_botoes, text='Upload Forms', command=lambda: self.__confirm_upload('Forms'))
        self.__upload_forms_button.pack(pady=10)

        self.__upload_mre_button = tk.Button(self.__frame_botoes, text='Upload MRE', command=lambda: self.__confirm_upload('MRE'))
        self.__upload_mre_button.pack(pady=10)

        self.__upload_sce_button = tk.Button(self.__frame_botoes, text='Upload SCE', command=lambda: self.__confirm_upload('SCE'))
        self.__upload_sce_button.pack(pady=10)

        self.__analyze_button = tk.Button(self.__frame_botoes, text='Resultado da Análise', command=self.__run_analyzer_Analise_de_Faltas)
        self.__analyze_button.pack(pady=10)

        self.__barra_progresso = ttk.Progressbar(self.__frame_botoes, orient="horizontal", length=300, mode="determinate")
        self.__barra_progresso.pack(pady=20)

    def __confirm_upload(self, tipo):
        resposta = messagebox.askyesno("Confirmação", "Deseja selecionar a planilha no seu computador?")
        if resposta:
            self.__upload_file_Analise_de_Faltas(tipo)

    def __Set_desconto(self, *args):
        self.__variavel_desconto = self.__variavel_escolha_desconto.get()
    
    def __Set_ano(self, *agrs):
        self.__variavel_ano = self.__variavel_escolha_ano.get()
    
    def __Set_mes(self, *args):
        self.__variavel_mes = self.__variavel_escolha_Mes.get()
    
    def __upload_file_Analise_de_Faltas(self, upload_type):
        file_path = filedialog.askopenfilename(
            title=f"Selecione a planilha {upload_type}",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        if file_path:
            if upload_type == 'Forms':
                self.__caminho_forms = file_path
            elif upload_type == 'MRE':
                self.__caminho_mre = file_path
            elif upload_type == 'SCE':
                self.__caminho_sce = file_path
            messagebox.showinfo('Sucesso', f'Arquivo {upload_type} carregado com sucesso.')
    
    def __run_analyzer_Analise_de_Faltas(self):
        if not (self.__caminho_forms and self.__caminho_mre and self.__caminho_sce):
            messagebox.showwarning('Aviso', 'Por favor, faça o Upload das 3 planilhas (Forms, MRE e SCE) antes de gerar.')
            return

        if self.__variavel_mes and self.__variavel_ano and self.__variavel_desconto:
            def tarefa_em_background():
                try:
                    faltas = Gerador_Analise_Faltas(self.__caminho_forms, self.__caminho_sce, self.__caminho_mre)
                    faltas.iniciar(self.__variavel_desconto, self.__variavel_mes, self.__variavel_ano)
                    messagebox.showinfo('Sucesso', 'Planilha de Faltas e Macro Hob geradas na pasta Downloads!')
                
                except ValueError as ve:
                    messagebox.showwarning('Aviso', str(ve))
                except Exception as e:
                    messagebox.showerror('Erro', f'Ocorreu um problema ao processar: {e}')
                finally:
                    self.__barra_progresso['value'] = 0
            
            self.__start_task(tarefa_em_background)
        else:
            messagebox.showerror('Erro', 'Escolha um Mês, um Ano e o tipo de Desconto.')

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