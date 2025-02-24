import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import shutil
import os
import subprocess
from datetime import datetime
import threading
import time

from ..gerador_de_lote.Gerador_de_lote import Gerador_de_lote

class Interface_gerador_lote():
    def __init__(self, root):
        self.__frame_botoes = tk.Frame(root)
        self.__frame_botoes.pack(pady=20)

        self.__anos = [str(ano) for ano in range(2023, datetime.now().year + 1)]

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

        self.__analyze_button = tk.Button(self.__frame_botoes, text='Resultado da Análise', command= self.__run_analyzer_gerador_lote)
        self.__analyze_button.pack(pady=10)

        self.__barra_progresso = ttk.Progressbar(self.__frame_botoes, orient="horizontal", length=300, mode="determinate")
        self.__barra_progresso.pack(pady=20)

        self.__frame_botoes.mainloop()

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

    def __limpar_texto_cpf(self, event):
        if self.__entrada_cpf.get() == "Insira um CPF":
            self.__entrada_cpf.delete(0, tk.END)
    
    def __ajustar_cpf_erro(self):
        self.__entrada_cpf.delete(0, tk.END)
        self.__entrada_cpf.insert(0, 'Insira um CPF')

    def __confirm_upload(self, tipo):
        resposta = messagebox.askyesno("Confirmação", "Caso este arquivo já tenha sido enviado, não é necessário enviá-lo novamente, a menos que seja uma atualização. Deseja enviar um novo?")
        if resposta:
            self.__upload_file_gerador_lote(tipo)
    
    def __Set_ano(self, *agrs):
        self.__variavel_ano_inicio = self.__variavel_escolha_ano_inicio.get()
        self.__variavel_ano_final = self.__variavel_escolha_ano_final.get()
    
    def __upload_file_gerador_lote(self, upload_type):
        file_path = filedialog.askopenfilename()
        if file_path:
            print(f'Arquivo selecionado para {upload_type}: {file_path}')
            
            destination_directory = 'utils/data'
            
            if upload_type == 'MRE':
                new_file_name = 'Mre.xlsx'
            else:
                new_file_name = os.path.basename(file_path)

            destination_path = os.path.join(destination_directory, new_file_name)
            shutil.copy(file_path, destination_path)
            print(f'Arquivo copiado e renomeado para: {destination_path}')
    
    def __run_analyzer_gerador_lote(self):
        try:
            if str(self.__entrada_cpf.get()) != '' and str(self.__entrada_cpf.get()) != 'Insira um CPF' and len(self.__entrada_cpf.get()) == 14 and self.__variavel_ano_inicio != None and self.__variavel_ano_final != None:
                lote = Gerador_de_lote()
                self.__start_task(lote.iniciar(str(self.__entrada_cpf.get()), self.__variavel_ano_inicio, self.__variavel_ano_final))
                messagebox.showinfo('Sucesso', 'Verifique sua pasta de Downloads')
                print('Gerador de Lote Financeiro executado com sucesso.')
                self.__variavel_ano_inicio = None
                self.__variavel_escolha_ano_inicio.set('Escolha o ano inicial')
                self.__variavel_ano_final = None
                self.__variavel_escolha_ano_final.set('Escolha o ano final')
                self.__ajustar_cpf_erro()
            else:
                messagebox.showerror('Erro', 'Verifique os campos')
                self.__variavel_ano_inicio = None
                self.__variavel_escolha_ano_inicio.set('Escolha o ano inicial')
                self.__variavel_ano_final = None
                self.__variavel_escolha_ano_final.set('Escolha o ano final')
                self.__ajustar_cpf_erro()
        except subprocess.CalledProcessError:
            messagebox.showerror('Erro', 'Erro ao executar o Gerador de lote financeiro.')
            self.__variavel_ano_inicio = None
            self.__variavel_escolha_ano_inicio.set('Escolha o ano inicial')
            self.__variavel_ano_final = None
            self.__variavel_escolha_ano_final.set('Escolha o ano final')
            self.__ajustar_cpf_erro()
    
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