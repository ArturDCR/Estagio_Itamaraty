import tkinter as tk
from tkinter import Toplevel
from PIL import Image, ImageTk

from utils.interface_grafica.Interface_dados import Interface_analise_de_dados as Dados
from utils.interface_grafica.Interface_faltas import Interface_analise_de_faltas as Faltas
from utils.interface_grafica.Interface_vigencia import Interface_vigencia as vigencias
from utils.interface_grafica.Interface_SouGov import Interface_SouGov as SouGov
from utils.interface_grafica.Interface_descontos import Interface_descontos as desconto
from utils.interface_grafica.Interface_recessos import Interface_recessos as recessos
from utils.interface_grafica.Interface_desligamentos import Interface_desligamentos as desligamentos

class Interface_principal:
    def __init__(self):
        self.__root = tk.Tk()
        self.__root.title('Central DTA')
        self.__root.resizable(False, False)

        self.__root.geometry('1200x400')

        self.__frame_botoes = tk.Frame(self.__root)
        self.__frame_botoes.pack(pady=20)

        self.__analisador_de_dados_button = tk.Button(self.__frame_botoes, text='Análise CIEE', command=lambda: self.__abrir_tela('Analisador de Dados CIEE'))
        self.__analisador_de_dados_button.pack(side=tk.LEFT, padx=10)

        self.__analisador_de_faltas_button = tk.Button(self.__frame_botoes, text='Análise de Faltas', command=lambda: self.__abrir_tela('Analisador de Faltas'))
        self.__analisador_de_faltas_button.pack(side=tk.LEFT, padx=10)

        self.__analisador_de_vigencia_button = tk.Button(self.__frame_botoes, text='Análise de Vigência', command=lambda: self.__abrir_tela('Analisador de Vigência'))
        self.__analisador_de_vigencia_button.pack(side=tk.LEFT, padx=10)

        self.__analisador_SouGov_button = tk.Button(self.__frame_botoes, text='Análise SouGov', command=lambda: self.__abrir_tela('Analisador SouGov'))
        self.__analisador_SouGov_button.pack(side=tk.LEFT, padx=10)

        self.__analisador_de_recessos_button = tk.Button(self.__frame_botoes, text='Análise de Recessos', command=lambda: self.__abrir_tela('Analisador de Recessos'))
        self.__analisador_de_recessos_button.pack(side=tk.LEFT, padx=10)

        self.__analisador_de_decontos_button = tk.Button(self.__frame_botoes, text='Análise de Descontos', command=lambda: self.__abrir_tela('Analisador de Descontos'))
        self.__analisador_de_decontos_button.pack(side=tk.LEFT, padx=10)

        self.__analisador_de_desligamentos_button = tk.Button(self.__frame_botoes, text='Gerador de Desligamentos', command=lambda: self.__abrir_tela('Gerador de Desligamentos'))
        self.__analisador_de_desligamentos_button.pack(side=tk.LEFT, padx=10)

        self.__imagem = Image.open("utils/interface_grafica/dados/MRE.jpg")
        self.__imagem_tk = ImageTk.PhotoImage(self.__imagem)

        self.__label_imagem = tk.Label(self.__root, image= self.__imagem_tk)
        self.__label_imagem.pack()

        self.__root.mainloop()
        
    def __abrir_tela(self, titulo):
        nova_tela = Toplevel(self.__root)
        nova_tela.resizable(False, False)

        nova_tela.title(titulo)

        nova_tela.geometry('854x480')

        if titulo == 'Analisador de Dados CIEE':
            Dados(nova_tela)

        elif titulo == 'Analisador de Faltas':
            nova_tela.geometry('854x700')
            Faltas(nova_tela)

        elif titulo == 'Analisador de Vigência':
            vigencias(nova_tela)

        elif titulo == 'Analisador SouGov':
            nova_tela.geometry('854x600')
            SouGov(nova_tela)

        elif titulo == 'Analisador de Descontos':
            desconto(nova_tela)

        elif titulo == 'Analisador de Recessos':
            recessos(nova_tela)

        elif titulo == 'Gerador de Desligamentos':
            desligamentos(nova_tela)