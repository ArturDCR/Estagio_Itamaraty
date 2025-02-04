import tkinter as tk
from tkinter import Toplevel
from PIL import Image, ImageTk

from utils.interface_grafica.Interface_conferencia_ciee import Interface_conferencia_ciee as Ciee
from utils.interface_grafica.Interface_faltas import Interface_analise_de_faltas as Faltas
from utils.interface_grafica.Interface_vigencia import Interface_vigencia as vigencias
from utils.interface_grafica.Interface_SouGov import Interface_SouGov as SouGov
from utils.interface_grafica.Interface_descontos import Interface_descontos as desconto
from utils.interface_grafica.Interface_recessos import Interface_recessos as recessos
from utils.interface_grafica.Interface_desligamentos import Interface_desligamentos as desligamentos
from utils.interface_grafica.Interface_declaracao import Interface_declaracao as declaracao
from utils.interface_grafica.Interface_mala_direta import Interface_mala_direta as mala_direta

class Interface_principal:
    def __init__(self):
        self.__root = tk.Tk()
        self.__root.title('Central DTA')
        self.__root.resizable(self.__root.winfo_screenwidth(), self.__root.winfo_screenheight())

        self.__frame_botoes = tk.Frame(self.__root)
        self.__frame_botoes.pack(pady=20)

        self.__conferencia_ciee_button = tk.Button(self.__frame_botoes, text='Conferência CIEE', command=lambda: self.__abrir_tela('Conferencia CIEE'))
        self.__conferencia_ciee_button.pack(side=tk.LEFT, padx=10)

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

        self.__gerador_de_declaracao_button = tk.Button(self.__frame_botoes, text='Gerador de Declarações', command=lambda: self.__abrir_tela('Gerador de Declarações'))
        self.__gerador_de_declaracao_button.pack(side=tk.LEFT, padx=10)
    
        self.__gerador_mala_direta_button = tk.Button(self.__frame_botoes, text='Gerador Mala Direta', command=lambda: self.__abrir_tela('Gerador Mala Direta'))
        self.__gerador_mala_direta_button.pack(side=tk.LEFT, padx=10)

        self.__imagem = Image.open("utils/interface_grafica/dados/MRE.jpg")
        self.__imagem_tk = ImageTk.PhotoImage(self.__imagem)

        self.__label_imagem = tk.Label(self.__root, image= self.__imagem_tk)
        self.__label_imagem.pack()

        self.__root.mainloop()
        
    def __abrir_tela(self, titulo):
        nova_tela = Toplevel(self.__root)
        nova_tela.resizable(self.__root.winfo_screenwidth(), self.__root.winfo_screenheight())

        nova_tela.title(titulo)

        if titulo == 'Conferencia CIEE':
            Ciee(nova_tela)

        elif titulo == 'Analisador de Faltas':
            Faltas(nova_tela)

        elif titulo == 'Analisador de Vigência':
            vigencias(nova_tela)

        elif titulo == 'Analisador SouGov':
            SouGov(nova_tela)

        elif titulo == 'Analisador de Descontos':
            desconto(nova_tela)

        elif titulo == 'Analisador de Recessos':
            recessos(nova_tela)

        elif titulo == 'Gerador de Desligamentos':
            desligamentos(nova_tela)
        
        elif titulo == 'Gerador de Declarações':
            declaracao(nova_tela)
        
        elif titulo == 'Gerador Mala Direta':
            mala_direta(nova_tela)