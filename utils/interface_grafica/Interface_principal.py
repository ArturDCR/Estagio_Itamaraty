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

        self.__largura = self.__root.winfo_screenwidth()
        self.__altura = self.__root.winfo_screenheight()

        self.__root.title('Central DTA')
        self.__root.geometry(f'{self.__largura}x{self.__altura}')

        self.__canvas = tk.Canvas(self.__root, width=self.__largura, height=self.__altura)
        self.__canvas.pack(fill="both", expand=True)

        self.__imagem_original = Image.open("utils/interface_grafica/dados/Meteoro.jpg")
        self.__imagem_tk = None

        self.__bg_id = self.__canvas.create_image(0, 0, anchor="nw", image=None)

        self.__redimensionar_imagem()

        self.__root.bind("<Configure>", self.__on_resize)

        self.__frame_botoes = tk.Frame(self.__root, bg="gray")
        self.__frame_botoes.place(relx=0.5, rely=0.0, anchor="n")

        botoes = [
            ('Conferência CIEE', 'Conferencia CIEE'),
            ('Análise de Faltas', 'Analisador de Faltas'),
            ('Análise de Vigência', 'Analisador de Vigência'),
            ('Análise SouGov', 'Analisador SouGov'),
            ('Análise de Recessos', 'Analisador de Recessos'),
            ('Análise de Descontos', 'Analisador de Descontos'),
            ('Gerador de Desligamentos', 'Gerador de Desligamentos'),
            ('Gerador de Declarações', 'Gerador de Declarações'),
            ('Gerador Mala Direta', 'Gerador Mala Direta')
        ]

        for texto, comando in botoes:
            tk.Button(self.__frame_botoes, text=texto, command=lambda c=comando: self.__abrir_tela(c)).pack(side=tk.LEFT, padx=10, pady=5)

        self.__root.mainloop()

    def __redimensionar_imagem(self):
        imagem_resized = self.__imagem_original.resize((self.__largura, self.__altura))
        self.__imagem_tk = ImageTk.PhotoImage(imagem_resized)
        self.__canvas.itemconfig(self.__bg_id, image=self.__imagem_tk)

    def __on_resize(self, event):
        self.__largura, self.__altura = event.width, event.height
        self.__redimensionar_imagem()
        self.__canvas.config(width=self.__largura, height=self.__altura)
        self.__canvas.coords(self.__bg_id, 0, 0)

    def __abrir_tela(self, titulo):
        nova_tela = Toplevel(self.__root)
        nova_tela.geometry(f'{850}x{600}')

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