import tkinter as tk
from tkinter import Toplevel

from utils.interface_grafica.Interface_conferencia_ciee import Interface_conferencia_ciee as Ciee
from utils.interface_grafica.Interface_faltas import Interface_analise_de_faltas as Faltas
from utils.interface_grafica.Interface_declaracao import Interface_declaracao as declaracao
from utils.interface_grafica.Interface_gerador_lote import Interface_gerador_lote as lote

class Interface_principal:
    def __init__(self):
        self.__root = tk.Tk()

        self.__largura = 850
        self.__altura = 600

        self.__root.title('Central DTA')
        self.__root.geometry(f'{self.__largura}x{self.__altura}')

        self.__root.resizable(False, False)

        self.__frame_botoes = tk.Frame(self.__root)
        self.__frame_botoes.place(relx=0.5, rely=0.0, anchor="n")

        botoes = [
            ('Conferência CIEE', 'Conferência CIEE'),
            ('Análise de Faltas', 'Análise de Faltas'),
            ('Gerador de Declarações', 'Gerador de Declarações'),
            ('Gerador Ficha Financeira em Lote', 'Gerador Ficha Financeira em Lote')
        ]

        for texto, comando in botoes:
            tk.Button(self.__frame_botoes, text=texto, command=lambda c=comando: self.__abrir_tela(c)).pack(side=tk.LEFT, padx=10, pady=5)

        self.__root.mainloop()

    def __abrir_tela(self, titulo):
        nova_tela = Toplevel(self.__root)
        nova_tela.geometry(f'{self.__largura}x{self.__altura}')
        nova_tela.resizable(False, False)
        nova_tela.title(titulo)

        if titulo == 'Conferência CIEE':
            Ciee(nova_tela)

        elif titulo == 'Análise de Faltas':
            Faltas(nova_tela)
        
        elif titulo == 'Gerador de Declarações':
            declaracao(nova_tela)

        elif titulo == 'Gerador Ficha Financeira em Lote':
            lote(nova_tela)