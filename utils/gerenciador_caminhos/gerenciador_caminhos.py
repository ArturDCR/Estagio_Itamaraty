import os
import sys

class GerenciadorCaminhos:
    @staticmethod
    def obter_caminho_recurso(caminho_relativo):
        try:
            base_path = sys._MEIPASS
        except AttributeError:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, caminho_relativo)