"""
caminhos.py
-----------
Módulo utilitário que resolve a "pasta base" do sistema: a pasta onde ficam
os arquivos de dados (equipamentos.xlsx, modelo_ficha_epi.xlsx e a pasta
modelos_x_cliente).

Usado por equipamentos.py, gerar_ficha.py e padroes_admissao.py, para que
os três procurem os arquivos de dados sempre no mesmo lugar.
"""

import os
import sys


def obter_pasta_base() -> str:
    """
    Retorna o caminho absoluto da pasta base do sistema:
    - Rodando como executável (.exe gerado pelo PyInstaller): a pasta onde
      está o .exe. No modo --onefile, sys.executable aponta para o .exe
      real, e não para a pasta temporária de extração (_MEIxxxxx), por isso
      os arquivos de dados devem ficar ao lado do .exe.
    - Rodando como script (ex.: pelo VSCode): a pasta onde está este arquivo,
      ou seja, a pasta raiz do projeto.
    """
    # O PyInstaller define o atributo "frozen" em sys quando o programa
    # está rodando empacotado como executável.
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)

    return os.path.dirname(os.path.abspath(__file__))