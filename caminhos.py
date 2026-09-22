"""
caminhos.py
-----------
Módulo utilitário responsável por descobrir a pasta onde os arquivos de
dados do sistema (equipamentos.xlsx, modelo_ficha_epi.xlsx, e a pasta
modelos_x_cliente) devem ser procurados — tanto quando o sistema roda como
script Python (`python main.py`) quanto quando roda como executável (.exe)
gerado pelo PyInstaller.
"""

import os
import sys


def obter_pasta_base() -> str:
    """
    Retorna a pasta "raiz" do sistema, onde ficam os arquivos de dados.

    - Rodando como script (`python main.py`): é a pasta onde estão os
      arquivos .py do projeto.
    - Rodando como executável gerado pelo PyInstaller (.exe), inclusive no
      modo "--onefile": é a pasta onde o .exe está — e NÃO a pasta
      temporária de extração que o modo --onefile cria a cada execução
      (sys._MEIPASS, algo como "C:\\Users\\...\\AppData\\Local\\Temp\\_MEIxxxxx").
      Usar a pasta temporária faria o sistema nunca encontrar
      equipamentos.xlsx, modelo_ficha_epi.xlsx e modelos_x_cliente, já que
      esses arquivos não fazem parte do .exe.

    Por isso, ao distribuir o .exe, os arquivos equipamentos.xlsx,
    modelo_ficha_epi.xlsx e a pasta modelos_x_cliente devem ficar na MESMA
    pasta que o .exe (veja o README.md para o passo a passo de geração do
    executável).
    """
    # `sys.frozen` só existe quando o programa foi empacotado (PyInstaller,
    # cx_Freeze, etc.) — é a forma padrão de detectar isso em tempo de execução.
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)

    return os.path.dirname(os.path.abspath(__file__))
