"""
equipamentos.py
----------------
Módulo responsável por carregar a lista de equipamentos (EPIs) e seus
respectivos CAs (Certificado de Aprovação) a partir do arquivo "equipamentos.xlsx".

Esse arquivo funciona como uma das bases de dados do sistema: ele é lido
sempre que o programa é executado, então qualquer equipamento novo que for
adicionado diretamente na planilha "equipamentos.xlsx" (na pasta raiz do
projeto) estará automaticamente disponível no sistema na próxima execução,
sem precisar alterar o código.
"""

import os
import openpyxl

# Caminho absoluto do arquivo de equipamentos, localizado na pasta raiz do
# projeto (mesma pasta deste arquivo .py).
#
# Importante: usamos o caminho absoluto baseado na localização deste arquivo
# (e não um caminho relativo simples, como "equipamentos.xlsx") porque o
# diretório de trabalho ("cwd") do Python pode ser diferente da pasta do
# projeto dependendo de como o programa é executado (ex.: ao rodar pelo
# botão "Run" do VSCode, pelo depurador, ou a partir de um atalho). Um
# caminho relativo simples falharia silenciosamente nesses casos, fazendo a
# lista de equipamentos aparecer vazia na interface.
PASTA_DO_PROJETO = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_EQUIPAMENTOS = os.path.join(PASTA_DO_PROJETO, "equipamentos.xlsx")


def carregar_equipamentos(caminho_arquivo: str = ARQUIVO_EQUIPAMENTOS) -> dict:
    """
    Lê o arquivo equipamentos.xlsx e retorna um dicionário no formato:
        {"Nome do Equipamento": "CA" (ou "" quando o CA não estiver preenchido)}

    Regras de leitura:
    - A primeira linha da planilha é o cabeçalho ("Equipamento" / "CA") e é ignorada.
    - Linhas totalmente vazias são ignoradas.
    - Quando o CA não estiver preenchido na planilha, o valor é mantido vazio ("").
    """
    equipamentos = {}

    # Caso o arquivo não seja encontrado (ex.: execução em outra pasta/máquina),
    # o sistema não deve travar: apenas retorna a lista vazia.
    if not os.path.exists(caminho_arquivo):
        return equipamentos

    planilha = openpyxl.load_workbook(caminho_arquivo, data_only=True)
    aba = planilha.active  # utiliza a primeira aba (planilha ativa) do arquivo

    # min_row=2 pula a linha de cabeçalho ("Equipamento", "CA")
    for linha in aba.iter_rows(min_row=2, values_only=True):
        if not linha or linha[0] in (None, ""):
            continue  # ignora linhas vazias

        nome_equipamento = str(linha[0]).strip()
        if not nome_equipamento:
            continue

        # Se a coluna do CA existir e tiver valor, usa esse valor; caso
        # contrário, mantém o CA vazio (conforme solicitado no projeto).
        valor_ca = linha[1] if len(linha) > 1 else None
        ca = str(valor_ca).strip() if valor_ca not in (None, "") else ""

        equipamentos[nome_equipamento] = ca

    return equipamentos
