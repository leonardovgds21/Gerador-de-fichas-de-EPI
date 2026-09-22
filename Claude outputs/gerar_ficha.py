"""
gerar_ficha.py
--------------
Módulo responsável por:
    1. Ler o modelo "modelo_ficha_epi.xlsx" (pasta raiz do projeto).
    2. Preencher os dados do colaborador e os itens de EPI informados na
       interface, seguindo o layout definido do modelo.
    3. Converter o resultado em PDF, salvando como "Ficha de EPI - <nome do
       colaborador>.pdf" na área de trabalho do usuário — assim, cada
       colaborador gera um PDF com nome próprio, sem precisar renomear
       manualmente nem sobrescrever o anterior.

Este módulo NUNCA altera o arquivo modelo_ficha_epi.xlsx original: os dados
são preenchidos em uma cópia temporária, que é convertida em PDF e depois
descartada.
"""

import os
import re
import shutil
import tempfile
import datetime
import openpyxl

# Caminho absoluto do modelo de ficha, dentro da pasta raiz do projeto (mesma
# pasta deste arquivo .py) — evita depender do diretório de trabalho de onde
# o programa foi executado (mesmo problema já corrigido em equipamentos.py).
PASTA_DO_PROJETO = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_MODELO_FICHA = os.path.join(PASTA_DO_PROJETO, "modelo_ficha_epi.xlsx")

NOME_ABA = "Ficha FRENTE"

# --- Mapeamento de células do modelo (definido pelo usuário) ---------------
CELULA_NOME = "A4"
CELULA_FUNCAO = "E4"
CELULA_ADMISSAO = "H4"
CELULA_DEMISSAO = "I4"

LINHA_INICIAL_ITENS = 10   # primeira linha da tabela de itens no modelo
COLUNA_ITEM_EQUIPAMENTO = "A"
COLUNA_ITEM_QUANTIDADE = "C"
COLUNA_ITEM_DATA = "D"
COLUNA_ITEM_CA = "F"

# Prefixo do nome do PDF gerado; o nome do colaborador é acrescentado a ele
# (ex.: "Ficha de EPI - João da Silva.pdf").
PREFIXO_NOME_ARQUIVO_SAIDA = "Ficha de EPI"

# Caracteres não permitidos em nomes de arquivo no Windows.
_CARACTERES_INVALIDOS_NOME_ARQUIVO = re.compile(r'[\\/:*?"<>|]')


def _converter_data(texto_data: str):
    """
    Tenta converter o texto da data (formato dd/mm/aaaa) para um objeto de
    data real, para que a célula do Excel utilize a formatação de data já
    existente no modelo. Se não for possível converter (campo vazio ou em
    outro formato), devolve o próprio texto digitado, sem travar o sistema.
    """
    texto_data = (texto_data or "").strip()
    if not texto_data:
        return ""
    try:
        return datetime.datetime.strptime(texto_data, "%d/%m/%Y").date()
    except ValueError:
        return texto_data  # mantém o texto exatamente como foi digitado


def _converter_quantidade(texto_quantidade: str):
    """Converte a quantidade para número inteiro quando possível (fica mais
    correta dentro da planilha); caso contrário, mantém o texto original."""
    texto_quantidade = (texto_quantidade or "").strip()
    if texto_quantidade.isdigit():
        return int(texto_quantidade)
    return texto_quantidade


def _sanitizar_nome_arquivo(texto: str) -> str:
    """
    Remove caracteres não permitidos em nomes de arquivo no Windows
    (\\ / : * ? " < > |) e espaços repetidos, para que o nome do colaborador
    possa ser usado com segurança no nome do PDF gerado.
    """
    texto = _CARACTERES_INVALIDOS_NOME_ARQUIVO.sub("", texto or "")
    return " ".join(texto.split())  # normaliza espaços repetidos/nas pontas


def _caminho_disponivel(caminho: str) -> str:
    """
    Se o caminho já existir (ex.: uma ficha já foi gerada antes para o mesmo
    colaborador), acrescenta um sufixo numérico (" (2)", " (3)", ...) até
    encontrar um nome de arquivo livre, para nunca sobrescrever um PDF
    gerado anteriormente.
    """
    if not os.path.exists(caminho):
        return caminho

    pasta, nome_arquivo = os.path.split(caminho)
    nome_base, extensao = os.path.splitext(nome_arquivo)
    contador = 2
    while True:
        novo_caminho = os.path.join(pasta, f"{nome_base} ({contador}){extensao}")
        if not os.path.exists(novo_caminho):
            return novo_caminho
        contador += 1


def _obter_pasta_area_de_trabalho() -> str:
    """
    Retorna o caminho da pasta "Área de Trabalho" (Desktop) do usuário no
    Windows. Primeiro tenta ler o caminho real no registro do Windows (cobre
    o caso da pasta ter sido redirecionada, por exemplo pelo OneDrive);
    se não for possível, usa o caminho padrão "<usuário>/Desktop".
    """
    try:
        import winreg  # disponível apenas no Windows
        chave = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders",
        )
        caminho, _ = winreg.QueryValueEx(chave, "Desktop")
        caminho = os.path.expandvars(caminho)  # resolve variáveis como %USERPROFILE%
        if os.path.isdir(caminho):
            return caminho
    except Exception:
        pass  # não é Windows, chave não encontrada, valor inválido, etc.

    # Caminho padrão de fallback.
    return os.path.join(os.path.expanduser("~"), "Desktop")


def preencher_e_gerar_pdf(dados_colaborador: dict, itens: list) -> str:
    """
    Preenche o modelo_ficha_epi.xlsx com os dados do colaborador e da lista
    de itens, e gera o arquivo "Ficha de EPI.pdf" na área de trabalho do
    usuário.

    dados_colaborador: dicionário com as chaves "nome", "funcao", "admissao",
        "demissao" e "data" (a mesma data vale para todos os itens).
    itens: lista de dicionários com as chaves "equipamento", "quantidade" e
        "ca" (o "ca" já vem calculado a partir da base equipamentos.xlsx).

    Retorna o caminho completo do PDF gerado.
    """
    if not os.path.exists(ARQUIVO_MODELO_FICHA):
        raise FileNotFoundError(
            f"Modelo de ficha não encontrado em: {ARQUIVO_MODELO_FICHA}"
        )

    planilha = openpyxl.load_workbook(ARQUIVO_MODELO_FICHA)
    aba = planilha[NOME_ABA]

    # --- Cabeçalho do colaborador -------------------------------------------
    # Nome e Função saem em maiúsculas na ficha (Admissão/Demissão ficam
    # como foram digitadas, já que são datas, não texto livre).
    aba[CELULA_NOME] = dados_colaborador.get("nome", "").upper()
    aba[CELULA_FUNCAO] = dados_colaborador.get("funcao", "").upper()
    aba[CELULA_ADMISSAO] = dados_colaborador.get("admissao", "")
    aba[CELULA_DEMISSAO] = dados_colaborador.get("demissao", "")

    # A data informada no cabeçalho é a mesma para todos os itens da lista.
    data_convertida = _converter_data(dados_colaborador.get("data", ""))

    # --- Itens de EPI (um por linha, a partir da linha 10) -------------------
    for indice, item in enumerate(itens):
        linha = LINHA_INICIAL_ITENS + indice
        aba[f"{COLUNA_ITEM_EQUIPAMENTO}{linha}"] = item.get("equipamento", "")
        aba[f"{COLUNA_ITEM_QUANTIDADE}{linha}"] = _converter_quantidade(item.get("quantidade", ""))
        aba[f"{COLUNA_ITEM_DATA}{linha}"] = data_convertida
        aba[f"{COLUNA_ITEM_CA}{linha}"] = item.get("ca", "")

    # --- Salva uma cópia temporária preenchida (o modelo original não é alterado) ---
    pasta_temp = tempfile.mkdtemp(prefix="ficha_epi_")
    caminho_xlsx_preenchido = os.path.join(pasta_temp, "Ficha de EPI - preenchida.xlsx")
    planilha.save(caminho_xlsx_preenchido)

    try:
        # --- Converte a cópia preenchida em PDF, direto na área de trabalho ---
        pasta_destino = _obter_pasta_area_de_trabalho()
        os.makedirs(pasta_destino, exist_ok=True)

        # Nome do PDF: "Ficha de EPI - <nome do colaborador>.pdf". Isso
        # permite gerar a ficha de vários colaboradores seguidos sem precisar
        # renomear manualmente nem sobrescrever o PDF de outro colaborador.
        nome_colaborador = _sanitizar_nome_arquivo(dados_colaborador.get("nome", ""))
        if nome_colaborador:
            nome_arquivo_pdf = f"{PREFIXO_NOME_ARQUIVO_SAIDA} - {nome_colaborador}.pdf"
        else:
            nome_arquivo_pdf = f"{PREFIXO_NOME_ARQUIVO_SAIDA}.pdf"

        caminho_pdf = os.path.join(pasta_destino, nome_arquivo_pdf)
        # Se já existir uma ficha com esse nome (ex.: gerada antes para o
        # mesmo colaborador), usa um nome com sufixo numérico em vez de
        # sobrescrever o arquivo existente.
        caminho_pdf = _caminho_disponivel(caminho_pdf)

        _converter_xlsx_para_pdf(caminho_xlsx_preenchido, caminho_pdf)
    finally:
        # Remove a pasta temporária (o xlsx preenchido não precisa ser mantido).
        shutil.rmtree(pasta_temp, ignore_errors=True)

    return caminho_pdf


def _converter_xlsx_para_pdf(caminho_xlsx: str, caminho_pdf: str):
    """
    Converte um arquivo .xlsx em .pdf utilizando o Microsoft Excel instalado
    na máquina (automação via COM, biblioteca pywin32). É necessário que o
    Excel esteja instalado no Windows para essa conversão funcionar.
    """
    try:
        import win32com.client
    except ImportError as erro:
        raise RuntimeError(
            "A geração do PDF depende da biblioteca 'pywin32' e do Microsoft "
            "Excel instalado no Windows. Instale com: pip install pywin32"
        ) from erro

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        pasta_trabalho = excel.Workbooks.Open(caminho_xlsx)
        try:
            # 0 = xlTypePDF (formato de exportação fixo em PDF)
            pasta_trabalho.ExportAsFixedFormat(0, caminho_pdf)
        finally:
            pasta_trabalho.Close(SaveChanges=False)
    finally:
        excel.Quit()
