# Gerador de Ficha de EPI

Sistema desktop, feito em Python, para preencher e gerar a **Ficha de Controle de
Fornecimento de EPI** dos colaboradores da GAP. O usuário preenche os dados do
colaborador e a lista de equipamentos entregues em uma interface gráfica simples
(Tkinter), e o sistema gera automaticamente um PDF pronto, a partir do modelo
oficial da ficha.

## O que o sistema faz

1. Você preenche os dados do colaborador (nome, função, admissão, demissão e a
   data da entrega dos EPIs).
2. Você monta a lista de equipamentos entregues — manualmente, item por item,
   ou automaticamente, escolhendo um "Padrão Admissão" (modelo já pronto de um
   cliente específico).
3. Ao clicar em **Gerar Ficha de EPI**, o sistema preenche o modelo oficial da
   ficha (`modelo_ficha_epi.xlsx`) com esses dados e gera um **PDF** com o nome
   do colaborador, salvo diretamente na Área de Trabalho.

O modelo original (`modelo_ficha_epi.xlsx`) nunca é alterado: o preenchimento
acontece sempre em uma cópia temporária, que é convertida em PDF e descartada.

## Estrutura de arquivos

| Arquivo / pasta              | Para que serve |
|-------------------------------|----------------|
| `main.py`                     | Ponto de entrada do sistema — é o arquivo que você executa. Contém toda a interface gráfica. |
| `equipamentos.py`             | Lê a planilha `equipamentos.xlsx` e monta a lista de equipamentos disponíveis (com seus CAs). |
| `equipamentos.xlsx`           | Base de dados de equipamentos: coluna A = nome do equipamento, coluna B = CA. |
| `padroes_admissao.py`         | Lê os arquivos da pasta `modelos_x_cliente/` para montar a lista de itens automaticamente quando um "Padrão Admissão" é selecionado. |
| `modelos_x_cliente/`          | Um arquivo `.xlsx` por cliente/site, com a relação padrão de uniformes e EPIs daquele cliente. |
| `gerar_ficha.py`              | Preenche o `modelo_ficha_epi.xlsx` com os dados informados e gera o PDF final. |
| `modelo_ficha_epi.xlsx`       | Modelo oficial da ficha de EPI (nunca é alterado pelo sistema). |
| `requirements.txt`            | Lista das bibliotecas Python necessárias. |

## Instalação

O sistema foi feito para rodar no Windows (a geração do PDF depende do
Microsoft Excel instalado na máquina).

1. Tenha o Python instalado (o Tkinter já vem incluso na instalação padrão do
   Python no Windows).
2. Instale as dependências, abrindo um terminal na pasta do projeto e
   rodando:

   ```
   pip install -r requirements.txt
   ```

   Isso instala:
   - `openpyxl` — para ler e escrever as planilhas Excel.
   - `pywin32` — para converter a planilha preenchida em PDF, usando o Excel
     instalado na máquina.

3. Tenha o **Microsoft Excel instalado** — ele é usado internamente (via
   automação COM) para gerar o PDF final. Sem o Excel instalado, o sistema
   funciona normalmente, mas o botão "Gerar Ficha de EPI" mostra uma mensagem
   de erro ao tentar converter para PDF.

## Como executar

Na pasta do projeto:

```
python main.py
```

A janela do sistema abre já ajustada ao tamanho necessário para mostrar toda
a interface, incluindo os botões do rodapé.

## Como usar

### 1. Dados do Colaborador (cabeçalho)

Preencha:

- **Nome do Colaborador** — único campo obrigatório.
- **Função**
- **Admissão**
- **Demissão**
- **Data** — data da entrega dos EPIs. É um campo único no cabeçalho (não por
  item), porque todos os itens da lista saem na ficha com essa mesma data.
  Formato esperado: `dd/mm/aaaa`.

### 2. Padrão Admissão (opcional, monta a lista automaticamente)

Se o colaborador está sendo admitido para um cliente que já tem um modelo
pronto (dentro de `modelos_x_cliente/`), selecione o nome do cliente no
combobox **Padrão Admissão**. A lista de itens é montada automaticamente, com
os equipamentos daquele padrão — os campos **Tamanho** e **Quantidade** ficam
em branco, pois são específicos de cada colaborador; preencha-os depois
usando o botão **Editar** de cada item (veja abaixo).

Se a lista já tiver itens quando você escolher um novo padrão, o sistema pede
confirmação antes de substituir tudo.

> A busca do nome do equipamento na base `equipamentos.xlsx` (para descobrir o
> CA) ignora acentuação e diferenças entre maiúsculas/minúsculas. Ainda assim,
> alguns itens dos modelos de cliente trazem descrições mais detalhadas (ex.:
> "Óculos Escuro (Steelpro - Aero)") que não têm correspondência exata na base
> de equipamentos — nesses casos o CA fica em branco e pode ser conferido/
> preenchido manualmente depois, se necessário.

### 3. Item de EPI (adicionar manualmente)

Para adicionar um item à mão:

1. Escolha o **Equipamento** no combobox (lista vinda de `equipamentos.xlsx`).
   O **CA** é preenchido automaticamente a partir do equipamento escolhido —
   não é um campo digitado.
2. Preencha **Tamanho** e **Quantidade**.
3. Clique em **Adicionar à lista**.

O item aparece na lista abaixo, com as colunas Equipamento, Tamanho,
Quantidade, CA e Ações.

### 4. Editar e excluir itens da lista

Cada linha da lista tem dois botões:

- **Editar** — carrega os dados daquele item de volta nos campos acima. O
  botão "Adicionar à lista" vira "Salvar alterações" enquanto você edita (há
  também um botão "Cancelar edição" para desistir sem salvar).
- **Excluir** — remove o item da lista, com uma confirmação antes.

### 5. Resetar

O botão **Resetar** (no rodapé, ao lado de "Gerar Ficha de EPI") limpa tudo:
os dados do cabeçalho, o padrão de admissão selecionado, os campos do item em
edição e a lista inteira — pede confirmação antes de apagar. Use quando for
começar a ficha de um novo colaborador do zero.

### 6. Gerar Ficha de EPI

Ao clicar em **Gerar Ficha de EPI**:

1. O sistema confere se o Nome do Colaborador foi preenchido e se há pelo
   menos 1 item na lista.
2. Preenche uma cópia do `modelo_ficha_epi.xlsx` com os dados:
   - Nome do Colaborador e Função são gravados **em maiúsculas**.
   - Cada item da lista vai em uma linha da tabela de itens, todos com a
     mesma Data informada no cabeçalho.
3. Converte essa cópia preenchida em PDF, usando o Excel instalado na
   máquina.
4. Salva o PDF na **Área de Trabalho**, com o nome:

   ```
   Ficha de EPI - <Nome do Colaborador>.pdf
   ```

   Caracteres não permitidos em nomes de arquivo do Windows (`\ / : * ? " < >
   |`) são removidos automaticamente do nome. Se já existir um PDF com esse
   nome (por exemplo, uma ficha gerada antes para o mesmo colaborador), o
   sistema não sobrescreve — salva com um sufixo numérico, como
   `Ficha de EPI - João da Silva (2).pdf`.

Ao final, uma mensagem mostra o caminho completo do PDF gerado.

## Mantendo as bases de dados atualizadas

- **Novo equipamento**: basta adicionar uma linha em `equipamentos.xlsx`
  (coluna A = nome do equipamento, coluna B = CA, ou em branco se não
  houver). O sistema lê essa planilha novamente toda vez que é aberto, então
  o novo equipamento já aparece na próxima execução — não é preciso mexer no
  código.
- **Novo padrão de cliente**: adicione um novo arquivo `.xlsx` dentro de
  `modelos_x_cliente/`, seguindo o mesmo padrão dos arquivos existentes (uma
  aba com o nome do cliente/site, cabeçalho "Relação de Uniformes/EPIs" na
  coluna E, e os equipamentos listados logo abaixo). O nome do arquivo (sem a
  extensão `.xlsx`) é o nome que aparece no combobox "Padrão Admissão".

## Solução de problemas

- **"Lista de equipamentos vazia"** ao abrir o sistema: verifique se
  `equipamentos.xlsx` existe na pasta do projeto, se não está aberto em outro
  programa, e se tem dados a partir da linha 2.
- **`ModuleNotFoundError: No module named 'openpyxl'`** (ou `win32com`):
  rode `pip install -r requirements.txt` na pasta do projeto.
- **Erro ao gerar a ficha, mencionando Excel/pywin32**: confirme que o
  Microsoft Excel está instalado na máquina — ele é necessário para converter
  a planilha preenchida em PDF.
- **"Não foi possível localizar o cabeçalho..."** ao escolher um Padrão
  Admissão: o arquivo daquele cliente em `modelos_x_cliente/` está fora do
  padrão esperado (sem o texto "Relação de Uniformes/EPIs" na coluna E, nas
  primeiras linhas da planilha).
