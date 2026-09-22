"""
main.py
-------
Sistema Gerador de Ficha de EPI.

Interface gráfica (Tkinter) que permite:
    1. Preencher os dados do colaborador (cabeçalho), incluindo a data única
       da ficha.
    2. Selecionar um "Padrão Admissão" (modelo de cliente, dentro de
       modelos_x_cliente) para montar a lista de itens automaticamente, ou
       adicionar/editar/excluir itens manualmente, um a um.
    3. Gerar a ficha de EPI: preenche o modelo_ficha_epi.xlsx com os dados
       informados e gera o PDF final ("Ficha de EPI.pdf") na área de
       trabalho do usuário.
"""

import tkinter as tk
from tkinter import ttk, messagebox

from equipamentos import carregar_equipamentos, ARQUIVO_EQUIPAMENTOS
from gerar_ficha import preencher_e_gerar_pdf
from padroes_admissao import listar_padroes_admissao, carregar_itens_do_padrao, normalizar


class GeradorFichaEPI:
    """Classe principal que monta e controla toda a interface gráfica."""

    # Texto de exemplo mostrado no campo Data até o usuário digitar algo.
    PLACEHOLDER_DATA = "dd/mm/aaaa"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Gerador de Ficha de EPI")

        # ------------------------------------------------------------------
        # Bases de dados do sistema
        # ------------------------------------------------------------------
        # Dicionário {"Nome do Equipamento": "CA"} lido do arquivo equipamentos.xlsx.
        self.equipamentos = carregar_equipamentos()

        # Se a planilha não foi encontrada ou não retornou nenhum equipamento,
        # avisa o usuário de imediato (em vez de deixar o combobox vazio sem
        # explicação), indicando o caminho exato onde o sistema procurou.
        if not self.equipamentos:
            messagebox.showwarning(
                "Lista de equipamentos vazia",
                "Não foi possível carregar nenhum equipamento do arquivo:\n\n"
                f"{ARQUIVO_EQUIPAMENTOS}\n\n"
                "Verifique se o arquivo equipamentos.xlsx existe nessa pasta, "
                "se ele não está aberto em outro programa e se possui dados "
                "a partir da linha 2 (coluna A = Equipamento, coluna B = CA)."
            )

        # Índice do dicionário de equipamentos "normalizado" (sem diferenciar
        # maiúsculas/minúsculas nem acentos), usado para casar os nomes de
        # equipamentos lidos dos padrões de admissão com a base equipamentos.xlsx.
        # Cada entrada guarda (nome_original_em_equipamentos, ca).
        self.equipamentos_normalizado = {
            normalizar(nome): (nome, ca) for nome, ca in self.equipamentos.items()
        }

        # Lista onde ficam armazenadas as informações informadas pelo cliente
        # (cada item é um dicionário com os dados de um EPI adicionado).
        self.lista_itens = []

        # Quando o usuário clica em "Editar" em um item da lista, o índice
        # desse item fica guardado aqui. Enquanto houver um valor aqui, o
        # botão "Adicionar à lista" passa a funcionar como "Salvar alterações".
        self.indice_em_edicao = None

        self._montar_interface()

        # Ajusta o tamanho da janela ao conteúdo real da interface, DEPOIS de
        # todos os widgets terem sido criados (cabeçalho, padrão admissão,
        # corpo, lista e rodapé). Isso corrige um problema em que o botão
        # "Gerar Ficha de EPI" (no rodapé) ficava fora da área visível até o
        # usuário redimensionar a janela manualmente: um tamanho fixo definido
        # antes de montar a interface pode ser menor do que o espaço que os
        # widgets realmente precisam.
        self.root.update_idletasks()
        largura = max(self.root.winfo_reqwidth(), 760)
        altura = max(self.root.winfo_reqheight(), 600)
        self.root.geometry(f"{largura}x{altura}")
        self.root.minsize(largura, altura)

    # ==========================================================================
    # MONTAGEM DA INTERFACE
    # ==========================================================================
    def _montar_interface(self):
        """Cria todos os frames e widgets da tela."""
        self._montar_cabecalho()
        self._montar_padrao_admissao()
        self._montar_corpo()
        self._montar_lista()
        self._montar_rodape()

    # --------------------------------------------------------------------
    # CABEÇALHO: dados do colaborador
    # --------------------------------------------------------------------
    def _montar_cabecalho(self):
        frame = ttk.LabelFrame(self.root, text="Dados do Colaborador")
        frame.pack(fill="x", padx=10, pady=(10, 5))

        # "*" indica campo obrigatório (apenas o Nome do Colaborador é obrigatório).
        ttk.Label(frame, text="Nome do Colaborador *").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.entry_nome = ttk.Entry(frame, width=35)
        self.entry_nome.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(frame, text="Função").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        self.entry_funcao = ttk.Entry(frame, width=25)
        self.entry_funcao.grid(row=0, column=3, padx=5, pady=5, sticky="w")

        ttk.Label(frame, text="Admissão").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.entry_admissao = ttk.Entry(frame, width=15)
        self.entry_admissao.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(frame, text="Demissão").grid(row=1, column=2, sticky="w", padx=5, pady=5)
        self.entry_demissao = ttk.Entry(frame, width=15)
        self.entry_demissao.grid(row=1, column=3, padx=5, pady=5, sticky="w")

        # Data: fica no cabeçalho (e não no item de EPI) porque é única para
        # toda a ficha — todos os itens adicionados à lista usam essa mesma data.
        ttk.Label(frame, text="Data").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.entry_data = ttk.Entry(frame, width=15)
        self.entry_data.insert(0, self.PLACEHOLDER_DATA)
        self.entry_data.grid(row=2, column=1, padx=5, pady=5, sticky="w")

    # --------------------------------------------------------------------
    # PADRÃO ADMISSÃO: seleção de um modelo de cliente (modelos_x_cliente)
    # que preenche a lista de itens automaticamente
    # --------------------------------------------------------------------
    def _montar_padrao_admissao(self):
        frame = ttk.LabelFrame(self.root, text="Padrão Admissão")
        frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame, text="Padrão Admissão").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        # Os itens do combobox são os nomes dos arquivos .xlsx dentro de
        # modelos_x_cliente (sem a extensão) — um por cliente/site.
        self.combo_padrao_admissao = ttk.Combobox(
            frame, width=40, state="readonly",
            values=listar_padroes_admissao()
        )
        self.combo_padrao_admissao.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        # Ao escolher um padrão, a lista de itens é montada automaticamente
        # (não é necessário clicar em nenhum botão adicional).
        self.combo_padrao_admissao.bind("<<ComboboxSelected>>", self._ao_selecionar_padrao_admissao)

    # --------------------------------------------------------------------
    # CORPO: campos do item de EPI a ser adicionado à lista
    # --------------------------------------------------------------------
    def _montar_corpo(self):
        frame = ttk.LabelFrame(self.root, text="Item de EPI")
        frame.pack(fill="x", padx=10, pady=5)

        # Equipamento: combobox alimentado pela lista lida de equipamentos.xlsx.
        ttk.Label(frame, text="Equipamento").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.combo_equipamento = ttk.Combobox(
            frame, width=32, state="readonly",
            values=sorted(self.equipamentos.keys())
        )
        self.combo_equipamento.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(frame, text="Tamanho").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        self.entry_tamanho = ttk.Entry(frame, width=12)
        self.entry_tamanho.grid(row=0, column=3, padx=5, pady=5, sticky="w")

        ttk.Label(frame, text="Quantidade").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.entry_quantidade = ttk.Entry(frame, width=12)
        self.entry_quantidade.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # Botão principal do corpo: adiciona o item à lista. Também é
        # reaproveitado como "Salvar alterações" durante uma edição (ver
        # _iniciar_edicao / _confirmar_edicao_ou_adicionar).
        self.btn_adicionar = ttk.Button(
            frame, text="Adicionar à lista", command=self._confirmar_edicao_ou_adicionar
        )
        self.btn_adicionar.grid(row=2, column=0, columnspan=2, padx=5, pady=8, sticky="w")

        # Botão de cancelar edição: só aparece quando o usuário está editando um item.
        self.btn_cancelar_edicao = ttk.Button(
            frame, text="Cancelar edição", command=self._cancelar_edicao
        )
        # Não é exibido (pack/grid) até que uma edição seja iniciada.

    # --------------------------------------------------------------------
    # LISTA: itens já adicionados, cada um com botões de Editar e Excluir
    # --------------------------------------------------------------------
    def _montar_lista(self):
        container = ttk.LabelFrame(self.root, text="Itens Adicionados")
        container.pack(fill="both", expand=True, padx=10, pady=5)

        # Cabeçalho da lista (títulos das colunas).
        cabecalho = ttk.Frame(container)
        cabecalho.pack(fill="x", padx=5, pady=(5, 0))
        for texto, largura in (("Equipamento", 28), ("Tamanho", 10),
                                ("Quantidade", 10), ("CA", 10), ("Ações", 16)):
            ttk.Label(cabecalho, text=texto, width=largura, font=("TkDefaultFont", 9, "bold")).pack(side="left")

        # Área rolável (canvas + scrollbar) que vai conter uma linha por item.
        canvas = tk.Canvas(container, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.frame_linhas = ttk.Frame(canvas)

        self.frame_linhas.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.frame_linhas, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=(5, 0), pady=5)
        scrollbar.pack(side="right", fill="y", pady=5)

    # --------------------------------------------------------------------
    # RODAPÉ: botão para gerar a ficha de EPI
    # --------------------------------------------------------------------
    def _montar_rodape(self):
        frame = ttk.Frame(self.root)
        frame.pack(fill="x", padx=10, pady=10)

        # Empacotados da direita para a esquerda: "Gerar Ficha de EPI" fica
        # na extremidade direita, e "Resetar" logo ao lado dele.
        ttk.Button(
            frame, text="Gerar Ficha de EPI", command=self._gerar_ficha_epi
        ).pack(side="right")

        ttk.Button(
            frame, text="Resetar", command=self._resetar_tudo
        ).pack(side="right", padx=(0, 8))

    # ==========================================================================
    # AÇÕES DA LISTA (adicionar / editar / excluir)
    # ==========================================================================
    def _ler_campos_corpo(self):
        """Lê e retorna os valores atuais dos campos do corpo (item de EPI)."""
        equipamento = self.combo_equipamento.get().strip()
        return {
            "equipamento": equipamento,
            "tamanho": self.entry_tamanho.get().strip(),
            "quantidade": self.entry_quantidade.get().strip(),
            # O CA não é digitado pelo usuário: é obtido automaticamente a
            # partir do equipamento selecionado, consultando a base lida de
            # equipamentos.xlsx (mesmo CA cadastrado para esse equipamento).
            "ca": self.equipamentos.get(equipamento, ""),
        }

    def _limpar_campos_corpo(self):
        """Limpa os campos do corpo após adicionar/salvar um item."""
        self.combo_equipamento.set("")
        self.entry_tamanho.delete(0, tk.END)
        self.entry_quantidade.delete(0, tk.END)

    def _confirmar_edicao_ou_adicionar(self):
        """
        Comportamento do botão "Adicionar à lista":
        - Se NÃO houver edição em andamento: adiciona um novo item à lista.
        - Se HOUVER edição em andamento (usuário clicou em "Editar" em um item):
          salva as alterações no item correspondente, em vez de criar um novo.
        """
        dados = self._ler_campos_corpo()

        # O equipamento é a informação essencial do item; sem ele não faz
        # sentido adicionar a linha na lista.
        if not dados["equipamento"]:
            messagebox.showwarning("Campo obrigatório", "Selecione um equipamento antes de adicionar à lista.")
            return

        if self.indice_em_edicao is None:
            # Modo normal: adiciona um novo item.
            self.lista_itens.append(dados)
        else:
            # Modo edição: substitui os dados do item que estava sendo editado.
            self.lista_itens[self.indice_em_edicao] = dados
            self._cancelar_edicao()  # volta o botão/rótulos ao estado normal

        self._limpar_campos_corpo()
        self._atualizar_lista_na_tela()

    def _ao_selecionar_padrao_admissao(self, event=None):
        """
        Chamado automaticamente quando o usuário escolhe um cliente no
        combobox "Padrão Admissão": lê a planilha do cliente correspondente
        (dentro de modelos_x_cliente) e monta a lista de itens com os
        equipamentos encontrados lá, deixando Tamanho e Quantidade em
        branco — o usuário preenche esses dois campos depois, usando o
        botão "Editar" de cada linha.

        Se já existirem itens na lista, pede confirmação antes de
        substituí-los, já que essa ação troca a lista inteira.
        """
        nome_padrao = self.combo_padrao_admissao.get()
        if not nome_padrao:
            return

        if self.lista_itens:
            confirmar = messagebox.askyesno(
                "Substituir lista de itens",
                f'Isso vai substituir os {len(self.lista_itens)} item(ns) atuais da lista '
                f'pelos itens do padrão "{nome_padrao}". Deseja continuar?'
            )
            if not confirmar:
                return

        try:
            nomes_equipamentos = carregar_itens_do_padrao(nome_padrao)
        except Exception as erro:
            messagebox.showerror("Erro ao carregar padrão de admissão", str(erro))
            return

        novos_itens = []
        for nome_bruto in nomes_equipamentos:
            # Busca o CA na base de equipamentos.xlsx ignorando maiúsculas/
            # minúsculas e acentos. Quando encontra, usa o nome "oficial" (o
            # mesmo cadastrado em equipamentos.xlsx) para manter consistência
            # com o combobox de Equipamento; quando não encontra, mantém o
            # nome como veio da planilha do cliente e deixa o CA em branco.
            nome_oficial, ca = self.equipamentos_normalizado.get(
                normalizar(nome_bruto), (nome_bruto, "")
            )
            novos_itens.append({
                "equipamento": nome_oficial,
                "tamanho": "",
                "quantidade": "",
                "ca": ca,
            })

        self.lista_itens = novos_itens
        self._cancelar_edicao()  # garante que nenhuma edição antiga fique pendente
        self._atualizar_lista_na_tela()

    def _iniciar_edicao(self, indice: int):
        """Chamado ao clicar em 'Editar' em uma linha: carrega os dados do
        item nos campos do corpo para que o usuário possa alterá-los."""
        item = self.lista_itens[indice]

        self.combo_equipamento.set(item["equipamento"])
        self.entry_tamanho.delete(0, tk.END)
        self.entry_tamanho.insert(0, item["tamanho"])
        self.entry_quantidade.delete(0, tk.END)
        self.entry_quantidade.insert(0, item["quantidade"])
        # O CA não é um campo editável: ele é recalculado automaticamente a
        # partir do equipamento sempre que o item é adicionado ou salvo.

        self.indice_em_edicao = indice
        self.btn_adicionar.configure(text="Salvar alterações")
        # Exibe o botão de cancelar edição, ao lado do botão principal.
        self.btn_cancelar_edicao.grid(row=2, column=2, columnspan=2, padx=5, pady=8, sticky="w")

    def _cancelar_edicao(self):
        """Sai do modo de edição sem salvar alterações e limpa os campos."""
        self.indice_em_edicao = None
        self.btn_adicionar.configure(text="Adicionar à lista")
        self.btn_cancelar_edicao.grid_forget()
        self._limpar_campos_corpo()

    def _excluir_item(self, indice: int):
        """Chamado ao clicar em 'Excluir' em uma linha: remove o item da lista."""
        item = self.lista_itens[indice]
        confirmar = messagebox.askyesno(
            "Confirmar exclusão",
            f"Remover o item \"{item['equipamento']}\" da lista?"
        )
        if not confirmar:
            return

        # Se o item removido era o que estava em edição, cancela a edição.
        if self.indice_em_edicao == indice:
            self._cancelar_edicao()

        del self.lista_itens[indice]
        self._atualizar_lista_na_tela()

    def _atualizar_lista_na_tela(self):
        """Redesenha todas as linhas da lista de itens na tela, cada uma com
        seus próprios botões de Editar e Excluir."""
        # Remove todas as linhas atualmente desenhadas.
        for widget in self.frame_linhas.winfo_children():
            widget.destroy()

        for indice, item in enumerate(self.lista_itens):
            linha = ttk.Frame(self.frame_linhas)
            linha.pack(fill="x", pady=1)

            ttk.Label(linha, text=item["equipamento"], width=28).pack(side="left")
            ttk.Label(linha, text=item["tamanho"], width=10).pack(side="left")
            ttk.Label(linha, text=item["quantidade"], width=10).pack(side="left")
            ttk.Label(linha, text=item["ca"], width=10).pack(side="left")

            # "indice=indice" no lambda evita o problema de late-binding do
            # Python (sem isso, todos os botões usariam o último índice do loop).
            ttk.Button(
                linha, text="Editar", width=7,
                command=lambda indice=indice: self._iniciar_edicao(indice)
            ).pack(side="left", padx=2)
            ttk.Button(
                linha, text="Excluir", width=7,
                command=lambda indice=indice: self._excluir_item(indice)
            ).pack(side="left", padx=2)

    # ==========================================================================
    # RESETAR FORMULÁRIO
    # ==========================================================================
    def _resetar_tudo(self):
        """
        Botão "Resetar": limpa todos os dados preenchidos no sistema —
        cabeçalho do colaborador, padrão de admissão selecionado, campos do
        item de EPI (cancelando qualquer edição em andamento) e a lista de
        itens — deixando a tela como se tivesse acabado de ser aberta.
        """
        confirmar = messagebox.askyesno(
            "Resetar formulário",
            "Isso vai apagar todos os dados preenchidos (colaborador e lista de itens). Deseja continuar?"
        )
        if not confirmar:
            return

        # Cabeçalho do colaborador.
        self.entry_nome.delete(0, tk.END)
        self.entry_funcao.delete(0, tk.END)
        self.entry_admissao.delete(0, tk.END)
        self.entry_demissao.delete(0, tk.END)
        self.entry_data.delete(0, tk.END)
        self.entry_data.insert(0, self.PLACEHOLDER_DATA)

        # Padrão Admissão selecionado.
        self.combo_padrao_admissao.set("")

        # Campos do corpo (e sai do modo de edição, se houver um em andamento).
        self._cancelar_edicao()

        # Lista de itens.
        self.lista_itens = []
        self._atualizar_lista_na_tela()

    # ==========================================================================
    # GERAÇÃO DA FICHA
    # ==========================================================================
    def _gerar_ficha_epi(self):
        """
        Ao clicar em "Gerar Ficha de EPI": lê o modelo_ficha_epi.xlsx,
        preenche o cabeçalho do colaborador e todos os itens da lista
        (usando a mesma data do cabeçalho para todos), e gera o arquivo
        final "Ficha de EPI.pdf" na área de trabalho do usuário.
        """
        nome = self.entry_nome.get().strip()
        if not nome:
            messagebox.showwarning("Campo obrigatório", "Informe o Nome do Colaborador.")
            return

        if not self.lista_itens:
            messagebox.showwarning(
                "Lista vazia", "Adicione ao menos um item de EPI à lista antes de gerar a ficha."
            )
            return

        data_texto = self.entry_data.get().strip()
        if data_texto == self.PLACEHOLDER_DATA:
            data_texto = ""  # usuário não alterou o texto de exemplo

        dados_colaborador = {
            "nome": nome,
            "funcao": self.entry_funcao.get().strip(),
            "admissao": self.entry_admissao.get().strip(),
            "demissao": self.entry_demissao.get().strip(),
            "data": data_texto,
        }

        try:
            caminho_pdf = preencher_e_gerar_pdf(dados_colaborador, self.lista_itens)
        except FileNotFoundError as erro:
            messagebox.showerror("Modelo não encontrado", str(erro))
            return
        except Exception as erro:
            # Cobre, por exemplo, falta do Microsoft Excel/pywin32 na máquina,
            # necessários para converter a planilha preenchida em PDF.
            messagebox.showerror(
                "Erro ao gerar a ficha",
                "Não foi possível gerar o PDF da ficha de EPI.\n\n"
                f"Detalhes: {erro}"
            )
            return

        messagebox.showinfo(
            "Ficha gerada com sucesso",
            f"A ficha de EPI foi gerada em:\n\n{caminho_pdf}"
        )


def main():
    """Ponto de entrada do sistema."""
    root = tk.Tk()
    GeradorFichaEPI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
