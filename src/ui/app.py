# -*- coding: utf-8 -*-

from config.settings import CARGOS, DISCIPLINAS, SUBCONTRATOS
from core.project import Portfolio
from ui.gui_app import GuiApp


class App:
    def __init__(self):
        self.funcionarios = []
        self.portfolio = Portfolio(list(DISCIPLINAS.values()))
        self.ui = GuiApp(app_controller=self)
        self.ultimo_df_consolidado = None
        self.ultimo_dashboards_lotes = None

    def get_resumo_equipe_df(self):
        """Garante que os dados mais recentes da UI estão no portfólio e retorna o resumo."""
        dados_ui = self.ui._coletar_dados_da_ui()
        if not dados_ui:
            return None
        self.portfolio.definir_dados_lotes(dados_ui["lotes"])
        return self.portfolio.gerar_relatorio_horas_por_funcionario()

    def run(self):
        self.ui.carregar_ou_iniciar_ui()
        self.ui.mainloop()

    def get_disciplinas(self):
        return sorted(list(DISCIPLINAS.values()))

    def get_cargos_disponiveis(self):
        return sorted(list(CARGOS.values()), reverse=True)

    def get_subcontratos_disponiveis(self):
        return sorted(list(SUBCONTRATOS.values()))

    def get_ultimo_df_consolidado(self):
        return self.ultimo_df_consolidado

    def get_ultimo_dashboards_lotes(self):
        return self.ultimo_dashboards_lotes

    def get_funcionarios_para_display(self):
        """Retorna a string formatada para a lista de gerenciamento de equipe."""
        return sorted(
            [f"{nome} [{disciplina}]" for nome, disciplina in self.funcionarios]
        )

    def get_nomes_funcionarios_por_disciplina(self, disciplina_filtro):
        """Retorna uma lista de NOMES de funcionários para uma dada disciplina."""
        return sorted(
            [
                nome
                for nome, disciplina in self.funcionarios
                if disciplina == disciplina_filtro
            ]
        )

    def adicionar_funcionario(self, nome, disciplina):
        nome_tratado = nome.strip()
        if not nome_tratado:
            return

        novo_funcionario = (nome_tratado, disciplina)

        # Verifica se um funcionário com o mesmo nome já existe
        if not any(f[0].lower() == nome_tratado.lower() for f in self.funcionarios):
            self.funcionarios.append(novo_funcionario)
            self.funcionarios.sort(key=lambda x: x[0])
            self.ui.atualizar_lista_funcionarios()
        else:
            print(f"Aviso: Funcionário com nome '{nome_tratado}' já existe.")

    def remover_funcionario(self, display_string):
        """Remove um funcionário baseado na sua string de exibição 'Nome [Disciplina]'."""
        try:
            # Faz o parsing da string "Nome [Disciplina]"
            nome, disciplina = display_string.rsplit(" [", 1)
            disciplina = disciplina.rstrip("]")

            funcionario_a_remover = (nome, disciplina)
            if funcionario_a_remover in self.funcionarios:
                self.funcionarios.remove(funcionario_a_remover)
                self.ui.atualizar_lista_funcionarios()

        except ValueError:
            print(
                f"Erro ao tentar remover: formato de string inválido '{display_string}'"
            )

    def processar_portfolio(self, lotes_data, horas_mes):
        """Processa todos os dados e atualiza a UI com os resultados."""
        self.portfolio.definir_configuracoes_gerais(
            horas_mes, self.funcionarios, self.get_cargos_disponiveis()
        )
        self.portfolio.definir_dados_lotes(lotes_data)

        self.ultimo_df_consolidado = self.portfolio.gerar_relatorio_alocacao_decimal()
        detalhes_tarefas = self.portfolio.gerar_relatorio_detalhado_por_tarefa()

        self.ultimo_dashboards_lotes = {}
        for lote in lotes_data:
            nome_lote = lote["nome"]
            df_lote = self.portfolio.gerar_relatorio_alocacao_decimal(
                nome_lote=lote["nome"]
            )
            self.ultimo_dashboards_lotes[nome_lote] = df_lote

        porcentagem_global, detalhes_lotes = (
            self.portfolio.verificar_porcentagem_horas_cargo(cargo_alvo=CARGOS["ESTAG"])
        )

        alerta_composicao = {"global": porcentagem_global, "lotes": detalhes_lotes}

        self.ui.atualizar_dashboards(
            self.ultimo_df_consolidado,
            self.ultimo_dashboards_lotes,
            detalhes_tarefas,
            alerta_composicao=alerta_composicao,
        )
