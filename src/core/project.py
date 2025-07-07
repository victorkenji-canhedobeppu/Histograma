# -*- coding: utf-8 -*-

import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta


class Portfolio:
    """
    Gerencia um portfólio de múltiplos projetos (Lotes), calcula a alocação
    de horas por pessoa e valida a capacidade dos recursos entre os projetos.
    """

    # ... (o resto da classe __init__, definir_configuracoes_gerais, etc. permanece igual) ...
    def __init__(self, disciplinas):
        self.disciplinas = disciplinas
        self.lotes_data = []
        self.funcionarios = []  # Armazenará tuplos (nome, cargo)
        self.horas_trabalhaveis_mes = 160

    def definir_configuracoes_gerais(self, horas_por_funcionario, funcionarios):
        """Define as configurações base, incluindo a lista de funcionários."""
        self.horas_trabalhaveis_mes = horas_por_funcionario
        self.funcionarios = funcionarios

    def definir_dados_lotes(self, lotes_data):
        """
        Armazena os dados de todos os lotes. A alocação agora é uma lista de
        dicionários com {'funcionario': (nome, cargo), 'horas_totais': valor}.
        """
        self.lotes_data = lotes_data

    def _get_date_range(self, lotes_filtrados=None):
        """Helper para obter o intervalo de datas do portfólio ou de lotes específicos."""
        lotes_a_verificar = (
            lotes_filtrados if lotes_filtrados is not None else self.lotes_data
        )
        try:
            datas_inicio = [
                datetime.strptime(lote["cronograma"][disc]["inicio"], "%d/%m/%Y")
                for lote in lotes_a_verificar
                for disc in self.disciplinas
                if lote.get("cronograma", {}).get(disc, {}).get("inicio")
            ]
            datas_fim = [
                datetime.strptime(lote["cronograma"][disc]["fim"], "%d/%m/%Y")
                for lote in lotes_a_verificar
                for disc in self.disciplinas
                if lote.get("cronograma", {}).get(disc, {}).get("fim")
            ]
            if not datas_inicio or not datas_fim:
                return None, None
            return min(datas_inicio), max(datas_fim)
        except ValueError:
            raise ValueError("Formato de data inválido. Use DD/MM/AAAA.")

    def gerar_relatorio_alocacao_decimal(self, nome_lote=None):
        """
        Gera um relatório com a alocação decimal de cada pessoa, por mês,
        detalhado por disciplina.
        """
        if not self.lotes_data or not self.funcionarios:
            return pd.DataFrame()

        lotes_a_processar = self.lotes_data
        if nome_lote:
            lotes_a_processar = [
                lote for lote in self.lotes_data if lote["nome"] == nome_lote
            ]
            if not lotes_a_processar:
                return pd.DataFrame()

        try:
            _, _ = self._get_date_range(lotes_a_processar)
        except ValueError as e:
            return pd.DataFrame([{"Erro": str(e)}])

        # A chave agora é (pessoa_tuplo, disciplina)
        decimal_por_disciplina = {}

        for lote in lotes_a_processar:
            for disc in self.disciplinas:
                for alocacao in lote.get("alocacao", {}).get(disc, []):
                    try:
                        inicio_tarefa = datetime.strptime(
                            lote["cronograma"][disc]["inicio"], "%d/%m/%Y"
                        )
                        fim_tarefa = datetime.strptime(
                            lote["cronograma"][disc]["fim"], "%d/%m/%Y"
                        )
                        horas_trabalhadas_tarefa = alocacao["horas_totais"]
                        pessoa_tuplo = alocacao["funcionario"]

                        # Define a chave de agregação
                        chave_alocacao = (pessoa_tuplo, disc)

                        num_meses = (
                            (fim_tarefa.year - inicio_tarefa.year) * 12
                            + (fim_tarefa.month - inicio_tarefa.month)
                            + 1
                        )
                        if num_meses <= 0:
                            continue
                        if self.horas_trabalhaveis_mes == 0:
                            continue

                        decimal_mensal = (
                            horas_trabalhadas_tarefa / self.horas_trabalhaveis_mes
                        ) / num_meses

                        data_corrente = inicio_tarefa.replace(day=1)
                        while data_corrente <= fim_tarefa:
                            mes_ano_key = data_corrente.strftime("%Y-%m")

                            # Inicializa o dicionário para a chave se não existir
                            if chave_alocacao not in decimal_por_disciplina:
                                decimal_por_disciplina[chave_alocacao] = {}
                            if (
                                mes_ano_key
                                not in decimal_por_disciplina[chave_alocacao]
                            ):
                                decimal_por_disciplina[chave_alocacao][
                                    mes_ano_key
                                ] = 0.0

                            decimal_por_disciplina[chave_alocacao][
                                mes_ano_key
                            ] += decimal_mensal
                            data_corrente += relativedelta(months=1)
                    except (ValueError, KeyError, ZeroDivisionError):
                        continue

        todos_meses_keys = sorted(
            list(
                set(
                    mes
                    for pessoa_meses in decimal_por_disciplina.values()
                    for mes in pessoa_meses.keys()
                )
            )
        )
        meses_display = [
            datetime.strptime(m, "%Y-%m").strftime("%b/%y") for m in todos_meses_keys
        ]

        relatorio_final = []
        for chave_alocacao, meses_data in decimal_por_disciplina.items():
            if not any(meses_data.values()):
                continue

            pessoa_tuplo, disciplina = chave_alocacao
            nome_funcionario, cargo_funcionario = pessoa_tuplo

            linha_relatorio = {
                "Funcionário": nome_funcionario,
                "Cargo": cargo_funcionario,
                "Disciplina": disciplina,
            }
            status = "OK"
            total_decimal = 0

            for i, mes_key in enumerate(todos_meses_keys):
                decimal_no_mes = meses_data.get(mes_key, 0.0)
                linha_relatorio[meses_display[i]] = f"{decimal_no_mes:.2f}".replace(
                    ".", ","
                )
                total_decimal += decimal_no_mes
                if decimal_no_mes > 1.0:
                    status = "Alocação Excedida"

            linha_relatorio["Total Decimal"] = f"{total_decimal:.2f}".replace(".", ",")
            linha_relatorio["Status"] = status
            relatorio_final.append(linha_relatorio)

        # Adiciona a nova coluna 'Disciplina'
        colunas = (
            ["Funcionário", "Cargo", "Disciplina"]
            + meses_display
            + ["Total Decimal", "Status"]
        )
        df_relatorio = pd.DataFrame(relatorio_final)

        # Se não houver dados, retorna um DF vazio com as colunas certas
        if df_relatorio.empty:
            return pd.DataFrame(columns=colunas)

        return df_relatorio.reindex(columns=colunas, fill_value="0,00")

    def gerar_relatorio_detalhado_por_tarefa(self):
        """
        Calcula a distribuição de horas mensais para cada tarefa individual.
        Retorna um dicionário para fácil consulta pela UI.
        Ex: {('1', 'Geometria', ('Nome', 'Cargo')): {'Jan/25': 40.0, 'Fev/25': 40.0}}
        """
        # (Nenhuma alteração neste método)
        detalhes_tarefas = {}
        if not self.lotes_data:
            return detalhes_tarefas

        for lote in self.lotes_data:
            for disc in self.disciplinas:
                # Usamos um contador para diferenciar tarefas da mesma pessoa na mesma disciplina/lote
                contador_tarefa = 0
                for alocacao in lote.get("alocacao", {}).get(disc, []):
                    try:
                        inicio_tarefa = datetime.strptime(
                            lote["cronograma"][disc]["inicio"], "%d/%m/%Y"
                        )
                        fim_tarefa = datetime.strptime(
                            lote["cronograma"][disc]["fim"], "%d/%m/%Y"
                        )
                        horas_totais_tarefa = alocacao["horas_totais"]
                        pessoa_tuplo = alocacao["funcionario"]

                        tarefa_key = (lote["nome"], disc, pessoa_tuplo, contador_tarefa)
                        contador_tarefa += 1

                        num_meses = (
                            (fim_tarefa.year - inicio_tarefa.year) * 12
                            + (fim_tarefa.month - inicio_tarefa.month)
                            + 1
                        )
                        if num_meses <= 0:
                            continue

                        horas_por_mes_tarefa = horas_totais_tarefa / num_meses

                        detalhes_tarefas[tarefa_key] = {}
                        data_corrente = inicio_tarefa.replace(day=1)
                        while data_corrente <= fim_tarefa:
                            mes_display = data_corrente.strftime("%b/%y")
                            detalhes_tarefas[tarefa_key][mes_display] = round(
                                horas_por_mes_tarefa, 2
                            )
                            data_corrente += relativedelta(months=1)
                    except (ValueError, KeyError, ZeroDivisionError):
                        continue
        return detalhes_tarefas
