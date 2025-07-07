# -*- coding: utf-8 -*-

import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta


class Portfolio:
    """
    Gerencia um portfólio de múltiplos projetos (Lotes), calcula a alocação
    de horas por pessoa e valida a capacidade dos recursos entre os projetos.
    """

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

    def _get_date_range(self):
        """Helper para obter o intervalo de datas de todo o portfólio."""
        try:
            datas_inicio = [
                datetime.strptime(lote["cronograma"][disc]["inicio"], "%d/%m/%Y")
                for lote in self.lotes_data
                for disc in self.disciplinas
                if lote.get("cronograma", {}).get(disc, {}).get("inicio")
            ]
            datas_fim = [
                datetime.strptime(lote["cronograma"][disc]["fim"], "%d/%m/%Y")
                for lote in self.lotes_data
                for disc in self.disciplinas
                if lote.get("cronograma", {}).get(disc, {}).get("fim")
            ]
            if not datas_inicio or not datas_fim:
                return None, None
            return min(datas_inicio), max(datas_fim)
        except ValueError:
            raise ValueError("Formato de data inválido. Use DD/MM/AAAA.")

    def gerar_relatorio_alocacao_decimal(self):
        """
        Gera um relatório com a alocação decimal de cada pessoa, por mês,
        conforme a nova fórmula de cálculo.
        """
        if not self.lotes_data or not self.funcionarios:
            return pd.DataFrame()

        try:
            data_inicio_portfolio, data_final_portfolio = self._get_date_range()
            if not data_inicio_portfolio:
                return pd.DataFrame()
        except ValueError as e:
            return pd.DataFrame([{"Erro": str(e)}])

        decimal_consolidado = {pessoa: {} for pessoa in self.funcionarios}

        for lote in self.lotes_data:
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
                            if mes_ano_key not in decimal_consolidado[pessoa_tuplo]:
                                decimal_consolidado[pessoa_tuplo][mes_ano_key] = 0.0
                            decimal_consolidado[pessoa_tuplo][
                                mes_ano_key
                            ] += decimal_mensal
                            data_corrente += relativedelta(months=1)
                    except (ValueError, KeyError, ZeroDivisionError):
                        continue

        todos_meses_keys = sorted(
            list(
                set(
                    mes
                    for pessoa_meses in decimal_consolidado.values()
                    for mes in pessoa_meses.keys()
                )
            )
        )
        meses_display = [
            datetime.strptime(m, "%Y-%m").strftime("%b/%y") for m in todos_meses_keys
        ]

        relatorio_final = []
        for pessoa_tuplo, meses_data in decimal_consolidado.items():
            if not any(meses_data.values()):
                continue

            linha_relatorio = {"Funcionário": pessoa_tuplo[0], "Cargo": pessoa_tuplo[1]}
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

        colunas = ["Funcionário", "Cargo"] + meses_display + ["Total Decimal", "Status"]
        df_relatorio = pd.DataFrame(relatorio_final)

        return df_relatorio.reindex(columns=colunas, fill_value="0,00")

    def gerar_relatorio_detalhado_por_tarefa(self):
        """
        Calcula a distribuição de horas mensais para cada tarefa individual.
        Retorna um dicionário para fácil consulta pela UI.
        Ex: {('1', 'Geometria', ('Nome', 'Cargo')): {'Jan/25': 40.0, 'Fev/25': 40.0}}
        """
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
