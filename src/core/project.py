# -*- coding: utf-8 -*-

import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from collections import defaultdict


class Portfolio:
    def __init__(self, disciplinas):
        self.disciplinas = disciplinas
        self.lotes_data = []
        self.funcionarios = []
        self.horas_trabalhaveis_mes = 160

    def definir_configuracoes_gerais(self, horas_por_funcionario, funcionarios):
        self.horas_trabalhaveis_mes = horas_por_funcionario
        self.funcionarios = funcionarios

    def definir_dados_lotes(self, lotes_data):
        self.lotes_data = lotes_data

    def gerar_relatorio_alocacao_decimal(self, nome_lote=None):
        if not self.lotes_data or not self.funcionarios:
            return pd.DataFrame()

        decimal_bruto_por_tarefa = defaultdict(lambda: defaultdict(float))
        for lote in self.lotes_data:
            for disc in self.disciplinas:
                for alocacao in lote.get("alocacao", {}).get(disc, []):
                    try:
                        cronograma_disc = lote.get("cronograma", {}).get(disc, {})
                        inicio_str = cronograma_disc.get("inicio")
                        fim_str = cronograma_disc.get("fim")
                        if not inicio_str or not fim_str:
                            continue

                        inicio_tarefa = datetime.strptime(inicio_str, "%d/%m/%Y")
                        fim_tarefa = datetime.strptime(fim_str, "%d/%m/%Y")

                        horas_trabalhadas = alocacao["horas_totais"]
                        pessoa_tuplo = alocacao["funcionario"]
                        chave_tarefa = (pessoa_tuplo, disc, lote["nome"])

                        num_meses = (
                            (fim_tarefa.year - inicio_tarefa.year) * 12
                            + (fim_tarefa.month - inicio_tarefa.month)
                            + 1
                        )
                        if num_meses <= 0 or self.horas_trabalhaveis_mes == 0:
                            continue

                        decimal_mensal = (
                            horas_trabalhadas / self.horas_trabalhaveis_mes
                        ) / num_meses

                        data_corrente = inicio_tarefa.replace(day=1)
                        while data_corrente <= fim_tarefa:
                            mes_ano_key = data_corrente.strftime("%Y-%m")
                            decimal_bruto_por_tarefa[chave_tarefa][
                                mes_ano_key
                            ] += decimal_mensal
                            data_corrente += relativedelta(months=1)
                    except (ValueError, ZeroDivisionError):
                        continue

        total_decimal_por_pessoa = defaultdict(lambda: defaultdict(float))
        for (pessoa, _, _), meses_data in decimal_bruto_por_tarefa.items():
            for mes, valor in meses_data.items():
                total_decimal_por_pessoa[pessoa][mes] += valor

        tarefas_a_processar = decimal_bruto_por_tarefa.items()
        if nome_lote:
            tarefas_a_processar = [
                (chave, meses)
                for (chave, meses) in decimal_bruto_por_tarefa.items()
                if chave[2] == nome_lote
            ]

        decimal_agrupado = defaultdict(lambda: defaultdict(float))
        for (pessoa, disc, _), meses_data in tarefas_a_processar:
            for mes, valor in meses_data.items():
                decimal_agrupado[(pessoa, disc)][mes] += valor

        todos_meses_keys = sorted(
            list(
                set(mes for meses in decimal_agrupado.values() for mes in meses.keys())
            )
        )

        relatorio_final = []
        for (pessoa, disc), meses_data in decimal_agrupado.items():
            linha = {"Funcionário": pessoa[0], "Cargo": pessoa[1], "Disciplina": disc}
            is_excedido = any(
                total_decimal_por_pessoa[pessoa].get(m, 0.0) > 1.0
                for m in todos_meses_keys
            )

            for mes in todos_meses_keys:
                linha[mes] = meses_data.get(mes, 0.0)

            # A coluna Status é adicionada internamente para a UI, mas não será incluída na exportação final
            linha["Status"] = "Alocação Excedida" if is_excedido else "OK"
            relatorio_final.append(linha)

        df = pd.DataFrame(relatorio_final)
        if df.empty:
            return pd.DataFrame()

        colunas_meses_numericas = [key for key in todos_meses_keys if key in df.columns]
        if colunas_meses_numericas:
            df["H.mês"] = df[colunas_meses_numericas].sum(axis=1)

        meses_display_map = {
            key: datetime.strptime(key, "%Y-%m").strftime("%b/%y")
            for key in todos_meses_keys
        }
        df.rename(columns=meses_display_map, inplace=True)

        colunas_para_formatar = list(meses_display_map.values())
        if "H.mês" in df.columns:
            colunas_para_formatar.append("H.mês")

        for col in colunas_para_formatar:
            if col in df.columns:
                df[col] = df[col].apply(
                    lambda x: f"{x:.2f}".replace(".", ",") if pd.notna(x) else "0,00"
                )

        # Define a nova ordem das colunas
        colunas_finais_ordenadas = ["Disciplina", "Funcionário", "Cargo"] + list(
            meses_display_map.values()
        )
        if "H.mês" in df.columns:
            colunas_finais_ordenadas.append("H.mês")

        # Adiciona a coluna Status apenas para a UI, ela não será exportada se não estiver na lista acima
        colunas_finais_com_status = colunas_finais_ordenadas + ["Status"]

        return df.reindex(columns=colunas_finais_com_status).fillna("")

    def gerar_relatorio_detalhado_por_tarefa(self):
        # ... (Nenhuma alteração neste método) ...
        detalhes_tarefas = {}
        if not self.lotes_data:
            return detalhes_tarefas
        for lote in self.lotes_data:
            for disc in self.disciplinas:
                contador_tarefa = 0
                for alocacao in lote.get("alocacao", {}).get(disc, []):
                    try:
                        cronograma_disc = lote.get("cronograma", {}).get(disc, {})
                        inicio_str = cronograma_disc.get("inicio")
                        fim_str = cronograma_disc.get("fim")
                        if not inicio_str or not fim_str:
                            continue
                        inicio_tarefa = datetime.strptime(inicio_str, "%d/%m/%Y")
                        fim_tarefa = datetime.strptime(fim_str, "%d/%m/%Y")
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
