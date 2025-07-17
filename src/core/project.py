# In core/project.py

import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from collections import defaultdict


class Portfolio:
    def __init__(self, disciplinas):
        self.disciplinas = disciplinas
        self.lotes_data = []
        self.funcionarios = []
        self.horas_trabalhaveis_mes = 160  # Default value

    def definir_configuracoes_gerais(self, horas_por_funcionario, funcionarios, cargos):
        self.horas_trabalhaveis_mes = horas_por_funcionario
        self.funcionarios = funcionarios  # List of (name, discipline) tuples
        self.cargos_disponiveis = cargos  # List of cargo strings

    def definir_dados_lotes(self, lotes_data):
        self.lotes_data = lotes_data

    def gerar_relatorio_horas_por_funcionario(
        self, nome_lote=None
    ):  # MODIFIED: Added nome_lote parameter
        if not self.lotes_data:
            print(
                "DEBUG: gerar_relatorio_horas_por_funcionario - No lotes_data, returning empty DataFrame."
            )
            return pd.DataFrame()

        horas_agregadas = defaultdict(float)

        lotes_a_processar = self.lotes_data
        if nome_lote:  # Filter if a specific lot name is provided
            lotes_a_processar = [
                lote for lote in self.lotes_data if lote.get("nome") == nome_lote
            ]
            if not lotes_a_processar:
                print(
                    f"DEBUG: Specific lote '{nome_lote}' not found in data for summary report."
                )
                return pd.DataFrame()

        for lote in lotes_a_processar:
            # Process employee allocations (disciplines)
            for disciplina_nome, dados_disc in lote.get("disciplinas", {}).items():
                for alocacao in dados_disc.get("alocacoes", []):
                    try:
                        nome_func, cargo_func = alocacao["funcionario"]
                        horas = alocacao.get("horas_totais", 0)

                        if nome_func and cargo_func and horas > 0:
                            chave = (nome_func, cargo_func, disciplina_nome)
                            horas_agregadas[chave] += horas
                    except (KeyError, ValueError) as e:
                        print(
                            f"DEBUG: Error processing allocation in gerar_relatorio_horas_por_funcionario: {e} - Data: {alocacao}"
                        )
                        continue

            # --- INCLUDE SUBCONTRACTS IN THIS REPORT TOO ---
            for subcontrato_data in lote.get("subcontratos", []):
                try:
                    nome_sub = subcontrato_data.get(
                        "nome_subcontrato", "Subcontrato Desconhecido"
                    )
                    horas = subcontrato_data.get("horas_totais", 0)

                    if nome_sub and horas > 0:
                        # Use a placeholder for func/cargo for subcontracts
                        chave = (
                            nome_sub,
                            "Subcontrato",
                            nome_sub,  # Use sub name as discipline for reporting
                        )
                        horas_agregadas[chave] += horas
                except (KeyError, ValueError) as e:
                    print(
                        f"DEBUG: Error processing subcontract in gerar_relatorio_horas_por_funcionario: {e} - Data: {subcontrato_data}"
                    )
                    continue
            # --- END SUBCONTRACTS INCLUSION ---

        if not horas_agregadas:
            print(
                "DEBUG: gerar_relatorio_horas_por_funcionario - No aggregated hours, returning empty DataFrame."
            )
            return pd.DataFrame()

        lista_relatorio = [
            {
                "Funcionário": nome,
                "Cargo": cargo,
                "Disciplina": disciplina,
                "Horas Totais Alocadas": horas,
            }
            for (nome, cargo, disciplina), horas in horas_agregadas.items()
        ]
        return pd.DataFrame(lista_relatorio).sort_values(
            by=["Funcionário", "Disciplina"]
        )

    def _processar_alocacao(self, inicio_str, fim_str, horas_totais_alocadas):
        if not inicio_str or not fim_str:
            return None

        try:
            horas_totais_alocadas = float(horas_totais_alocadas)
        except ValueError:
            return None

        if horas_totais_alocadas <= 0:
            return None

        try:
            inicio = datetime.strptime(inicio_str, "%d/%m/%Y")
            fim = datetime.strptime(fim_str, "%d/%m/%Y")
        except ValueError as e:
            print(
                f"DEBUG: _processar_alocacao - Date parsing error: {e} (Inicio: '{inicio_str}', Fim: '{fim_str}')."
            )
            return None

        if fim < inicio:
            return None

        num_meses = (fim.year - inicio.year) * 12 + (fim.month - inicio.month) + 1

        if num_meses <= 0 or self.horas_trabalhaveis_mes <= 0:
            return None

        decimal_mensal = round(
            horas_totais_alocadas / (self.horas_trabalhaveis_mes * num_meses), 2
        )

        decimais = {}
        data_corrente = inicio.replace(day=1)
        while data_corrente <= fim:
            mes_key = data_corrente.strftime("%Y-%m")
            decimais[mes_key] = decimal_mensal
            data_corrente += relativedelta(months=1)

        return decimais

    def gerar_relatorio_alocacao_decimal(self, nome_lote=None):
        if not self.lotes_data:
            print("DEBUG: No lotes_data available. Returning empty DataFrame.")
            return pd.DataFrame()

        decimal_bruto_por_tarefa = defaultdict(lambda: defaultdict(float))

        lotes_a_processar = self.lotes_data
        if nome_lote:
            lotes_a_processar = [
                lote for lote in self.lotes_data if lote["nome"] == nome_lote
            ]
            if not lotes_a_processar:
                print(f"DEBUG: Specific lote '{nome_lote}' not found in data.")
                return pd.DataFrame()

        for lote in lotes_a_processar:
            # --- PROCESS EMPLOYEE ALLOCATIONS (DISCIPLINES) ---
            for disc, dados_disc in lote.get("disciplinas", {}).items():
                cronograma = dados_disc.get("cronograma", {})

                if not cronograma.get("inicio") or not cronograma.get("fim"):
                    continue

                for alocacao in dados_disc.get("alocacoes", []):
                    try:
                        nome_func = alocacao.get("funcionario", ["", ""])[0]
                        cargo_func = alocacao.get("funcionario", ["", ""])[1]
                        horas_totais = alocacao.get("horas_totais", 0.0)

                        if not nome_func.strip():
                            continue

                        employee_registered = any(
                            f[0] == nome_func for f in self.funcionarios
                        )
                        if not employee_registered:
                            continue

                        if not cargo_func.strip():
                            continue

                        if horas_totais <= 0:
                            continue

                        decimais = self._processar_alocacao(
                            cronograma.get("inicio"),
                            cronograma.get("fim"),
                            horas_totais,
                        )

                        if decimais:
                            chave_agregacao = (disc, cargo_func, nome_func)
                            for mes, valor in decimais.items():
                                decimal_bruto_por_tarefa[chave_agregacao][mes] += valor

                    except (KeyError, ValueError, IndexError, ZeroDivisionError) as e:
                        print(
                            f"ERROR: Exception processing employee allocation in {lote.get('nome', 'Unknown Lote')}/{disc}: {e} for data {alocacao.get('funcionario')}, {alocacao.get('horas_totais')}"
                        )
                        continue

            # --- PROCESS SUBCONTRACT ALLOCATIONS ---
            # Assuming 'subcontratos' is a list of dicts at the lot level
            for subcontrato_data in lote.get("subcontratos", []):
                try:
                    sub_nome = subcontrato_data.get(
                        "nome_subcontrato", "Subcontrato Desconhecido"
                    )
                    sub_cronograma = subcontrato_data.get("cronograma", {})
                    sub_horas_totais = subcontrato_data.get("horas_totais", 0.0)

                    if not sub_nome.strip() or not sub_horas_totais > 0:
                        continue

                    if not sub_cronograma.get("inicio") or not sub_cronograma.get(
                        "fim"
                    ):
                        continue

                    decimais = self._processar_alocacao(
                        sub_cronograma.get("inicio"),
                        sub_cronograma.get("fim"),
                        sub_horas_totais,
                    )

                    if decimais:
                        # For subcontracts, use "Subcontrato" as a placeholder for cargo
                        # and the sub_nome itself as the "Funcionário" for grouping
                        # and sub_nome as the "Discipline" if it helps in reporting
                        chave_agregacao = (
                            sub_nome,
                            "Subcontrato",
                            sub_nome,
                        )  # Discipline, Cargo, Func
                        for mes, valor in decimais.items():
                            decimal_bruto_por_tarefa[chave_agregacao][mes] += valor

                except (KeyError, ValueError, IndexError, ZeroDivisionError) as e:
                    print(
                        f"ERROR: Exception processing subcontract in {lote.get('nome', 'Unknown Lote')}: {e} for data {subcontrato_data}"
                    )
                    continue
            # --- END SUBCONTRACTS INCLUSION ---

        if not decimal_bruto_por_tarefa:
            return pd.DataFrame()

        total_decimal_por_pessoa = defaultdict(lambda: defaultdict(float))
        for (disc, cargo, nome), meses in decimal_bruto_por_tarefa.items():
            for mes, valor in meses.items():
                total_decimal_por_pessoa[nome][mes] += valor

        todos_meses_keys = sorted(
            list(
                set(
                    mes
                    for meses in decimal_bruto_por_tarefa.values()
                    for mes in meses.keys()
                )
            )
        )
        if not todos_meses_keys:
            print("DEBUG: No months generated, returning empty DataFrame.")
            return pd.DataFrame()

        relatorio_final = []
        for (disc, cargo_ou_sub, nome), meses_data in decimal_bruto_por_tarefa.items():
            linha = {"Disciplina": disc, "Cargo": cargo_ou_sub, "Funcionário": nome}

            # Only check for over-allocation for actual registered employees
            # Subcontratos are generally not "over-allocated" in the same way as internal staff
            if any(
                f[0] == nome for f in self.funcionarios
            ):  # Check if 'nome' is an actual employee
                is_excedido = any(
                    total_decimal_por_pessoa.get(nome, {}).get(m, 0.0) > 1.0001
                    for m in todos_meses_keys
                )
            else:  # If it's a subcontract, it's not "excedido" in terms of internal staff capacity
                is_excedido = False

            for mes in todos_meses_keys:
                linha[mes] = meses_data.get(mes, 0.0)

            linha["Status"] = "Alocação Excedida" if is_excedido else "OK"
            relatorio_final.append(linha)

        df = pd.DataFrame(relatorio_final)
        if df.empty:
            print(
                "DEBUG: Final DataFrame is empty after processing. Returning empty DataFrame."
            )
            return pd.DataFrame()

        colunas_meses_numericas = [key for key in todos_meses_keys if key in df.columns]
        if colunas_meses_numericas:
            df["H.total"] = df[colunas_meses_numericas].sum(axis=1)

        meses_display_map = {
            key: datetime.strptime(key, "%Y-%m").strftime("%b/%y")
            for key in todos_meses_keys
        }
        df.rename(columns=meses_display_map, inplace=True)

        colunas_finais = ["Disciplina", "Cargo", "Funcionário"] + sorted(
            list(meses_display_map.values()),
            key=lambda x: datetime.strptime(x, "%b/%y"),
        )
        if "H.total" in df.columns:
            colunas_finais.append("H.total")
        colunas_finais.append("Status")

        final_df = df.reindex(columns=colunas_finais).fillna("")

        return final_df

    def gerar_relatorio_detalhado_por_tarefa(self):
        return {}  # Still returns empty, as per previous code

    def verificar_porcentagem_horas_cargo(self, cargo_alvo):

        if not self.lotes_data:
            return 0.0, {}

        horas_cargo_alvo_global = 0.0
        horas_totais_projeto_global = 0.0
        detalhes_por_lote = {}

        for lote in self.lotes_data:
            nome_lote = lote.get("nome", "Desconhecido")
            horas_cargo_alvo_lote = 0.0
            horas_totais_lote = 0.0

            # Iterate over disciplines (employee allocations)
            for disc, dados_disc in lote.get("disciplinas", {}).items():
                for alocacao in dados_disc.get("alocacoes", []):
                    try:
                        horas = alocacao.get("horas_totais", 0)
                        # Ensure 'funcionario' is present and has at least two elements (name, cargo)
                        cargo_func = alocacao.get("funcionario", ["", ""])[1]

                        if horas > 0:  # Only count positive hours
                            horas_totais_lote += horas
                            if cargo_func == cargo_alvo:
                                horas_cargo_alvo_lote += horas
                    except (KeyError, ValueError, IndexError) as e:
                        print(
                            f"DEBUG: Error in verificar_porcentagem_horas_cargo processing employee allocation: {e} - Data: {alocacao}"
                        )
                        continue

            # --- PROCESS SUBCONTRACTS FOR TOTAL HOURS (BUT NOT FOR CARGO_ALVO PERCENTAGE) ---
            for sub_aloc in lote.get("subcontratos", []):
                try:
                    horas = sub_aloc.get("horas_totais", 0)
                    if horas > 0:  # Only count positive hours
                        horas_totais_lote += horas
                except (KeyError, ValueError) as e:
                    print(
                        f"DEBUG: Error in verificar_porcentagem_horas_cargo processing subcontract: {e} - Data: {sub_aloc}"
                    )
                    continue
            # --- END SUBCONTRACTS INCLUSION ---

            if horas_totais_lote > 0:
                percent_lote = (horas_cargo_alvo_lote / horas_totais_lote) * 100
                detalhes_por_lote[nome_lote] = round(percent_lote, 2)
            else:
                detalhes_por_lote[nome_lote] = 0.0

            horas_cargo_alvo_global += horas_cargo_alvo_lote
            horas_totais_projeto_global += horas_totais_lote

        if horas_totais_projeto_global == 0:
            porcentagem_global = 0.0
        else:
            porcentagem_global = (
                horas_cargo_alvo_global / horas_totais_projeto_global
            ) * 100

        return round(porcentagem_global, 2), detalhes_por_lote
