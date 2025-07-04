# src/core/scheduler.py
# Este módulo contém a lógica de negócio para gerar o cronograma.
# Ele é independente da interface do usuário.

import pandas as pd
from datetime import datetime


class ProjectScheduler:
    """
    Calcula a alocação de horas do projeto com base nas entradas e regras de negócio.
    """

    def __init__(self, input_data):
        """
        Inicializa o agendador com os dados fornecidos pela UI.

        Args:
            input_data (dict): Um dicionário contendo todos os parâmetros de entrada.
        """
        self.data = input_data
        self.categories = input_data["categories"]
        self.employee_classes = [
            "Projetista/Estagiário",
            "Eng. Júnior",
            "Eng. Pleno",
            "Eng. Sênior",
        ]

        # Valida e converte os dados de entrada para os tipos corretos
        try:
            self.total_hours_per_month_per_person = int(
                input_data["globals"]["total_hours"]
            )
            if self.total_hours_per_month_per_person <= 0:
                raise ValueError("Horas/Mês/Pessoa deve ser um número positivo.")
            self.num_lots = int(input_data["globals"]["num_lots"])
            self.max_percentage_per_class = {
                cls: float(perc) / 100.0
                for cls, perc in input_data["globals"]["max_percentages"].items()
            }
        except (ValueError, TypeError) as e:
            raise ValueError(
                f"Erro nos dados de entrada globais. Verifique se os números são válidos. Detalhe: {e}"
            )

    def _parse_and_validate_lots(self):
        """
        Analisa a string 'applicable_lots' para cada categoria, convertendo-a
        numa lista de inteiros validados.
        """
        total_lots_range = set(range(1, self.num_lots + 1))
        for cat_name, cat_data in self.categories.items():
            lots_str = cat_data.get("applicable_lots", "Todos").strip().lower()
            if not lots_str or lots_str == "todos":
                cat_data["parsed_lots"] = list(total_lots_range)
            else:
                try:
                    parsed = {int(x.strip()) for x in lots_str.split(",")}
                    # Valida se os lotes estão dentro do intervalo global
                    if not parsed.issubset(total_lots_range):
                        raise ValueError(
                            f"Lotes inválidos para a categoria '{cat_name}'. Os lotes devem estar entre 1 e {self.num_lots}."
                        )
                    cat_data["parsed_lots"] = sorted(list(parsed))
                except Exception:
                    raise ValueError(
                        f"Formato de 'Lotes Aplicáveis' inválido para a categoria '{cat_name}'. Use números separados por vírgula (ex: 1, 3)."
                    )

    def generate_schedule(self):
        """
        Orquestra a geração do cronograma.

        Returns:
            pandas.DataFrame: Um DataFrame contendo o cronograma detalhado.
            list: Uma lista de avisos ou erros de validação.
        """
        # Converte as datas de string para datetime e valida os lotes
        self._parse_dates()
        self._parse_and_validate_lots()

        # Encontra o intervalo de tempo total do projeto
        all_dates = []
        for cat_data in self.categories.values():
            if cat_data["start_date"] and cat_data["end_date"]:
                all_dates.extend([cat_data["start_date"], cat_data["end_date"]])

        if not all_dates:
            return pd.DataFrame(), ["Nenhuma data de categoria foi fornecida."]

        project_start_date = min(all_dates)
        project_end_date = max(all_dates)

        # Gera um range de meses para o projeto inteiro
        monthly_periods = pd.date_range(
            start=project_start_date, end=project_end_date, freq="MS"
        )

        schedule_records = []

        # Itera sobre cada mês do projeto
        for month_start in monthly_periods:
            # Itera sobre cada categoria de trabalho
            for cat_name, cat_data in self.categories.items():
                # Pula se a categoria não tiver datas, lotes definidos ou horas inseridas
                if not (
                    cat_data["start_date"]
                    and cat_data["end_date"]
                    and cat_data["parsed_lots"]
                    and any(cat_data["total_hours"].values())
                ):
                    continue

                # Verifica se a categoria está ativa neste mês
                if self._is_active_in_month(cat_data, month_start):

                    concurrent_lots_for_cat = len(cat_data["parsed_lots"])
                    if concurrent_lots_for_cat == 0:
                        continue

                    # Itera sobre cada classe de funcionário
                    for emp_class in self.employee_classes:
                        # A ENTRADA AGORA É O TOTAL DE HORAS PARA A CLASSE NA CATEGORIA
                        total_hours_for_class = cat_data["total_hours"].get(
                            emp_class, 0
                        )

                        if total_hours_for_class > 0:
                            # Calcula o "equivalente em pessoas" para fins de relatório
                            headcount_equivalent = (
                                total_hours_for_class
                                / self.total_hours_per_month_per_person
                            )

                            # Regra: divide as horas totais da classe igualmente entre os lotes aplicáveis
                            allocated_hours_per_lot = (
                                total_hours_for_class / concurrent_lots_for_cat
                            )

                            # Regra: Arredonda para 0.5 ou inteiro
                            allocated_hours_per_lot = (
                                round(allocated_hours_per_lot * 2) / 2
                            )

                            # Itera sobre os lotes específicos para esta categoria para criar um registro para cada
                            for lot_num in cat_data["parsed_lots"]:
                                schedule_records.append(
                                    {
                                        "Mês": month_start.strftime("%Y-%m"),
                                        "Lote": lot_num,
                                        "Categoria": cat_name,
                                        "Classe de Funcionário": emp_class,
                                        "Equivalente Pessoas": round(
                                            headcount_equivalent, 2
                                        ),
                                        "Horas Alocadas (por Lote)": allocated_hours_per_lot,
                                    }
                                )

        if not schedule_records:
            return pd.DataFrame(), [
                "Nenhuma alocação pôde ser gerada com os dados fornecidos."
            ]

        schedule_df = pd.DataFrame(schedule_records)

        # Realiza as validações finais
        validation_warnings = self._validate_schedule(schedule_df)

        return schedule_df, validation_warnings

    def _parse_dates(self):
        """Converte as strings de data do dicionário de entrada em objetos datetime."""
        for cat_name, cat_data in self.categories.items():
            try:
                start_str = cat_data.get("start_date")
                end_str = cat_data.get("end_date")
                # errors='coerce' transforma datas inválidas em NaT (Not a Time)
                cat_data["start_date"] = (
                    pd.to_datetime(start_str, errors="coerce") if start_str else None
                )
                cat_data["end_date"] = (
                    pd.to_datetime(end_str, errors="coerce") if end_str else None
                )
            except Exception as e:
                raise ValueError(
                    f"Formato de data inválido para '{cat_name}'. Use AAAA-MM-DD. Detalhe: {e}"
                )

    def _is_active_in_month(self, cat_data, month_start):
        """Verifica se uma categoria está ativa em um determinado mês."""
        month_end = month_start + pd.offsets.MonthEnd(1)
        # pd.isnull verifica se a data é NaT ou None
        if pd.isnull(cat_data["start_date"]) or pd.isnull(cat_data["end_date"]):
            return False
        return (
            cat_data["start_date"] <= month_end and cat_data["end_date"] >= month_start
        )

    def _validate_schedule(self, df):
        """
        Valida o cronograma gerado contra as regras de negócio.
        """
        warnings = []
        if df.empty:
            return warnings

        # 1. Validação: Percentual máximo de horas por classe
        # A validação agora é feita sobre o total de horas alocadas
        total_hours_project = df["Horas Alocadas (por Lote)"].sum()
        if total_hours_project > 0:
            hours_per_class = df.groupby("Classe de Funcionário")[
                "Horas Alocadas (por Lote)"
            ].sum()
            percentage_per_class = hours_per_class / total_hours_project

            for emp_class, percentage in percentage_per_class.items():
                max_perc = self.max_percentage_per_class.get(emp_class, 0)
                if percentage > max_perc:
                    warnings.append(
                        f"AVISO: A classe '{emp_class}' excedeu o limite de {max_perc:.1%}, "
                        f"utilizando {percentage:.1%} do total de horas."
                    )

        # 2. Validação de alocação de pessoal (>100%)
        # Esta validação muda de escopo. Como o input agora é o total de horas,
        # a responsabilidade de não sobrecarregar uma equipe dentro de UMA categoria
        # é do usuário ao inserir as horas.
        # Podemos, no entanto, verificar se um TIPO de funcionário (ex: Eng. Pleno)
        # está sobrecarregado somando seu trabalho em MÚLTIPLAS categorias no mesmo mês.

        # Agrupa por Mês e Classe de Funcionário, somando o equivalente de pessoas.
        # Se a soma for maior que o número de pessoas que você tem na vida real, é um sinal de alerta.
        # Esta é uma validação mais complexa e depende de um input que não temos (total de pessoas por classe).
        # Por enquanto, a validação de sobrecarga fica implícita na entrada do usuário.

        return warnings
