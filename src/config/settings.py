DISCIPLINAS = {
    "GEO": "Geometria",
    "TER": "Terraplenagem",
    "DRE": "Drenagem",
    "PAV": "Pavimento",
    "SIN": "Sinalização",
    "GEOL": "Geologia",
    "GEOT": "Geotecnia",
    "EST": "Estruturas",
    "COO": "Coordenação",
    "ETH": "Estudos Hidrológicos",
    "ETA": "Estudo de Alternativa",
    "TOP": "Topografia",
    "PAPE": "Passarelas e de Pedestres",
    "GEGE": "Estudos Geológicos-Geotécnicos",
    "PAS": "Paisagismo",
    "DES": "Desapropriação",
    "PRO": "Parada de Ônibus",
    "OBC": "Obras Complementares",
    "INT": "Interferência",
}

CARGOS = {
    "ESTAG": "Estagiário/Projetista",
    "ENG_JR": "Eng. Júnior",
    "ENG_PL": "Eng. Pleno",
    "ENG_SR": "Eng. Sênior",
    "COORD": "Coordenador",
    "DES": "Desenhista",
    "COORD_EQP": "Coordenador de Equipe",
    "ENG_CONS": "Eng. Consultor",
    "BIM_MAN": "BIM Manager",
}

SUBCONTRATOS = {
    "OAE": "OAE",
    "ILU": "Iluminação",
    "GEO": "Geotecnia",
    "OBA_CIV": "Obra Civil",
}

OUTROS = {
    "ANTEC": "Análise Técnica",
    "CII": "Consolidação e Identificação de Interferências",
    "OBE": "Obras Emergenciais",
    "PAEB": "Prospecção de Áreas de Empréstimo e Bota-fora",
    "AEV": "Estudo de Alternativas e Engenharia de Valores",
    "ACS": "Acompanhamento de Sondagens",
}

LISTA_TAREFAS_GERAIS = list(OUTROS.values())


MAPEAMENTO_TAREFA_DISCIPLINA = {
    "Estudos Geológicos-Geotécnicos": ["Geotecnia", "Geologia"],
    "Estudos Hidrológicos": ["Drenagem"],
    "Estudo de Alternativa": ["Geometria"],
    "Topografia": ["Geometria"],
    "Passarelas e de Pedestres": ["OAE"],
    "Paisagismo": ["Sinalização"],
    "Desapropriação": ["Geometria"],
    "Parada de Ônibus": ["Geometria"],
    "Obras Complementares": ["Sinalização"],
    "Interferência": ["Sinalização"],
}
