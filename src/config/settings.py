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
    "PAPE": "Passarelas e pedestres",
    "GEGE": "Estudos Gelógios-Geotécnicos",
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
}

SUBCONTRATOS = {
    "OAE": "OAE",
    "ILU": "Iluminação",
    "GEO": "Geotecnia",
    "OBA_CIV": "Obra Civil",
    "OUTROS": "Outros",
}

OUTROS = {
    "ANTEC": "Análise Técnica",
    "CII": "Consolidação e Identificação de Interferências",
    "OBE": "Obras Emergenciais",
    "PAEB": "Prospecção de Áreas de Empréstimo e Bota-fora",
    "AEV": "Estudo de Alternativas e Engenharia de Valores",
    "ACS": "Acompanhamento de Sondagens",
}


MAPEAMENTO_TAREFA_DISCIPLINA = {
    "Estudos geológicos-geotécnicos": ["Geotecnia", "Geologia"],
    "Estudos hidrológicos": ["Drenagem"],
    "Estudo de Alternativa": ["Geometria"],
    "Projeto de Topografia": ["Geometria"],
    "Projeto de Passarelas e de Pedestres": ["OAE"],
    "Projeto de Paisagismo": ["Sinalização"],
    "Projeto de Desapropriação": ["Geometria"],
    "Projeto de Parada de Ônibus": ["Geometria"],
    "Projeto de Obras complementares": ["Sinalização"],
    "Projeto de Interferência": ["Sinalização"],
}
