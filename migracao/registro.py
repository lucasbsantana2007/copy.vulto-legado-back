"""Lê o registro_decisoes.md (a fonte única das decisões) para as abas Decisões das planilhas geradas."""

import os
import re

REGISTRO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "registro_decisoes.md")
COLUNAS = ["#", "O que encontramos", "Onde (arquivo / tabela / campo)", "Quantos registros afeta", "O que decidimos fazer",
           "Por quê", "Impacto no número final", "Tipo"]
LARGURAS = [4, 50, 34, 18, 50, 40, 36, 14]


def decisoes(de: int, ate: int) -> list[list]:
    """Linhas da tabela do registro com o número entre `de` e `ate`, sem a marcação do Markdown."""
    linhas = []
    with open(REGISTRO, encoding="utf-8") as f:
        for linha in f:
            m = re.match(r"\| (\d+) \|", linha)
            if not m or not de <= int(m.group(1)) <= ate:
                continue
            celulas = [c.strip() for c in linha.strip().strip("|").split(" | ")]
            linhas.append([int(celulas[0])] + [re.sub(r"\*\*|`", "", c) for c in celulas[1:]])
    return linhas
