"""Pastas da migração, relativas ao repositório.

- data/            dados originais da Vulto (vulto.db e planilhas): as tabelas e abas originais nunca são alteradas
- migracao/entrada planilhas que o time limpou à mão (é aqui que se edita)
- migracao/saida   o que os scripts geram (planilhas para conferir, CSV, log e o banco intermediário)
"""

import os

MIGRACAO = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(MIGRACAO)
DADOS = os.path.join(REPO, "data")
ENTRADA = os.path.join(MIGRACAO, "entrada")
SAIDA = os.path.join(MIGRACAO, "saida")

VULTO_DB = os.path.join(DADOS, "vulto.db")
PADRONIZADO = os.path.join(SAIDA, "vulto_padronizado.db")
