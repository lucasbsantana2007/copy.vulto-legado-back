"""
Roda a padronização completa e gera, na pasta de saída:

  clientes_corrigidos.xlsx, catalogo_corrigido.xlsx, vendas_batel_corrigidas.xlsx,
  atendimento_corrigido.xlsx, inventario_corrigido.xlsx,
  registro_decisoes_numeros.xlsx  (quantos registros cada decisão afetou)
  vulto_dadoscorrigidos.db        (as mesmas tabelas em SQLite, para o backend)

Uso:
    python rodar_tudo.py <pasta data da Vulto> [pasta de saída]

Os arquivos originais não são alterados. Ficam de fora as decisões 38 a 41 (ainda pendentes).
"""
import os
import sqlite3
import sys

import pandas as pd

import atendimento
import catalogo
import clientes
import inventario
import vendas_batel
from comum import salvar_excel

DECISOES = {
    1: "CPFs incompletos", 2: "CPFs repetidos", 3: "Pontos vazios", 4: "Categorias inconsistentes",
    5: "E-mails inválidos ou mal formatados", 6: "Custos ausentes", 7: "EAN ausente ou incompleto",
    8: "Produtos aguardando revisão de preço", 9: "Produtos duplicados", 10: "Margens abaixo do mínimo",
    11: "Fornecedores ausentes", 12: "Tamanhos inconsistentes", 13: "Códigos de CD ausentes",
    14: "Datas incorretas (Batel)", 15: "Valores armazenados como texto (Batel)", 16: "CPFs sem zeros à esquerda (Batel)",
    17: "Vendedores com apelidos diferentes", 18: "Formas de pagamento inconsistentes",
    19: "Cancelamentos identificados por texto riscado", 20: "Lançamentos no Vulto+ não estruturados",
    21: "Datas, canais e atendentes inconsistentes", 22: "Contatos incompletos e nomes inconsistentes",
    23: "Assuntos e tags redundantes", 24: "Problemas de acentuação", 25: "Pedidos ausentes ou mal formatados",
    26: "Status inconsistentes", 27: "Atendimentos resolvidos, mas ainda abertos", 28: "Peças avariadas",
    29: "Provadores não contabilizados", 30: "Recontagem inexistente", 31: "Contagem somente por modelo (Savassi)",
    32: "Estoque reservado no CD", 33: "Datas de contagem diferentes", 34: "Ausência de registros de avaria",
    35: "SKUs e códigos incompatíveis", 36: "Produtos fora do catálogo atual", 37: "Códigos de CD com dígitos invertidos",
}


def main(pasta, saida):
    os.makedirs(saida, exist_ok=True)
    caminho = lambda nome: os.path.join(saida, nome)
    registros = []

    cli, reg = clientes.processar(pasta)
    registros.append(reg)
    salvar_excel(caminho("clientes_corrigidos.xlsx"), {"Clientes": cli["clientes"], "Alertas por e-mail": cli["alertas"],
                                                      "Ocorrências": reg.df(), "Resumo": reg.resumo()})
    print("clientes ok")

    cat, reg = catalogo.processar(pasta)
    registros.append(reg)
    salvar_excel(caminho("catalogo_corrigido.xlsx"), {"Catálogo": cat["catalogo"], "Cadastro mestre": cat["mestre"],
                                                     "Listas": cat["listas"], "Ocorrências": reg.df(),
                                                     "Resumo": reg.resumo()},
                 destaques={"Catálogo": catalogo.destaque_catalogo, "Cadastro mestre": catalogo.destaque_catalogo})
    print("catálogo ok")

    bat, reg = vendas_batel.processar(pasta)
    registros.append(reg)
    salvar_excel(caminho("vendas_batel_corrigidas.xlsx"), {"Vendas": bat["vendas"], "Totais por semana": bat["totais"],
                                                          "Ocorrências": reg.df(), "Resumo": reg.resumo()},
                 destaques={"Vendas": vendas_batel.destaque})
    print("vendas Batel ok")

    ate, reg = atendimento.processar(pasta)
    registros.append(reg)
    salvar_excel(caminho("atendimento_corrigido.xlsx"), {"Atendimentos": ate["atendimentos"],
                                                        "Ocorrências": reg.df(), "Resumo": reg.resumo()})
    print("atendimento ok")

    inv, reg = inventario.processar(pasta, cat["mestre"], cat["skus_2024"])
    registros.append(reg)
    abas = {"Estoque consolidado": inv["consolidado"], "Resumo por loja": inv["resumo_lojas"]}
    for loja in inv["consolidado"].loja.unique():
        abas[loja.upper()] = inv["consolidado"][inv["consolidado"].loja == loja]
    abas.update({"AVARIA PINHEIROS": inv["avaria"], "Ocorrências": reg.df(), "Resumo": reg.resumo()})
    salvar_excel(caminho("inventario_corrigido.xlsx"), abas,
                 destaques={k: inventario.destaque for k in abas if k.isupper() or k == "Estoque consolidado"})
    print("inventário ok")

    # ------------------------------------------------------------ números para o Registro de Decisões
    ocorr = pd.concat([r.df() for r in registros], ignore_index=True)
    detalhe = pd.concat([r.resumo() for r in registros], ignore_index=True)
    por_decisao = (detalhe.groupby("decisao").registros.sum().reset_index()
                   .assign(problema=lambda x: x.decisao.map(DECISOES)))
    por_decisao = por_decisao[["decisao", "problema", "registros"]]
    salvar_excel(caminho("registro_decisoes_numeros.xlsx"), {"Por decisão": por_decisao, "Detalhe": detalhe})

    # ------------------------------------------------------------ banco SQLite
    db = caminho("vulto_dadoscorrigidos.db")
    if os.path.exists(db):
        os.remove(db)
    with sqlite3.connect(db) as c:
        tabelas = {"clientes": cli["clientes"], "clientes_alertas_email": cli["alertas"],
                   "catalogo": cat["catalogo"], "catalogo_mestre": cat["mestre"], "catalogo_listas": cat["listas"],
                   "vendas_batel_marco": bat["vendas"], "atendimentos": ate["atendimentos"],
                   "estoque_set26": inv["consolidado"], "estoque_avaria": inv["avaria"],
                   "ocorrencias": ocorr, "registro_decisoes": por_decisao}
        for nome, df in tabelas.items():
            df.astype({col: str for col in df.columns if df[col].dtype == object}).replace({"None": None, "nan": None}) \
                .to_sql(nome, c, index=False)
    print(f"\nArquivos em {os.path.abspath(saida)}")
    print(por_decisao.to_string(index=False))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "../vulto-legado-back/data",
         sys.argv[2] if len(sys.argv) > 2 else "saida")
