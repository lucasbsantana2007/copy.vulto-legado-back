"""Confere as tabelas padronizadas antes de gravar no vulto.db. Chamado pelo carregar.py: se encontrar um erro,
para e não grava nada (a gente não passa para a migração com tabela errada).

Também roda sozinho, conferindo o que já está no vulto.db:  python migracao/validar.py
"""

import re
import sqlite3
import sys

import pandas as pd

import caminhos

# tabelas originais da Vulto: nunca mudam (são o "sistema da Vulto" da comparação)
ORIGINAIS = {"vendas": 18598, "itens_venda": 31110, "compras": 14948, "pontos_mov": 12127, "estoque": 7130,
             "entradas_estoque": 17331, "lojas": 7, "funcionarios": 24, "config": 6}
STATUS = {"Concluída", "Cancelada", "Estornada", "Chargeback", "Aguardando pagamento", "Sem status"}
CANAIS = {"Loja física", "Site", "App"}
CATEGORIAS = {"Sem categoria", "Bronze", "Prata", "Ouro", "Vip"}
LOCAIS = {"Pinheiros": "pinheiros", "Eldorado": "eldorado", "Savassi": "savassi", "Batel": "batel", "CD Cajamar": "cd_cajamar"}


def valor(texto):
    return float(str(texto).replace(",", "."))


def conferir(t):
    erros, avisos = [], []

    def checa(condicao, mensagem, aviso=False):
        if not condicao:
            (avisos if aviso else erros).append(mensagem)

    # ---------------------------------------------------------------- originais intactas
    with sqlite3.connect(caminhos.VULTO_DB) as c:
        for tabela, linhas in ORIGINAIS.items():
            n = c.execute(f"select count(*) from {tabela}").fetchone()[0]
            checa(n == linhas, f"{tabela} original tem {n} linhas (deveria ter {linhas}): alguém mexeu na tabela original")
        utc = c.execute("select count(*) from vendas where criado_em like '%Z'").fetchone()[0]
        checa(utc == 8738, "vendas original foi alterada (as datas em UTC sumiram)")

    # ---------------------------------------------------------------- vendas
    v = t["vendas_padronizadas"]
    ids = set(v.id_venda.astype(str))
    checa(v.id_venda.astype(str).is_unique, "vendas_padronizadas: id_venda repetido")
    checa(set(v.status) <= STATUS, f"vendas_padronizadas: status fora do padrão {set(v.status) - STATUS}")
    checa(set(v.canal) <= CANAIS, f"vendas_padronizadas: canal fora do padrão {set(v.canal) - CANAIS}")
    checa(set(v.entra_no_faturamento) <= {"Sim", "Não", "Pendente"}, "vendas_padronizadas: entra_no_faturamento fora do padrão")
    pend = (v.entra_no_faturamento == "Pendente").sum()
    checa(pend == 0, f"vendas_padronizadas: {pend} vendas ainda 'Pendente' (decisão em aberto)", aviso=True)
    checa(v.mes.str.fullmatch(r"\d{4}-\d{2}").all(), "vendas_padronizadas: mes fora do formato AAAA-MM")
    checa(v.valor_total.map(lambda x: re.fullmatch(r"-?\d+,\d{2}", str(x)) is not None).all(),
          "vendas_padronizadas: valor_total fora do formato 559,70")
    checa(not ((v.tipo == "Teste") & (v.entra_no_faturamento == "Sim")).any(), "venda de teste entrando no faturamento")
    checa(not ((v.status != "Concluída") & (v.status != "Sem status") & (v.entra_no_faturamento == "Sim")).any(),
          "venda não paga entrando no faturamento")

    # ---------------------------------------------------------------- produtos e itens
    p = t["produtos_padronizados"]
    skus = set(p.sku)
    checa(p.sku.is_unique, "produtos_padronizados: SKU repetido")
    checa(p[p.sku != "DIVERSO"].preco.notna().all(), "produtos_padronizados: produto sem preço")
    i = t["itens_venda_padronizados"]
    checa(set(i.venda_id.astype(str)) <= ids, "itens_venda_padronizados: item de venda que não existe")
    testes = set(v.loc[v.tipo == "Teste", "id_venda"].astype(str))
    fora = i[~i.sku.isin(skus) & ~i.venda_id.astype(str).isin(testes)]
    checa(fora.empty, f"itens_venda_padronizados: {len(fora)} itens com SKU fora do catálogo {sorted(set(fora.sku))[:5]}")
    batel = i[i.venda_id.astype(str).str.startswith("BATEL-")].assign(t=lambda d: d.qtd.astype(int) * d.preco_unit.map(valor))
    soma = batel.groupby(batel.venda_id.astype(str)).t.sum().round(2)
    ref = v.set_index(v.id_venda.astype(str)).valor_total.map(valor)
    checa(((soma - ref.reindex(soma.index)).abs() < 0.01).all(), "vendas da Batel: soma dos itens diferente do valor da venda")

    # ---------------------------------------------------------------- clientes e pontos
    cl = t["clientes_padronizados"]
    checa(cl.cpf.is_unique, "clientes_padronizados: CPF repetido")
    checa(cl.cpf.astype(str).str.fullmatch(r"\d{3}\.\d{3}\.\d{3}-\d{2}").all(), "clientes_padronizados: CPF fora do formato")
    checa(set(cl.categoria) <= CATEGORIAS, f"clientes_padronizados: categoria fora do padrão {set(cl.categoria) - CATEGORIAS}")
    checa((pd.to_numeric(cl.saldo_valido_01_10_2026) >= 0).all(), "clientes_padronizados: saldo negativo")
    checa(set(t["extrato_pontos"].cpf) <= set(cl.cpf), "extrato_pontos: movimento de CPF que não está no cadastro")

    # ---------------------------------------------------------------- estoque
    e = t["estoque_01_10"]
    h = t["estoque_historico"]
    checa((h.saldo >= 0).all(), "estoque_historico: saldo negativo em algum dia")
    checa(set(h.sku) <= skus, f"estoque_historico: SKU fora do catálogo {sorted(set(h.sku) - skus)[:5]}")
    final = h.groupby(["local", "sku"]).qtd.sum()
    oficial = e.melt(id_vars="sku", value_vars=list(LOCAIS.values()), var_name="col", value_name="qtd")
    oficial["local"] = oficial.col.map({v: k for k, v in LOCAIS.items()})
    oficial = oficial.set_index(["local", "sku"]).qtd.astype(int)
    oficial = oficial[oficial != 0]
    diferentes = (final.reindex(oficial.index).fillna(0) - oficial).abs().sum() + final[~final.index.isin(oficial.index)].abs().sum()
    checa(diferentes == 0, f"estoque_historico não termina no estoque oficial de 01/10 ({diferentes} peças de diferença)")

    # ---------------------------------------------------------------- atendimentos
    a = t["atendimentos_padronizados"]
    checa(a.id.is_unique, "atendimentos_padronizados: id repetido")
    ligadas = a[a.id_venda.astype(str) != ""]
    checa(set(ligadas.id_venda.astype(str)) <= ids, "atendimentos_padronizados: venda ligada que não existe")
    com_cli = a[a.cliente_cpf.astype(str) != ""]
    checa(set(com_cli.cliente_cpf) <= set(cl.cpf), "atendimentos_padronizados: cliente ligado que não está no cadastro")

    # ---------------------------------------------------------------- registro
    r = t["registro_decisoes"]
    abertas = r[r.o_que_decidimos_fazer.str.contains("A decidir", na=False)].numero.tolist()
    checa(not abertas, f"registro de decisões: em aberto {abertas}", aviso=True)

    for a_ in avisos:
        print("  aviso:", a_)
    if erros:
        for e_ in erros:
            print("  ERRO:", e_)
        sys.exit("\nConferência falhou: nada foi gravado no vulto.db. Corrija e rode de novo.")
    print(f"  conferência ok ({len(t)} tabelas; originais intactas)")


if __name__ == "__main__":
    with sqlite3.connect(caminhos.VULTO_DB) as c:
        nomes = [n for (n,) in c.execute("select name from sqlite_master where type='table'") if n not in ORIGINAIS
                 and n != "sqlite_sequence"]
        conferir({n: pd.read_sql(f"select * from {n}", c) for n in nomes})
