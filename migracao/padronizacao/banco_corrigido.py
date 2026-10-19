"""
Gera o vulto_corrigido.db: o vulto.db com as MESMAS tabelas e as MESMAS colunas, só com os dados consertados.

O arquivo é uma cópia do original (a estrutura não muda); o conteúdo de cada tabela é trocado pela versão corrigida.
O que nunca existiu no vulto.db (planilhas, histórico de estoque, Registro de Decisões) vai para outro arquivo,
planilhas_e_relatorios.db (ver rodar_tudo.py).

Padrões adotados em todas as tabelas
  datas    'AAAA-MM-DD HH:MM:SS' no horário de Brasília (entradas_estoque: 'AAAA-MM-DD')      decisão 44
  valores  em reais, com ponto decimal e 2 casas ('129.90')                                     decisão 38
  canal    'loja', 'site' ou 'app'                                                              decisão 45
  status   'pago', 'cancelado', 'estornado', 'chargeback' ou 'aguardando'                       decisão 39
  cliente  CPF com 11 dígitos quando dá para identificar                                         decisão 50
  SKU      o código do catálogo (VLT-CAT-NNN-COR-TAM)                                            decisão 35
"""
import os
import shutil
import sqlite3

import pandas as pd

from comum import DATA_CORTE, LOJAS, chave, data_vulto, digitos, dinheiro, identificar, sku_canonico, vazio

CANAL = {"Loja física": "loja", "Site": "site", "App": "app"}
STATUS = {"Paga": "pago", "Cancelada": "cancelado", "Estornada": "estornado", "Chargeback": "chargeback",
          "Pagamento não confirmado": "aguardando"}
OBS_PONTOS = {
    "DOBRO_DEVIDO": "dobro da compra online: o sistema creditou só 1x (decisão 53)",
    "SALDO_PLANILHA": "saldo da planilha Vulto+ congelado em 01/01/2025 (decisão 3)",
    "CREDITO_BATEL": "venda manual da Batel em mar/26, sem pontos lançados (decisão 58)",
    "VENCIMENTO": "pontos vencidos: 12 meses, Ouro 18 (decisão 54)",
    "SALDO_ZERADO": "resgate maior que o saldo: saldo zerado (decisão 56)",
}


def dois(x):
    return None if x is None or (isinstance(x, float) and x != x) else f"{float(x):.2f}"


def iso(d, formato="%Y-%m-%d %H:%M:%S"):
    return None if d is None or pd.isna(d) else pd.Timestamp(d).strftime(formato)


def gerar(pasta, destino, ven, pts, est, clientes):
    """Copia o vulto.db e reescreve cada tabela. Devolve um resumo do que mudou por tabela."""
    original = os.path.join(pasta, "vulto.db")
    if os.path.exists(destino):
        os.remove(destino)
    shutil.copyfile(original, destino)
    with sqlite3.connect(original) as c:
        antes = {t: pd.read_sql(f"select * from {t}", c) for t in
                 ("vendas", "itens_venda", "compras", "pontos_mov", "estoque", "entradas_estoque", "config")}

    vendas, removidas, novos_ids = tabela_vendas(ven["vendas"])
    itens = tabela_itens(ven["itens"], removidas, novos_ids)
    compras = tabela_compras(antes["compras"], ven["vendas"], removidas, clientes)
    pontos = tabela_pontos(pts, antes["pontos_mov"])
    estoque = tabela_estoque(est["posicao"])
    entradas = tabela_entradas(antes["entradas_estoque"], est["recebimentos_criados"])
    config = antes["config"].copy()
    config.loc[config.chave == "valor_ponto", "valor"] = "0.05"                     # decisão 40

    novas = {"vendas": vendas, "itens_venda": itens, "compras": compras, "pontos_mov": pontos, "estoque": estoque,
             "entradas_estoque": entradas, "config": config}
    with sqlite3.connect(destino) as c:
        for nome, df in novas.items():
            colunas = [r[1] for r in c.execute(f"pragma table_info({nome})")]
            c.execute(f"delete from {nome}")
            c.executemany(f"insert into {nome} ({', '.join(colunas)}) values ({', '.join('?' * len(colunas))})",
                          valores(df, colunas))
        for (t,) in c.execute("select name from sqlite_sequence").fetchall():
            c.execute(f"update sqlite_sequence set seq = (select coalesce(max(id), 0) from {t}) where name = ?", (t,))
    resumo = []
    for nome, df in novas.items():
        a = antes[nome]
        com_id = "id" in a.columns
        resumo.append(dict(tabela=nome, linhas_antes=len(a), linhas_depois=len(df),
                           removidas=len(set(a.id) - set(df.id)) if com_id else None,
                           incluidas=len(set(df.id) - set(a.id)) if com_id else None))
    resumo += [dict(tabela=t, linhas_antes=None, linhas_depois=None, removidas=0, incluidas=0)
               for t in ("lojas", "funcionarios")]
    return pd.DataFrame(resumo)


INTEIRAS = {"id", "loja_id", "qtd", "compra_id", "venda_id"}


def valores(df, colunas):
    """Linhas prontas para o sqlite: vazio vira NULL e as colunas de número inteiro voltam a ser inteiras."""
    def conv(x, col):
        if x is None or (isinstance(x, float) and x != x) or x is pd.NaT:
            return None
        if col in INTEIRAS:
            return int(x)
        return x.item() if hasattr(x, "item") else x
    return [tuple(conv(x, col) for x, col in zip(linha, colunas)) for linha in df[colunas].itertuples(index=False)]


def tabela_vendas(v):
    """Tira pedidos de teste (43) e gravações duplicadas (42); inclui as vendas da Batel que só estavam na planilha
    (41) com ids novos; padroniza o resto."""
    sai = v.teste.astype(bool) | v.duplicada_de.notna()
    removidas = set(v.loc[sai, "id"])
    v = v[~sai].copy()
    extras = v.id.astype(str).str.startswith("BATEL-")
    proximo = int(pd.to_numeric(v.loc[~extras, "id"]).max()) + 1
    novos = v[extras].sort_values("data_hora")
    novos_ids = dict(zip(novos.id, range(proximo, proximo + len(novos))))
    v.loc[novos.index, "id"] = list(novos_ids.values())
    v.loc[novos.index, "obs"] = "venda manual da Batel (PDV fora do ar): " + novos.obs.fillna("").astype(str)
    ref = []
    for r in v.itertuples():
        if not pd.isna(r.cliente_cpf):
            ref.append(r.cliente_cpf)
        else:
            tipo, val = identificar(r.cliente_ref)
            ref.append(None if tipo == "anonimo" else (val or str(r.cliente_ref).strip()))
    out = pd.DataFrame(dict(
        id=v.id.astype(int), pedido_externo=v.pedido_externo, canal=v.canal.map(CANAL), loja_id=v.loja_id.astype(int),
        cliente_ref=ref, vendedor=v.vendedor, total=v.total_corrigido.map(dois), desconto=v.desconto_valor.map(dois),
        frete=v.frete_valor.map(dois), status=v.status.map(STATUS), criado_em=v.data_hora.map(iso), obs=v.obs,
        compra_id=v.compra_id))
    return out.sort_values("id"), removidas, novos_ids


def tabela_itens(itens, removidas, novos_ids):
    """Itens das vendas que ficaram; os das vendas da Batel que só estavam na planilha seguem o id novo da venda."""
    i = itens[~itens.venda_id.isin(removidas)]
    venda = i.venda_id.map(lambda x: novos_ids.get(x, x))
    return pd.DataFrame(dict(id=i.id, venda_id=venda, sku=i.sku, descricao=i.descricao, qtd=i.qtd,
                             preco_unit=i.preco_unit.map(dois))).sort_values("id")


def tabela_compras(compras, vendas, removidas, clientes):
    """Valor em reais (centavos do app ÷ 100), canal e loja do app corrigidos, cliente pelo CPF. Sai a compra que
    só existia por causa de um pedido de teste ou duplicado (os pontos dela também saíram)."""
    ligadas_removidas = set(vendas.loc[vendas.id.isin(removidas), "compra_id"].dropna())
    ficam = set(vendas.loc[~vendas.id.isin(removidas), "compra_id"].dropna())
    c = compras[~compras.id.isin(ligadas_removidas - ficam)].copy()
    centavos = (c.origem == "app") & ~c.valor.astype(str).str.contains(r"[.,]")
    c["valor_num"] = c.valor.map(lambda x: dinheiro(x) or 0.0)
    c.loc[centavos, "valor_num"] = c.loc[centavos, "valor_num"] / 100
    canal = c.origem.map({"pdv": "loja", "site": "site", "app": "app"})
    loja = c.loja_id.where(c.origem != "app", 6)
    por_email = dict(zip(clientes.email.dropna(), clientes.loc[clientes.email.notna(), "cpf"]))
    tel = clientes.dropna(subset=["telefone"]).drop_duplicates("telefone", keep=False)
    por_tel = dict(zip(tel.telefone, tel.cpf))
    ref = []
    for x in c.cliente_ref:
        tipo, val = identificar(x)
        cpf = val if tipo == "cpf" else por_email.get(val) if tipo == "email" else por_tel.get(val)
        ref.append(cpf or (None if tipo == "anonimo" else (val or str(x).strip())))
    return pd.DataFrame(dict(id=c.id, cliente_ref=ref, valor=c.valor_num.map(dois), canal=canal, loja_id=loja,
                             criado_em=c.criado_em.map(lambda s: iso(data_vulto(s))), origem=c.origem))


def tabela_pontos(pts, original):
    """O extrato corrigido no formato do pontos_mov (C crédito, D resgate, E estorno, AJ ajuste). A soma por CPF
    (C + AJ + E − D) é exatamente o saldo em 01/10/2026."""
    ex = pts["extrato"]
    obs_original = dict(zip(original.id, original.obs))
    linhas, novas = [], []
    for r in ex.itertuples():
        mantida = str(r.origem).startswith("pontos_mov ") and r.tipo in ("CREDITO", "RESGATE", "ESTORNO", "AJUSTE")
        tipo = {"CREDITO": "C", "RESGATE": "D", "ESTORNO": "E", "AJUSTE": "AJ", "DOBRO_DEVIDO": "C",
                "SALDO_PLANILHA": "AJ", "RESGATE_PAPEL": "D", "CREDITO_BATEL": "C"}[r.tipo]
        qtd = -r.qtd if tipo == "D" else r.qtd
        compra = None if pd.isna(r.compra_id) else int(r.compra_id)
        if mantida:
            pid = int(str(r.origem).split()[1])
            linhas.append(dict(id=pid, cpf=r.cpf, tipo=tipo, qtd=int(qtd), compra_id=compra, criado_em=iso(r.data),
                               obs=obs_original.get(pid)))
        else:
            obs = OBS_PONTOS.get(r.tipo) or f"resgate anotado só no papel, aba Resgates ({r.obs}) (decisão 55)"
            novas.append(dict(cpf=r.cpf, tipo=tipo, qtd=int(qtd), compra_id=compra, criado_em=iso(r.data), obs=obs))
    for e in pts["eventos"].itertuples():
        novas.append(dict(cpf=e.cpf, tipo="AJ", qtd=int(e.qtd), compra_id=None, criado_em=iso(e.data),
                          obs=OBS_PONTOS[e.tipo]))
    novas = pd.DataFrame(novas).sort_values(["criado_em", "cpf"])
    proximo = int(original.id.max()) + 1
    novas["id"] = range(proximo, proximo + len(novas))
    out = pd.concat([pd.DataFrame(linhas), novas], ignore_index=True).sort_values("id")
    # conferência: o extrato fecha com o saldo de cada cliente
    soma = out.assign(s=out.qtd.where(out.tipo != "D", -out.qtd)).groupby("cpf").s.sum()
    esperado = pts["saldos"].set_index("cpf").saldo
    diferenca = (soma.reindex(esperado.index).fillna(0) - esperado).abs()
    assert diferenca.max() == 0, f"extrato não fecha com o saldo: {diferenca[diferenca > 0].head()}"
    return out


def tabela_estoque(posicao):
    """Posição em 01/10/2026 (decisões 59 a 63). As lojas 5 e 6 deixam de ter estoque: as peças estão no CD."""
    p = posicao[posicao.loja_id.isin([1, 2, 3, 4, 7])]
    return pd.DataFrame(dict(sku=p.sku, loja_id=p.loja_id, qtd=p.qtd.astype(int),
                             atualizado_em=iso(DATA_CORTE))).sort_values(["loja_id", "sku"])


def tabela_entradas(e, criados):
    """SKU do catálogo, data no formato ISO, custo em número, fornecedor 'CD' (ou a loja de origem) nos
    recebimentos de transferência; inclui os recebimentos que a loja não anotou (decisão 60). As anotações de
    envio do CD continuam como estavam (qtd vazia, destino no texto): o envio em si está no histórico."""
    e = e.copy()
    e["sku"] = e.sku.map(lambda s: sku_canonico(s) or str(s).strip().upper())
    e["data"] = e.data.map(lambda d: iso(data_vulto(d), "%Y-%m-%d"))
    e["custo_unit"] = e.custo_unit.map(lambda x: dois(dinheiro(x)) if not vazio(x) else None)
    receb = e.qtd.notna() & e.fornecedor.fillna("").str.strip().str.upper().isin(["", "CD"]) & (e.loja_id != 7)
    origem = e.obs.map(lambda o: next((LOJAS[i] for i in (1, 2, 3, 4) if f"veio da {LOJAS[i].lower()}" in chave(o)), "CD"))
    e.loc[receb, "fornecedor"] = origem[receb]
    if len(criados):
        proximo = int(e.id.max()) + 1
        novos = pd.DataFrame(dict(
            id=range(proximo, proximo + len(criados)), sku=criados.sku, loja_id=criados.destino_id, qtd=criados.qtd,
            custo_unit=None, nf=None, fornecedor=criados.origem_id.map(lambda i: "CD" if i == 7 else LOJAS[i]),
            data=criados.data.map(lambda d: iso(d, "%Y-%m-%d")),
            obs=criados.ref.map(lambda r: f"recebimento não anotado pela loja (envio {r}) (decisão 60)"),
            usuario="correcao_migracao"))
        e = pd.concat([e, novos], ignore_index=True)
    return e[["id", "sku", "loja_id", "qtd", "custo_unit", "nf", "fornecedor", "data", "obs", "usuario"]]
