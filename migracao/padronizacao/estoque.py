"""
Estoque com histórico e posição em 01/10/2026 — decisões 59 a 64 do Registro de Decisões.

Entrada : vulto.db (entradas_estoque, estoque) + vendas e itens corrigidos (saída de vendas.py) +
          contagem física consolidada (saída de inventario.py) + cadastro mestre do catálogo
Saída   : movimentações (o histórico que vai para a tabela movimentacoes da Lorenzi), posição por loja e SKU
          em 01/10/2026, comparação com o estoque do sistema da Vulto e a qualidade do histórico por loja.

Como o histórico é montado
  1. Entradas de fornecedor, transferências (pareando o lado do CD com o lado da loja) e vendas/devoluções viram
     movimentos com data.
  2. A contagem física é a verdade na data em que foi feita. O saldo inicial (01/01/2025) é o que faz o histórico
     chegar na contagem; se isso deixasse o saldo negativo em algum dia, o saldo inicial sobe até o mínimo que
     evita estoque negativo e a diferença vira um ajuste de inventário na data da contagem.
  3. Da contagem até 30/09, os movimentos registrados levam o saldo até 01/10/2026.
"""
import os
import sqlite3
from collections import defaultdict
from datetime import datetime, time, timedelta

import pandas as pd

from comum import DATA_CORTE, LOJAS, Registro, chave, data_vulto, ler_transferencia, sku_canonico, vazio

INICIO = datetime(2025, 1, 1)
CD = 7
LOJA_POR_CONTAGEM = {"Pinheiros": 1, "Eldorado": 2, "Savassi": 3, "Batel": 4, "CD Cajamar": 7}
# decisão 33 (Arthur) + 62: Pinheiros contou depois do fechamento de 20/09; o CD mandou a posição do WMS de 28/09;
# Eldorado, Savassi e Batel não anotaram a data — considerada 28/09 (fim do dia), igual ao CD.
DATA_CONTAGEM = {1: datetime(2026, 9, 20, 23, 59), 2: datetime(2026, 9, 28, 23, 59), 3: datetime(2026, 9, 28, 23, 59),
                 4: datetime(2026, 9, 28, 23, 59), 7: datetime(2026, 9, 28, 23, 59)}
JANELA_TRANSFERENCIA = 2              # decisão 60: quando os dois lados anotam, é no mesmo dia (picos em 7 e 14 dias são a
                                       # reposição da semana seguinte, não a mesma transferência)
HORA_ENTRADA = time(8, 0)              # entradas só têm a data: entram antes das vendas do dia


def ler(pasta):
    with sqlite3.connect(os.path.join(pasta, "vulto.db")) as c:
        entradas = pd.read_sql("select * from entradas_estoque", c)
        sistema = pd.read_sql("select * from estoque", c)
    return entradas, sistema


def processar(pasta, vendas, itens, contagem, mestre):
    reg = Registro("Estoque")
    entradas, sistema = ler(pasta)
    entradas["sku_original"] = entradas.sku
    entradas["sku"] = entradas.sku.map(lambda s: sku_canonico(s) or str(s).strip().upper())
    entradas["dia"] = entradas.data.map(lambda d: data_vulto(d))
    entradas["quando"] = entradas.dia.map(lambda d: datetime.combine(d.date(), HORA_ENTRADA) if d is not None else None)
    movs = []

    def mov(loja, sku, quando, tipo, qtd, origem, decisao=None, ref=None):
        movs.append(dict(loja_id=loja, sku=sku, data_hora=quando, tipo=tipo, qtd=int(qtd), origem=origem,
                         decisao=decisao, ref=ref))

    # ------------------------------------------------------------------ entradas de fornecedor
    forn = entradas[entradas.qtd.notna() & ~entradas.fornecedor.fillna("").str.strip().str.upper().isin(["CD", ""])]
    for r in forn.itertuples():
        mov(r.loja_id, r.sku, r.quando, "ENTRADA", r.qtd, f"entradas_estoque {r.id}", ref=f"NF {r.nf} {r.fornecedor}")

    # ------------------------------------------------------------------ transferências (decisão 60)
    envios, recebimentos = [], []
    for r in entradas[entradas.qtd.isna()].itertuples():          # quem manda anota no texto, sem quantidade
        destino, qtd = ler_transferencia(r.obs)
        if destino is None:
            reg.add(60, "Envio sem destino legível", f"entradas_estoque {r.id}", "obs", r.obs, None)
            continue
        envios.append(dict(origem_id=r.loja_id, destino_id=destino, sku=r.sku, qtd=qtd, quando=r.quando, ref=r.id))
    receb = entradas[entradas.qtd.notna() & entradas.fornecedor.fillna("").str.strip().str.upper().isin(["CD", ""])
                     & (entradas.loja_id != CD)]
    for r in receb.itertuples():
        k = chave(r.obs)
        origem_id = next((i for nome, i in [("pinheiros", 1), ("eldorado", 2), ("savassi", 3), ("batel", 4)]
                          if f"veio da {nome}" in k), CD)
        recebimentos.append(dict(origem_id=origem_id, destino_id=r.loja_id, sku=r.sku, qtd=int(r.qtd), quando=r.quando,
                                 ref=r.id))
    pares, so_envio, so_receb, lag = parear(pd.DataFrame(envios), pd.DataFrame(recebimentos))
    for p in pares:
        mov(p["origem_id"], p["sku"], p["envio"], "TRANSFERENCIA_SAIDA", -p["qtd"], f"entradas_estoque {p['ref_envio']}", 60)
        mov(p["destino_id"], p["sku"], p["receb"], "TRANSFERENCIA_ENTRADA", p["qtd"], f"entradas_estoque {p['ref_receb']}", 60)
    recebimentos_criados = []
    for e in so_envio:
        chegada = e["quando"] + lag
        recebimentos_criados.append(dict(origem_id=e["origem_id"], destino_id=e["destino_id"], sku=e["sku"],
                                         qtd=e["qtd"], data=chegada, ref=e["ref"]))
        mov(e["origem_id"], e["sku"], e["quando"], "TRANSFERENCIA_SAIDA", -e["qtd"], f"entradas_estoque {e['ref']}", 60)
        mov(e["destino_id"], e["sku"], chegada, "TRANSFERENCIA_ENTRADA", e["qtd"], f"entradas_estoque {e['ref']}", 60,
            "recebimento não anotado pela loja: data estimada")
        reg.add(60, "Só quem enviou anotou: entrada na loja criada", f"entradas_estoque {e['ref']}", "qtd", None, e["qtd"],
                f"{LOJAS[e['origem_id']]} → {LOJAS[e['destino_id']]}, chegada estimada em +{lag.days} dia(s)")
    for rc in so_receb:
        saida = rc["quando"] - lag
        mov(rc["origem_id"], rc["sku"], saida, "TRANSFERENCIA_SAIDA", -rc["qtd"], f"entradas_estoque {rc['ref']}", 60,
            "envio não anotado: baixa criada")
        mov(rc["destino_id"], rc["sku"], rc["quando"], "TRANSFERENCIA_ENTRADA", rc["qtd"], f"entradas_estoque {rc['ref']}", 60)
        reg.add(60, "Só a loja anotou o recebimento: baixa no CD criada", f"entradas_estoque {rc['ref']}", "qtd", None,
                -rc["qtd"], f"{LOJAS[rc['origem_id']]} → {LOJAS[rc['destino_id']]}: o CD nunca baixou essa peça")
    for p in pares:
        reg.add(60, "Envio e recebimento pareados (uma transferência só)", f"entradas_estoque {p['ref_receb']}", "qtd",
                p["qtd"], p["qtd"], f"envio {p['ref_envio']} em {p['envio']:%d/%m}, recebida em {p['receb']:%d/%m}")

    # ------------------------------------------------------------------ vendas, trocas e devoluções
    v = vendas[vendas.baixa_estoque & vendas.tipo.isin(["Venda", "Troca", "Devolução"])].set_index("id")
    it = itens[itens.venda_id.isin(v.index)]
    nao_identificadas = it[it.sku == "DIVERSO"]                  # decisão 72: não baixa estoque de peça nenhuma
    it = it[it.sku != "DIVERSO"]
    for r in it.itertuples():
        venda = v.loc[r.venda_id]
        tipo = "DEVOLUCAO" if r.qtd < 0 else ("TROCA" if venda.tipo == "Troca" else "VENDA")
        mov(int(venda.loja_estoque), r.sku, venda.data_hora, tipo, -r.qtd, f"venda {r.venda_id}",
            49 if venda.ship_from_store else (59 if venda.loja_id in (5, 6) else None))
    for loja, g in nao_identificadas.groupby(nao_identificadas.venda_id.map(v.loja_estoque)):
        reg.add(72, "Peça não identificada: sem baixa de estoque", LOJAS[int(loja)], "qtd", int(g.qtd.sum()), 0,
                "A diferença aparece no ajuste de inventário da contagem")
    for loja_virtual in (5, 6):
        n = (v.loja_id == loja_virtual).sum()
        reg.add(59, f"Vendas do {LOJAS[loja_virtual]} baixam do CD", LOJAS[loja_virtual], "loja", loja_virtual, CD,
                f"{n} vendas: a loja {loja_virtual} só existe no sistema; as peças ficam no CD")
    m = pd.DataFrame(movs)
    m = m[m.data_hora.notna()]

    # ------------------------------------------------------------------ contagem física e saldo inicial (decisão 61)
    cont = contagem.assign(loja_id=contagem.loja.map(LOJA_POR_CONTAGEM))
    cont = cont.groupby(["loja_id", "sku"]).qtd.sum()
    finais, linhas_ajuste, qualidade = [], [], defaultdict(lambda: defaultdict(int))
    chaves = set(zip(m.loja_id, m.sku)) | set(cont.index)
    m = m.sort_values(["loja_id", "sku", "data_hora", "qtd"], ascending=[True, True, True, False])
    grupos = {k: g for k, g in m.groupby(["loja_id", "sku"])}
    for loja, sku in chaves:
        g = grupos.get((loja, sku), m.iloc[0:0])
        data_cont = DATA_CONTAGEM[loja]
        antes = g[g.data_hora <= data_cont]
        acumulado = antes.qtd.cumsum()
        contado = int(cont.get((loja, sku), 0))
        pela_contagem = contado - int(antes.qtd.sum())                # saldo inicial que fecha com a contagem
        minimo = max(0, -int(acumulado.min())) if len(acumulado) else 0
        inicial = max(pela_contagem, minimo)
        ajuste = pela_contagem - inicial                              # ≤ 0: o que o sistema não registrou
        if inicial:
            linhas_ajuste.append(dict(loja_id=loja, sku=sku, data_hora=INICIO, tipo="SALDO_INICIAL", qtd=inicial,
                                      origem="reconstruído", decisao=61, ref=None))
        if ajuste:
            linhas_ajuste.append(dict(loja_id=loja, sku=sku, data_hora=data_cont, tipo="AJUSTE_INVENTARIO", qtd=ajuste,
                                      origem="contagem física set/26", decisao=61,
                                      ref=f"contado {contado}, sistema levaria a {contado - ajuste}"))
        depois = g[(g.data_hora > data_cont) & (g.data_hora < DATA_CORTE)]
        final = contado + int(depois.qtd.sum())
        if final < 0:
            linhas_ajuste.append(dict(loja_id=loja, sku=sku, data_hora=DATA_CORTE - timedelta(minutes=1),
                                      tipo="AJUSTE_INVENTARIO", qtd=-final, origem="saldo negativo depois da contagem",
                                      decisao=62, ref="venda de peça que a contagem não achou"))
            reg.add(62, "Venda depois da contagem de peça contada como zero", f"{LOJAS[loja]} {sku}", "qtd", final, 0)
            final = 0
        finais.append(dict(loja_id=loja, sku=sku, contado=contado, data_contagem=data_cont,
                           movimento_pos_contagem=int(depois.qtd.sum()), qtd=final))
        q = qualidade[loja]
        q["skus"] += 1
        q["pecas_contadas"] += contado
        q["ajuste_pecas"] += abs(ajuste)
        q["skus_com_ajuste"] += int(ajuste != 0)
        q["historico_negativo"] += int(minimo > pela_contagem)
    m = pd.concat([m, pd.DataFrame(linhas_ajuste)], ignore_index=True)
    m = m[m.data_hora < DATA_CORTE].sort_values(["loja_id", "sku", "data_hora", "qtd"], ascending=[True, True, True, False])
    m["saldo_resultante"] = m.groupby(["loja_id", "sku"]).qtd.cumsum()
    m["loja"] = m.loja_id.map(LOJAS)

    # ------------------------------------------------------------------ posição em 01/10/2026 e o sistema (decisão 63)
    pos = pd.DataFrame(finais)
    pos["loja"] = pos.loja_id.map(LOJAS)
    nomes = dict(zip(mestre.sku, mestre.nome))
    pos["produto"] = pos.sku.map(nomes)
    sistema["sku_canonico"] = sistema.sku.map(lambda s: sku_canonico(s) or str(s).strip().upper())
    sistema["loja_fisica"] = sistema.loja_id.where(~sistema.loja_id.isin([5, 6]), CD)
    sis = sistema.groupby(["loja_fisica", "sku_canonico"]).qtd.sum()
    pos["qtd_sistema_vulto"] = [int(sis.get((l, s), 0)) for l, s in zip(pos.loja_id, pos.sku)]
    ja = set(zip(pos.loja_id, pos.sku))
    so_no_sistema = [dict(loja_id=l, sku=s, contado=0, data_contagem=DATA_CONTAGEM.get(l), movimento_pos_contagem=0,
                          qtd=0, loja=LOJAS[l], produto=nomes.get(s), qtd_sistema_vulto=int(q))
                     for (l, s), q in sis.items() if (l, s) not in ja]
    pos = pd.concat([pos, pd.DataFrame(so_no_sistema)], ignore_index=True)
    pos = pos[(pos.qtd != 0) | (pos.qtd_sistema_vulto != 0)].sort_values(["loja_id", "sku"])
    pos["diferenca"] = pos.qtd - pos.qtd_sistema_vulto
    neg = sistema[sistema.qtd < 0]
    for r in neg.itertuples():
        reg.add(63, "Saldo negativo no sistema da Vulto (descartado)", f"{LOJAS[r.loja_id]} {r.sku}", "qtd", r.qtd, None,
                "A posição vem da contagem + movimentos, não do saldo do sistema")
    dup = sistema.groupby(["loja_id", "sku_canonico"]).size()
    for (l, s), n in dup[dup > 1].items():
        reg.add(63, "Mesma peça com dois códigos no estoque do sistema", f"{LOJAS[l]} {s}", "sku", n, 1,
                "Código curto do CD e SKU longo somados")

    resumo = resumo_lojas(pos, qualidade, vendas, itens)
    for loja, q in qualidade.items():
        reg.add(61, "Histórico reconstruído", LOJAS[loja], "skus", None, q["skus"],
                f"{q['skus_com_ajuste']} SKUs precisaram de ajuste de inventário ({q['ajuste_pecas']} peças)")
    duas_semanas = posicao_em(m, datetime(2026, 9, 17, 23, 59))
    return dict(movimentacoes=m, posicao=pos, resumo=resumo, exemplo_17set=duas_semanas,
                recebimentos_criados=pd.DataFrame(recebimentos_criados), transferencias=dict(
        pareadas=len(pares), so_envio=len(so_envio), so_recebimento=len(so_receb), dias_entre_envio_e_chegada=lag.days)), reg


def parear(envios, recebimentos):
    """Casa cada envio com o recebimento da mesma peça, mesmo destino e mesma quantidade, até 2 dias depois."""
    if envios.empty or recebimentos.empty:
        return [], envios.to_dict("records"), recebimentos.to_dict("records"), timedelta(days=2)
    pares, usados_r = [], set()
    receb_por_chave = {k: g.sort_values("quando") for k, g in recebimentos.groupby(["origem_id", "destino_id", "sku", "qtd"])}
    so_envio = []
    for e in envios.sort_values("quando").itertuples():
        cands = receb_por_chave.get((e.origem_id, e.destino_id, e.sku, e.qtd))
        achou = None
        if cands is not None:
            for r in cands.itertuples():
                if r.Index in usados_r:
                    continue
                dias = (r.quando - e.quando).days
                if -1 <= dias <= JANELA_TRANSFERENCIA:
                    achou = r
                    break
        if achou is None:
            so_envio.append(e._asdict())
            continue
        usados_r.add(achou.Index)
        pares.append(dict(origem_id=e.origem_id, destino_id=e.destino_id, sku=e.sku, qtd=e.qtd, envio=e.quando,
                          receb=achou.quando, ref_envio=e.ref, ref_receb=achou.ref))
    so_receb = recebimentos[~recebimentos.index.isin(usados_r)].to_dict("records")
    lag = timedelta(days=int(pd.Series([(p["receb"] - p["envio"]).days for p in pares]).median())) if pares else timedelta(days=2)
    return pares, so_envio, so_receb, lag


def posicao_em(m, quando):
    """Saldo de cada loja numa data qualquer — é o que a tela de posição por data da Lorenzi vai responder."""
    ate = m[m.data_hora <= quando].groupby(["loja", "sku"]).qtd.sum()
    return ate[ate != 0].groupby(level=0).sum().rename("pecas").reset_index()


def resumo_lojas(pos, qualidade, vendas, itens):
    """Por loja: nosso número × sistema, e a margem de erro da data de contagem que ninguém anotou."""
    linhas = []
    v = vendas[vendas.baixa_estoque].set_index("id")
    it = itens[itens.venda_id.isin(v.index)].assign(loja=lambda x: x.venda_id.map(v.loja_estoque),
                                                     dia=lambda x: x.venda_id.map(v.data_hora))
    for loja, g in pos.groupby("loja_id"):
        q = qualidade.get(loja, {})
        incerteza = 0
        if loja in (2, 3, 4):                                  # contagem sem data: vendas de 20 a 28/09 como margem
            janela = it[(it.loja == loja) & (it.dia > datetime(2026, 9, 20, 23, 59)) & (it.dia <= DATA_CONTAGEM[loja])]
            incerteza = int(janela.qtd.sum())
        linhas.append(dict(loja=LOJAS[loja], pecas_01out=int(g.qtd.sum()), skus_com_estoque=int((g.qtd > 0).sum()),
                           pecas_sistema_vulto=int(g.qtd_sistema_vulto.sum()),
                           saldos_negativos_no_sistema=int((g.qtd_sistema_vulto < 0).sum()),
                           margem_data_contagem_pecas=incerteza,
                           ajuste_inventario_pecas=int(q.get("ajuste_pecas", 0)),
                           skus_com_ajuste=int(q.get("skus_com_ajuste", 0)),
                           skus_com_historico_incompleto=int(q.get("historico_negativo", 0))))
    total = pd.DataFrame(linhas)
    total.loc[len(total)] = ["Total"] + [total[c].sum() for c in total.columns[1:]]
    return total
