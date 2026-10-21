"""Histórico do estoque da Vulto, dia a dia, de 01/01/2025 a 01/10/2026 (decisões 44 e 45).

A Vulto não guardava histórico: só o saldo atual (que é ficção). O histórico é reconstruído com as mesmas regras
e contagens do estoque_01_10.py, e termina exatamente no estoque oficial de 01/10/2026:

  1. Antes da contagem, cada movimento registrado vira uma linha: vendas e devoluções (vendas_padronizadas, com as
     vendas manuais da Batel na data real), compras do CD (NF repetida conta uma vez) e transferências (cada uma
     uma vez só: o CD baixa no envio e a loja soma no recebimento; se só um lado anotou, vale esse lado).
  2. A contagem física é a verdade no dia em que foi feita. O saldo inicial (01/01/2025) é o que faz o histórico
     chegar nela; se isso deixasse o saldo negativo em algum dia, o saldo inicial sobe até o mínimo que evita estoque
     negativo e a diferença vira um ajuste de inventário no dia da contagem (é o que o sistema não registrou).
  3. Da contagem até 30/09 valem os movimentos do estoque_01_10.py (os mesmos do número oficial).

Com isso a posição em qualquer data é a soma das linhas até ela (consulta pronta em consultas.sql).

Saída: tabela estoque_historico no vulto_padronizado.db (o carregar.py grava no vulto.db) e o resumo por local.

Uso: python migracao/estoque_historico.py   (depois de padroniza_vendas.py e produtos.py)
"""

import sqlite3
from collections import defaultdict
from datetime import date, datetime

import pandas as pd

import caminhos
import produtos
from estoque_01_10 import (LOCAIS, LOJA_DO_ID, SFS, Normalizador, carregar_catalogo, contagens, data_entrada,
                           destino_no_texto, movimentos, quantidade_no_texto)

INICIO = date(2025, 1, 1)
VIRADA = date(2026, 10, 1)
JANELA = 10                     # dias entre o envio do CD e o recebimento da loja (decisão 29)
NAO_SAI = ("Cancelada", "Estornada", "Chargeback", "Aguardando pagamento")


def vendas_ate_contagem(conn, local_da_contagem):
    """Vendas, trocas e devoluções até o dia da contagem de cada local: (local, dia, sku, qtd, tipo, origem)."""
    df = pd.read_sql(
        "select v.id_venda, v.data, v.loja_id, v.status, v.tipo, v.duplicada_de, v.motivo, v.obs, i.sku, i.qtd "
        "from vendas_padronizadas v join itens_venda i on i.venda_id = v.id_venda", conn)
    movs = []
    for v in df.itertuples():
        if v.tipo == "Teste" or v.duplicada_de not in ("", None) or v.status in NAO_SAI:
            continue
        if str(v.motivo).startswith("Relançamento da venda manual"):     # conta pela planilha (decisão 17)
            continue
        if v.sku in (produtos.GENERICO, "TESTE-001"):                    # peça não identificada (decisão 37)
            continue
        dia = datetime.strptime(v.data, "%d/%m/%Y").date()
        local = LOJA_DO_ID.get(v.loja_id, "CD Cajamar")
        if v.loja_id == 5 and SFS.search(v.obs or ""):
            local = "Pinheiros"
        if INICIO <= dia <= local_da_contagem[local]:
            tipo = {"Devolução": "Devolução", "Troca": "Troca (peça nova)"}.get(v.tipo, "Venda")
            movs.append((local, dia, v.sku, -int(v.qtd), tipo, f"venda {v.id_venda}"))
    return movs


def entradas_ate_contagem(conn, local_da_contagem, resolver):
    """Compras do CD e transferências até o dia da contagem, cada transferência uma vez só."""
    entradas = pd.read_sql("select * from entradas_estoque order by id", conn)
    movs, envios, recebidos, compras_vistas = [], [], [], set()
    for e in entradas.itertuples():
        dia, obs, sku = data_entrada(e.data), str(e.obs or ""), resolver(e.sku)
        if dia is None or sku is None:
            continue
        if e.loja_id == 7 and e.usuario in ("marina.compras", "joao.fiscal"):
            chave = (sku, e.qtd, e.nf)
            if e.nf and chave in compras_vistas:                         # a mesma NF lançada duas vezes (decisão 31)
                continue
            compras_vistas.add(chave)
            movs.append(("CD Cajamar", dia, sku, int(e.qtd), "Compra (NF)", f"entradas_estoque {e.id}"))
        elif pd.isna(e.qtd):                                             # envio: CD ou loja, quantidade no texto
            origem = LOJA_DO_ID.get(e.loja_id, "CD Cajamar")
            envios.append(dict(id=e.id, dia=dia, sku=sku, qtd=quantidade_no_texto(obs) or 1, de=origem,
                               para=destino_no_texto(obs)))
        else:                                                            # recebimento anotado pela loja
            para = LOJA_DO_ID.get(e.loja_id)
            de = destino_no_texto(obs) if "veio da" in obs.lower() else "CD Cajamar"
            recebidos.append(dict(id=e.id, dia=dia, sku=sku, qtd=int(e.qtd), de=de or "CD Cajamar", para=para))
    livres = defaultdict(list)
    for r in recebidos:
        livres[(r["sku"], r["de"], r["para"])].append(r)
    for env in envios:
        if not env["para"]:
            continue
        par = next((r for r in livres[(env["sku"], env["de"], env["para"])]
                    if 0 <= (r["dia"] - env["dia"]).days <= JANELA and not r.get("par")), None)
        if par:
            par["par"] = env["id"]
        movs.append((env["de"], env["dia"], env["sku"], -env["qtd"], f"Envio p/ {env['para']}", f"entradas_estoque {env['id']}"))
        if par is None:                                                  # a loja não anotou: vale o envio
            movs.append((env["para"], env["dia"], env["sku"], env["qtd"], f"Recebido de {env['de']} (só o envio anotado)",
                         f"entradas_estoque {env['id']}"))
    for r in recebidos:
        if not r["para"]:
            continue
        movs.append((r["para"], r["dia"], r["sku"], r["qtd"], f"Recebido de {r['de']}", f"entradas_estoque {r['id']}"))
        if not r.get("par"):                                             # quem mandou não anotou: baixa nele
            movs.append((r["de"], r["dia"], r["sku"], -r["qtd"], f"Envio p/ {r['para']} (só o recebimento anotado)",
                         f"entradas_estoque {r['id']}"))
    return [m for m in movs if INICIO <= m[1] <= local_da_contagem[m[0]]]


def main():
    catalogo, cores = carregar_catalogo()
    contado, avaria_pinheiros, _ = contagens(catalogo)
    depois, _ = movimentos(Normalizador(catalogo, cores))             # da contagem até 30/09: os do número oficial
    dia_contagem = {local: v[2] for local, v in LOCAIS.items()}
    conn = sqlite3.connect(caminhos.PADRONIZADO)
    antes = vendas_ate_contagem(conn, dia_contagem) + entradas_ate_contagem(conn, dia_contagem, produtos.resolvedor())

    por_chave = defaultdict(list)
    for m in antes:
        por_chave[(m[0], m[2])].append(m)
    pos_contagem = defaultdict(list)
    for m in depois:
        pos_contagem[(m[0], m[2])].append(m)
    chaves = set(por_chave) | set(pos_contagem) | {(local, sku) for local, linhas in contado.items() for sku in linhas} \
        | {("Pinheiros", sku) for sku in avaria_pinheiros}

    linhas, resumo = [], defaultdict(lambda: defaultdict(int))
    for local, sku in sorted(chaves):
        movs = sorted(por_chave.get((local, sku), []), key=lambda m: (m[1], m[3] < 0))
        qtd_contada = contado[local].get(sku, {}).get("qtd", 0)              # estoque vendável contado
        # a peça com avaria estava na loja no dia da contagem: o histórico chega no contado físico e a avaria sai
        # em seguida (Pinheiros: contada; Eldorado, Savassi e Batel: estimada, decisão 54)
        avaria = avaria_pinheiros.get(sku, 0) if local == "Pinheiros" else contado[local].get(sku, {}).get("avaria_estimada", 0)
        qtd_fisica = qtd_contada + avaria
        acumulado, menor = 0, 0
        for m in movs:
            acumulado += m[3]
            menor = min(menor, acumulado)
        pela_contagem = qtd_fisica - acumulado                           # saldo inicial que fecha com a contagem
        inicial = max(pela_contagem, -menor, 0)
        ajuste = qtd_fisica - (inicial + acumulado)                     # ≤ 0: o que o sistema não registrou
        serie = []
        if inicial:
            serie.append((INICIO, inicial, "Saldo inicial (reconstruído)", "decisão 44"))
        serie += [(m[1], m[3], m[4], m[5]) for m in movs]
        if ajuste:
            serie.append((dia_contagem[local], ajuste, "Ajuste de inventário (contagem física)",
                          f"contado {qtd_fisica}; os movimentos registrados levariam a {qtd_fisica - ajuste}"))
        if avaria:
            serie.append((dia_contagem[local], -avaria,
                          "Avaria (contada na Pinheiros)" if local == "Pinheiros" else "Avaria estimada (decisão 54)",
                          "fora do estoque vendável"))
        # da contagem a 30/09: se a peça vendeu mais do que foi contado, o oficial para em zero (decisão 36); o ajuste
        # entra no primeiro dia em que o saldo ficaria negativo, para o histórico nunca ter estoque negativo
        corrente, zerado = qtd_contada, False
        for m in sorted(pos_contagem.get((local, sku), []), key=lambda m: (m[1], m[3] < 0)):
            if corrente + m[3] < 0 and not zerado:
                falta = -min(corrente + sum(x[3] for x in pos_contagem[(local, sku)] if (x[1], x[3] < 0) >= (m[1], m[3] < 0)), 0)
                serie.append((m[1], falta, "Ajuste: vendeu mais que o contado (estoque não fica negativo)", "decisão 36"))
                corrente += falta
                zerado = True
            serie.append((m[1], m[3], m[4], m[5]))
            corrente += m[3]
        final = corrente
        saldo = 0
        for dia, qtd, tipo, origem in serie:
            saldo += qtd
            linhas.append(dict(local=local, loja_id=LOCAIS[local][0], sku=sku, data=dia.isoformat(), tipo=tipo, qtd=qtd,
                               saldo=saldo, origem=origem))
        r = resumo[local]
        r["skus"] += 1
        r["saldo_inicial"] += inicial
        r["ajuste_inventario"] += ajuste
        r["skus_com_ajuste"] += ajuste != 0
        r["avaria"] += avaria
        r["estoque_01_10"] += max(final, 0)

    historico = pd.DataFrame(linhas)
    historico.to_sql("estoque_historico", conn, if_exists="replace", index=False)
    conn.close()
    resumo = pd.DataFrame([dict(local=k, **v) for k, v in resumo.items()])
    print(resumo.to_string(index=False))
    print(f"{len(historico)} linhas | estoque em 01/10/2026: {resumo.estoque_01_10.sum()} peças")
    return historico, resumo


if __name__ == "__main__":
    main()
