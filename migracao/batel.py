"""Vendas da Batel feitas sem sistema (PDV fora do ar de 02 a 22/03/2026) — decisão 17.

A planilha manual (limpa pelo time, migracao/entrada) é a fonte dessas vendas: cada linha vira uma venda
(BATEL-<Nº>) e cada peça da aba Itens Vendidos vira um item com o SKU do catálogo. Quando o sistema voltou, a
gerente relançou parte delas (vendedor "gerente", com a data do relançamento e às vezes com o valor arredondado):
o relançamento que casa com uma linha da planilha não conta (a venda já conta pela planilha); o que não casa com
nenhuma linha é uma venda real que não foi anotada e fica.

Usado por padroniza_vendas.py.
"""

import os
import re
import unicodedata
from datetime import datetime

import pandas as pd

import caminhos

PLANILHA = os.path.join(caminhos.ENTRADA, "vendas_batel_MANUAL_marco (2).xlsx")
TOLERANCIA = 1.00          # a gerente às vezes arredondou o valor ao relançar (ex.: 329,80 -> 330,00)


def chave_texto(texto):
    return " ".join(unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode().lower().split())


def so_digitos(texto):
    return re.sub(r"\D", "", str(texto or ""))


def hora(texto):
    """'15:46:00' -> 15:46. Sem hora anotada: meio-dia (a venda fica no dia certo)."""
    m = re.match(r"(\d{1,2}):(\d{2})", str(texto or ""))
    return datetime.strptime(f"{m.group(1)}:{m.group(2)}", "%H:%M").time() if m else datetime.strptime("12:00", "%H:%M").time()


def ler():
    """(vendas, itens) da planilha manual."""
    vendas = pd.read_excel(PLANILHA, sheet_name="Vendas Março", header=3, dtype=str)
    vendas = vendas[vendas["Nº"].fillna("").str.fullmatch(r"\d+")].copy()
    vendas["numero"] = vendas["Nº"].astype(int)
    vendas["valor"] = vendas["Valor (R$)"].astype(float).round(2)
    vendas["cpf"] = vendas["CPF"].map(lambda c: so_digitos(c).zfill(11) if so_digitos(c) else "")
    vendas["momento"] = [datetime.combine(pd.Timestamp(d).date(), hora(h)) for d, h in zip(vendas["Data"], vendas["Hora"])]
    itens = pd.read_excel(PLANILHA, sheet_name="Itens Vendidos", header=3, dtype=str).dropna(subset=["Nº venda"])
    itens["numero"] = itens["Nº venda"].astype(int)
    itens["qtd"] = itens["Qtd"].astype(int)
    return vendas, itens


def casar_relancamentos(relancadas, vendas):
    """relancadas: [(id_venda, valor, cpf)] do sistema. Devolve {id_venda: Nº da planilha}.
    Casa pelo valor (até R$ 1 de diferença) e, quando os dois têm, pelo CPF; o valor exato tem preferência."""
    concluidas = vendas[vendas["Status"] == "Concluída"]
    usadas, casadas = set(), {}
    for vid, valor, cpf in sorted(relancadas, key=lambda r: r[0]):
        c = concluidas[~concluidas.numero.isin(usadas) & ((concluidas.valor - valor).abs() <= TOLERANCIA)]
        if cpf:
            c = c[(c.cpf == "") | (c.cpf == cpf)]
        if c.empty:
            continue
        melhor = c.assign(d=(c.valor - valor).abs(), mesmo_cpf=(c.cpf == cpf) & (cpf != "")) \
                  .sort_values(["mesmo_cpf", "d"], ascending=[False, True]).iloc[0]
        usadas.add(melhor.numero)
        casadas[vid] = int(melhor.numero)
    return casadas


def itens_com_sku(itens, vendas, produtos):
    """Cada peça da planilha com o SKU do catálogo e o preço (o preço do catálogo, ajustado para a soma dos itens
    fechar com o valor da venda). Devolve (linhas, nao_achados)."""
    import sys
    sys.path.append(os.path.join(caminhos.MIGRACAO, "padronizacao"))  # por último: os scripts do time têm prioridade
    from comum import tamanho_padrao

    indice = {}
    for p in produtos.itertuples():
        if p.tamanho:
            indice.setdefault((chave_texto(p.produto), chave_texto(p.cor), tamanho_padrao(p.tamanho)), []).append(p)
    preco_de = dict(zip(produtos.sku, produtos.preco))
    valor_da_venda = dict(zip(vendas.numero, vendas.valor))
    linhas, nao_achados = [], []
    for numero, grupo in itens.groupby("numero", sort=True):
        pecas = []
        for i in grupo.itertuples():
            nome, cor, tam = chave_texto(i.Produto), chave_texto(i.Cor), tamanho_padrao(i.Tamanho)
            achados = indice.get((nome, cor, tam), [])
            if not achados:
                # nome abreviado na planilha ('Cropped canelado'): produtos cujo nome começa assim, na mesma cor e
                # tamanho; se houver mais de um, vale o de preço igual ao da venda (só quando a venda tem uma peça)
                parecidos = [p for (n, c, t), ps in indice.items() if n.startswith(nome) and c == cor and t == tam for p in ps]
                if len(grupo) == 1 and len(parecidos) > 1:
                    parecidos = [p for p in parecidos if abs((p.preco or 0) * i.qtd - valor_da_venda[numero]) < 0.01]
                achados = parecidos if len({p.modelo for p in parecidos}) == 1 else []
            achados = sorted(achados, key=lambda p: p.status != "Ativo")     # ativo primeiro
            if achados:
                p = achados[0]
                pecas.append([p.sku, f"{p.produto} {p.sku.split('-')[3]}", i.qtd, preco_de.get(p.sku) or 0])
            else:
                nao_achados.append((numero, i.Produto, i.Cor, i.Tamanho))
                pecas.append(["DIVERSO", f"Peça não identificada: {i.Produto} ({i.Cor}, {i.Tamanho})", i.qtd, 0])
        base = sum(q * p for _, _, q, p in pecas)
        fator = valor_da_venda[numero] / base if base else 0
        for peca in pecas:
            peca[3] = round(peca[3] * fator, 2)
        if pecas:                                          # centavos de arredondamento ficam no último item
            soma = sum(q * p for _, _, q, p in pecas[:-1])
            pecas[-1][3] = round((valor_da_venda[numero] - soma) / pecas[-1][2], 2)
        linhas += [(f"BATEL-{numero}", sku, desc, q, p) for sku, desc, q, p in pecas]
    return linhas, nao_achados
